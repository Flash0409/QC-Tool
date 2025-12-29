import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Menu
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from datetime import datetime
import shutil
import tempfile
import re
import os
import json
import numpy as np
import getpass
import sys
import subprocess
import shlex
from difflib import SequenceMatcher

def get_app_base_dir():
    """
    Returns the directory where the app is running from.
    Works for both .py and PyInstaller .exe
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class CircuitInspector:
    def __init__(self, root):
        self.root = root
        self.root.title("Quality Inspection Tool")
        self.root.geometry("1400x900")

        # Data / files
        self.pdf_document = None
        self.current_pdf_path = None
        self.current_page = 0
        self.project_name = ""
        self.sales_order_no = ""
        self.cabinet_id = ""
        self.annotations = []
        base = get_app_base_dir()
        self.master_excel_file = os.path.join(base, "Emerson.xlsx")
        self.recent_projects_file = os.path.join(base, "recent_projects.json")

        self.excel_file = None
        self.working_excel_path = None
        self.checklist_file = self.excel_file
        self.zoom_level = 1.0
        self.current_sr_no = 1
        self.current_page_image = None
        self.tool_mode = None  # None, "pen", or "text"
        self.pen_points = []
        self.session_refs = set()
        self.project_dirs = {}

        # Fixed column mapping
        self.punch_sheet_name = 'Punch Sheet'
        self.punch_cols = {
            'sr_no': 'A',
            'ref_no': 'B',
            'desc': 'C',
            'category': 'D',
            'checked_name': 'E',
            'checked_date': 'F',
            'implemented_name': 'G',
            'implemented_date': 'H',
            'closed_name': 'I',
            'closed_date': 'J'
        }
        
        self.interphase_sheet_name = 'Interphase'
        self.interphase_cols = {
            'ref_no': 'B',
            'description': 'C',
            'status': 'D',
        }

        self.header_cells = {
            "Interphase": {
                "project_name": "C4",
                "sales_order": "C6",
                "cabinet_id": "F6"
            },
            "Punch Sheet": {
                "project_name": "C2",
                "sales_order": "C4",
                "cabinet_id": "H4"
            }
        }

        self.categories = []
        self.category_file = os.path.join(os.path.dirname(get_app_base_dir()), "assets", "categories.json")
        self.load_categories()

        # Drawing / selection state
        self.drawing = False
        self.drawing_type = None  # 'ok', 'error', 'pen', 'text'
        self.rect_start_x = None
        self.rect_start_y = None
        self.temp_rect_id = None
        self.temp_line_ids = []  # Store temporary pen line IDs
        self.selected_annotation = None

        self.setup_ui()
        self.current_sr_no = self.get_next_sr_no()
        
        # Load recent projects on startup
        self.load_recent_projects_ui()

    # ================================================================
    # COORDINATE CONVERSION HELPERS
    # ================================================================
    
    def page_to_display_scale(self):
        return 2.0 * self.zoom_level

    def display_to_page_coords(self, pts):
        """
        Convert display-space coordinates to page-space coordinates.
        Handles:
        - Single point: (x, y) -> (x/scale, y/scale)
        - List of points: [(x1,y1), ...] -> [(x1/scale, y1/scale), ...]
        """
        scale = self.page_to_display_scale()
        
        # Handle single point tuple
        if isinstance(pts, tuple) and len(pts) == 2:
            if not isinstance(pts[0], (list, tuple)):
                return (pts[0] / scale, pts[1] / scale)
        
        # Handle list of points
        return [(x / scale, y / scale) for x, y in pts]

    def page_to_display_coords(self, pts):
        """
        Convert page-space coordinates to display-space coordinates.
        Handles:
        - Single point: (x, y) -> (x*scale, y*scale)
        - List of points: [(x1,y1), ...] -> [(x1*scale, y1*scale), ...]
        """
        scale = self.page_to_display_scale()
        
        # Handle single point tuple
        if isinstance(pts, tuple) and len(pts) == 2:
            if not isinstance(pts[0], (list, tuple)):
                return (pts[0] * scale, pts[1] * scale)
        
        # Handle list of points
        return [(x * scale, y * scale) for x, y in pts]

    def bbox_page_to_display(self, bbox_page):
        scale = self.page_to_display_scale()
        x1, y1, x2, y2 = bbox_page
        return (x1 * scale, y1 * scale, x2 * scale, y2 * scale)

    def bbox_display_to_page(self, bbox_display):
        scale = self.page_to_display_scale()
        x1, y1, x2, y2 = bbox_display
        return (x1 / scale, y1 / scale, x2 / scale, y2 / scale)

    # ================================================================
    # TOOL MODE CONTROL
    # ================================================================

    def set_tool_mode(self, mode):
        """Set the current tool mode: None, 'pen', or 'text'"""
        if self.tool_mode == mode:
            # Toggle off if clicking same tool
            self.tool_mode = None
            self.root.config(cursor="")
            self.mode_indicator.config(text="Normal Mode", fg='#2ecc71')
            self._flash_status("Normal mode - Left drag: OK | Right drag: Error")
        else:
            self.tool_mode = mode
            if mode == "pen":
                self.root.config(cursor="pencil")
                self.mode_indicator.config(text="✏️ Pen Tool Active", fg='#e74c3c')
                self._flash_status("✏ Pen Tool Active - Right-click to exit")
            elif mode == "text":
                self.root.config(cursor="xterm")
                self.mode_indicator.config(text="📝 Text Tool Active", fg='#3498db')
                self._flash_status("📝 Text Tool Active - Right-click to exit")
        
        self.pen_points.clear()
        self.clear_temp_drawings()


    def clear_temp_drawings(self):
        """Clear temporary drawing elements from canvas"""
        for line_id in self.temp_line_ids:
            try:
                self.canvas.delete(line_id)
            except:
                pass
        self.temp_line_ids.clear()
        
        if self.temp_rect_id:
            try:
                self.canvas.delete(self.temp_rect_id)
            except:
                pass
            self.temp_rect_id = None

    # ================================================================
    # MOUSE EVENT HANDLERS - COMPLETELY REWRITTEN
    # ================================================================

    def on_left_press(self, event):
        """Handle left mouse button press"""
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please load a PDF first")
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # -------- PEN TOOL --------
        if self.tool_mode == "pen":
            self.drawing = True
            self.drawing_type = "pen"
            self.pen_points.clear()
            self.clear_temp_drawings()
            self.pen_points.append((x, y))
            return

        # -------- TEXT TOOL --------
        if self.tool_mode == "text":
            self.drawing = True
            self.drawing_type = "text"
            # Store position for text placement
            self.text_pos_x = x
            self.text_pos_y = y
            return

        # -------- NORMAL MODE: OK Rectangle --------
        self.drawing = True
        self.drawing_type = 'ok'
        self.rect_start_x = x
        self.rect_start_y = y

    def on_left_drag(self, event):
        """Handle left mouse button drag"""
        if not self.drawing:
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # -------- PEN TOOL DRAWING --------
        if self.drawing_type == "pen":
            if len(self.pen_points) > 0:
                last_x, last_y = self.pen_points[-1]
                # Draw line segment on canvas
                line_id = self.canvas.create_line(
                    last_x, last_y, x, y,
                    fill="red", width=3,
                    capstyle=tk.ROUND, smooth=True
                )
                self.temp_line_ids.append(line_id)
            self.pen_points.append((x, y))
            return

        # -------- OK RECTANGLE PREVIEW --------
        if self.drawing_type == 'ok':
            if self.temp_rect_id:
                self.canvas.delete(self.temp_rect_id)
            self.temp_rect_id = self.canvas.create_rectangle(
                self.rect_start_x, self.rect_start_y, x, y,
                outline='green', width=3, dash=(5, 5)
            )

    def on_left_release(self, event):
        """Handle left mouse button release"""
        if not self.pdf_document or not self.drawing:
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # -------- PEN TOOL FINISH --------
        if self.drawing_type == "pen":
            if len(self.pen_points) >= 2:
                # Convert to page coordinates and save
                points_page = self.display_to_page_coords(self.pen_points)
                self.annotations.append({
                    'type': 'pen',
                    'page': self.current_page,
                    'points': points_page,
                    'timestamp': datetime.now().isoformat()
                })
            self.pen_points.clear()
            self.clear_temp_drawings()
            self.drawing = False
            self.drawing_type = None
            self.display_page()
            return

        # -------- TEXT TOOL FINISH --------
        if self.drawing_type == "text":
            txt = simpledialog.askstring("Text", "Enter text:", parent=self.root)
            if txt and txt.strip():
                pos_page = self.display_to_page_coords((self.text_pos_x, self.text_pos_y))
                self.annotations.append({
                    'type': 'text',
                    'page': self.current_page,
                    'pos_page': pos_page,
                    'text': txt.strip(),
                    'timestamp': datetime.now().isoformat()
                })
                self.display_page()
            self.drawing = False
            self.drawing_type = None
            return

        # -------- OK RECTANGLE FINISH --------
        if self.drawing_type == 'ok':
            self.clear_temp_drawings()

            x1 = min(self.rect_start_x, x)
            y1 = min(self.rect_start_y, y)
            x2 = max(self.rect_start_x, x)
            y2 = max(self.rect_start_y, y)

            # Minimum size check
            if abs(x2 - x1) < 10 or abs(y2 - y1) < 10:
                self.drawing = False
                self.drawing_type = None
                return

            bbox_display = (x1, y1, x2, y2)
            bbox_page = self.bbox_display_to_page(bbox_display)

            self.annotations.append({
                'type': 'ok',
                'page': self.current_page,
                'bbox_page': bbox_page,
                'timestamp': datetime.now().isoformat()
            })
            self.display_page()

        self.drawing = False
        self.drawing_type = None

    # ================================================================
    # UPDATED: on_right_press - Right-click exits pen/text mode
    # ================================================================

    def on_right_press(self, event):
        """Handle right mouse button press.
        - If in pen/text mode: exit to normal mode
        - If in normal mode: start error rectangle
        """
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please load a PDF first")
            return

        # -------- EXIT TOOL MODE ON RIGHT CLICK --------
        if self.tool_mode in ("pen", "text"):
            self.tool_mode = None
            self.root.config(cursor="")
            self.pen_points.clear()
            self.clear_temp_drawings()
            self.drawing = False
            self.drawing_type = None
            
            # Visual feedback
            self.root.after(50, lambda: self._flash_status("Normal mode - Right drag to mark errors"))
            return

        # -------- NORMAL MODE: Start Error Rectangle --------
        self.drawing = True
        self.drawing_type = 'error'
        self.rect_start_x = self.canvas.canvasx(event.x)
        self.rect_start_y = self.canvas.canvasy(event.y)


    def _flash_status(self, message):
        """Show a temporary status message"""
        # Create a temporary label that fades away
        status_label = tk.Label(
            self.root, 
            text=message, 
            bg='#27ae60', 
            fg='white', 
            font=('Arial', 11, 'bold'),
            padx=20, 
            pady=10
        )
        status_label.place(relx=0.5, rely=0.1, anchor='center')
        
        # Remove after 1.5 seconds
        self.root.after(1500, status_label.destroy)


    def on_right_drag(self, event):
        """Handle right mouse button drag"""
        if not self.drawing or self.drawing_type != 'error':
            return

        if self.temp_rect_id:
            try:
                self.canvas.delete(self.temp_rect_id)
            except:
                pass

        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        self.temp_rect_id = self.canvas.create_rectangle(
            self.rect_start_x, self.rect_start_y,
            current_x, current_y,
            outline='orange', width=3, dash=(5, 5)
        )

    def on_right_release(self, event):
        """Handle right mouse button release"""
        if not self.drawing or self.drawing_type != 'error':
            return

        self.clear_temp_drawings()

        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)

        x1 = min(self.rect_start_x, end_x)
        y1 = min(self.rect_start_y, end_y)
        x2 = max(self.rect_start_x, end_x)
        y2 = max(self.rect_start_y, end_y)

        if abs(x2 - x1) < 10 or abs(y2 - y1) < 10:
            self.drawing = False
            self.drawing_type = None
            return

        bbox_display = (x1, y1, x2, y2)
        bbox_page = self.bbox_display_to_page(bbox_display)

        # Show category menu
        menu = Menu(self.root, tearoff=0)

        for cat in self.categories:
            if cat.get("mode") == "template":
                menu.add_command(
                    label=f"🔧 {cat['name']}",
                    command=lambda c=cat, bp=bbox_page: self.handle_template_category(c, bp)
                )
            elif cat.get("mode") == "parent":
                cat_menu = Menu(menu, tearoff=0)
                for sub in cat.get("subcategories", []):
                    cat_menu.add_command(
                        label=sub["name"],
                        command=lambda c=cat, s=sub, bp=bbox_page: self.handle_subcategory(c, s, bp)
                    )
                menu.add_cascade(label=f"🔧 {cat['name']}", menu=cat_menu)

        menu.add_separator()
        menu.add_command(
            label="📝 Custom Action Point",
            command=lambda bp=bbox_page: self.log_custom_error(bp, None)
        )

        menu.tk_popup(event.x_root, event.y_root)
        self.drawing = False
        self.drawing_type = None

    # ================================================================
    # DISPLAY PAGE - WITH PEN AND TEXT RENDERING
    # ================================================================

    def display_page(self):
        """Render the current PDF page with all annotations"""
        if not self.pdf_document:
            self.canvas.delete("all")
            self.page_label.config(text="Page: 0/0")
            return

        try:
            page = self.pdf_document[self.current_page]
            mat = fitz.Matrix(self.page_to_display_scale(), self.page_to_display_scale())
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.current_page_image = np.array(img)
            draw = ImageDraw.Draw(img, 'RGBA')

            # Try to load a font for text
            try:
                font_size = max(12, int(14 * self.zoom_level))
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                font = ImageFont.load_default()

            for ann in self.annotations:
                if ann.get('page') != self.current_page:
                    continue

                ann_type = ann.get('type')

                # -------- RECTANGLE ANNOTATIONS (ok/error) --------
                if ann_type in ('ok', 'error') and 'bbox_page' in ann:
                    x1d, y1d, x2d, y2d = self.bbox_page_to_display(ann['bbox_page'])
                    is_selected = (self.selected_annotation is ann)
                    w = int(5 * self.zoom_level) if is_selected else int(3 * self.zoom_level)

                    if ann_type == 'ok':
                        draw.rectangle([x1d, y1d, x2d, y2d],
                                       fill=(0, 255, 0, 80),
                                       outline='blue' if is_selected else 'green',
                                       width=w)
                    else:
                        draw.rectangle([x1d, y1d, x2d, y2d],
                                       fill=(255, 165, 0, 120),
                                       outline='blue' if is_selected else 'orange',
                                       width=w)

                    if ann.get('closed_by'):
                        cx = x1d + 8
                        cy = y1d + 8
                        draw.ellipse([cx - 6, cy - 6, cx + 6, cy + 6], fill=(0, 128, 0, 200))

                # -------- PEN STROKES --------
                elif ann_type == 'pen' and 'points' in ann:
                    points_page = ann['points']
                    if len(points_page) >= 2:
                        points_display = self.page_to_display_coords(points_page)
                        stroke_width = max(2, int(3 * self.zoom_level))
                        for i in range(len(points_display) - 1):
                            x1, y1 = points_display[i]
                            x2, y2 = points_display[i + 1]
                            draw.line([x1, y1, x2, y2], fill='red', width=stroke_width)

                # -------- TEXT ANNOTATIONS --------
                elif ann_type == 'text' and 'pos_page' in ann:
                    pos_page = ann['pos_page']
                    pos_display = self.page_to_display_coords(pos_page)
                    text = ann.get('text', '')
                    if text:
                        # Draw text background for visibility
                        try:
                            bbox = draw.textbbox(pos_display, text, font=font)
                            padding = 2
                            draw.rectangle(
                                [bbox[0] - padding, bbox[1] - padding,
                                 bbox[2] + padding, bbox[3] + padding],
                                fill=(255, 255, 200, 200)
                            )
                        except:
                            pass
                        draw.text(pos_display, text, fill='red', font=font)

            self.photo = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
            self.page_label.config(text=f"Page: {self.current_page + 1}/{len(self.pdf_document)}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to display page: {e}")

    # ================================================================
    # SAVE SESSION - WITH PROPER SERIALIZATION
    # ================================================================

    def save_session(self):
        """Save current session to JSON file"""
        if not self.pdf_document:
            messagebox.showwarning("No PDF", "Load a PDF first before saving a session.")
            return

        if not hasattr(self, 'project_dirs') or not self.project_dirs.get("sessions"):
            messagebox.showerror("Error", "Project directories not set up. Load a PDF first.")
            return

        save_path = os.path.join(
            self.project_dirs["sessions"],
            f"{self.cabinet_id}_annotations.json"
        )

        data = {
            'project_name': self.project_name,
            'sales_order_no': self.sales_order_no,
            'cabinet_id': getattr(self, 'cabinet_id', ''),
            'pdf_path': self.current_pdf_path,
            'current_page': self.current_page,
            'zoom_level': self.zoom_level,
            'current_sr_no': self.current_sr_no,
            'session_refs': list(self.session_refs),
            'annotations': []
        }

        for ann in self.annotations:
            entry = ann.copy()

            # Serialize bbox_page (rectangles)
            if 'bbox_page' in entry:
                entry['bbox_page'] = [float(x) for x in entry['bbox_page']]

            # Serialize points (pen strokes) - convert tuples to lists
            if 'points' in entry:
                entry['points'] = [[float(x), float(y)] for x, y in entry['points']]

            # Serialize pos_page (text position)
            if 'pos_page' in entry:
                pos = entry['pos_page']
                entry['pos_page'] = [float(pos[0]), float(pos[1])]

            data['annotations'].append(entry)

        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
            messagebox.showinfo("Saved", f"Session saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save session: {e}")

    # ================================================================
    # LOAD SESSION - WITH PROPER DESERIALIZATION
    # ================================================================

    def load_session(self):
        """Load session from JSON file via file dialog"""
        path = filedialog.askopenfilename(
            title="Load Session JSON",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        self.load_session_from_path(path)

    def load_session_from_path(self, path):
        """Load session from a specific JSON file path"""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Session Load Error", f"Failed to load session:\n{e}")
            return

        # Restore basic state
        self.project_name = data.get('project_name', self.project_name)
        self.sales_order_no = data.get('sales_order_no', self.sales_order_no)
        self.cabinet_id = data.get('cabinet_id', getattr(self, "cabinet_id", ""))
        self.current_page = data.get('current_page', 0)
        self.zoom_level = data.get('zoom_level', 1.0)
        self.current_sr_no = data.get('current_sr_no', self.current_sr_no)

        # Restore session refs
        self.session_refs = set(data.get('session_refs', []))

        # Restore annotations with proper type conversion
        self.annotations = []
        for entry in data.get('annotations', []):
            ann = entry.copy()

            # Deserialize bbox_page (convert list back to tuple)
            if 'bbox_page' in ann:
                ann['bbox_page'] = tuple(float(x) for x in ann['bbox_page'])

            # Deserialize points (pen strokes - convert lists to tuples)
            if 'points' in ann:
                ann['points'] = [(float(p[0]), float(p[1])) for p in ann['points']]

            # Deserialize pos_page (text position - convert list to tuple)
            if 'pos_page' in ann:
                pos = ann['pos_page']
                ann['pos_page'] = (float(pos[0]), float(pos[1]))

            self.annotations.append(ann)

            # Add ref_no to session refs
            if ann.get('ref_no'):
                self.session_refs.add(str(ann['ref_no']).strip())

        self.display_page()
        messagebox.showinfo("Loaded", f"Session loaded with {len(self.annotations)} annotations.\nMake sure the same PDF is open.")

    # ================================================================
    # EXPORT ANNOTATED PDF - WITH PEN AND TEXT
    # ================================================================

    def export_annotated_pdf(self):
        """Export PDF with all annotations including pen strokes and text"""
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please load a PDF first")
            return

        if not hasattr(self, 'project_dirs') or not self.project_dirs.get("annotated_drawings"):
            messagebox.showerror("Error", "Project directories not set up.")
            return

        try:
            save_path = os.path.join(
                self.project_dirs["annotated_drawings"],
                f"{self.cabinet_id.replace(' ', '_')}_Annotated.pdf"
            )

            # Create output PDF
            out_doc = fitz.open()
            for pnum in range(len(self.pdf_document)):
                out_doc.insert_pdf(self.pdf_document, from_page=pnum, to_page=pnum)

            # Open Excel for SR No lookup
            wb = None
            ws = None
            if self.excel_file and os.path.exists(self.excel_file):
                try:
                    wb = load_workbook(self.excel_file, data_only=True)
                    ws = wb[self.punch_sheet_name]
                except:
                    pass

            # Draw annotations
            for ann in self.annotations:
                p = ann.get('page')
                if p is None or p < 0 or p >= len(out_doc):
                    continue

                target_page = out_doc[p]
                ann_type = ann.get('type')

                # -------- RECTANGLE ANNOTATIONS --------
                if ann_type in ('ok', 'error') and 'bbox_page' in ann:
                    x1, y1, x2, y2 = ann['bbox_page']
                    rect = self.transform_bbox_for_rotation((x1, y1, x2, y2), target_page)

                    if ann_type == 'ok':
                        target_page.draw_rect(rect, color=(0, 1, 0), width=2)
                    else:
                        target_page.draw_rect(rect, color=(1, 0.55, 0), width=2)

                    sr_text = None
                    row = ann.get('excel_row')

                    if row:
                        try:
                            sr_val = self.read_cell(ws, row, self.punch_cols['sr_no'])
                            if sr_val is not None:
                                sr_text = f"Sr {sr_val}"
                        except:
                            sr_text = None

                    if sr_text:
                        text_pos = fitz.Point(rect.x0, max(rect.y0 - 12, rect.y0))
                        try:
                            target_page.insert_text(text_pos, sr_text, fontsize=8)
                        except:
                            pass

                # -------- PEN STROKES --------
                elif ann_type == 'pen' and 'points' in ann:
                    points = ann['points']
                    if len(points) >= 2:
                        # Transform all points for rotation
                        transformed_points = [
                            self.transform_point_for_rotation(pt, target_page) 
                            for pt in points
                        ]
                        
                        for i in range(len(transformed_points) - 1):
                            p1 = transformed_points[i]
                            p2 = transformed_points[i + 1]
                            target_page.draw_line(p1, p2, color=(1, 0, 0), width=2)

                # -------- TEXT ANNOTATIONS --------
                elif ann_type == 'text' and 'pos_page' in ann:
                    pos = ann['pos_page']
                    text = ann.get('text', '')
                    if text:
                        # Transform text position for rotation
                        text_point = self.transform_point_for_rotation(pos, target_page)
                        try:
                            target_page.insert_text(
                                text_point, text,
                                fontsize=10, color=(1, 0, 0),
                                rotate=target_page.rotation
                            )
                        except:
                            pass

            if wb:
                wb.close()

            out_doc.save(save_path)
            out_doc.close()

            messagebox.showinfo("Success", f"Annotated PDF saved to:\n{save_path}")

        except PermissionError:
            messagebox.showerror("Error", "Close the target file (if open) and try again.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export annotated PDF:\n{e}")


    def transform_bbox_for_rotation(self, rect, page):
        """Transform bbox for page rotation"""
        r = page.rotation
        w = page.rect.width
        h = page.rect.height
        x1, y1, x2, y2 = rect

        if r == 0:
            return fitz.Rect(x1, y1, x2, y2)
        if r == 90:
            return fitz.Rect(y1, w - x2, y2, w - x1)
        if r == 180:
            return fitz.Rect(w - x2, h - y2, w - x1, h - y1)
        if r == 270:
            return fitz.Rect(h - y2, x1, h - y1, x2)

        return fitz.Rect(x1, y1, x2, y2)


    def transform_point_for_rotation(self, point, page):
        """Transform a single point (x, y) or tuple for page rotation
        
        Args:
            point: tuple (x, y) representing a point coordinate
            page: fitz page object with rotation info
            
        Returns:
            fitz.Point object with transformed coordinates
        """
        r = page.rotation
        w = page.rect.width
        h = page.rect.height
        x, y = point

        if r == 0:
            return fitz.Point(x, y)
        elif r == 90:
            return fitz.Point(y, w - x)
        elif r == 180:
            return fitz.Point(w - x, h - y)
        elif r == 270:
            return fitz.Point(h - y, x)
        
        return fitz.Point(x, y)


    def get_text_position_above_rect(self, rect, page):
        """Get the correct position for text above a rectangle based on page rotation
        
        Args:
            rect: fitz.Rect object (already transformed)
            page: fitz page object with rotation info
            
        Returns:
            fitz.Point object for text position
        """
        r = page.rotation
        offset = 12  # Distance from rectangle edge
        
        if r == 0:
            # Normal orientation - text above top edge
            return fitz.Point(rect.x0, max(rect.y0 - offset, 5))
        elif r == 90:
            # 90° rotation - text to the left of left edge
            return fitz.Point(max(rect.x0 - offset, 5), rect.y0)
        elif r == 180:
            # 180° rotation - text below bottom edge
            return fitz.Point(rect.x0, min(rect.y1 + offset, page.rect.height - 5))
        elif r == 270:
            # 270° rotation - text to the right of right edge
            return fitz.Point(min(rect.x1 + offset, page.rect.width - 5), rect.y0)
        
        # Default fallback
        return fitz.Point(rect.x0, max(rect.y0 - offset, 5))

    # ================================================================
    # UI SETUP
    # ================================================================

    def setup_ui(self):
        toolbar = tk.Frame(self.root, bg='#2c3e50', height=60)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        tools_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Review Checklist", command=self.review_checklist_now)
        tools_menu.add_command(label="Punch Closing Mode", command=self.punch_closing_mode)
        tools_menu.add_separator()
        tools_menu.add_command(label="Save Interphase Excel", command=self.save_interphase_excel)
        tools_menu.add_command(label="Open Excel", command=self.open_excel)

        btn_style = {'bg': '#3498db', 'fg': 'white', 'padx': 15, 'pady': 8, 'font': ('Arial', 10)}

        tk.Button(toolbar, text="📁 Load PDF", command=self.load_pdf, **btn_style).pack(side=tk.LEFT, padx=5, pady=10)
        
        # Recent Projects Dropdown
        recent_frame = tk.Frame(toolbar, bg='#2c3e50')
        recent_frame.pack(side=tk.LEFT, padx=5)
        
        tk.Label(recent_frame, text="Recent:", bg='#2c3e50', fg='white', font=('Arial', 9)).pack(side=tk.LEFT, padx=(0, 3))
        
        self.recent_var = tk.StringVar(value="Select Project...")
        self.recent_dropdown = tk.OptionMenu(recent_frame, self.recent_var, "Select Project...", command=self.load_recent_project)
        self.recent_dropdown.config(bg='#34495e', fg='white', font=('Arial', 9), width=25)
        self.recent_dropdown.pack(side=tk.LEFT)

        self.page_label = tk.Label(toolbar, text="Page: 0/0", bg='#2c3e50', fg='white', font=('Arial', 10))
        self.page_label.pack(side=tk.LEFT, padx=5)

        tk.Button(toolbar, text="◀ Prev", command=self.prev_page, bg='#95a5a6', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Next ▶", command=self.next_page, bg='#95a5a6', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=5)

        tk.Button(toolbar, text="🔍+", command=self.zoom_in, bg='#27ae60', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=(20, 2))
        tk.Button(toolbar, text="🔍-", command=self.zoom_out, bg='#27ae60', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=2)

        # Tool buttons with toggle behavior
        self.pen_btn = tk.Button(toolbar, text="✏ Pen Tool", command=lambda: self.set_tool_mode("pen"), **btn_style)
        self.pen_btn.pack(side=tk.LEFT, padx=5)

        self.text_btn = tk.Button(toolbar, text="📝 Text Box", command=lambda: self.set_tool_mode("text"), **btn_style)
        self.text_btn.pack(side=tk.LEFT, padx=5)

        tk.Button(toolbar, text="📥 Export PDF", command=self.export_annotated_pdf, bg='#e67e22', fg='white', padx=10, pady=8).pack(side=tk.RIGHT, padx=5, pady=10)
        tk.Button(toolbar, text="💾 Save Session", command=self.save_session, bg='#2c3e50', fg='white', padx=10, pady=8).pack(side=tk.RIGHT, padx=5, pady=10)
        tk.Button(toolbar, text="📂 Load Session", command=self.load_session, bg='#34495e', fg='white', padx=10, pady=8).pack(side=tk.RIGHT, padx=5, pady=10)

        self.root.bind_all("<Control-Shift-p>", lambda e: self.punch_closing_mode())

        # Canvas with scrollbars
        canvas_frame = tk.Frame(self.root)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        v_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.canvas = tk.Canvas(canvas_frame, bg='#ecf0f1',
                               yscrollcommand=v_scrollbar.set,
                               xscrollcommand=h_scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scrollbar.config(command=self.canvas.yview)
        h_scrollbar.config(command=self.canvas.xview)

        # Bind mouse events
        self.canvas.bind("<ButtonPress-1>", self.on_left_press)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_release)

        self.canvas.bind("<ButtonPress-3>", self.on_right_press)
        self.canvas.bind("<B3-Motion>", self.on_right_drag)
        self.canvas.bind("<ButtonRelease-3>", self.on_right_release)

        self.canvas.bind("<Control-Button-1>", self.on_ctrl_click)
        self.root.bind("<Delete>", self.delete_selected_annotation)

        self.canvas.bind("<Double-Button-1>", self.on_double_left_zoom)
        self.canvas.bind("<Double-Button-3>", self.on_double_right_zoom)
        self.root.bind("<Escape>", self.exit_tool_mode)

        # Instructions bar
        instructions = tk.Frame(self.root, bg='#34495e', height=50)
        instructions.pack(side=tk.BOTTOM, fill=tk.X)
        inst_text = "🖱️ Left: OK (Green) | Right: Error (Orange) | Ctrl+Click: Select | Delete: Remove | Right-Click: Exit Tool Mode"
        tk.Label(instructions, text=inst_text, bg='#34495e', fg='white', font=('Arial', 10), pady=15).pack()

    # ================================================================
    # RECENT PROJECTS MANAGEMENT
    # ================================================================
    
    def save_recent_project(self):
        """Save current project to recent projects list"""
        if not self.current_pdf_path or not self.excel_file:
            return
        
        try:
            # Load existing recent projects
            recent_projects = []
            if os.path.exists(self.recent_projects_file):
                with open(self.recent_projects_file, 'r', encoding='utf-8') as f:
                    recent_projects = json.load(f)
            
            # Create new entry
            session_path = os.path.join(
                self.project_dirs.get("sessions", ""),
                f"{self.cabinet_id}_annotations.json"
            ) if hasattr(self, 'project_dirs') else None
            
            new_entry = {
                'cabinet_id': self.cabinet_id,
                'project_name': self.project_name,
                'sales_order_no': self.sales_order_no,
                'pdf_path': self.current_pdf_path,
                'excel_path': self.excel_file,
                'session_path': session_path if session_path and os.path.exists(session_path) else None,
                'last_accessed': datetime.now().isoformat()
            }
            
            # Remove old entry with same cabinet_id if exists
            recent_projects = [p for p in recent_projects if p.get('cabinet_id') != self.cabinet_id]
            
            # Add new entry at the beginning
            recent_projects.insert(0, new_entry)
            
            # Keep only projects from last 7 days and max 20 entries
            week_ago = datetime.now().timestamp() - (7 * 24 * 60 * 60)
            recent_projects = [
                p for p in recent_projects 
                if datetime.fromisoformat(p['last_accessed']).timestamp() > week_ago
            ][:20]
            
            # Save updated list
            with open(self.recent_projects_file, 'w', encoding='utf-8') as f:
                json.dump(recent_projects, f, indent=2)
            
            # Update dropdown
            self.update_recent_dropdown()
            
        except Exception as e:
            print(f"Error saving recent project: {e}")
    
    def load_recent_projects_ui(self):
        """Load and display recent projects in dropdown"""
        self.update_recent_dropdown()
    
    def update_recent_dropdown(self):
        """Update the recent projects dropdown menu"""
        try:
            if not os.path.exists(self.recent_projects_file):
                return
            
            with open(self.recent_projects_file, 'r', encoding='utf-8') as f:
                recent_projects = json.load(f)
            
            # Clear existing menu
            menu = self.recent_dropdown['menu']
            menu.delete(0, 'end')
            
            if not recent_projects:
                menu.add_command(label="No recent projects", command=lambda: None)
                return
            
            # Add each project
            for proj in recent_projects:
                label = f"{proj.get('cabinet_id', 'Unknown')} - {proj.get('project_name', 'Unknown')}"
                menu.add_command(
                    label=label,
                    command=lambda p=proj: self.load_recent_project(p)
                )
                
        except Exception as e:
            print(f"Error updating recent dropdown: {e}")
    
    def load_recent_project(self, project_data):
        """Load a recent project with its PDF, Excel, and session"""
        if isinstance(project_data, str):
            # Called from OptionMenu callback - ignore
            return
        
        try:
            # Verify files exist
            pdf_path = project_data.get('pdf_path')
            excel_path = project_data.get('excel_path')
            session_path = project_data.get('session_path')
            
            if not pdf_path or not os.path.exists(pdf_path):
                messagebox.showerror("Error", "PDF file not found. It may have been moved or deleted.")
                return
            
            if not excel_path or not os.path.exists(excel_path):
                messagebox.showerror("Error", "Excel file not found. It may have been moved or deleted.")
                return
            
            # Load PDF
            self.pdf_document = fitz.open(pdf_path)
            self.current_pdf_path = pdf_path
            self.current_page = 0
            self.annotations = []
            self.zoom_level = 1.0
            self.tool_mode = None
            self.root.config(cursor="")
            
            # Set project details
            self.cabinet_id = project_data.get('cabinet_id', '')
            self.project_name = project_data.get('project_name', '')
            self.sales_order_no = project_data.get('sales_order_no', '')
            
            # Prepare folders
            self.prepare_project_folders()
            
            # Set Excel file
            self.excel_file = excel_path
            self.working_excel_path = excel_path
            
            # Get next SR number
            self.current_sr_no = self.get_next_sr_no()
            
            # Load session if available
            if session_path and os.path.exists(session_path):
                self.load_session_from_path(session_path)
            else:
                self.display_page()
            
            # Update recent projects (moves this to top)
            self.save_recent_project()
            
            messagebox.showinfo(
                "Project Loaded",
                f"Loaded: {self.cabinet_id}\n"
                f"Project: {self.project_name}\n"
                f"PDF: {len(self.pdf_document)} pages\n"
                f"Session: {'Restored' if session_path and os.path.exists(session_path) else 'New'}"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load recent project:\n{e}")

    # ================================================================
    # CELL HELPERS
    # ================================================================

    def split_cell(self, cell_ref):
        """Splits 'F6' -> (6, 'F')"""
        m = re.match(r"([A-Z]+)(\d+)", cell_ref)
        if not m:
            raise ValueError(f"Invalid cell reference: {cell_ref}")
        col, row = m.groups()
        return int(row), col

    def _resolve_merged_target(self, ws, row, col_idx):
        for merged in ws.merged_cells.ranges:
            if merged.min_row <= row <= merged.max_row and merged.min_col <= col_idx <= merged.max_col:
                return merged.min_row, merged.min_col
        return row, col_idx

    def write_cell(self, ws, row, col, value):
        if isinstance(col, str):
            col_idx = column_index_from_string(col)
        else:
            col_idx = int(col)
        target_row, target_col = self._resolve_merged_target(ws, int(row), col_idx)
        ws.cell(row=target_row, column=target_col).value = value

    def read_cell(self, ws, row, col):
        if isinstance(col, str):
            col_idx = column_index_from_string(col)
        else:
            col_idx = int(col)
        target_row, target_col = self._resolve_merged_target(ws, int(row), col_idx)
        return ws.cell(row=target_row, column=target_col).value
    
    def exit_tool_mode(self, event=None):
        """Exit pen/text tool mode and return to normal"""
        if self.tool_mode:
            self.tool_mode = None
            self.root.config(cursor="")
            self.mode_indicator.config(text="Normal Mode", fg='#2ecc71')
            self.pen_points.clear()
            self.clear_temp_drawings()
            self.drawing = False
            self.drawing_type = None
            self._flash_status("Normal mode")

    # ================================================================
    # PDF LOADING
    # ================================================================

    def load_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Select Circuit Diagram PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            try:
                self.pdf_document = fitz.open(file_path)
                self.current_pdf_path = file_path
                self.current_page = 0
                self.annotations = []
                self.zoom_level = 1.0
                self.tool_mode = None
                self.root.config(cursor="")
                self.current_sr_no = self.get_next_sr_no()
                self.display_page()
                messagebox.showinfo("Success", f"Loaded PDF with {len(self.pdf_document)} pages")
                self.ask_project_details()
                self.prepare_project_folders()

                # Session working Excel
                try:
                    self.working_excel_path = os.path.join(
                        self.project_dirs["working_excel"],
                        f"{self.cabinet_id.replace(' ', '_')}_Working.xlsx"
                    )

                    if os.path.exists(self.working_excel_path):
                        resume = messagebox.askyesno(
                            "Resume Inspection",
                            f"Existing working Excel found:\n\n{os.path.basename(self.working_excel_path)}\n\nResume previous inspection?"
                        )
                        if not resume:
                            shutil.copy2(self.master_excel_file, self.working_excel_path)
                    else:
                        shutil.copy2(self.master_excel_file, self.working_excel_path)

                    self.excel_file = self.working_excel_path

                except Exception as e:
                    messagebox.showerror("Excel Error", f"Failed to prepare working Excel:\n{e}")
                    return

                self.write_project_details_to_excel()

                # Auto load session if exists
                session_path = self.get_session_path_for_pdf()
                if session_path:
                    resume = messagebox.askyesno(
                        "Resume Session",
                        f"Existing session found for this drawing:\n\n"
                        f"{os.path.basename(session_path)}\n\n"
                        "Do you want to resume it?"
                    )
                    if resume:
                        self.load_session_from_path(session_path)
                
                # Save to recent projects
                self.save_recent_project()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")

    def get_next_sr_no(self):
        try:
            if not self.excel_file or not os.path.exists(self.excel_file):
                return 1
            wb = load_workbook(self.excel_file, read_only=True)
            ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active
            last_sr_no = 0
            row_num = 8
            while row_num <= ws.max_row + 5:
                val = self.read_cell(ws, row_num, self.punch_cols['sr_no'])
                if val is None:
                    break
                try:
                    last_sr_no = int(val)
                except:
                    pass
                row_num += 1
            wb.close()
            return last_sr_no + 1
        except Exception:
            return 1

    # ================================================================
    # ZOOM FUNCTIONS
    # ================================================================

    def zoom_in(self):
        if self.zoom_level < 3.0:
            self.zoom_level += 0.25
            self.display_page()

    def zoom_out(self):
        if self.zoom_level > 0.5:
            self.zoom_level -= 0.25
            self.display_page()

    def zoom_at_point(self, canvas_x, canvas_y, zoom_delta):
        if not self.pdf_document:
            return

        old_zoom = self.zoom_level
        new_zoom = max(0.5, min(3.0, old_zoom + zoom_delta))

        if new_zoom == old_zoom:
            return

        self.zoom_level = new_zoom
        self.display_page()

        scale = new_zoom / old_zoom
        bbox = self.canvas.bbox("all")
        if not bbox:
            return

        self.canvas.xview_moveto((canvas_x * scale) / max(1, bbox[2]))
        self.canvas.yview_moveto((canvas_y * scale) / max(1, bbox[3]))

    def on_double_left_zoom(self, event):
        self.drawing = False
        self.temp_rect_id = None
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.zoom_at_point(x, y, +0.25)

    def on_double_right_zoom(self, event):
        self.drawing = False
        self.temp_rect_id = None
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.zoom_at_point(x, y, -0.25)

    # ================================================================
    # PAGE NAVIGATION
    # ================================================================

    def prev_page(self):
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.display_page()

    def next_page(self):
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.display_page()

    # ================================================================
    # SELECTION / DELETE
    # ================================================================

    def on_ctrl_click(self, event):
        if not self.pdf_document:
            return
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        for ann in reversed(self.annotations):
            if ann.get('page') != self.current_page:
                continue

            ann_type = ann.get('type')

            # Check rectangles
            if 'bbox_page' in ann:
                x1d, y1d, x2d, y2d = self.bbox_page_to_display(ann['bbox_page'])
                if ann_type == 'ok':
                    cx = (x1d + x2d) / 2
                    cy = (y1d + y2d) / 2
                    radius = max(abs(x2d - x1d), abs(y2d - y1d)) / 2
                    inside = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5 <= radius
                else:
                    inside = (x1d <= x <= x2d) and (y1d <= y <= y2d)

                if inside:
                    self.selected_annotation = None if self.selected_annotation is ann else ann
                    self.display_page()
                    return

            # Check pen strokes (bounding box)
            elif ann_type == 'pen' and 'points' in ann:
                points_display = self.page_to_display_coords(ann['points'])
                if len(points_display) >= 2:
                    xs = [p[0] for p in points_display]
                    ys = [p[1] for p in points_display]
                    if min(xs) - 10 <= x <= max(xs) + 10 and min(ys) - 10 <= y <= max(ys) + 10:
                        self.selected_annotation = None if self.selected_annotation is ann else ann
                        self.display_page()
                        return

            # Check text annotations
            elif ann_type == 'text' and 'pos_page' in ann:
                pos_display = self.page_to_display_coords(ann['pos_page'])
                if abs(x - pos_display[0]) < 50 and abs(y - pos_display[1]) < 20:
                    self.selected_annotation = None if self.selected_annotation is ann else ann
                    self.display_page()
                    return

        self.selected_annotation = None
        self.display_page()

    def delete_selected_annotation(self, event=None):
        if not self.selected_annotation:
            messagebox.showinfo("No Selection", "Please select an annotation first by Ctrl+Clicking on it")
            return
        if messagebox.askyesno("Delete Annotation", "Are you sure you want to delete this annotation?"):
            try:
                self.annotations.remove(self.selected_annotation)
            except:
                pass
            self.selected_annotation = None
            self.display_page()
            messagebox.showinfo("Deleted", "Annotation removed successfully")

    # ================================================================
    # INTERPHASE UPDATE
    # ================================================================

    def update_interphase_status_for_ref(self, ref_no, status='NOK'):
        try:
            wb = load_workbook(self.excel_file)
            if self.interphase_sheet_name not in wb.sheetnames:
                wb.close()
                return False
            ws = wb[self.interphase_sheet_name]
            ref_col = self.interphase_cols['ref_no']
            status_col = self.interphase_cols['status']
            updated_any = False
            max_r = ws.max_row if ws.max_row else 2000
            for r in range(1, max_r + 1):
                cell_val = self.read_cell(ws, r, ref_col)
                if cell_val is None:
                    continue
                if str(cell_val).strip() == str(ref_no).strip():
                    self.write_cell(ws, r, status_col, status)
                    updated_any = True
            if updated_any:
                wb.save(self.excel_file)
            wb.close()
            return updated_any
        except Exception as e:
            try:
                wb.close()
            except:
                pass
            print("update_interphase_status_for_ref error:", e)
            return False

    # ================================================================
    # ERROR LOGGING
    # ================================================================

    def log_error(self, component_type, error_name, error_template, bbox_page, tag_name):
        """Logs an error punch into Excel and stores annotation."""
        punch_text = error_template

        if not punch_text:
            messagebox.showerror("Error", "Punch description is empty.")
            return

        ref_no = simpledialog.askstring("Reference No", "Enter Reference Number:", parent=self.root)
        if not ref_no:
            return

        ref_no = str(ref_no).strip()
        self.session_refs.add(ref_no)

        try:
            wb = load_workbook(self.excel_file)
            ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active

            row_num = 8
            while True:
                val = self.read_cell(ws, row_num, self.punch_cols['sr_no'])
                if val is None:
                    break
                row_num += 1

            prev_sr = None
            if row_num > 8:
                prev_sr = self.read_cell(ws, row_num - 1, self.punch_cols['sr_no'])

            try:
                sr_no_assigned = int(prev_sr) + 1 if prev_sr is not None else 1
            except:
                sr_no_assigned = 1

            self.write_cell(ws, row_num, self.punch_cols['sr_no'], sr_no_assigned)
            self.write_cell(ws, row_num, self.punch_cols['ref_no'], ref_no)
            self.write_cell(ws, row_num, self.punch_cols['desc'], punch_text)
            self.write_cell(ws, row_num, self.punch_cols['category'], component_type)

            try:
                uname = os.getlogin()
            except:
                uname = getpass.getuser()

            self.write_cell(ws, row_num, self.punch_cols['checked_name'], uname)
            self.write_cell(ws, row_num, self.punch_cols['checked_date'], datetime.now().strftime("%Y-%m-%d"))

            wb.save(self.excel_file)
            wb.close()

            updated = self.update_interphase_status_for_ref(ref_no, status='NOK')
            if updated:
                print(f"Interphase: marked ref {ref_no} as NOK")

            ann = {
                'type': 'error',
                'page': self.current_page,
                'bbox_page': bbox_page,
                'component': component_type,
                'subcategory': error_name,
                'punch_text': punch_text,
                'ref_no': ref_no,
                'excel_row': row_num,
                'sr_no': sr_no_assigned,
                'implemented': False,
                'implemented_name': None,
                'implemented_date': None,
                'implementation_remark': None,
                'timestamp': datetime.now().isoformat()
            }

            self.annotations.append(ann)
            self.current_sr_no = self.get_next_sr_no()
            self.display_page()

            self.root.after(100, lambda: messagebox.showinfo("Logged", f"Punch logged:\n{punch_text}"))

        except PermissionError:
            messagebox.showerror("Error", "Close the Excel file before writing to it.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to log punch:\n{e}")

    def log_custom_error(self, bbox_page, tag_name):
        try:
            custom_action = simpledialog.askstring(
                "Custom Action Point",
                "Enter the action point / punch description:",
                parent=self.root
            )
            if not custom_action:
                return

            custom_category = simpledialog.askstring(
                "Custom Category",
                "Enter the category:",
                parent=self.root
            )
            if not custom_category:
                return

            ref_no = simpledialog.askstring("Reference No", "Enter Reference No:", parent=self.root)
            if not ref_no:
                messagebox.showwarning("Reference Required", "Reference No is required for custom action points.")
                return

            ref_no = str(ref_no).strip()
            self.session_refs.add(ref_no)

            wb = load_workbook(self.excel_file)
            ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active

            row_num = 8
            while True:
                val = self.read_cell(ws, row_num, self.punch_cols['sr_no'])
                if val is None:
                    break
                row_num += 1

            prev_sr = None
            if row_num > 8:
                prev_sr = self.read_cell(ws, row_num - 1, self.punch_cols['sr_no'])

            try:
                sr_no_assigned = int(prev_sr) + 1 if prev_sr is not None else 1
            except:
                sr_no_assigned = 1

            self.write_cell(ws, row_num, self.punch_cols['sr_no'], sr_no_assigned)
            self.write_cell(ws, row_num, self.punch_cols['ref_no'], ref_no)
            self.write_cell(ws, row_num, self.punch_cols['desc'], custom_action)
            self.write_cell(ws, row_num, self.punch_cols['category'], custom_category)

            try:
                uname = os.getlogin()
            except:
                uname = getpass.getuser()

            self.write_cell(ws, row_num, self.punch_cols['checked_name'], uname)
            self.write_cell(ws, row_num, self.punch_cols['checked_date'], datetime.now().strftime("%Y-%m-%d"))

            wb.save(self.excel_file)
            wb.close()

            ann = {
                'type': 'error',
                'page': self.current_page,
                'bbox_page': bbox_page,
                'component': custom_category,
                'tag_name': tag_name,
                'error': 'Custom',
                'punch_text': custom_action,
                'ref_no': ref_no,
                'excel_row': row_num,
                'sr_no': sr_no_assigned,
                'timestamp': datetime.now().isoformat()
            }
            self.annotations.append(ann)

            self.current_sr_no = self.get_next_sr_no()
            self.display_page()

            self.root.after(100, lambda: messagebox.showinfo("Logged", f"Custom punch logged:\n{custom_action}"))

            updated = self.update_interphase_status_for_ref(ref_no, status='NOK')
            if updated:
                print(f"Interphase: marked ref {ref_no} as NOK")

        except PermissionError:
            messagebox.showerror("Error", "Close the Excel file before writing to it.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to log custom error:\n{e}")

    # ================================================================
    # CATEGORY HANDLERS
    # ================================================================

    def handle_template_category(self, category, bbox_page):
        punch_text = self.run_template(category, tag_name=None)
        if not punch_text:
            return

        self.log_error(
            component_type=category["name"],
            error_name=None,
            error_template=punch_text,
            bbox_page=bbox_page,
            tag_name=None
        )

    def handle_subcategory(self, category, subcategory, bbox_page):
        punch_text = self.run_template(subcategory, tag_name=None)
        if not punch_text:
            return

        self.log_error(
            component_type=category["name"],
            error_name=subcategory["name"],
            error_template=punch_text,
            bbox_page=bbox_page,
            tag_name=None
        )

    def run_template(self, template_def, tag_name=None):
        """Execute a template definition at runtime."""
        values = {}

        if tag_name:
            values["tag"] = tag_name

        for inp in template_def.get("inputs", []):
            val = simpledialog.askstring("Input Required", inp["label"], parent=self.root)
            if not val:
                return None
            values[inp["name"]] = val.strip()

        try:
            return template_def["template"].format(**values)
        except KeyError as e:
            messagebox.showerror("Template Error", f"Missing placeholder: {e}")
            return None

    def load_categories(self):
        try:
            if os.path.exists(self.category_file):
                with open(self.category_file, "r", encoding="utf-8") as f:
                    self.categories = json.load(f)
                    print("Categories loaded:", self.categories)
            else:
                print(f"Warning: categories.json not found at {self.category_file}")
                self.categories = []
        except Exception as e:
            print(f"Error loading categories: {e}")
            self.categories = []

    # ================================================================
    # PUNCH CLOSING MODE
    # ================================================================

    def punch_closing_mode(self):
        punches = self.read_open_punches_from_excel()

        if not punches:
            messagebox.showinfo("No Punches", "No open punches found in Excel.")
            return

        punches.sort(key=lambda p: (not p['implemented'], p['sr_no']))

        dlg = tk.Toplevel(self.root)
        dlg.title("Punch Closing Mode")
        dlg.geometry("800x420")
        dlg.transient(self.root)
        dlg.grab_set()

        idx_label = tk.Label(dlg, text="", font=('Arial', 10, 'bold'))
        idx_label.pack(pady=(10, 0))

        text_widget = tk.Text(dlg, wrap=tk.WORD, height=16)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        text_widget.config(state=tk.DISABLED)

        pos = [0]

        def show_item():
            p = punches[pos[0]]
            idx_label.config(text=f"Item {pos[0]+1}/{len(punches)} | SR {p['sr_no']} | Ref {p['ref_no']}")

            text_widget.config(state=tk.NORMAL)
            text_widget.delete("1.0", tk.END)
            text_widget.insert(tk.END, p['punch_text'])
            text_widget.insert(tk.END, f"\n\n---\nCategory: {p['category']}\nImplemented: {'YES' if p['implemented'] else 'NO'}\n")

            ann = next((a for a in self.annotations if a.get('sr_no') == p['sr_no']), None)
            if ann and ann.get('implementation_remark'):
                text_widget.insert(tk.END, "\n---\nImplementation Remarks:\n" + ann['implementation_remark'])

            text_widget.config(state=tk.DISABLED)

        show_item()

        def close_punch():
            p = punches[pos[0]]

            try:
                default_user = os.getlogin()
            except:
                default_user = getpass.getuser()

            name = simpledialog.askstring("Closed By", "Enter your name:", initialvalue=default_user, parent=dlg)
            if not name:
                return

            try:
                wb = load_workbook(self.excel_file)
                ws = wb[self.punch_sheet_name]

                self.write_cell(ws, p['row'], self.punch_cols['closed_name'], name)
                self.write_cell(ws, p['row'], self.punch_cols['closed_date'], datetime.now().strftime("%Y-%m-%d"))

                wb.save(self.excel_file)
                wb.close()

            except Exception as e:
                messagebox.showerror("Excel Error", str(e))
                return

            ann = next((a for a in self.annotations if a.get('excel_row') == p['row']), None)
            if ann:
                ann['type'] = 'ok'
                ann['closed_by'] = name
                ann['closed_date'] = datetime.now().strftime("%Y-%m-%d")

            self.display_page()

            if pos[0] < len(punches) - 1:
                pos[0] += 1
                show_item()
            else:
                messagebox.showinfo("Done", "All punches closed.")
                dlg.destroy()

        def next_item():
            if pos[0] < len(punches) - 1:
                pos[0] += 1
                show_item()

        def prev_item():
            if pos[0] > 0:
                pos[0] -= 1
                show_item()

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(fill=tk.X, pady=8)

        tk.Button(btn_frame, text="◀ Prev", command=prev_item, width=10).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="OKAY", command=close_punch, bg="#2ecc71", fg="white", width=12).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Next ▶", command=next_item, width=10).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side=tk.RIGHT, padx=6)

    def read_open_punches_from_excel(self):
        """Reads punch sheet and returns list of open punches."""
        punches = []

        if not self.excel_file or not os.path.exists(self.excel_file):
            return punches

        wb = load_workbook(self.excel_file, data_only=True)
        ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active

        row = 8
        while True:
            sr = self.read_cell(ws, row, self.punch_cols['sr_no'])
            if sr is None:
                break

            closed = self.read_cell(ws, row, self.punch_cols['closed_name'])
            if closed:
                row += 1
                continue

            implemented = bool(self.read_cell(ws, row, self.punch_cols['implemented_name']))

            punches.append({
                'sr_no': sr,
                'row': row,
                'ref_no': self.read_cell(ws, row, self.punch_cols['ref_no']),
                'punch_text': self.read_cell(ws, row, self.punch_cols['desc']),
                'category': self.read_cell(ws, row, self.punch_cols['category']),
                'implemented': implemented
            })

            row += 1

        wb.close()
        return punches

    # ================================================================
    # PROJECT DETAILS
    # ================================================================

    def ask_project_details(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Project Details")
        dlg.geometry("420x260")
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="Cabinet ID").pack(anchor="w", padx=20, pady=(15, 0))
        cabinet_var = tk.StringVar(value=getattr(self, "cabinet_id", ""))
        tk.Entry(dlg, textvariable=cabinet_var).pack(fill="x", padx=20)

        tk.Label(dlg, text="Project Name").pack(anchor="w", padx=20, pady=(10, 0))
        project_var = tk.StringVar(value=self.project_name)
        tk.Entry(dlg, textvariable=project_var).pack(fill="x", padx=20)

        tk.Label(dlg, text="Sales Order Number").pack(anchor="w", padx=20, pady=(10, 0))
        so_var = tk.StringVar(value=self.sales_order_no)
        tk.Entry(dlg, textvariable=so_var).pack(fill="x", padx=20)

        def on_ok():
            self.cabinet_id = cabinet_var.get().strip()
            self.project_name = project_var.get().strip()
            self.sales_order_no = so_var.get().strip()
            dlg.destroy()

        tk.Button(dlg, text="OK", command=on_ok, bg="#2ecc71", fg="white").pack(pady=20)
        dlg.wait_window()

    def write_project_details_to_excel(self):
        if not self.excel_file or not os.path.exists(self.excel_file):
            return

        try:
            wb = load_workbook(self.excel_file)

            for sheet_name, cells in self.header_cells.items():
                if sheet_name not in wb.sheetnames:
                    continue

                ws = wb[sheet_name]

                if getattr(self, "project_name", ""):
                    r, c = self.split_cell(cells["project_name"])
                    self.write_cell(ws, r, c, self.project_name)

                if getattr(self, "sales_order_no", ""):
                    r, c = self.split_cell(cells["sales_order"])
                    self.write_cell(ws, r, c, self.sales_order_no)

                if getattr(self, "cabinet_id", ""):
                    r, c = self.split_cell(cells["cabinet_id"])
                    self.write_cell(ws, r, c, self.cabinet_id)

            wb.save(self.excel_file)
            wb.close()

        except PermissionError:
            messagebox.showerror("Excel Locked", "Please close the Excel file before entering project details.")
        except Exception as e:
            messagebox.showerror("Excel Error", f"Failed to write project details:\n{e}")

    def prepare_project_folders(self):
        if not self.project_name:
            raise ValueError("Project name not set")

        safe_project = "".join(
            c for c in self.project_name if c.isalnum() or c in (" ", "_", "-")
        ).strip().replace(" ", "_")

        base_dir = get_app_base_dir()
        project_root = os.path.join(base_dir, safe_project)

        folders = {
            "root": project_root,
            "working_excel": os.path.join(project_root, "Working_Excel"),
            "interphase_export": os.path.join(project_root, "Interphase_Export"),
            "annotated_drawings": os.path.join(project_root, "Annotated_Drawings"),
            "sessions": os.path.join(project_root, "Sessions")
        }

        for p in folders.values():
            os.makedirs(p, exist_ok=True)

        self.project_dirs = folders

    def get_session_path_for_pdf(self):
        if not self.current_pdf_path:
            return None

        session_path = os.path.join(
            self.project_dirs.get("sessions", ""),
            f"{self.cabinet_id}_annotations.json"
        )

        return session_path if os.path.exists(session_path) else None

    # ================================================================
    # CHECKLIST FUNCTIONS
    # ================================================================

    def review_checklist_now(self):
        if not self.excel_file or not os.path.exists(self.excel_file):
            messagebox.showerror("Excel Missing", "Working Excel file not found.")
            return

        self.checklist_file = self.excel_file

        try:
            self.review_checklist_before_save(self.checklist_file, self.session_refs)
        except Exception as e:
            messagebox.showerror("Checklist Error", f"Checklist review failed:\n{e}")

    def gather_checklist_matches(self, checklist_path, refs_set):
        """Returns Interphase rows where Reference No is NOT in refs_set."""
        wb = load_workbook(checklist_path)
        if self.interphase_sheet_name not in wb.sheetnames:
            wb.close()
            raise ValueError("Interphase sheet not found")

        ws = wb[self.interphase_sheet_name]
        ref_col = self.interphase_cols['ref_no']
        desc_col = self.interphase_cols['description']
        status_col = self.interphase_cols['status']

        matches = []
        max_row = ws.max_row if ws.max_row else 2000

        for r in range(11, max_row + 1):
            ref_val = self.read_cell(ws, r, ref_col)
            if ref_val is None:
                continue

            ref_str = str(ref_val).strip()

            if ref_str in refs_set:
                continue

            status_val = self.read_cell(ws, r, status_col)
            status_str = str(status_val).strip().lower() if status_val is not None else ''

            if status_str in ('ok', 'nok', 'n/a', 'na', 'not applicable'):
                continue

            desc_val = self.read_cell(ws, r, desc_col) or ''
            matches.append((r, ref_str, str(desc_val)))

        wb.close()
        return {'ref_col': ref_col, 'desc_col': desc_col, 'status_col': status_col}, matches

    def review_checklist_before_save(self, checklist_path, refs_set):
        """Dialog for reviewing and marking checklist items."""
        try:
            cols, matches = self.gather_checklist_matches(checklist_path, refs_set)
        except Exception as e:
            raise

        if not matches:
            messagebox.showinfo("Checklist", "No matching Interphase items found for references NOT in session.")
            return

        wb = load_workbook(checklist_path)
        ws = wb[self.interphase_sheet_name]
        status_col = cols['status_col']

        dlg = tk.Toplevel(self.root)
        dlg.title("Interphase Checklist Review")
        dlg.geometry("900x420")
        dlg.transient(self.root)
        dlg.grab_set()

        idx_label = tk.Label(dlg, text="", font=('Arial', 10, 'bold'))
        idx_label.pack(pady=(10, 0))

        txt_frame = tk.Frame(dlg)
        txt_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        text_widget = tk.Text(txt_frame, wrap=tk.WORD, height=12)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.config(state=tk.DISABLED)

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(fill=tk.X, pady=8)

        pos = [0]

        def show_item(p):
            r, ref_str, desc = matches[p]
            idx_label.config(text=f"Item {p+1} / {len(matches)}  |  Row: {r}  | Ref: {ref_str}")
            text_widget.config(state=tk.NORMAL)
            text_widget.delete('1.0', tk.END)
            text_widget.insert(tk.END, f"Description:\n\n{desc}")
            text_widget.config(state=tk.DISABLED)

        show_item(pos[0])

        def do_action_set_status(status_value):
            r, ref_str, desc = matches[pos[0]]

            try:
                self.write_cell(ws, r, status_col, status_value)
                wb.save(checklist_path)
            except PermissionError:
                messagebox.showerror("Excel Locked", f"Please close checklist file and try again.")
                return
            except Exception as e:
                messagebox.showerror("Excel Error", f"Failed writing to checklist: {e}")
                return

            if pos[0] < len(matches) - 1:
                pos[0] += 1
                show_item(pos[0])
            else:
                messagebox.showinfo("Done", "Checklist review finished.")
                dlg.destroy()

        def on_ok():
            do_action_set_status("OK")

        def on_nok():
            do_action_set_status("NOK")

        def on_na():
            do_action_set_status("N/A")

        def on_prev():
            if pos[0] > 0:
                pos[0] -= 1
                show_item(pos[0])

        def on_next():
            if pos[0] < len(matches) - 1:
                pos[0] += 1
                show_item(pos[0])

        tk.Button(btn_frame, text="◀ Prev", command=on_prev, width=10).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="OK", command=on_ok, bg='#2ecc71', fg='white', width=12).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="NOK", command=on_nok, bg='#e74c3c', fg='white', width=12).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Not applicable", command=on_na, bg='#f39c12', fg='white', width=16).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Next ▶", command=on_next, width=10).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Cancel", command=lambda: dlg.destroy()).pack(side=tk.RIGHT, padx=6)

        dlg.wait_window()
        wb.close()

    # ================================================================
    # EXCEL HELPERS
    # ================================================================

    def save_interphase_excel(self):
        if not self.current_pdf_path:
            messagebox.showwarning("No PDF", "Load a PDF first.")
            return

        if not self.excel_file or not os.path.exists(self.excel_file):
            messagebox.showerror("Missing File", "Working Excel file not found.")
            return

        save_path = os.path.join(
            self.project_dirs["interphase_export"],
            f"{self.cabinet_id.replace(' ', '_')}_Interphase.xlsx"
        )

        try:
            shutil.copy2(self.excel_file, save_path)
            messagebox.showinfo("Saved", f"Final Interphase Excel saved:\n\n{save_path}")
        except PermissionError:
            messagebox.showerror("File Open", "Close the Excel file and try again.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel:\n{e}")

    def open_excel(self):
        if not self.excel_file or not os.path.exists(self.excel_file):
            messagebox.showwarning("No Excel", "No working Excel file found.")
            return

        try:
            if os.name == 'nt':
                os.startfile(self.excel_file)
            else:
                if sys.platform == 'darwin':
                    cmd = f"open {shlex.quote(self.excel_file)}"
                else:
                    cmd = f"xdg-open {shlex.quote(self.excel_file)}"
                subprocess.Popen(cmd, shell=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel: {e}")

    # ================================================================
    # FUZZY MATCH HELPER
    # ================================================================

    def find_row_by_sr_or_text(self, sr_no, punch_text, min_ratio=0.60):
        try:
            wb = load_workbook(self.excel_file, read_only=True)
            ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active
            row = 8

            while True:
                cell = self.read_cell(ws, row, self.punch_cols['sr_no'])
                if cell is None:
                    if self.read_cell(ws, row, self.punch_cols['desc']) is None:
                        break
                    else:
                        row += 1
                        continue
                try:
                    if int(cell) == int(sr_no):
                        wb.close()
                        return (row, 1.0, 'sr_exact')
                except:
                    if str(cell).strip() == str(sr_no).strip():
                        wb.close()
                        return (row, 1.0, 'sr_exact')
                row += 1
                if row > 2000:
                    break

            best_row = None
            best_ratio = 0.0
            row = 8

            while True:
                txt = self.read_cell(ws, row, self.punch_cols['desc'])
                if txt is None:
                    if row > 2000:
                        break
                    row += 1
                    continue
                try:
                    ratio = SequenceMatcher(None, str(punch_text).strip().lower(), str(txt).strip().lower()).ratio()
                except:
                    ratio = 0.0
                if ratio > best_ratio:
                    best_ratio = ratio
                    best_row = row
                row += 1
                if row > 2000:
                    break

            wb.close()
            if best_row and best_ratio >= min_ratio:
                return (best_row, best_ratio, 'fuzzy_text')
            return (None, best_ratio, None)
        except Exception as e:
            try:
                wb.close()
            except:
                pass
            return (None, 0.0, None)


# ================================================================
# MAIN ENTRY POINT
# ================================================================

def main():
    root = tk.Tk()
    app = CircuitInspector(root)
    root.mainloop()


if __name__ == "__main__":
    main()
