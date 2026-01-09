"""
PDF HIGHLIGHTER TEST APPLICATION
A complete standalone app to test the highlighter functionality before integration
NOW WITH AUTOMATIC STRAIGHTENING OF ANNOTATIONS
"""

import tkinter as tk
from tkinter import messagebox, simpledialog, Menu, filedialog
from PIL import Image, ImageDraw, ImageFont, ImageTk
import fitz  # PyMuPDF
from datetime import datetime
import os
import math

class HighlighterTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Highlighter Test Application - Auto-Straighten")
        self.root.geometry("1400x900")
        self.root.configure(bg='#0f172a')
        
        # PDF state
        self.pdf_document = None
        self.current_page = 0
        self.zoom_level = 1.0
        self.photo = None
        self.current_page_image = None
        
        # Highlighter state
        self.active_highlighter = None
        self.highlighter_colors = {
            'green': {'rgb': (0, 255, 0), 'rgba': (0, 255, 0, 100), 'name': '✓ OK'},
            'orange': {'rgb': (255, 165, 0), 'rgba': (255, 165, 0, 120), 'name': '✗ Error'},
            'yellow': {'rgb': (255, 255, 0), 'rgba': (255, 255, 0, 80), 'name': '⚠ Review'},
            'pink': {'rgb': (255, 105, 180), 'rgba': (255, 105, 180, 80), 'name': '💡 Note'},
            'blue': {'rgb': (100, 149, 237), 'rgba': (100, 149, 237, 80), 'name': 'ℹ Info'}
        }
        self.highlight_points = []
        self.temp_line_ids = []
        
        # Drawing state
        self.drawing = False
        self.drawing_type = None
        self.annotations = []
        
        # Tool state
        self.tool_mode = None  # 'pen' or 'text'
        self.pen_points = []
        self.text_pos_x = 0
        self.text_pos_y = 0
        
        # Straightening toggle
        self.auto_straighten = tk.BooleanVar(value=True)
        
        self.setup_ui()
    
    def straighten_path(self, points):
        """
        Convert a freehand path into a straight line from start to end.
        Returns a list with just the start and end points.
        """
        if len(points) < 2:
            return points
        
        # Simply return start and end points for a perfectly straight line
        return [points[0], points[-1]]
    
    def simplify_path(self, points, tolerance=5.0):
        """
        Simplify a path using the Ramer-Douglas-Peucker algorithm.
        This reduces the number of points while preserving the shape.
        """
        if len(points) < 3:
            return points
        
        # Find the point with the maximum distance from the line segment
        def perpendicular_distance(point, line_start, line_end):
            x0, y0 = point
            x1, y1 = line_start
            x2, y2 = line_end
            
            dx = x2 - x1
            dy = y2 - y1
            
            if dx == 0 and dy == 0:
                return math.sqrt((x0 - x1)**2 + (y0 - y1)**2)
            
            return abs(dy * x0 - dx * y0 + x2 * y1 - y2 * x1) / math.sqrt(dx**2 + dy**2)
        
        dmax = 0
        index = 0
        end = len(points)
        
        for i in range(1, end - 1):
            d = perpendicular_distance(points[i], points[0], points[-1])
            if d > dmax:
                index = i
                dmax = d
        
        # If max distance is greater than tolerance, recursively simplify
        if dmax > tolerance:
            results1 = self.simplify_path(points[:index + 1], tolerance)
            results2 = self.simplify_path(points[index:], tolerance)
            return results1[:-1] + results2
        else:
            return [points[0], points[-1]]
        
    def setup_ui(self):
        """Setup the user interface"""
        self.setup_toolbar()
        self.setup_canvas()
        self.setup_status_bar()
        
    def setup_toolbar(self):
        """Create modern highlighter toolbar"""
        toolbar = tk.Frame(self.root, bg='#1e293b', height=80)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        # Left section - File operations
        left_frame = tk.Frame(toolbar, bg='#1e293b')
        left_frame.pack(side=tk.LEFT, padx=10, pady=10)
        
        btn_style = {
            'bg': '#3b82f6',
            'fg': 'white',
            'padx': 12,
            'pady': 10,
            'font': ('Segoe UI', 9, 'bold'),
            'relief': tk.FLAT,
            'cursor': 'hand2'
        }
        
        tk.Button(left_frame, text="📁 Open PDF", command=self.load_pdf, **btn_style).pack(side=tk.LEFT, padx=3)
        tk.Button(left_frame, text="💾 Save", command=self.save_annotations, **btn_style).pack(side=tk.LEFT, padx=3)
        
        export_btn_style = btn_style.copy()
        export_btn_style['bg'] = '#10b981'
        tk.Button(left_frame, text="📄 Export PDF", command=self.export_pdf, **export_btn_style).pack(side=tk.LEFT, padx=3)
        
        tk.Button(left_frame, text="🗑️ Clear All", command=self.clear_all, **btn_style).pack(side=tk.LEFT, padx=3)
        
        # Straighten toggle
        straighten_frame = tk.Frame(toolbar, bg='#1e293b')
        straighten_frame.pack(side=tk.LEFT, padx=10)
        
        tk.Checkbutton(
            straighten_frame,
            text="📏 Auto-Straighten",
            variable=self.auto_straighten,
            bg='#1e293b',
            fg='white',
            selectcolor='#3b82f6',
            font=('Segoe UI', 9, 'bold'),
            activebackground='#1e293b',
            activeforeground='white'
        ).pack(pady=10)
        
        # Center section - HIGHLIGHTER PALETTE
        highlighter_frame = tk.Frame(toolbar, bg='#1e293b')
        highlighter_frame.pack(side=tk.LEFT, padx=30)
        
        tk.Label(highlighter_frame, text="Highlighters:", bg='#1e293b', fg='#94a3b8', 
                 font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 10))
        
        # Store highlighter buttons for toggling
        self.highlighter_btns = {}
        
        # Create highlighter buttons
        for color_key, color_info in self.highlighter_colors.items():
            rgb = color_info['rgb']
            hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
            
            btn = tk.Button(
                highlighter_frame,
                text=color_info['name'],
                command=lambda ck=color_key: self.select_highlighter(ck),
                bg=hex_color,
                fg='white' if color_key != 'yellow' else 'black',
                font=('Segoe UI', 10, 'bold'),
                relief=tk.RAISED,
                borderwidth=3,
                padx=12,
                pady=8,
                cursor='hand2',
                width=10
            )
            btn.pack(side=tk.LEFT, padx=3)
            self.highlighter_btns[color_key] = btn
        
        # Active highlighter indicator
        self.highlighter_indicator = tk.Label(
            toolbar,
            text="🖱️ No Highlighter",
            bg='#475569',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            padx=20,
            pady=10,
            relief=tk.FLAT
        )
        self.highlighter_indicator.pack(side=tk.LEFT, padx=15)
        
        # Navigation controls
        nav_frame = tk.Frame(toolbar, bg='#1e293b')
        nav_frame.pack(side=tk.LEFT, padx=10)
        
        self.page_label = tk.Label(nav_frame, text="Page: 0/0", bg='#1e293b', 
                                  fg='white', font=('Segoe UI', 10, 'bold'))
        self.page_label.pack(side=tk.LEFT, padx=10)
        
        nav_btn_style = btn_style.copy()
        nav_btn_style['bg'] = '#64748b'
        
        tk.Button(nav_frame, text="◀", command=self.prev_page, width=3, **nav_btn_style).pack(side=tk.LEFT, padx=2)
        tk.Button(nav_frame, text="▶", command=self.next_page, width=3, **nav_btn_style).pack(side=tk.LEFT, padx=2)
        
        # Zoom controls
        zoom_frame = tk.Frame(nav_frame, bg='#1e293b')
        zoom_frame.pack(side=tk.LEFT, padx=10)
        
        zoom_btn_style = btn_style.copy()
        zoom_btn_style['bg'] = '#10b981'
        
        tk.Button(zoom_frame, text="🔍+", command=self.zoom_in, width=4, **zoom_btn_style).pack(side=tk.LEFT, padx=2)
        tk.Button(zoom_frame, text="🔍−", command=self.zoom_out, width=4, **zoom_btn_style).pack(side=tk.LEFT, padx=2)
        
        # Tool buttons
        tool_frame = tk.Frame(toolbar, bg='#1e293b')
        tool_frame.pack(side=tk.LEFT, padx=10)
        
        tool_btn_style = btn_style.copy()
        tool_btn_style['bg'] = '#334155'
        
        self.pen_btn = tk.Button(tool_frame, text="✏️ Pen", 
                                 command=lambda: self.set_tool_mode("pen"),
                                 **tool_btn_style)
        self.pen_btn.pack(side=tk.LEFT, padx=2)
        
        self.text_btn = tk.Button(tool_frame, text="🅰️ Text", 
                                  command=lambda: self.set_tool_mode("text"),
                                  **tool_btn_style)
        self.text_btn.pack(side=tk.LEFT, padx=2)
        
        # Stats
        stats_frame = tk.Frame(toolbar, bg='#1e293b')
        stats_frame.pack(side=tk.RIGHT, padx=10)
        
        self.stats_label = tk.Label(
            stats_frame,
            text="Annotations: 0",
            bg='#1e293b',
            fg='#94a3b8',
            font=('Segoe UI', 9, 'bold')
        )
        self.stats_label.pack(pady=10)
    
    def setup_canvas(self):
        """Setup the canvas with scrollbars"""
        canvas_frame = tk.Frame(self.root, bg='#f1f5f9')
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        v_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.canvas = tk.Canvas(canvas_frame, bg='#f8fafc',
                               yscrollcommand=v_scrollbar.set,
                               xscrollcommand=h_scrollbar.set,
                               highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        v_scrollbar.config(command=self.canvas.yview)
        h_scrollbar.config(command=self.canvas.xview)
        
        # Bind mouse events
        self.canvas.bind("<ButtonPress-1>", self.on_left_press)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_release)
        self.root.bind("<Escape>", lambda e: self.deactivate_all())
    
    def setup_status_bar(self):
        """Setup the status bar"""
        status_bar = tk.Frame(self.root, bg='#334155', height=40)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        instructions_text = "🖍️ Select Highlighter → Drag to Highlight | 📏 Auto-Straighten: ON | Esc: Deactivate"
        tk.Label(status_bar, text=instructions_text, bg='#334155', fg='#e2e8f0', 
                 font=('Segoe UI', 9), pady=10).pack()
    
    def select_highlighter(self, color_key):
        """Select or deselect a highlighter color"""
        # Toggle off if clicking active highlighter
        if self.active_highlighter == color_key:
            self.active_highlighter = None
            self.root.config(cursor="")
            
            # Reset all button styles
            for key, btn in self.highlighter_btns.items():
                rgb = self.highlighter_colors[key]['rgb']
                hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
                btn.config(bg=hex_color, relief=tk.RAISED, borderwidth=3)
            
            self.highlighter_indicator.config(
                text="🖱️ No Highlighter",
                bg='#475569'
            )
            return
        
        # Deactivate pen/text tools if active
        if self.tool_mode:
            self.tool_mode = None
            self.pen_btn.config(bg='#334155', relief=tk.FLAT)
            self.text_btn.config(bg='#334155', relief=tk.FLAT)
        
        # Activate new highlighter
        self.active_highlighter = color_key
        
        # Update button styles
        for key, btn in self.highlighter_btns.items():
            rgb = self.highlighter_colors[key]['rgb']
            hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
            
            if key == color_key:
                btn.config(bg='#1e293b', relief=tk.SUNKEN, borderwidth=5)
            else:
                btn.config(bg=hex_color, relief=tk.RAISED, borderwidth=3)
        
        # Update indicator
        color_info = self.highlighter_colors[color_key]
        rgb = color_info['rgb']
        hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
        
        self.highlighter_indicator.config(
            text=f"🖍️ {color_info['name']} Active",
            bg=hex_color,
            fg='white' if color_key != 'yellow' else 'black'
        )
        
        # Change cursor to indicate highlighter mode
        self.root.config(cursor="pencil")
    
    def set_tool_mode(self, mode):
        """Set tool mode (pen or text)"""
        if self.active_highlighter:
            self.select_highlighter(self.active_highlighter)  # Deactivate highlighter
        
        if self.tool_mode == mode:
            self.tool_mode = None
            if mode == "pen":
                self.pen_btn.config(bg='#334155', relief=tk.FLAT)
            else:
                self.text_btn.config(bg='#334155', relief=tk.FLAT)
        else:
            self.tool_mode = mode
            if mode == "pen":
                self.pen_btn.config(bg='#3b82f6', relief=tk.SUNKEN)
                self.text_btn.config(bg='#334155', relief=tk.FLAT)
            else:
                self.text_btn.config(bg='#3b82f6', relief=tk.SUNKEN)
                self.pen_btn.config(bg='#334155', relief=tk.FLAT)
    
    def deactivate_all(self):
        """Deactivate all tools and highlighters"""
        if self.active_highlighter:
            self.select_highlighter(self.active_highlighter)
        if self.tool_mode:
            self.set_tool_mode(self.tool_mode)
    
    def on_left_press(self, event):
        """Handle mouse press - start highlighting if highlighter is active"""
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please load a PDF first")
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # HIGHLIGHTER MODE
        if self.active_highlighter:
            self.drawing = True
            self.drawing_type = "highlight"
            self.highlight_points = [(x, y)]
            self.clear_temp_drawings()
            return

        # PEN TOOL
        if self.tool_mode == "pen":
            self.drawing = True
            self.drawing_type = "pen"
            self.pen_points = [(x, y)]
            self.clear_temp_drawings()
            return

        # TEXT TOOL
        if self.tool_mode == "text":
            self.drawing = True
            self.drawing_type = "text"
            self.text_pos_x = x
            self.text_pos_y = y
            return
    
    def on_left_drag(self, event):
        """Handle mouse drag - draw highlight preview"""
        if not self.drawing:
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        # HIGHLIGHTER DRAWING
        if self.drawing_type == "highlight":
            if len(self.highlight_points) > 0:
                last_x, last_y = self.highlight_points[-1]
                
                # Get highlighter color
                rgba = self.highlighter_colors[self.active_highlighter]['rgba']
                rgb = self.highlighter_colors[self.active_highlighter]['rgb']
                hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
                
                # Draw thick line segment
                line_id = self.canvas.create_line(
                    last_x, last_y, x, y,
                    fill=hex_color,
                    width=max(15, int(15 * self.zoom_level)),
                    capstyle=tk.ROUND,
                    smooth=True
                )
                self.temp_line_ids.append(line_id)
            
            self.highlight_points.append((x, y))
            return

        # PEN TOOL DRAWING
        if self.drawing_type == "pen":
            if len(self.pen_points) > 0:
                last_x, last_y = self.pen_points[-1]
                line_id = self.canvas.create_line(
                    last_x, last_y, x, y,
                    fill="red", width=3,
                    capstyle=tk.ROUND, smooth=True
                )
                self.temp_line_ids.append(line_id)
            self.pen_points.append((x, y))
            return
    
    def on_left_release(self, event):
        """Handle mouse release - finalize highlight"""
        if not self.pdf_document or not self.drawing:
            return

        # HIGHLIGHTER FINISH
        if self.drawing_type == "highlight":
            if len(self.highlight_points) >= 2:
                # Apply straightening if enabled
                if self.auto_straighten.get():
                    processed_points = self.straighten_path(self.highlight_points)
                else:
                    processed_points = self.highlight_points.copy()
                
                # Calculate bounding box for the highlight
                xs = [p[0] for p in processed_points]
                ys = [p[1] for p in processed_points]
                bbox = (min(xs), min(ys), max(xs), max(ys))
                
                # Create annotation
                annotation = {
                    'type': 'highlight',
                    'color': self.active_highlighter,
                    'page': self.current_page,
                    'bbox': bbox,
                    'points': processed_points,
                    'timestamp': datetime.now().isoformat()
                }
                
                # Show action menu for orange highlighter
                if self.active_highlighter == 'orange':
                    action = messagebox.askyesno(
                        "Error Highlight",
                        "Would you like to add a note to this error highlight?"
                    )
                    if action:
                        note = simpledialog.askstring("Note", "Enter error note:", parent=self.root)
                        if note:
                            annotation['note'] = note
                
                self.annotations.append(annotation)
                self.update_stats()
            
            self.highlight_points = []
            self.clear_temp_drawings()
            self.drawing = False
            self.drawing_type = None
            self.display_page()
            return

        # PEN TOOL FINISH
        if self.drawing_type == "pen":
            if len(self.pen_points) >= 2:
                # Apply straightening if enabled
                if self.auto_straighten.get():
                    processed_points = self.straighten_path(self.pen_points)
                else:
                    processed_points = self.pen_points.copy()
                
                annotation = {
                    'type': 'pen',
                    'page': self.current_page,
                    'points': processed_points,
                    'timestamp': datetime.now().isoformat()
                }
                self.annotations.append(annotation)
                self.update_stats()
            
            self.pen_points = []
            self.clear_temp_drawings()
            self.drawing = False
            self.drawing_type = None
            self.display_page()
            return

        # TEXT TOOL FINISH
        if self.drawing_type == "text":
            txt = simpledialog.askstring("Text", "Enter text:", parent=self.root)
            if txt and txt.strip():
                annotation = {
                    'type': 'text',
                    'page': self.current_page,
                    'pos': (self.text_pos_x, self.text_pos_y),
                    'text': txt.strip(),
                    'timestamp': datetime.now().isoformat()
                }
                self.annotations.append(annotation)
                self.update_stats()
                self.display_page()
            
            self.drawing = False
            self.drawing_type = None
            return
    
    def clear_temp_drawings(self):
        """Clear temporary drawing lines"""
        for line_id in self.temp_line_ids:
            self.canvas.delete(line_id)
        self.temp_line_ids = []
    
    def load_pdf(self):
        """Load a PDF file"""
        filepath = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            self.pdf_document = fitz.open(filepath)
            self.current_page = 0
            self.annotations = []
            self.display_page()
            messagebox.showinfo("Success", f"Loaded PDF with {len(self.pdf_document)} pages")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF:\n{e}")
    
    def display_page(self):
        """Render the current PDF page with all annotations"""
        if not self.pdf_document:
            self.canvas.delete("all")
            self.page_label.config(text="Page: 0/0")
            return

        try:
            page = self.pdf_document[self.current_page]
            mat = fitz.Matrix(self.zoom_level, self.zoom_level)
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            draw = ImageDraw.Draw(img, 'RGBA')

            # Try to load font
            try:
                font = ImageFont.truetype("arial.ttf", max(12, int(14 * self.zoom_level)))
            except:
                font = ImageFont.load_default()

            for ann in self.annotations:
                if ann.get('page') != self.current_page:
                    continue

                ann_type = ann.get('type')

                # HIGHLIGHT STROKES
                if ann_type == 'highlight' and 'points' in ann:
                    points = ann['points']
                    if len(points) >= 2:
                        color_key = ann.get('color', 'yellow')
                        rgba = self.highlighter_colors[color_key]['rgba']
                        
                        # Draw thick semi-transparent strokes
                        stroke_width = max(15, int(15 * self.zoom_level))
                        for i in range(len(points) - 1):
                            x1, y1 = points[i]
                            x2, y2 = points[i + 1]
                            draw.line([x1, y1, x2, y2], fill=rgba, width=stroke_width)

                # PEN STROKES
                elif ann_type == 'pen' and 'points' in ann:
                    points = ann['points']
                    if len(points) >= 2:
                        stroke_width = max(2, int(3 * self.zoom_level))
                        for i in range(len(points) - 1):
                            x1, y1 = points[i]
                            x2, y2 = points[i + 1]
                            draw.line([x1, y1, x2, y2], fill='red', width=stroke_width)

                # TEXT ANNOTATIONS
                elif ann_type == 'text' and 'pos' in ann:
                    pos = ann['pos']
                    text = ann.get('text', '')
                    if text:
                        try:
                            bbox = draw.textbbox(pos, text, font=font)
                            padding = 2
                            draw.rectangle(
                                [bbox[0] - padding, bbox[1] - padding,
                                 bbox[2] + padding, bbox[3] + padding],
                                fill=(255, 255, 200, 200)
                            )
                        except:
                            pass
                        draw.text(pos, text, fill='red', font=font)

            self.photo = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
            self.page_label.config(text=f"Page: {self.current_page + 1}/{len(self.pdf_document)}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to display page: {e}")
    
    def prev_page(self):
        """Go to previous page"""
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.display_page()
    
    def next_page(self):
        """Go to next page"""
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.display_page()
    
    def zoom_in(self):
        """Zoom in"""
        self.zoom_level = min(3.0, self.zoom_level + 0.2)
        self.display_page()
    
    def zoom_out(self):
        """Zoom out"""
        self.zoom_level = max(0.5, self.zoom_level - 0.2)
        self.display_page()
    
    def update_stats(self):
        """Update annotation statistics"""
        total = len(self.annotations)
        by_color = {}
        for ann in self.annotations:
            if ann['type'] == 'highlight':
                color = ann.get('color', 'unknown')
                by_color[color] = by_color.get(color, 0) + 1
        
        stats_text = f"Annotations: {total}"
        if by_color:
            color_stats = ", ".join([f"{self.highlighter_colors[k]['name']}: {v}" for k, v in by_color.items()])
            stats_text += f" ({color_stats})"
        
        self.stats_label.config(text=stats_text)
    
    def save_annotations(self):
        """Save annotations to a file"""
        if not self.annotations:
            messagebox.showinfo("Info", "No annotations to save")
            return
        
        filepath = filedialog.asksaveasfilename(
            title="Save Annotations",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            with open(filepath, 'w') as f:
                f.write("PDF HIGHLIGHTER TEST - ANNOTATIONS\n")
                f.write("=" * 50 + "\n\n")
                
                for i, ann in enumerate(self.annotations, 1):
                    f.write(f"Annotation {i}:\n")
                    f.write(f"  Type: {ann['type']}\n")
                    f.write(f"  Page: {ann['page'] + 1}\n")
                    
                    if ann['type'] == 'highlight':
                        color_name = self.highlighter_colors[ann['color']]['name']
                        f.write(f"  Color: {color_name}\n")
                        if 'note' in ann:
                            f.write(f"  Note: {ann['note']}\n")
                    elif ann['type'] == 'text':
                        f.write(f"  Text: {ann['text']}\n")
                    
                    f.write(f"  Timestamp: {ann['timestamp']}\n")
                    f.write("\n")
            
            messagebox.showinfo("Success", f"Saved {len(self.annotations)} annotations to {filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save annotations:\n{e}")
    
    def clear_all(self):
        """Clear all annotations"""
        if not self.annotations:
            return
        
        if messagebox.askyesno("Confirm", "Clear all annotations?"):
            self.annotations = []
            self.update_stats()
            self.display_page()
    
    def export_pdf(self):
        """Export PDF with all annotations embedded"""
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please load a PDF first")
            return
        
        if not self.annotations:
            response = messagebox.askyesno("No Annotations", 
                                          "There are no annotations. Export original PDF anyway?")
            if not response:
                return
        
        filepath = filedialog.asksaveasfilename(
            title="Export Annotated PDF",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            initialfile="annotated_output.pdf"
        )
        
        if not filepath:
            return
        
        try:
            # Create a copy of the PDF document
            output_doc = fitz.open()
            
            for page_num in range(len(self.pdf_document)):
                # Get the original page
                src_page = self.pdf_document[page_num]
                
                # Insert it into output document
                output_doc.insert_pdf(self.pdf_document, from_page=page_num, to_page=page_num)
                out_page = output_doc[page_num]
                
                # Get page dimensions
                page_rect = out_page.rect
                
                # Filter annotations for this page
                page_annotations = [ann for ann in self.annotations if ann.get('page') == page_num]
                
                for ann in page_annotations:
                    ann_type = ann.get('type')
                    
                    # HIGHLIGHT ANNOTATIONS
                    if ann_type == 'highlight' and 'points' in ann:
                        points = ann['points']
                        if len(points) >= 2:
                            color_key = ann.get('color', 'yellow')
                            rgb = self.highlighter_colors[color_key]['rgb']
                            # Normalize RGB to 0-1 range for PyMuPDF
                            color = (rgb[0]/255, rgb[1]/255, rgb[2]/255)
                            
                            # Draw highlight as ink annotation (freehand drawing)
                            # Convert points to PDF coordinates - each point needs to be a list/tuple
                            current_stroke = []
                            for point in points:
                                # Convert canvas coordinates to PDF coordinates
                                pdf_x = float(point[0] / self.zoom_level)
                                pdf_y = float(point[1] / self.zoom_level)
                                current_stroke.append((pdf_x, pdf_y))
                            
                            # Create ink annotation with list of strokes
                            if current_stroke and len(current_stroke) >= 2:
                                ink_list = [current_stroke]  # Wrap in list
                                annot = out_page.add_ink_annot(ink_list)
                                annot.set_colors(stroke=color)
                                annot.set_border(width=15)  # Thick highlighter stroke
                                annot.set_opacity(0.4)  # Semi-transparent
                                annot.update()
                                
                                # Add note if it exists (for orange errors)
                                if 'note' in ann:
                                    annot.set_info(content=ann['note'])
                                    annot.update()
                    
                    # PEN ANNOTATIONS
                    elif ann_type == 'pen' and 'points' in ann:
                        points = ann['points']
                        if len(points) >= 2:
                            current_stroke = []
                            for point in points:
                                pdf_x = float(point[0] / self.zoom_level)
                                pdf_y = float(point[1] / self.zoom_level)
                                current_stroke.append((pdf_x, pdf_y))
                            
                            if current_stroke and len(current_stroke) >= 2:
                                ink_list = [current_stroke]  # Wrap in list
                                annot = out_page.add_ink_annot(ink_list)
                                annot.set_colors(stroke=(1, 0, 0))  # Red
                                annot.set_border(width=3)
                                annot.update()
                    
                    # TEXT ANNOTATIONS
                    elif ann_type == 'text' and 'pos' in ann:
                        pos = ann['pos']
                        text = ann.get('text', '')
                        if text:
                            # Convert position to PDF coordinates
                            pdf_x = pos[0] / self.zoom_level
                            pdf_y = pos[1] / self.zoom_level
                            
                            # Create a small rectangle for the text annotation
                            rect = fitz.Rect(pdf_x, pdf_y, pdf_x + 200, pdf_y + 30)
                            
                            # Add free text annotation
                            annot = out_page.add_freetext_annot(
                                rect,
                                text,
                                fontsize=12,
                                text_color=(1, 0, 0),
                                fill_color=(1, 1, 0.78)  # Light yellow background
                            )
                            annot.update()
            
            # Save the output PDF
            output_doc.save(filepath)
            output_doc.close()
            
            messagebox.showinfo("Success", 
                              f"Exported annotated PDF with {len(self.annotations)} annotations!\n\n"
                              f"Saved to: {filepath}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF:\n{e}")
            import traceback
            traceback.print_exc()

def main():
    root = tk.Tk()
    app = HighlighterTestApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
