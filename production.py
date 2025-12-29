"""
Complete Integration - Replace your entire production.py with this
Integrates handover system + visual navigation to annotations
"""

import tkinter as tk
from tkinter import messagebox, simpledialog, Menu
from PIL import Image, ImageTk, ImageDraw
import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from datetime import datetime
import os
import sys
import json
import getpass
import re
from handover_database import HandoverDB


def get_app_base_dir():
    """Returns the directory where the app is running from"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class CircuitInspector:
    def __init__(self, root):
        self.root = root
        self.root.title("Production Tool")
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
        
        # Initialize handover database
        self.handover_db = HandoverDB(os.path.join(base, "handover_db.json"))

        self.excel_file = None
        self.working_excel_path = None
        self.checklist_file = self.excel_file
        self.zoom_level = 1.0
        self.current_sr_no = 1
        self.current_page_image = None
        self.session_refs = set()
        
        # Visual navigation for production mode
        self.production_arrow_id = None
        self.production_highlight_id = None
        self.production_dialog_open = False

        # Column mapping
        self.punch_sheet_name = 'Punch Sheet'
        self.punch_cols = {
            'sr_no': 'A',
            'ref_no': 'B',
            'desc': 'C',
            'category': 'D',
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

        # Drawing state
        self.drawing = False
        self.drawing_type = None
        self.rect_start_x = None
        self.rect_start_y = None
        self.temp_rect_id = None
        self.selected_annotation = None

        self.setup_ui()
        self.current_sr_no = self.get_next_sr_no()

    # ================================================================
    # CELL HELPERS
    # ================================================================
    def split_cell(self, cell_ref):
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

    # ================================================================
    # UI SETUP
    # ================================================================
    def setup_ui(self):
        toolbar = tk.Frame(self.root, bg='#2c3e50', height=60)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        
        # File Menu
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open PDF", command=self.load_pdf)
        file_menu.add_command(label="Load from Handover", command=self.load_from_handover_queue)
        file_menu.add_separator()
        file_menu.add_command(label="Open Excel", command=self.open_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Tools Menu
        tools_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="🏭 Production Mode", command=self.production_mode)
        tools_menu.add_separator()
        tools_menu.add_command(label="✅ Complete Rework & Handback", command=self.complete_rework_handback)

        btn_style = {'bg': '#3498db', 'fg': 'white', 'padx': 15, 'pady': 8, 'font': ('Arial', 10)}

        tk.Button(toolbar, text="📁 Load PDF", command=self.load_pdf, **btn_style).pack(side=tk.LEFT, padx=5, pady=10)
        
        tk.Button(toolbar, text="📦 Load from Queue", command=self.load_from_handover_queue,
                 bg='#9b59b6', fg='white', padx=15, pady=8, font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5, pady=10)

        self.page_label = tk.Label(toolbar, text="Page: 0/0", bg='#2c3e50', fg='white', font=('Arial', 10))
        self.page_label.pack(side=tk.LEFT, padx=5)
        
        tk.Button(toolbar, text="◀ Prev", command=self.prev_page, bg='#95a5a6', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Next ▶", command=self.next_page, bg='#95a5a6', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=5)

        tk.Button(toolbar, text="🔍+", command=self.zoom_in, bg='#27ae60', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=(20, 2))
        tk.Button(toolbar, text="🔍-", command=self.zoom_out, bg='#27ae60', fg='white', padx=10, pady=8).pack(side=tk.LEFT, padx=2)

        tk.Button(toolbar, text="🏭 Production Mode", command=self.production_mode,
                 bg='#e67e22', fg='white', padx=15, pady=8, font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)
        
        tk.Button(toolbar, text="✅ Handback to Quality", command=self.complete_rework_handback,
                 bg='#16a085', fg='white', padx=15, pady=8, font=('Arial', 10, 'bold')).pack(side=tk.LEFT, padx=5)

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

        self.canvas.bind("<ButtonPress-1>", self.on_left_press)
        self.canvas.bind("<ButtonPress-3>", self.on_right_press)
        self.canvas.bind("<Double-Button-1>", self.on_double_left_zoom)
        self.canvas.bind("<Double-Button-3>", self.on_double_right_zoom)

    # ================================================================
    # LOAD FROM HANDOVER QUEUE
    # ================================================================
    def load_from_handover_queue(self):
        """Load item from production handover queue"""
        pending_items = self.handover_db.get_pending_production_items()
        
        if not pending_items:
            messagebox.showinfo("No Items", 
                              "✓ No items in production queue.\n"
                              "All items have been processed!",
                              icon='info')
            return
        
        # Create selection dialog
        dlg = tk.Toplevel(self.root)
        dlg.title("Production Queue - Select Item")
        dlg.geometry("900x500")
        dlg.transient(self.root)
        dlg.grab_set()
        
        # Header
        header = tk.Frame(dlg, bg='#9b59b6', height=50)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(header, text="📦 Production Queue", bg='#9b59b6', fg='white',
                font=('Arial', 14, 'bold')).pack(pady=12)
        
        # Listbox
        list_frame = tk.Frame(dlg)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        tk.Label(list_frame, text="Select item to load:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=(0, 8))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, font=('Consolas', 9),
                            yscrollcommand=scrollbar.set,
                            selectmode=tk.SINGLE, height=15)
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Populate
        for item in pending_items:
            status_icon = "⚙️" if item['status'] == 'in_progress' else "📦"
            display = (
                f"{status_icon} {item['cabinet_id']:20} | {item['project_name']:30} | "
                f"Open Punches: {item['open_punches']:3} | By: {item['handed_over_by']:15}"
            )
            listbox.insert(tk.END, display)
        
        def load_selected():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select an item first.")
                return
            
            item = pending_items[selection[0]]
            dlg.destroy()
            self.load_handover_item(item)
        
        # Buttons
        btn_frame = tk.Frame(dlg)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        tk.Button(btn_frame, text="📂 Load Selected", command=load_selected,
                 bg='#3498db', fg='white', font=('Arial', 10, 'bold'),
                 padx=20, pady=10).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy,
                 padx=20, pady=10).pack(side=tk.RIGHT, padx=5)
        
        listbox.bind('<Double-Button-1>', lambda e: load_selected())

    def load_handover_item(self, item):
        """Load a specific handover item"""
        try:
            # Verify files exist
            if not os.path.exists(item['pdf_path']):
                messagebox.showerror("File Not Found", 
                                   f"PDF not found:\n{item['pdf_path']}")
                return
            
            if not os.path.exists(item['excel_path']):
                messagebox.showerror("File Not Found", 
                                   f"Excel not found:\n{item['excel_path']}")
                return
            
            # Load PDF
            self.pdf_document = fitz.open(item['pdf_path'])
            self.current_pdf_path = item['pdf_path']
            self.current_page = 0
            self.zoom_level = 1.0
            
            # Set project details
            self.cabinet_id = item['cabinet_id']
            self.project_name = item['project_name']
            self.sales_order_no = item['sales_order_no']
            
            # Prepare folders
            self.prepare_project_folders()
            
            # Set Excel
            self.excel_file = item['excel_path']
            self.working_excel_path = item['excel_path']
            
            # Load session if available
            if item.get('session_path') and os.path.exists(item['session_path']):
                self.load_session_from_path(item['session_path'])
            else:
                self.annotations = []
                self.session_refs.clear()
            
            # Mark as in progress
            try:
                username = os.getlogin()
            except:
                username = getpass.getuser()
            
            self.handover_db.update_production_status(
                item['cabinet_id'],
                status='in_progress',
                user=username
            )
            
            self.display_page()
            
            messagebox.showinfo(
                "Item Loaded",
                f"✓ Loaded from Production Queue:\n\n"
                f"Cabinet: {self.cabinet_id}\n"
                f"Project: {self.project_name}\n"
                f"Open Punches: {item['open_punches']}\n\n"
                f"Use '🏭 Production Mode' to review and complete punches.",
                icon='info'
            )
            
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load item:\n{e}")

    # ================================================================
    # COMPLETE REWORK & HANDBACK
    # ================================================================
    def complete_rework_handback(self):
        """Complete rework and handback to Quality"""
        if not self.pdf_document or not self.excel_file:
            messagebox.showwarning("No Item Loaded", 
                                 "Please load an item from the production queue first.")
            return
        
        # Check if item is from handover
        item = self.handover_db.get_item_by_cabinet_id(self.cabinet_id, "quality_to_production")
        if not item:
            messagebox.showwarning("Not from Queue", 
                                 "This item was not loaded from the production queue.\n"
                                 "Only items from the handover queue can be handed back.")
            return
        
        # Count open punches
        open_count = self.count_open_punches()
        
        if open_count > 0:
            proceed = messagebox.askyesno(
                "Open Punches Remaining",
                f"⚠️ There are still {open_count} open punch(es).\n\n"
                "Are you sure you want to handback to Quality?\n\n"
                "It's recommended to complete all rework first.",
                icon='warning'
            )
            if not proceed:
                return
        
        # Get remarks
        remarks = simpledialog.askstring(
            "Production Remarks",
            "Enter any remarks or notes for Quality team:\n\n"
            "(Optional - Leave blank if none)",
            parent=self.root
        )
        
        try:
            username = os.getlogin()
        except:
            username = getpass.getuser()
        
        handback_data = {
            "cabinet_id": self.cabinet_id,
            "project_name": self.project_name,
            "sales_order_no": self.sales_order_no,
            "pdf_path": self.current_pdf_path,
            "excel_path": self.excel_file,
            "session_path": self.get_session_path_for_pdf(),
            "rework_completed_by": username,
            "rework_completed_date": datetime.now().isoformat(),
            "production_remarks": remarks or "No remarks"
        }
        
        success = self.handover_db.add_production_handback(handback_data)
        
        if success:
            messagebox.showinfo(
                "Handback Complete",
                f"✓ Successfully handed back to Quality:\n\n"
                f"Cabinet: {self.cabinet_id}\n"
                f"Project: {self.project_name}\n\n"
                f"Quality team will verify and close this item.",
                icon='info'
            )
            
            # Clear current work
            self.pdf_document = None
            self.current_pdf_path = None
            self.excel_file = None
            self.annotations = []
            self.canvas.delete("all")
            self.page_label.config(text="Page: 0/0")
            self.root.title("Production Tool")
        else:
            messagebox.showerror("Error", "Failed to handback item to Quality.")

    def count_open_punches(self):
        """Count open punches in current Excel"""
        try:
            if not self.excel_file or not os.path.exists(self.excel_file):
                return 0
            
            wb = load_workbook(self.excel_file, data_only=True)
            ws = wb[self.punch_sheet_name] if self.punch_sheet_name in wb.sheetnames else wb.active
            
            open_count = 0
            row = 8
            
            while row <= ws.max_row + 5:
                sr = self.read_cell(ws, row, self.punch_cols['sr_no'])
                if sr is None:
                    break
                
                closed = self.read_cell(ws, row, self.punch_cols['closed_name'])
                if not closed:
                    open_count += 1
                
                row += 1
            
            wb.close()
            return open_count
            
        except Exception as e:
            print(f"Error counting open punches: {e}")
            return 0

    # ================================================================
    # ENHANCED PRODUCTION MODE WITH VISUAL NAVIGATION
    # ================================================================
    def production_mode(self):
        """Enhanced production mode with visual navigation to annotations"""
        if not self.pdf_document or not self.excel_file:
            messagebox.showwarning("No Item", "Please load a PDF and Excel file first.")
            return
        
        punches = self.read_open_punches_from_excel()

        if not punches:
            messagebox.showinfo("No Punches", "✓ All punches are closed!")
            return

        punches.sort(key=lambda p: (p['implemented'], p['sr_no']))

        dlg = tk.Toplevel(self.root)
        dlg.title("Production Mode")
        dlg.geometry("800x500")
        dlg.transient(self.root)
        
        self.production_dialog_open = True

        # Header
        header = tk.Frame(dlg, bg='#e67e22', height=50)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(header, text="🏭 PRODUCTION MODE", bg='#e67e22', fg='white',
                font=('Arial', 14, 'bold')).pack(pady=12)

        idx_label = tk.Label(dlg, text="", font=('Arial', 11, 'bold'))
        idx_label.pack(pady=(10, 0))
        
        # Info frame
        info_frame = tk.Frame(dlg, bg='#ecf0f1')
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        sr_label = tk.Label(info_frame, text="", font=('Arial', 10), bg='#ecf0f1')
        sr_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        ref_label = tk.Label(info_frame, text="", font=('Arial', 10), bg='#ecf0f1')
        ref_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        impl_label = tk.Label(info_frame, text="", font=('Arial', 10), bg='#ecf0f1')
        impl_label.pack(side=tk.LEFT, padx=10, pady=5)

        text_widget = tk.Text(dlg, wrap=tk.WORD, height=14, font=('Arial', 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
        text_widget.config(state=tk.DISABLED)

        pos = [0]

        def show_item():
            p = punches[pos[0]]

            idx_label.config(text=f"Item {pos[0]+1} of {len(punches)}")
            sr_label.config(text=f"SR No: {p['sr_no']}")
            ref_label.config(text=f"Ref: {p['ref_no']}")
            impl_status = "✓ Implemented" if p['implemented'] else "⚠ Not Implemented"
            impl_color = '#27ae60' if p['implemented'] else '#e74c3c'
            impl_label.config(text=impl_status, fg=impl_color)

            text_widget.config(state=tk.NORMAL)
            text_widget.delete("1.0", tk.END)
            text_widget.insert(tk.END, p['punch_text'])
            text_widget.insert(tk.END, f"\n\n───────────────────\nCategory: {p['category']}\n")

            ann = next((a for a in self.annotations if a.get('excel_row') == p['row']), None)
            if ann and ann.get('implementation_remark'):
                text_widget.insert(tk.END, "\n───────────────────\nPrevious Remarks:\n")
                text_widget.insert(tk.END, ann['implementation_remark'])

            text_widget.config(state=tk.DISABLED)
            
            # Navigate to annotation on PDF
            self.navigate_to_punch(p['sr_no'], p['punch_text'])

        show_item()

        def mark_implemented():
            p = punches[pos[0]]

            try:
                default_user = os.getlogin()
            except:
                default_user = getpass.getuser()

            name = simpledialog.askstring("Implemented By", "Enter your name:",
                                         initialvalue=default_user, parent=dlg)
            if not name:
                return

            remark = simpledialog.askstring("Remarks (optional)", 
                                           "Add remarks (optional):", parent=dlg)

            try:
                wb = load_workbook(self.excel_file)
                ws = wb[self.punch_sheet_name]

                self.write_cell(ws, p['row'], self.punch_cols['implemented_name'], name)
                self.write_cell(ws, p['row'], self.punch_cols['implemented_date'],
                              datetime.now().strftime("%Y-%m-%d"))

                wb.save(self.excel_file)
                wb.close()

            except Exception as e:
                messagebox.showerror("Excel Error", str(e))
                return

            ann = next((a for a in self.annotations if a.get('sr_no') == p['sr_no']), None)
            if ann:
                ann['implemented'] = True
                ann['implemented_name'] = name
                ann['implemented_date'] = datetime.now().isoformat()
                ann['implementation_remark'] = remark

            if pos[0] < len(punches) - 1:
                pos[0] += 1
                show_item()
            else:
                messagebox.showinfo("Complete", "✓ All punches reviewed!")
                self.clear_production_visuals()
                self.production_dialog_open = False
                dlg.destroy()

        def next_item():
            if pos[0] < len(punches) - 1:
                pos[0] += 1
                show_item()

        def prev_item():
            if pos[0] > 0:
                pos[0] -= 1
                show_item()
        
        def on_close():
            self.clear_production_visuals()
            self.production_dialog_open = False
            dlg.destroy()

        dlg.protocol("WM_DELETE_WINDOW", on_close)

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(fill=tk.X, pady=10)

        tk.Button(btn_frame, text="◀ Prev", command=prev_item, width=10,
                 padx=10, pady=8).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="✓ DONE", command=mark_implemented, 
                 bg="#27ae60", fg="white", width=14, font=('Arial', 10, 'bold'),
                 padx=10, pady=8).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Next ▶", command=next_item, width=10,
                 padx=10, pady=8).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Close", command=on_close,
                 padx=10, pady=8).pack(side=tk.RIGHT, padx=6)

    def navigate_to_punch(self, sr_no, punch_text):
        """Navigate to annotation and highlight it visually"""
        # Find matching annotation
        target_ann = None
        
        # First try SR No match
        for ann in self.annotations:
            if ann.get('sr_no') == sr_no and ann.get('type') == 'error':
                target_ann = ann
                break
        
        # If not found, try fuzzy text match
        if not target_ann:
            best_match = None
            best_score = 0
            for ann in self.annotations:
                if ann.get('type') == 'error' and ann.get('punch_text'):
                    # Simple text similarity
                    ann_text = str(ann['punch_text']).lower()
                    search_text = str(punch_text).lower()
                    if search_text in ann_text or ann_text in search_text:
                        score = len(set(search_text.split()) & set(ann_text.split()))
                        if score > best_score:
                            best_score = score
                            best_match = ann
            if best_match:
                target_ann = best_match
        
        # Clear previous visuals
        self.clear_production_visuals()
        
        if target_ann:
            # Navigate to page
            if target_ann.get('page') is not None:
                self.current_page = target_ann['page']
                self.display_page()
            
            # Draw arrow and highlight
            if 'bbox_page' in target_ann:
                self.highlight_annotation(target_ann)

    def highlight_annotation(self, annotation):
        """Draw visual indicators for the annotation"""
        bbox_display = self.bbox_page_to_display(annotation['bbox_page'])
        x1, y1, x2, y2 = bbox_display
        
        # Calculate center
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2
        
        # Draw pulsing highlight
        padding = 20
        self.production_highlight_id = self.canvas.create_rectangle(
            x1 - padding, y1 - padding,
            x2 + padding, y2 + padding,
            outline='#e74c3c', width=5, dash=(10, 5)
        )
        
        # Draw arrow pointing to it
        arrow_start_x = cx - 100
        arrow_start_y = cy - 100
        self.production_arrow_id = self.canvas.create_line(
            arrow_start_x, arrow_start_y, cx - 15, cy - 15,
            arrow=tk.LAST, fill='#e74c3c', width=4
        )
        
        # Add text label
        self.canvas.create_text(
            arrow_start_x - 10, arrow_start_y - 10,
            text=f"SR {annotation.get('sr_no', '?')}",
            fill='#e74c3c', font=('Arial', 12, 'bold'),
            anchor='se'
        )
        
        # Scroll to make it visible
        self.canvas.yview_moveto(max(0, (y1 - 100) / max(1, self.canvas.bbox("all")[3])))
        self.canvas.xview_moveto(max(0, (x1 - 100) / max(1, self.canvas.bbox("all")[2])))

    def clear_production_visuals(self):
        """Clear production mode visual indicators"""
        if self.production_arrow_id:
            try:
                self.canvas.delete(self.production_arrow_id)
            except:
                pass
            self.production_arrow_id = None
        
        if self.production_highlight_id:
            try:
                self.canvas.delete(self.production_highlight_id)
            except:
                pass
            self.production_highlight_id = None

    def read_open_punches_from_excel(self):
        """Read open punches from Excel"""
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
    # PDF HELPERS
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
                self.current_sr_no = self.get_next_sr_no()
                self.display_page()
                messagebox.showinfo("Success", f"Loaded PDF with {len(self.pdf_document)} pages")
                self.ask_project_details()
                self.prepare_project_folders()
                
                try:
                    self.working_excel_path = os.path.join(
                        self.project_dirs["working_excel"],
                        f"{self.cabinet_id.replace(' ', '_')}_Working.xlsx"
                    )

                    if os.path.exists(self.working_excel_path):
                        resume = messagebox.askyesno(
                            "Resume Inspection",
                            f"Existing working Excel found:\n\n{os.path.basename(self.working_excel_path)}\n\n"
                            "Resume previous inspection?"
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
                
                session_path = self.get_session_path_for_pdf()
                if session_path:
                    resume = messagebox.askyesno(
                        "Resume Session",
                        f"Existing session found for this drawing:\n\n{os.path.basename(session_path)}\n\n"
                        "Do you want to resume it?"
                    )
                    if resume:
                        self.load_session_from_path(session_path)

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

    def page_to_display_scale(self):
        return 2.0 * self.zoom_level

    def bbox_page_to_display(self, bbox_page):
        scale = self.page_to_display_scale()
        x1, y1, x2, y2 = bbox_page
        return (x1 * scale, y1 * scale, x2 * scale, y2 * scale)

    def bbox_display_to_page(self, bbox_display):
        scale = self.page_to_display_scale()
        x1, y1, x2, y2 = bbox_display
        return (x1 / scale, y1 / scale, x2 / scale, y2 / scale)

    def display_page(self):
        if not self.pdf_document:
            self.canvas.delete("all")
            self.page_label.config(text="Page: 0/0")
            return
        
        try:
            page = self.pdf_document[self.current_page]
            mat = fitz.Matrix(self.page_to_display_scale(), self.page_to_display_scale())
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.current_page_image = None
            draw = ImageDraw.Draw(img, 'RGBA')
            
            for ann in self.annotations:
                if ann.get('page') != self.current_page or 'bbox_page' not in ann:
                    continue
                x1d, y1d, x2d, y2d = self.bbox_page_to_display(ann['bbox_page'])
                is_selected = (self.selected_annotation is ann)
                w = int(5 * self.zoom_level) if is_selected else int(3 * self.zoom_level)
                
                if ann.get('type') == 'ok':
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
            
            self.photo = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
            self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
            self.page_label.config(text=f"Page: {self.current_page + 1}/{len(self.pdf_document)}")
            
            # Restore production visuals if dialog is open
            if self.production_dialog_open and hasattr(self, '_last_highlighted_ann'):
                self.highlight_annotation(self._last_highlighted_ann)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to display page: {e}")

    def on_left_press(self, event):
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please load a PDF first")
            return
        self.drawing = True
        self.drawing_type = 'ok'
        self.rect_start_x = self.canvas.canvasx(event.x)
        self.rect_start_y = self.canvas.canvasy(event.y)

    def on_right_press(self, event):
        if not self.pdf_document:
            messagebox.showwarning("Warning", "Please load a PDF first")
            return
        self.drawing = True
        self.drawing_type = 'error'
        self.rect_start_x = self.canvas.canvasx(event.x)
        self.rect_start_y = self.canvas.canvasy(event.y)

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

    def prev_page(self):
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.display_page()

    def next_page(self):
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.display_page()

    def zoom_in(self):
        if self.zoom_level < 3.0:
            self.zoom_level += 0.25
            self.display_page()

    def zoom_out(self):
        if self.zoom_level > 0.5:
            self.zoom_level -= 0.25
            self.display_page()

    # ================================================================
    # PROJECT MANAGEMENT
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
            messagebox.showerror("Excel Locked", 
                               "Please close the Excel file before entering project details.")
        except Exception as e:
            messagebox.showerror("Excel Error", 
                               f"Failed to write project details:\n{e}")

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

    def open_excel(self):
        if not self.excel_file or not os.path.exists(self.excel_file):
            messagebox.showwarning("No Excel", "No Excel file loaded.")
            return
        try:
            if os.name == 'nt':
                os.startfile(self.excel_file)
            else:
                import subprocess
                import shlex
                if sys.platform == 'darwin':
                    cmd = f"open {shlex.quote(self.excel_file)}"
                else:
                    cmd = f"xdg-open {shlex.quote(self.excel_file)}"
                subprocess.Popen(cmd, shell=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel: {e}")

    def get_session_path_for_pdf(self):
        if not self.current_pdf_path or not hasattr(self, 'project_dirs'):
            return None

        session_path = os.path.join(
            self.project_dirs.get("sessions", ""),
            f"{self.cabinet_id}_annotations.json"
        )

        return session_path if os.path.exists(session_path) else None

    def load_session_from_path(self, path):
        """Load annotation session from JSON file"""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Session Load Error", f"Failed to load session:\n{e}")
            return

        self.project_name = data.get('project_name', self.project_name)
        self.sales_order_no = data.get('sales_order_no', self.sales_order_no)
        self.cabinet_id = data.get('cabinet_id', getattr(self, "cabinet_id", ""))
        self.current_page = data.get('current_page', 0)
        self.zoom_level = data.get('zoom_level', 1.0)
        self.current_sr_no = data.get('current_sr_no', self.current_sr_no)

        self.annotations = []
        self.session_refs.clear()

        for entry in data.get('annotations', []):
            ann = entry.copy()
            if 'bbox_page' in ann:
                ann['bbox_page'] = tuple(float(x) for x in ann['bbox_page'])
            self.annotations.append(ann)

            if ann.get('ref_no'):
                self.session_refs.add(str(ann['ref_no']).strip())

        self.display_page()
        messagebox.showinfo("Session Loaded", 
                          f"Loaded {len(self.annotations)} annotations.")


def main():
    root = tk.Tk()
    app = CircuitInspector(root)
    root.mainloop()


if __name__ == "__main__":
    main()
