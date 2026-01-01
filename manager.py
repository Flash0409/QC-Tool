"""
Manager UI - Complete Dashboard, Analytics, and Category Management System
UPDATED: Financial year support, date displays, integrated search with filters
Run this as: python manager_ui.py
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from PIL import Image, ImageTk
import json
import os
import sys
from datetime import datetime, timedelta
from collections import defaultdict
import sqlite3
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import calendar


def get_app_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_financial_year():
    """Get current financial year in format 2026-27 (starts October 1st)"""
    today = datetime.now()
    if today.month >= 10:  # October onwards
        return f"{today.year}-{str(today.year + 1)[-2:]}"
    else:
        return f"{today.year - 1}-{str(today.year)[-2:]}"


def get_week_number():
    """Get current week number"""
    today = datetime.now()
    return today.isocalendar()[1]


class ManagerDatabase:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS cabinets (
            cabinet_id TEXT PRIMARY KEY,
            project_name TEXT,
            sales_order_no TEXT,
            total_pages INTEGER DEFAULT 0,
            annotated_pages INTEGER DEFAULT 0,
            total_punches INTEGER DEFAULT 0,
            open_punches INTEGER DEFAULT 0,
            implemented_punches INTEGER DEFAULT 0,
            closed_punches INTEGER DEFAULT 0,
            status TEXT DEFAULT 'quality_inspection',
            created_date TEXT,
            last_updated TEXT)''')
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS category_occurrences (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cabinet_id TEXT,
            project_name TEXT,
            category TEXT,
            subcategory TEXT,
            occurrence_date TEXT)''')
        
        conn.commit()
        conn.close()
    
    def get_all_projects(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''SELECT project_name, COUNT(DISTINCT cabinet_id) as count,
                          MAX(last_updated) as updated
                          FROM cabinets
                          GROUP BY project_name
                          ORDER BY updated DESC''')
        projects = [{'project_name': r[0], 'cabinet_count': r[1], 'last_updated': r[2]} 
                   for r in cursor.fetchall()]
        conn.close()
        return projects
    
    def get_cabinets_by_project(self, project_name):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''SELECT cabinet_id, project_name, total_pages, annotated_pages,
                          total_punches, open_punches, implemented_punches, closed_punches, status
                          FROM cabinets
                          WHERE project_name = ?
                          ORDER BY last_updated DESC''', (project_name,))
        cols = ['cabinet_id', 'project_name', 'total_pages', 'annotated_pages',
                'total_punches', 'open_punches', 'implemented_punches', 'closed_punches', 'status']
        cabinets = [dict(zip(cols, row)) for row in cursor.fetchall()]
        conn.close()
        return cabinets
    
    def search_projects(self, search_term):
        """Search projects by name"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''SELECT project_name, COUNT(DISTINCT cabinet_id) as count,
                          MAX(last_updated) as updated
                          FROM cabinets
                          WHERE project_name LIKE ?
                          GROUP BY project_name
                          ORDER BY updated DESC''', (f'%{search_term}%',))
        projects = [{'project_name': r[0], 'cabinet_count': r[1], 'last_updated': r[2]} 
                   for r in cursor.fetchall()]
        conn.close()
        return projects
    
    def get_category_stats(self, start_date=None, end_date=None, project_name=None):
        """Get category stats with flexible date filtering"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        query = 'SELECT category, subcategory, COUNT(*) as count FROM category_occurrences WHERE 1=1'
        params = []
        
        if start_date:
            query += ' AND occurrence_date >= ?'
            params.append(start_date)
        
        if end_date:
            query += ' AND occurrence_date <= ?'
            params.append(end_date)
        
        if project_name:
            query += ' AND project_name = ?'
            params.append(project_name)
        
        query += ' GROUP BY category, subcategory ORDER BY count DESC'
        cursor.execute(query, params)
        stats = [{'category': r[0], 'subcategory': r[1], 'count': r[2]} 
                for r in cursor.fetchall()]
        conn.close()
        return stats
    
    def get_all_project_names(self):
        """Get list of all unique project names"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT DISTINCT project_name FROM cabinets ORDER BY project_name')
        projects = [row[0] for row in cursor.fetchall()]
        conn.close()
        return projects
    
    def get_cabinet_statistics(self):
        """Get cabinet counts for different periods with proper financial year"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        today = datetime.now().date()
        week_start = today - timedelta(days=today.weekday())
        month_start = today.replace(day=1)
        
        # Financial year starts on October 1st
        current_year = today.year
        if today.month >= 10:
            fy_start = datetime(current_year, 10, 1).date()
        else:
            fy_start = datetime(current_year - 1, 10, 1).date()
        
        stats = {}
        
        # Daily count
        cursor.execute('''SELECT COUNT(DISTINCT cabinet_id) FROM cabinets 
                         WHERE DATE(created_date) = ?''', (today.isoformat(),))
        stats['daily'] = cursor.fetchone()[0]
        
        # Weekly count
        cursor.execute('''SELECT COUNT(DISTINCT cabinet_id) FROM cabinets 
                         WHERE DATE(created_date) >= ?''', (week_start.isoformat(),))
        stats['weekly'] = cursor.fetchone()[0]
        
        # Monthly count
        cursor.execute('''SELECT COUNT(DISTINCT cabinet_id) FROM cabinets 
                         WHERE DATE(created_date) >= ?''', (month_start.isoformat(),))
        stats['monthly'] = cursor.fetchone()[0]
        
        # Financial Yearly count
        cursor.execute('''SELECT COUNT(DISTINCT cabinet_id) FROM cabinets 
                         WHERE DATE(created_date) >= ?''', (fy_start.isoformat(),))
        stats['yearly'] = cursor.fetchone()[0]
        
        conn.close()
        return stats


class ManagerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Manager Dashboard")
        self.root.geometry("1600x900")
        
        base_dir = get_app_base_dir()
        self.db = ManagerDatabase(os.path.join(base_dir, "manager.db"))
        self.category_file = os.path.join(os.path.dirname(base_dir), "assets", "categories.json")
        self.categories = self.load_categories()  # This will return the loaded list
        
        self.setup_ui()
        self.show_dashboard()
    
    def load_categories(self):
        try:
            if os.path.exists(self.category_file):
                with open(self.category_file, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    # Don't modify existing categories - keep them as-is
                    return loaded
        except Exception:
            pass
        return []
    
    def save_categories(self):
        try:
            os.makedirs(os.path.dirname(self.category_file), exist_ok=True)
            with open(self.category_file, "w", encoding="utf-8") as f:
                json.dump(self.categories, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save:\n{e}")
    
    def setup_ui(self):
        # Navigation
        nav = tk.Frame(self.root, bg='#1e293b', height=70)
        nav.pack(side=tk.TOP, fill=tk.X)
        nav.pack_propagate(False)
        
        tk.Label(nav, text="📊 Manager Dashboard", bg='#1e293b', fg='white',
                font=('Segoe UI', 18, 'bold')).pack(side=tk.LEFT, padx=30, pady=15)
        
        btn_style = {'font': ('Segoe UI', 11, 'bold'), 'relief': tk.FLAT,
                    'cursor': 'hand2', 'padx': 25, 'pady': 12}
        
        self.nav_btns = {}
        self.nav_btns['dashboard'] = tk.Button(nav, text="🏠 Dashboard",
                                               command=self.show_dashboard,
                                               bg='#3b82f6', fg='white', **btn_style)
        self.nav_btns['dashboard'].pack(side=tk.LEFT, padx=5)
        
        self.nav_btns['analytics'] = tk.Button(nav, text="📈 Analytics",
                                               command=self.show_analytics,
                                               bg='#334155', fg='white', **btn_style)
        self.nav_btns['analytics'].pack(side=tk.LEFT, padx=5)
        
        self.nav_btns['categories'] = tk.Button(nav, text="🏷️ Categories",
                                                command=self.show_categories,
                                                bg='#334155', fg='white', **btn_style)
        self.nav_btns['categories'].pack(side=tk.LEFT, padx=5)
        
        # Content frame
        self.content = tk.Frame(self.root, bg='#f8fafc')
        self.content.pack(fill=tk.BOTH, expand=True)
    
    def set_active_nav(self, key):
        for k, btn in self.nav_btns.items():
            btn.config(bg='#3b82f6' if k == key else '#334155')
    
    def clear_content(self):
        for w in self.content.winfo_children():
            w.destroy()
    
    # ============ DASHBOARD - WITH PROPER DATE DISPLAYS AND SEARCH ============
    def show_dashboard(self):
        self.set_active_nav('dashboard')
        self.clear_content()
        
        # Centered container with 70% width
        center_container = tk.Frame(self.content, bg='#f8fafc')
        center_container.place(relx=0.5, rely=0, anchor='n', relwidth=0.7, relheight=1.0)
        
        # Statistics Cards at the top
        stats_frame = tk.Frame(center_container, bg='#f8fafc')
        stats_frame.pack(fill=tk.X, padx=30, pady=(20, 10))
        
        stats = self.db.get_cabinet_statistics()
        today = datetime.now()
        
        # Create 4 stat cards with proper labels
        stat_cards = [
            (today.strftime("%B %d"), stats['daily'], "#3b82f6"),  # December 31
            (f"Week {get_week_number()}", stats['weekly'], "#8b5cf6"),  # Week 52
            (today.strftime("%B"), stats['monthly'], "#10b981"),  # December
            (f"FY {get_financial_year()}", stats['yearly'], "#f59e0b")  # FY 2024-25
        ]
        
        for label, count, color in stat_cards:
            card = tk.Frame(stats_frame, bg='white', relief=tk.SOLID, borderwidth=1)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            
            tk.Label(card, text=label, font=('Segoe UI', 11, 'bold'), 
                    bg='white', fg='#64748b').pack(pady=(15, 5))
            tk.Label(card, text=str(count), font=('Segoe UI', 28, 'bold'),
                    bg='white', fg=color).pack(pady=(0, 5))
            tk.Label(card, text="Cabinets", font=('Segoe UI', 9),
                    bg='white', fg='#94a3b8').pack(pady=(0, 15))
        
        projects = self.db.get_all_projects()
        
        if not projects:
            empty_container = tk.Frame(center_container, bg='#f8fafc')
            empty_container.pack(expand=True, fill=tk.BOTH)
            center_frame = tk.Frame(empty_container, bg='#f8fafc')
            center_frame.place(relx=0.5, rely=0.5, anchor='center')
            
            tk.Label(center_frame, text="No projects found", 
                    font=('Segoe UI', 16, 'bold'), fg='#1e293b', bg='#f8fafc').pack(pady=10)
            tk.Label(center_frame, text="Projects will appear here once Quality Inspection tool syncs data.",
                    font=('Segoe UI', 11), fg='#64748b', bg='#f8fafc').pack(pady=5)
            return
        
        # Scrollable container - CREATE THIS FIRST
        canvas_container = tk.Frame(center_container, bg='#f8fafc')
        canvas_container.pack(expand=True, fill=tk.BOTH, padx=30, pady=(0, 20))
        
        canvas = tk.Canvas(canvas_container, bg='#f8fafc', highlightthickness=0)
        scrollbar = tk.Scrollbar(canvas_container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg='#f8fafc')
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        scroll_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Header with search bar - NOW scroll_frame is defined
        header = tk.Frame(center_container, bg='#f8fafc')
        header.pack(fill=tk.X, padx=30, pady=(10, 10))
        header.pack_forget()  # Remove from pack order
        header.pack(fill=tk.X, padx=30, pady=(10, 10), before=canvas_container)  # Pack before canvas
        
        tk.Label(header, text="Projects Overview", font=('Segoe UI', 16, 'bold'),
                bg='#f8fafc').pack(side=tk.LEFT)
        
        # Search bar on the right
        search_frame = tk.Frame(header, bg='white', relief=tk.SOLID, borderwidth=1)
        search_frame.pack(side=tk.RIGHT, padx=10)
        
        tk.Label(search_frame, text="🔍", bg='white', font=('Segoe UI', 12)).pack(side=tk.LEFT, padx=(10, 5))
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30,
                               font=('Segoe UI', 10), relief=tk.FLAT, bg='white')
        search_entry.pack(side=tk.LEFT, padx=(0, 10), pady=8)
        
        def on_search(*args):
            search_term = search_var.get().strip()
            if search_term and search_term != "Search projects...":
                filtered_projects = self.db.search_projects(search_term)
            else:
                filtered_projects = self.db.get_all_projects()
            self.update_project_list(scroll_frame, filtered_projects)
        
        search_var.trace('w', on_search)
        search_entry.insert(0, "Search projects...")
        search_entry.config(fg='#94a3b8')
        
        def on_focus_in(event):
            if search_entry.get() == "Search projects...":
                search_entry.delete(0, tk.END)
                search_entry.config(fg='#1e293b')
        
        def on_focus_out(event):
            if not search_entry.get():
                search_entry.insert(0, "Search projects...")
                search_entry.config(fg='#94a3b8')
        
        search_entry.bind("<FocusIn>", on_focus_in)
        search_entry.bind("<FocusOut>", on_focus_out)
        
        # Store references for search updates
        self.dashboard_scroll_frame = scroll_frame
        self.update_project_list(scroll_frame, projects)
    
    def update_project_list(self, scroll_frame, projects):
        """Update the project list in the dashboard"""
        for w in scroll_frame.winfo_children():
            w.destroy()
        
        if not projects:
            tk.Label(scroll_frame, text="No matching projects found",
                    font=('Segoe UI', 12), fg='#64748b', bg='#f8fafc').pack(pady=50)
            return
        
        for proj in projects:
            self.create_project_card(scroll_frame, proj)
    
    def create_project_card(self, parent, project):
        card = tk.Frame(parent, bg='white', relief=tk.SOLID, borderwidth=1)
        card.pack(fill=tk.X, pady=10, padx=5)
        
        header = tk.Frame(card, bg='#eff6ff', cursor='hand2')
        header.pack(fill=tk.X)
        
        expand_var = tk.BooleanVar(value=False)
        
        indicator = tk.Label(header, text="▶", font=('Segoe UI', 12, 'bold'),
                           bg='#eff6ff', fg='#3b82f6', width=3)
        indicator.pack(side=tk.LEFT)
        
        tk.Label(header, text=project['project_name'], font=('Segoe UI', 13, 'bold'),
                bg='#eff6ff').pack(side=tk.LEFT, pady=15, padx=10)
        
        # Cabinet count on the right
        tk.Label(header, text=f"📦 {project['cabinet_count']} Cabinet(s)",
                font=('Segoe UI', 11, 'bold'), bg='#eff6ff', fg='#3b82f6').pack(side=tk.RIGHT, padx=20)
        
        dropdown = tk.Frame(card, bg='white')
        
        def toggle():
            if expand_var.get():
                dropdown.pack_forget()
                indicator.config(text="▶")
                expand_var.set(False)
            else:
                self.populate_cabinets(dropdown, project['project_name'])
                dropdown.pack(fill=tk.BOTH, padx=15, pady=10)
                indicator.config(text="▼")
                expand_var.set(True)
        
        header.bind("<Button-1>", lambda e: toggle())
        indicator.bind("<Button-1>", lambda e: toggle())
    
    def populate_cabinets(self, parent, project_name):
        for w in parent.winfo_children():
            w.destroy()
        
        cabinets = self.db.get_cabinets_by_project(project_name)
        
        if not cabinets:
            tk.Label(parent, text="No cabinets", bg='white').pack(pady=20)
            return
        
        # Header
        hdr = tk.Frame(parent, bg='#f1f5f9')
        hdr.pack(fill=tk.X, pady=5)
        
        headers = [
            ("Cabinet", 18), ("Drawing %", 10), ("Total Punches", 12),
            ("Implemented", 12), ("Closed", 8), ("Status", 25), ("Debug", 8)
        ]
        
        for text, w in headers:
            tk.Label(hdr, text=text, font=('Segoe UI', 9, 'bold'),
                    bg='#f1f5f9', width=w, anchor='w').pack(side=tk.LEFT, padx=3)
        
        # Rows
        for cab in cabinets:
            row = tk.Frame(parent, bg='white')
            row.pack(fill=tk.X, pady=2)
            
            # Cabinet ID
            tk.Label(row, text=cab['cabinet_id'], font=('Segoe UI', 9),
                    bg='white', width=18, anchor='w').pack(side=tk.LEFT, padx=3)
            
            # Drawing completion percentage
            pct = (cab['annotated_pages']/cab['total_pages']*100) if cab['total_pages'] else 0
            tk.Label(row, text=f"{pct:.0f}%", font=('Segoe UI', 9, 'bold'),
                    bg='white', fg='#10b981' if pct==100 else '#f59e0b',
                    width=10).pack(side=tk.LEFT, padx=3)
            
            # Total Punches
            tk.Label(row, text=str(cab['total_punches']), font=('Segoe UI', 9),
                    bg='white', width=12, anchor='center').pack(side=tk.LEFT, padx=3)
            
            # Implemented
            tk.Label(row, text=str(cab['implemented_punches']), font=('Segoe UI', 9),
                    bg='white', width=12, anchor='center').pack(side=tk.LEFT, padx=3)
            
            # Closed
            tk.Label(row, text=str(cab['closed_punches']), font=('Segoe UI', 9),
                    bg='white', width=8, anchor='center').pack(side=tk.LEFT, padx=3)
            
            # Status
            status_map = {
                'quality_inspection': ('🔍 Quality Inspection', '#3b82f6'),
                'handed_to_production': ('📦 Handed to Production', '#8b5cf6'),
                'in_progress': ('🔧 Production Rework', '#f59e0b'),
                'being_closed_by_quality': ('✅ Being Closed', '#10b981'),
                'closed': ('✓ Closed', '#64748b')
            }
            
            status_text, status_color = status_map.get(
                cab['status'],
                (cab['status'].replace('_', ' ').title(), '#64748b')
            )
            
            status_label = tk.Label(row, text=status_text, font=('Segoe UI', 9, 'bold'),
                                   bg=status_color, fg='white', padx=10, pady=4,
                                   anchor='w', width=25)
            status_label.pack(side=tk.LEFT, padx=3)
            
            # Debug button
            debug_btn = tk.Button(row, text="🔍", 
                                 command=lambda c=cab: self.show_cabinet_debug(c),
                                 bg='#f59e0b', fg='white', font=('Segoe UI', 8, 'bold'),
                                 width=5, relief=tk.FLAT, cursor='hand2')
            debug_btn.pack(side=tk.LEFT, padx=3)
    
    def show_cabinet_debug(self, cabinet):
        """Show debug information for a cabinet"""
        conn = sqlite3.connect(self.db.db_path)
        cursor = conn.cursor()
        
        # Get actual counts from category_occurrences
        cursor.execute('''SELECT COUNT(*) FROM category_occurrences 
                         WHERE cabinet_id = ?''', (cabinet['cabinet_id'],))
        actual_total = cursor.fetchone()[0]
        
        # Get detailed breakdown
        debug_info = f"""Cabinet Debug Information
        
Cabinet ID: {cabinet['cabinet_id']}
Project: {cabinet['project_name']}

=== Stored in cabinets table ===
Total Punches: {cabinet['total_punches']}
Open: {cabinet['open_punches']}
Implemented: {cabinet['implemented_punches']}
Closed: {cabinet['closed_punches']}
Sum: {cabinet['open_punches'] + cabinet['implemented_punches'] + cabinet['closed_punches']}

=== Actual from category_occurrences ===
Total Logged Punches: {actual_total}

=== Discrepancy ===
Difference: {cabinet['total_punches'] - actual_total}

"""
        
        # Get punch details
        cursor.execute('''SELECT category, subcategory, occurrence_date 
                         FROM category_occurrences 
                         WHERE cabinet_id = ?
                         ORDER BY occurrence_date DESC''', (cabinet['cabinet_id'],))
        punches = cursor.fetchall()
        
        if punches:
            debug_info += "\n=== Logged Punches ===\n"
            for idx, (cat, subcat, date) in enumerate(punches, 1):
                debug_info += f"{idx}. {cat}"
                if subcat:
                    debug_info += f" → {subcat}"
                debug_info += f" ({date})\n"
        
        conn.close()
        
        # Show in a dialog
        debug_window = tk.Toplevel(self.root)
        debug_window.title(f"Debug: {cabinet['cabinet_id']}")
        debug_window.geometry("600x500")
        
        text_widget = tk.Text(debug_window, wrap=tk.WORD, font=('Courier New', 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_widget.insert('1.0', debug_info)
        text_widget.config(state=tk.DISABLED)
        
        tk.Button(debug_window, text="Close", command=debug_window.destroy,
                 bg='#3b82f6', fg='white', font=('Segoe UI', 10, 'bold'),
                 padx=20, pady=8).pack(pady=10)
    
    # ============ ANALYTICS - INTEGRATED SEARCH WITH FILTERS ============
    def show_analytics(self):
        self.set_active_nav('analytics')
        self.clear_content()
        
        # Header
        header = tk.Frame(self.content, bg='#f8fafc')
        header.pack(fill=tk.X, padx=30, pady=(20, 10))
        
        tk.Label(header, text="Category Analytics", font=('Segoe UI', 16, 'bold'),
                bg='#f8fafc').pack(side=tk.LEFT)
        
        # Integrated Search Bar with Filters
        search_control_frame = tk.Frame(self.content, bg='white', relief=tk.SOLID, borderwidth=1)
        search_control_frame.pack(fill=tk.X, padx=30, pady=(0, 10))
        
        # Main search bar
        search_bar_frame = tk.Frame(search_control_frame, bg='white')
        search_bar_frame.pack(fill=tk.X, padx=20, pady=(15, 10))
        
        tk.Label(search_bar_frame, text="🔍", bg='white', 
                font=('Segoe UI', 14)).pack(side=tk.LEFT, padx=(5, 10))
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_bar_frame, textvariable=search_var, width=50,
                               font=('Segoe UI', 11), relief=tk.FLAT, bg='#f8fafc')
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 10))
        
        # Suggestion dropdown
        suggestion_frame = tk.Frame(search_control_frame, bg='white')
        suggestion_listbox = tk.Listbox(suggestion_frame, height=5, font=('Segoe UI', 10),
                                       relief=tk.FLAT, bg='#f8fafc', borderwidth=0)
        
        all_projects = self.db.get_all_project_names()
        
        def update_suggestions(*args):
            search_text = search_var.get().lower()
            suggestion_listbox.delete(0, tk.END)
            
            if search_text:
                matches = [p for p in all_projects if search_text in p.lower()]
                if matches:
                    for match in matches[:5]:  # Show top 5 matches
                        suggestion_listbox.insert(tk.END, match)
                    suggestion_frame.pack(fill=tk.X, padx=20, pady=(0, 10))
                else:
                    suggestion_frame.pack_forget()
            else:
                suggestion_frame.pack_forget()
        
        def select_suggestion(event):
            if suggestion_listbox.curselection():
                selected = suggestion_listbox.get(suggestion_listbox.curselection())
                search_var.set(selected)
                suggestion_frame.pack_forget()
                apply_filters()
        
        search_var.trace('w', update_suggestions)
        suggestion_listbox.pack(fill=tk.X, padx=10, pady=5)
        suggestion_listbox.bind('<<ListboxSelect>>', select_suggestion)
        
        # Filter buttons
        filter_frame = tk.Frame(search_control_frame, bg='white')
        filter_frame.pack(fill=tk.X, padx=20, pady=(0, 15))
        
        tk.Label(filter_frame, text="Filter by:", font=('Segoe UI', 10, 'bold'),
                bg='white').pack(side=tk.LEFT, padx=(0, 10))
        
        # Date filter options
        date_filter_var = tk.StringVar(value="all")
        
        filter_buttons = [
            ("All Time", "all"),
            ("Today", "today"),
            ("This Week", "week"),
            ("This Month", "month"),
            ("This Year", "year"),
            ("Custom Date", "custom")
        ]
        
        for text, value in filter_buttons:
            btn = tk.Radiobutton(filter_frame, text=text, variable=date_filter_var,
                               value=value, bg='white', font=('Segoe UI', 9),
                               indicatoron=False, padx=15, pady=5,
                               selectcolor='#3b82f6', fg='#1e293b',
                               activebackground='#3b82f6', activeforeground='white',
                               relief=tk.FLAT, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=2)
        
        # Level selection (Category/Subcategory)
        tk.Label(filter_frame, text=" | View:", font=('Segoe UI', 10, 'bold'),
                bg='white').pack(side=tk.LEFT, padx=(20, 10))
        
        level_var = tk.StringVar(value="category")
        tk.Radiobutton(filter_frame, text="Category", variable=level_var,
                      value="category", bg='white', font=('Segoe UI', 9),
                      indicatoron=False, padx=15, pady=5,
                      selectcolor='#10b981', fg='#1e293b',
                      activebackground='#10b981', activeforeground='white',
                      relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=2)
        
        tk.Radiobutton(filter_frame, text="Subcategory", variable=level_var,
                      value="subcategory", bg='white', font=('Segoe UI', 9),
                      indicatoron=False, padx=15, pady=5,
                      selectcolor='#10b981', fg='#1e293b',
                      activebackground='#10b981', activeforeground='white',
                      relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=2)
        
        # Export button
        tk.Button(filter_frame, text="📥 Export",
                 command=lambda: self.export_excel_filtered(),
                 bg='#8b5cf6', fg='white', font=('Segoe UI', 9, 'bold'),
                 padx=15, pady=5, relief=tk.FLAT, cursor='hand2').pack(side=tk.RIGHT, padx=5)
        
        # Custom date picker frame (hidden by default)
        custom_date_frame = tk.Frame(search_control_frame, bg='#f8fafc')
        
        tk.Label(custom_date_frame, text="From:", bg='#f8fafc',
                font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(20, 5))
        start_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        start_date_entry = tk.Entry(custom_date_frame, textvariable=start_date_var,
                                    width=12, font=('Segoe UI', 9))
        start_date_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Label(custom_date_frame, text="To:", bg='#f8fafc',
                font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(20, 5))
        end_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        end_date_entry = tk.Entry(custom_date_frame, textvariable=end_date_var,
                                  width=12, font=('Segoe UI', 9))
        end_date_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(custom_date_frame, text="Apply", command=lambda: apply_filters(),
                 bg='#3b82f6', fg='white', font=('Segoe UI', 9, 'bold'),
                 padx=10, pady=3, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=10)
        
        def show_custom_date(*args):
            if date_filter_var.get() == "custom":
                custom_date_frame.pack(fill=tk.X, padx=20, pady=(0, 15))
            else:
                custom_date_frame.pack_forget()
                apply_filters()
        
        date_filter_var.trace('w', show_custom_date)
        level_var.trace('w', lambda *args: apply_filters())
        
        # Chart frame
        self.chart_frame = tk.Frame(self.content, bg='white')
        self.chart_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=(0, 20))
        
        # Store variables for export
        self.analytics_search_var = search_var
        self.analytics_date_filter = date_filter_var
        self.analytics_level = level_var
        self.analytics_start_date = start_date_var
        self.analytics_end_date = end_date_var
        
        def apply_filters():
            project_filter = search_var.get().strip() if search_var.get() != "Search projects or select filters..." else None
            date_filter = date_filter_var.get()
            level = level_var.get()
            
            # Calculate date range
            start_date = None
            end_date = None
            
            if date_filter == "today":
                start_date = datetime.now().date().isoformat()
                end_date = start_date
            elif date_filter == "week":
                today = datetime.now().date()
                start_date = (today - timedelta(days=today.weekday())).isoformat()
                end_date = today.isoformat()
            elif date_filter == "month":
                today = datetime.now().date()
                start_date = today.replace(day=1).isoformat()
                end_date = today.isoformat()
            elif date_filter == "year":
                today = datetime.now().date()
                # Financial year starts October 1st
                if today.month >= 10:
                    start_date = datetime(today.year, 10, 1).date().isoformat()
                else:
                    start_date = datetime(today.year - 1, 10, 1).date().isoformat()
                end_date = today.isoformat()
            elif date_filter == "custom":
                start_date = start_date_var.get()
                end_date = end_date_var.get()
            
            self.update_chart_with_filters(start_date, end_date, project_filter, level)
        
        # Initial load
        apply_filters()
        
        # Placeholder behavior
        search_entry.insert(0, "Search projects or select filters...")
        search_entry.config(fg='#94a3b8')
        
        def on_focus_in(event):
            if search_entry.get() == "Search projects or select filters...":
                search_entry.delete(0, tk.END)
                search_entry.config(fg='#1e293b')
        
        def on_focus_out(event):
            if not search_entry.get():
                search_entry.insert(0, "Search projects or select filters...")
                search_entry.config(fg='#94a3b8')
        
        search_entry.bind("<FocusIn>", on_focus_in)
        search_entry.bind("<FocusOut>", on_focus_out)
        search_entry.bind("<Return>", lambda e: apply_filters())
    
    def update_chart_with_filters(self, start_date, end_date, project, level):
        """Update chart with filtered data"""
        # Clear previous chart
        for w in self.chart_frame.winfo_children():
            w.destroy()
        plt.close('all')
        
        stats = self.db.get_category_stats(start_date, end_date, project)
        
        if not stats:
            empty_frame = tk.Frame(self.chart_frame, bg='white')
            empty_frame.place(relx=0.5, rely=0.5, anchor='center')
            
            tk.Label(empty_frame, text="No data available for the selected filters.",
                    font=('Segoe UI', 12), fg='#64748b', bg='white').pack(pady=5)
            tk.Label(empty_frame, 
                    text="Category data will appear once Quality Inspection logs punches.",
                    font=('Segoe UI', 10), fg='#94a3b8', bg='white').pack(pady=5)
            return
        
        counts = defaultdict(int)
        if level == "category":
            for item in stats:
                counts[item['category']] += item['count']
        else:
            for item in stats:
                key = f"{item['category']} → {item['subcategory'] or 'N/A'}"
                counts[key] += item['count']
        
        sorted_items = sorted(counts.items(), key=lambda x: x[1], reverse=True)[:15]
        labels = [item[0] for item in sorted_items]
        values = [item[1] for item in sorted_items]
        
        total = sum(values)
        cumulative = []
        cum = 0
        for v in values:
            cum += v
            cumulative.append((cum/total)*100)
        
        fig = Figure(figsize=(14, 7), facecolor='white')
        ax1 = fig.add_subplot(111)
        ax2 = ax1.twinx()
        
        bars = ax1.bar(range(len(labels)), values, color='#3b82f6', alpha=0.7)
        line = ax2.plot(range(len(labels)), cumulative, color='#ef4444',
                       marker='o', linewidth=2, markersize=6)
        ax2.axhline(y=80, color='#10b981', linestyle='--', linewidth=1.5, alpha=0.7)
        
        ax1.set_xlabel('Category', fontsize=11, fontweight='bold')
        ax1.set_ylabel('Frequency', fontsize=11, fontweight='bold', color='#3b82f6')
        ax2.set_ylabel('Cumulative %', fontsize=11, fontweight='bold', color='#ef4444')
        
        # Add filter info to title
        filter_text = f"{level.title()} Analysis"
        if project:
            filter_text += f" - {project}"
        
        ax1.set_title(f'Pareto Chart - {filter_text}',
                     fontsize=14, fontweight='bold')
        
        ax1.set_xticks(range(len(labels)))
        ax1.set_xticklabels(labels, rotation=45, ha='right', fontsize=9)
        ax1.tick_params(axis='y', labelcolor='#3b82f6')
        ax2.tick_params(axis='y', labelcolor='#ef4444')
        ax2.set_ylim(0, 105)
        ax1.grid(axis='y', alpha=0.3, linestyle='--')
        
        fig.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def export_excel_filtered(self):
        """Export with current filters - creates two separate Excel files with dynamic columns"""
        # Ask user to select location for first file
        file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Category Analysis Excel"
        )
        if not file:
            return
        
        # Generate second filename
        base_name = file.rsplit('.', 1)[0]
        category_file = f"{base_name}_Category_Analysis.xlsx"
        subcategory_file = f"{base_name}_Subcategory_Analysis.xlsx"
        
        # Get current filters
        project_filter = self.analytics_search_var.get().strip() if self.analytics_search_var.get() != "Search projects or select filters..." else None
        date_filter = self.analytics_date_filter.get()
        
        start_date = None
        end_date = None
        
        today = datetime.now().date()
        
        if date_filter == "today":
            start_date = today.isoformat()
            end_date = start_date
        elif date_filter == "week":
            start_date = (today - timedelta(days=today.weekday())).isoformat()
            end_date = today.isoformat()
        elif date_filter == "month":
            start_date = today.replace(day=1).isoformat()
            end_date = today.isoformat()
        elif date_filter == "year":
            if today.month >= 10:
                start_date = datetime(today.year, 10, 1).date().isoformat()
            else:
                start_date = datetime(today.year - 1, 10, 1).date().isoformat()
            end_date = today.isoformat()
        elif date_filter == "custom":
            start_date = self.analytics_start_date.get()
            end_date = self.analytics_end_date.get()
        
        # Get raw data from database with dates
        conn = sqlite3.connect(self.db.db_path)
        cursor = conn.cursor()
        
        query = '''SELECT category, subcategory, occurrence_date, cabinet_id, project_name 
                   FROM category_occurrences WHERE 1=1'''
        params = []
        
        if start_date:
            query += ' AND occurrence_date >= ?'
            params.append(start_date)
        
        if end_date:
            query += ' AND occurrence_date <= ?'
            params.append(end_date)
        
        if project_filter:
            query += ' AND project_name = ?'
            params.append(project_filter)
        
        cursor.execute(query, params)
        raw_data = cursor.fetchall()
        conn.close()
        
        if not raw_data:
            messagebox.showwarning("No Data", "No data available for the selected filters.")
            return
        
        # Determine column structure based on filter
        if project_filter:
            # Cabinet-wise columns
            self.export_by_cabinets(category_file, subcategory_file, raw_data, project_filter)
        elif date_filter == "year":
            # Month-wise columns
            self.export_by_months(category_file, subcategory_file, raw_data, start_date, end_date)
        else:
            # Day-wise columns (for today, week, month, custom)
            self.export_by_days(category_file, subcategory_file, raw_data, start_date, end_date, date_filter)
        
        messagebox.showinfo("Export Complete", 
                          f"Data exported to:\n\n1. {category_file}\n2. {subcategory_file}")
    
    def export_by_cabinets(self, category_file, subcategory_file, raw_data, project_name):
        """Export with cabinet-wise columns"""
        # Get unique cabinets
        cabinets = sorted(list(set(row[3] for row in raw_data)))
        
        # Category Analysis
        wb_cat = Workbook()
        ws_cat = wb_cat.active
        ws_cat.title = "Category Analysis"
        
        # Headers
        ws_cat['A1'] = "Category"
        for idx, cabinet in enumerate(cabinets, start=2):
            ws_cat.cell(1, idx, cabinet)
        ws_cat.cell(1, len(cabinets) + 2, "Total")
        
        # Style headers
        header_fill = PatternFill(start_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for col in range(1, len(cabinets) + 3):
            cell = ws_cat.cell(1, col)
            cell.fill = header_fill
            cell.font = header_font
        
        # Aggregate data by category and cabinet
        category_cabinet_counts = defaultdict(lambda: defaultdict(int))
        for row in raw_data:
            category, _, _, cabinet, _ = row
            category_cabinet_counts[category][cabinet] += 1
        
        # Write data
        row_num = 2
        for category in sorted(category_cabinet_counts.keys()):
            ws_cat.cell(row_num, 1, category)
            total = 0
            for idx, cabinet in enumerate(cabinets, start=2):
                count = category_cabinet_counts[category][cabinet]
                ws_cat.cell(row_num, idx, count if count > 0 else 0)
                total += count
            ws_cat.cell(row_num, len(cabinets) + 2, total)
            row_num += 1
        
        wb_cat.save(category_file)
        
        # Subcategory Analysis
        wb_sub = Workbook()
        ws_sub = wb_sub.active
        ws_sub.title = "Subcategory Analysis"
        
        # Headers
        ws_sub['A1'] = "Category"
        ws_sub['B1'] = "Subcategory"
        for idx, cabinet in enumerate(cabinets, start=3):
            ws_sub.cell(1, idx, cabinet)
        ws_sub.cell(1, len(cabinets) + 3, "Total")
        
        # Style headers
        for col in range(1, len(cabinets) + 4):
            cell = ws_sub.cell(1, col)
            cell.fill = header_fill
            cell.font = header_font
        
        # Aggregate data by category, subcategory, and cabinet
        subcat_cabinet_counts = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
        for row in raw_data:
            category, subcategory, _, cabinet, _ = row
            subcat_cabinet_counts[category][subcategory or 'N/A'][cabinet] += 1
        
        # Write data
        row_num = 2
        for category in sorted(subcat_cabinet_counts.keys()):
            for subcategory in sorted(subcat_cabinet_counts[category].keys()):
                ws_sub.cell(row_num, 1, category)
                ws_sub.cell(row_num, 2, subcategory)
                total = 0
                for idx, cabinet in enumerate(cabinets, start=3):
                    count = subcat_cabinet_counts[category][subcategory][cabinet]
                    ws_sub.cell(row_num, idx, count if count > 0 else 0)
                    total += count
                ws_sub.cell(row_num, len(cabinets) + 3, total)
                row_num += 1
        
        wb_sub.save(subcategory_file)
    
    def export_by_months(self, category_file, subcategory_file, raw_data, start_date, end_date):
        """Export with month-wise columns"""
        # Generate list of months in the range
        start = datetime.fromisoformat(start_date)
        end = datetime.fromisoformat(end_date)
        
        months = []
        current = start.replace(day=1)
        while current <= end:
            months.append(current.strftime("%B %Y"))
            # Move to next month
            if current.month == 12:
                current = current.replace(year=current.year + 1, month=1)
            else:
                current = current.replace(month=current.month + 1)
        
        # Category Analysis
        wb_cat = Workbook()
        ws_cat = wb_cat.active
        ws_cat.title = "Category Analysis"
        
        # Headers
        ws_cat['A1'] = "Category"
        for idx, month in enumerate(months, start=2):
            ws_cat.cell(1, idx, month)
        ws_cat.cell(1, len(months) + 2, "Total")
        
        # Style headers
        header_fill = PatternFill(start_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for col in range(1, len(months) + 3):
            cell = ws_cat.cell(1, col)
            cell.fill = header_fill
            cell.font = header_font
        
        # Aggregate data by category and month
        category_month_counts = defaultdict(lambda: defaultdict(int))
        for row in raw_data:
            category, _, date_str, _, _ = row
            if date_str:
                date_obj = datetime.fromisoformat(date_str)
                month_key = date_obj.strftime("%B %Y")
                category_month_counts[category][month_key] += 1
        
        # Write data
        row_num = 2
        for category in sorted(category_month_counts.keys()):
            ws_cat.cell(row_num, 1, category)
            total = 0
            for idx, month in enumerate(months, start=2):
                count = category_month_counts[category][month]
                ws_cat.cell(row_num, idx, count if count > 0 else 0)
                total += count
            ws_cat.cell(row_num, len(months) + 2, total)
            row_num += 1
        
        wb_cat.save(category_file)
        
        # Subcategory Analysis
        wb_sub = Workbook()
        ws_sub = wb_sub.active
        ws_sub.title = "Subcategory Analysis"
        
        # Headers
        ws_sub['A1'] = "Category"
        ws_sub['B1'] = "Subcategory"
        for idx, month in enumerate(months, start=3):
            ws_sub.cell(1, idx, month)
        ws_sub.cell(1, len(months) + 3, "Total")
        
        # Style headers
        for col in range(1, len(months) + 4):
            cell = ws_sub.cell(1, col)
            cell.fill = header_fill
            cell.font = header_font
        
        # Aggregate data by category, subcategory, and month
        subcat_month_counts = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
        for row in raw_data:
            category, subcategory, date_str, _, _ = row
            if date_str:
                date_obj = datetime.fromisoformat(date_str)
                month_key = date_obj.strftime("%B %Y")
                subcat_month_counts[category][subcategory or 'N/A'][month_key] += 1
        
        # Write data
        row_num = 2
        for category in sorted(subcat_month_counts.keys()):
            for subcategory in sorted(subcat_month_counts[category].keys()):
                ws_sub.cell(row_num, 1, category)
                ws_sub.cell(row_num, 2, subcategory)
                total = 0
                for idx, month in enumerate(months, start=3):
                    count = subcat_month_counts[category][subcategory][month]
                    ws_sub.cell(row_num, idx, count if count > 0 else 0)
                    total += count
                ws_sub.cell(row_num, len(months) + 3, total)
                row_num += 1
        
        wb_sub.save(subcategory_file)
    
    def export_by_days(self, category_file, subcategory_file, raw_data, start_date, end_date, filter_type):
        """Export with day-wise columns"""
        # Generate list of days in the range
        start = datetime.fromisoformat(start_date)
        end = datetime.fromisoformat(end_date)
        
        days = []
        current = start
        while current <= end:
            days.append(current.strftime("%Y-%m-%d"))
            current += timedelta(days=1)
        
        # Category Analysis
        wb_cat = Workbook()
        ws_cat = wb_cat.active
        ws_cat.title = "Category Analysis"
        
        # Headers
        ws_cat['A1'] = "Category"
        for idx, day in enumerate(days, start=2):
            ws_cat.cell(1, idx, day)
        ws_cat.cell(1, len(days) + 2, "Total")
        
        # Style headers
        header_fill = PatternFill(start_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for col in range(1, len(days) + 3):
            cell = ws_cat.cell(1, col)
            cell.fill = header_fill
            cell.font = header_font
            # Auto-adjust column width for dates
            if col > 1:
                ws_cat.column_dimensions[ws_cat.cell(1, col).column_letter].width = 12
        
        # Aggregate data by category and day
        category_day_counts = defaultdict(lambda: defaultdict(int))
        for row in raw_data:
            category, _, date_str, _, _ = row
            if date_str:
                date_key = date_str[:10]  # Extract YYYY-MM-DD
                category_day_counts[category][date_key] += 1
        
        # Write data
        row_num = 2
        for category in sorted(category_day_counts.keys()):
            ws_cat.cell(row_num, 1, category)
            total = 0
            for idx, day in enumerate(days, start=2):
                count = category_day_counts[category][day]
                ws_cat.cell(row_num, idx, count if count > 0 else 0)
                total += count
            ws_cat.cell(row_num, len(days) + 2, total)
            row_num += 1
        
        wb_cat.save(category_file)
        
        # Subcategory Analysis
        wb_sub = Workbook()
        ws_sub = wb_sub.active
        ws_sub.title = "Subcategory Analysis"
        
        # Headers
        ws_sub['A1'] = "Category"
        ws_sub['B1'] = "Subcategory"
        for idx, day in enumerate(days, start=3):
            ws_sub.cell(1, idx, day)
        ws_sub.cell(1, len(days) + 3, "Total")
        
        # Style headers
        for col in range(1, len(days) + 4):
            cell = ws_sub.cell(1, col)
            cell.fill = header_fill
            cell.font = header_font
            # Auto-adjust column width for dates
            if col > 2:
                ws_sub.column_dimensions[ws_sub.cell(1, col).column_letter].width = 12
        
        # Aggregate data by category, subcategory, and day
        subcat_day_counts = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
        for row in raw_data:
            category, subcategory, date_str, _, _ = row
            if date_str:
                date_key = date_str[:10]  # Extract YYYY-MM-DD
                subcat_day_counts[category][subcategory or 'N/A'][date_key] += 1
        
        # Write data
        row_num = 2
        for category in sorted(subcat_day_counts.keys()):
            for subcategory in sorted(subcat_day_counts[category].keys()):
                ws_sub.cell(row_num, 1, category)
                ws_sub.cell(row_num, 2, subcategory)
                total = 0
                for idx, day in enumerate(days, start=3):
                    count = subcat_day_counts[category][subcategory][day]
                    ws_sub.cell(row_num, idx, count if count > 0 else 0)
                    total += count
                ws_sub.cell(row_num, len(days) + 3, total)
                row_num += 1
        
        wb_sub.save(subcategory_file)
    
    # ============ CATEGORIES - BIGGER BUTTONS WITH TEXT ============
    def show_categories(self):
        self.set_active_nav('categories')
        self.clear_content()
        
        # Centered container
        center_container = tk.Frame(self.content, bg='#f8fafc')
        center_container.place(relx=0.5, rely=0, anchor='n', relwidth=0.7, relheight=1.0)
        
        # Header
        header = tk.Frame(center_container, bg='#f8fafc')
        header.pack(fill=tk.X, padx=30, pady=(20, 10))
        
        tk.Label(header, text="Category Management", font=('Segoe UI', 16, 'bold'),
                bg='#f8fafc').pack(side=tk.LEFT)
        
        tk.Button(header, text="➕ Add Category", command=self.add_category,
                 bg='#10b981', fg='white', font=('Segoe UI', 10, 'bold'),
                 padx=15, pady=8).pack(side=tk.RIGHT)
        
        if not self.categories:
            empty_container = tk.Frame(center_container, bg='#f8fafc')
            empty_container.pack(expand=True, fill=tk.BOTH)
            center_frame = tk.Frame(empty_container, bg='#f8fafc')
            center_frame.place(relx=0.5, rely=0.5, anchor='center')
            
            tk.Label(center_frame, text="No categories defined",
                    font=('Segoe UI', 16, 'bold'), fg='#1e293b', bg='#f8fafc').pack(pady=10)
            tk.Label(center_frame, text="Click 'Add Category' to create your first punch category.",
                    font=('Segoe UI', 11), fg='#64748b', bg='#f8fafc').pack(pady=5)
            return
        
        # Scrollable container
        canvas_container = tk.Frame(center_container, bg='#f8fafc')
        canvas_container.pack(expand=True, fill=tk.BOTH, padx=30, pady=(0, 20))
        
        canvas = tk.Canvas(canvas_container, bg='#f8fafc', highlightthickness=0)
        scrollbar = tk.Scrollbar(canvas_container, command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg='#f8fafc')
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        scroll_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        for cat in self.categories:
            self.create_category_card(scroll_frame, cat)
    
    def create_category_card(self, parent, category):
        card = tk.Frame(parent, bg='white', relief=tk.SOLID, borderwidth=1)
        card.pack(fill=tk.X, pady=8, padx=5)
        
        header = tk.Frame(card, bg='#dbeafe')
        header.pack(fill=tk.X)
        
        tk.Label(header, text=category['name'], font=('Segoe UI', 12, 'bold'),
                bg='#dbeafe', fg='#1e40af').pack(side=tk.LEFT, padx=15, pady=10)
        
        # Determine mode - check if category has mode field, otherwise infer
        mode = category.get('mode')
        if not mode:
            # Infer mode from structure for existing categories
            mode = 'parent' if category.get('subcategories') else 'template'
        
        mode_text = "📝 Template" if mode == 'template' else "📁 Parent"
        tk.Label(header, text=mode_text, font=('Segoe UI', 9),
                bg='#dbeafe', fg='#64748b').pack(side=tk.LEFT, padx=10)
        
        btn_frame = tk.Frame(header, bg='#dbeafe')
        btn_frame.pack(side=tk.RIGHT, padx=10)
        
        # Bigger buttons with text labels
        tk.Button(btn_frame, text="✏️ Edit", command=lambda: self.edit_category(category),
                 bg='#3b82f6', fg='white', font=('Segoe UI', 9, 'bold'),
                 padx=12, pady=6, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=3)
        
        tk.Button(btn_frame, text="🗑️ Delete", command=lambda: self.delete_category(category),
                 bg='#ef4444', fg='white', font=('Segoe UI', 9, 'bold'),
                 padx=12, pady=6, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=3)
        
        if mode == 'parent':
            tk.Button(btn_frame, text="➕ Add Sub", command=lambda: self.add_subcategory(category),
                     bg='#10b981', fg='white', font=('Segoe UI', 9, 'bold'),
                     padx=12, pady=6, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=3)
        elif mode == 'template':
            # Add Test button for template categories
            tk.Button(btn_frame, text="▶️ Test", command=lambda: self.handle_template_category(category),
                     bg='#8b5cf6', fg='white', font=('Segoe UI', 9, 'bold'),
                     padx=12, pady=6, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=3)
        
        # Subcategories
        if category.get('subcategories'):
            sub_frame = tk.Frame(card, bg='white')
            sub_frame.pack(fill=tk.X, padx=20, pady=10)
            
            for sub in category['subcategories']:
                sub_row = tk.Frame(sub_frame, bg='#f8fafc')
                sub_row.pack(fill=tk.X, pady=3)
                
                tk.Label(sub_row, text=f" ↳ {sub['name']}", font=('Segoe UI', 10),
                        bg='#f8fafc', anchor='w').pack(side=tk.LEFT, fill=tk.X,
                                                      expand=True, padx=10, pady=8)
                
                sub_btn_frame = tk.Frame(sub_row, bg='#f8fafc')
                sub_btn_frame.pack(side=tk.RIGHT, padx=10)
                
                # Test button for subcategories
                tk.Button(sub_btn_frame, text="▶️ Test",
                         command=lambda c=category, s=sub: self.handle_subcategory(c, s),
                         bg='#8b5cf6', fg='white', font=('Segoe UI', 8, 'bold'),
                         padx=10, pady=5, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=2)
                
                # Bigger subcategory buttons with text
                tk.Button(sub_btn_frame, text="✏️ Edit",
                         command=lambda c=category, s=sub: self.edit_subcategory(c, s),
                         bg='#3b82f6', fg='white', font=('Segoe UI', 8, 'bold'),
                         padx=10, pady=5, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=2)
                
                tk.Button(sub_btn_frame, text="🗑️ Delete",
                         command=lambda c=category, s=sub: self.delete_subcategory(c, s),
                         bg='#ef4444', fg='white', font=('Segoe UI', 8, 'bold'),
                         padx=10, pady=5, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=2)
    
    # ============================================================
    # TEMPLATE DATA COLLECTION
    # ============================================================
    def collect_template_data(self, mandatory=True, existing=None):
        """Collect or edit inputs + template."""
        min_inputs = 1 if mandatory else 0
        default_inputs = len(existing.get("inputs", [])) if existing else min_inputs
        num_inputs = simpledialog.askinteger(
            "Expected Inputs",
            "How many inputs are required?",
            parent=self.root,
            minvalue=min_inputs,
            maxvalue=10,
            initialvalue=default_inputs
        )
        if num_inputs is None:
            return None
        
        inputs = []
        for i in range(num_inputs):
            default = existing["inputs"][i] if existing and i < len(existing.get("inputs", [])) else None
            name = simpledialog.askstring(
                "Input Name",
                f"Internal name for input #{i+1}",
                parent=self.root,
                initialvalue=default.get("name") if default else ""
            )
            if not name:
                return None
            
            label = simpledialog.askstring(
                "Input Label",
                f"Question asked to user for '{name}'",
                parent=self.root,
                initialvalue=default.get("label") if default else ""
            )
            if not label:
                return None
            
            inputs.append({"name": name.strip(), "label": label.strip()})
        
        placeholder_text = ", ".join(f"{{{i['name']}}}" for i in inputs)
        default_template = existing.get("template") if existing else ""
        template = simpledialog.askstring(
            "Punch Text Template",
            f"Enter punch text template.\nAvailable placeholders:\n{placeholder_text}",
            parent=self.root,
            initialvalue=default_template
        )
        
        if mandatory and not template:
            messagebox.showerror("Required", "Template is mandatory")
            return None
        
        return {"inputs": inputs, "template": template.strip() if template else None}
    
    # ============================================================
    # CATEGORY / SUBCATEGORY CRUD HELPERS
    # ============================================================
    def create_category(self):
        name = simpledialog.askstring("New Category", "Enter category name:", parent=self.root)
        if not name:
            return None
        
        category = {
            "name": name.strip(),
            "mode": None,
            "inputs": [],
            "template": None,
            "subcategories": []
        }
        
        use_template = messagebox.askyesno(
            "Category Type",
            "Does this category directly generate punch text?\n\nYES → Template category\nNO → Parent category",
            parent=self.root
        )
        
        if use_template:
            category["mode"] = "template"
            data = self.collect_template_data(mandatory=False)
            if data:
                category.update(data)
        else:
            category["mode"] = "parent"
        
        return category
    
    def run_template(self, template_def, tag_name=None):
        """
        Execute a template definition at runtime:
        - ask for inputs
        - optionally inject tag_name
        - return final punch text
        """
        values = {}
        if tag_name:
            values["tag"] = tag_name
        
        for inp in template_def.get("inputs", []):
            val = simpledialog.askstring(
                "Input Required",
                inp["label"],
                parent=self.root
            )
            if not val:
                return None
            values[inp["name"]] = val.strip()
        
        try:
            return template_def["template"].format(**values)
        except KeyError as e:
            messagebox.showerror("Template Error", f"Missing placeholder: {e}")
            return None
    
    def add_category(self):
        cat = self.create_category()
        if not cat:
            return
        
        if any(c["name"].lower() == cat["name"].lower() for c in self.categories):
            messagebox.showwarning("Duplicate", "Category already exists")
            return
        
        self.categories.append(cat)
        self.save_categories()
        self.show_categories()
    
    def edit_category(self, category):
        # Check if this is a new-style category with mode field and inputs
        mode = category.get('mode')
        
        # If no mode field, infer from structure
        if not mode:
            mode = 'parent' if category.get('subcategories') else 'template'
            category['mode'] = mode  # Add mode field
        
        # TEMPLATE CATEGORY
        if mode == 'template':
            # Check if it has the full input structure or just a simple template
            if category.get('inputs'):
                # New style with inputs - use full template definition editor
                updated = self.edit_template_definition(
                    "Edit Category",
                    existing=category,
                    require_inputs=True
                )
                if not updated:
                    return
                category.clear()
                category.update(updated)
                category["mode"] = "template"
            else:
                # Old style - just edit name and simple template
                new_name = simpledialog.askstring(
                    "Edit Category",
                    "Enter new category name:",
                    initialvalue=category["name"],
                    parent=self.root
                )
                if not new_name:
                    return
                
                new_template = simpledialog.askstring(
                    "Edit Template",
                    "Enter punch text template:",
                    initialvalue=category.get("template", ""),
                    parent=self.root
                )
                if new_template is None:
                    return
                
                category["name"] = new_name.strip()
                category["template"] = new_template.strip()
        
        # PARENT CATEGORY
        elif mode == "parent":
            new_name = simpledialog.askstring(
                "Edit Category",
                "Enter new category name:",
                initialvalue=category["name"],
                parent=self.root
            )
            if not new_name:
                return
            category["name"] = new_name.strip()
        
        self.save_categories()
        self.show_categories()
    
    def edit_template_definition(self, title, existing, require_inputs):
        """Edit a template definition (used for both categories and subcategories)"""
        # Edit name
        new_name = simpledialog.askstring(
            title,
            "Enter new name:",
            initialvalue=existing.get("name", ""),
            parent=self.root
        )
        if not new_name:
            return None
        
        # Collect template data
        data = self.collect_template_data(mandatory=require_inputs, existing=existing)
        if not data:
            return None
        
        result = {"name": new_name.strip()}
        result.update(data)
        return result
    
    def delete_category(self, category):
        if not messagebox.askyesno("Confirm", f"Delete category '{category['name']}'?"):
            return
        self.categories.remove(category)
        self.save_categories()
        self.show_categories()
    
    def handle_template_category(self, category, bbox_page=None):
        """Handle template category execution"""
        # Check if it has inputs or just a simple template
        if category.get('inputs'):
            # New style with inputs
            punch_text = self.run_template(category, tag_name=None)
            if not punch_text:
                return
        else:
            # Old style - just show the template as-is
            punch_text = category.get('template', 'No template defined')
        
        # For manager UI, just show the generated text
        messagebox.showinfo("Generated Punch Text", 
                          f"Category: {category['name']}\n\nPunch Text:\n{punch_text}")
    
    def handle_subcategory(self, category, subcategory, bbox_page=None):
        """Handle subcategory execution"""
        # Check if it has inputs or just a simple template
        if subcategory.get('inputs'):
            # New style with inputs
            punch_text = self.run_template(subcategory, tag_name=None)
            if not punch_text:
                return
        else:
            # Old style - just show the template as-is
            punch_text = subcategory.get('template', 'No template defined')
        
        # For manager UI, just show the generated text
        messagebox.showinfo("Generated Punch Text",
                          f"Category: {category['name']}\nSubcategory: {subcategory['name']}\n\nPunch Text:\n{punch_text}")
    
    def add_subcategory(self, category):
        name = simpledialog.askstring("New Subcategory", "Enter subcategory name:", parent=self.root)
        if not name:
            return
        
        data = self.collect_template_data(mandatory=True)
        if not data:
            return
        
        if 'subcategories' not in category:
            category['subcategories'] = []
        
        category["subcategories"].append({"name": name.strip(), **data})
        self.save_categories()
        self.show_categories()
    
    def edit_subcategory(self, category, subcategory):
        # Check if subcategory has the full input structure or just simple template
        if subcategory.get('inputs'):
            # New style with inputs - use full template definition editor
            updated = self.edit_template_definition(
                "Edit Subcategory",
                existing=subcategory,
                require_inputs=True
            )
            if not updated:
                return
            subcategory.clear()
            subcategory.update(updated)
        else:
            # Old style - just edit name and simple template
            new_name = simpledialog.askstring(
                "Edit Subcategory",
                "Enter new name:",
                initialvalue=subcategory['name'],
                parent=self.root
            )
            if not new_name:
                return
            
            new_template = simpledialog.askstring(
                "Edit Template",
                "Enter new template:",
                initialvalue=subcategory.get('template', ''),
                parent=self.root
            )
            if new_template is None:
                return
            
            subcategory['name'] = new_name.strip()
            subcategory['template'] = new_template
        
        self.save_categories()
        self.show_categories()
    
    def delete_subcategory(self, category, sub):
        if not messagebox.askyesno("Confirm", f"Delete subcategory '{sub['name']}'?"):
            return
        
        if 'subcategories' in category:
            category['subcategories'].remove(sub)
            self.save_categories()
            self.show_categories()


def main():
    root = tk.Tk()
    app = ManagerUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
