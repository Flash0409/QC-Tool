"""
Manager UI - Complete Dashboard, Analytics, and Category Management System
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

def get_app_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

class ManagerDatabase:
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS cabinets (
            cabinet_id TEXT PRIMARY KEY, project_name TEXT, sales_order_no TEXT,
            total_pages INTEGER DEFAULT 0, annotated_pages INTEGER DEFAULT 0,
            total_punches INTEGER DEFAULT 0, open_punches INTEGER DEFAULT 0,
            implemented_punches INTEGER DEFAULT 0, closed_punches INTEGER DEFAULT 0,
            status TEXT DEFAULT 'quality_inspection', created_date TEXT, last_updated TEXT)''')
        
        cursor.execute('''CREATE TABLE IF NOT EXISTS category_occurrences (
            id INTEGER PRIMARY KEY AUTOINCREMENT, cabinet_id TEXT, project_name TEXT,
            category TEXT, subcategory TEXT, occurrence_date TEXT)''')
        
        conn.commit()
        conn.close()
    
    def get_all_projects(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''SELECT project_name, COUNT(DISTINCT cabinet_id) as count,
                         MAX(last_updated) as updated FROM cabinets 
                         GROUP BY project_name ORDER BY updated DESC''')
        projects = [{'project_name': r[0], 'cabinet_count': r[1], 'last_updated': r[2]} 
                   for r in cursor.fetchall()]
        conn.close()
        return projects
    
    def get_cabinets_by_project(self, project_name):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''SELECT cabinet_id, project_name, total_pages, annotated_pages,
                         total_punches, open_punches, implemented_punches, closed_punches, status
                         FROM cabinets WHERE project_name = ? ORDER BY last_updated DESC''',
                      (project_name,))
        cols = ['cabinet_id', 'project_name', 'total_pages', 'annotated_pages',
               'total_punches', 'open_punches', 'implemented_punches', 'closed_punches', 'status']
        cabinets = [dict(zip(cols, row)) for row in cursor.fetchall()]
        conn.close()
        return cabinets
    
    def get_category_stats(self, period='all', project_name=None):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        query = 'SELECT category, subcategory, COUNT(*) as count FROM category_occurrences WHERE 1=1'
        params = []
        
        if period == 'weekly':
            query += ' AND occurrence_date >= ?'
            params.append((datetime.now() - timedelta(days=7)).isoformat())
        elif period == 'monthly':
            query += ' AND occurrence_date >= ?'
            params.append((datetime.now() - timedelta(days=30)).isoformat())
        elif period == 'yearly':
            query += ' AND occurrence_date >= ?'
            params.append((datetime.now() - timedelta(days=365)).isoformat())
        
        if project_name:
            query += ' AND project_name = ?'
            params.append(project_name)
        
        query += ' GROUP BY category, subcategory ORDER BY count DESC'
        cursor.execute(query, params)
        stats = [{'category': r[0], 'subcategory': r[1], 'count': r[2]} for r in cursor.fetchall()]
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
        self.categories = self.load_categories()
        
        self.setup_ui()
        self.show_dashboard()
    
    def load_categories(self):
        try:
            if os.path.exists(self.category_file):
                with open(self.category_file, "r") as f:
                    return json.load(f)
        except: pass
        return []
    
    def save_categories(self):
        try:
            os.makedirs(os.path.dirname(self.category_file), exist_ok=True)
            with open(self.category_file, "w") as f:
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
            command=self.show_dashboard, bg='#3b82f6', fg='white', **btn_style)
        self.nav_btns['dashboard'].pack(side=tk.LEFT, padx=5)
        
        self.nav_btns['analytics'] = tk.Button(nav, text="📈 Analytics",
            command=self.show_analytics, bg='#334155', fg='white', **btn_style)
        self.nav_btns['analytics'].pack(side=tk.LEFT, padx=5)
        
        self.nav_btns['categories'] = tk.Button(nav, text="🏷️ Categories",
            command=self.show_categories, bg='#334155', fg='white', **btn_style)
        self.nav_btns['categories'].pack(side=tk.LEFT, padx=5)
        
        self.content = tk.Frame(self.root, bg='#f8fafc')
        self.content.pack(fill=tk.BOTH, expand=True)
    
    def set_active_nav(self, key):
        for k, btn in self.nav_btns.items():
            btn.config(bg='#3b82f6' if k == key else '#334155')
    
    def clear_content(self):
        for w in self.content.winfo_children():
            w.destroy()
    
    # ============ DASHBOARD ============
    def show_dashboard(self):
        self.set_active_nav('dashboard')
        self.clear_content()
        
        header = tk.Frame(self.content, bg='#f8fafc')
        header.pack(fill=tk.X, padx=30, pady=20)
        tk.Label(header, text="Projects Overview", font=('Segoe UI', 16, 'bold'),
                bg='#f8fafc').pack(side=tk.LEFT)
        
        projects = self.db.get_all_projects()
        
        if not projects:
            tk.Label(self.content, text="No projects found.\nAdd test data or sync from Quality tool.",
                    font=('Segoe UI', 12), fg='#64748b', bg='#f8fafc').pack(expand=True)
            tk.Button(self.content, text="Generate Sample Data", command=self.generate_sample_data,
                     bg='#3b82f6', fg='white', font=('Segoe UI', 10, 'bold'),
                     padx=20, pady=10).pack(pady=20)
            return
        
        canvas = tk.Canvas(self.content, bg='#f8fafc', highlightthickness=0)
        scrollbar = tk.Scrollbar(self.content, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg='#f8fafc')
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=30)
        scrollbar.pack(side="right", fill="y")
        
        for proj in projects:
            self.create_project_card(scroll_frame, proj)
    
    def create_project_card(self, parent, project):
        card = tk.Frame(parent, bg='white', relief=tk.SOLID, borderwidth=1)
        card.pack(fill=tk.X, pady=10)
        
        header = tk.Frame(card, bg='#eff6ff', cursor='hand2')
        header.pack(fill=tk.X)
        
        expand_var = tk.BooleanVar(value=False)
        indicator = tk.Label(header, text="▶", font=('Segoe UI', 12, 'bold'),
                           bg='#eff6ff', fg='#3b82f6', width=3)
        indicator.pack(side=tk.LEFT)
        
        tk.Label(header, text=project['project_name'], font=('Segoe UI', 13, 'bold'),
                bg='#eff6ff').pack(side=tk.LEFT, pady=15, padx=10)
        tk.Label(header, text=f"{project['cabinet_count']} Cabinet(s)",
                font=('Segoe UI', 10), bg='#eff6ff', fg='#64748b').pack(side=tk.LEFT)
        
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
        for text, w in [("Cabinet", 15), ("Draw%", 8), ("Total", 6), ("Open", 6), 
                        ("Impl", 6), ("Closed", 6), ("Status", 20)]:
            tk.Label(hdr, text=text, font=('Segoe UI', 9, 'bold'),
                    bg='#f1f5f9', width=w, anchor='w').pack(side=tk.LEFT, padx=3)
        
        # Rows
        for cab in cabinets:
            row = tk.Frame(parent, bg='white')
            row.pack(fill=tk.X, pady=2)
            
            tk.Label(row, text=cab['cabinet_id'], font=('Segoe UI', 9),
                    bg='white', width=15, anchor='w').pack(side=tk.LEFT, padx=3)
            
            pct = (cab['annotated_pages']/cab['total_pages']*100) if cab['total_pages'] else 0
            tk.Label(row, text=f"{pct:.0f}%", font=('Segoe UI', 9, 'bold'),
                    bg='white', fg='#10b981' if pct==100 else '#f59e0b',
                    width=8).pack(side=tk.LEFT, padx=3)
            
            for val in [cab['total_punches'], cab['open_punches'], 
                       cab['implemented_punches'], cab['closed_punches']]:
                tk.Label(row, text=str(val), font=('Segoe UI', 9),
                        bg='white', width=6).pack(side=tk.LEFT, padx=3)
            
            status_map = {
                'quality_inspection': ('🔍 Quality', '#3b82f6'),
                'handed_to_production': ('📦 Production', '#8b5cf6'),
                'in_rework': ('🔧 Rework', '#f59e0b'),
                'being_closed_by_quality': ('✅ Closing', '#10b981'),
                'closed': ('✓ Closed', '#64748b')
            }
            txt, color = status_map.get(cab['status'], (cab['status'], '#64748b'))
            tk.Label(row, text=txt, font=('Segoe UI', 9, 'bold'),
                    bg=color, fg='white', padx=8, pady=3).pack(side=tk.LEFT, padx=3)
    
    def generate_sample_data(self):
        """Generate sample data for testing"""
        conn = sqlite3.connect(self.db.db_path)
        cursor = conn.cursor()
        
        sample_data = [
            ('CAB-001', 'Solar Plant Alpha', 'SO-2024-001', 50, 48, 15, 3, 10, 2, 'quality_inspection'),
            ('CAB-002', 'Solar Plant Alpha', 'SO-2024-001', 45, 45, 20, 0, 15, 5, 'handed_to_production'),
            ('CAB-003', 'Wind Farm Beta', 'SO-2024-002', 60, 30, 25, 8, 12, 5, 'in_rework'),
            ('CAB-004', 'Hydro Project Gamma', 'SO-2024-003', 40, 40, 18, 2, 14, 2, 'being_closed_by_quality'),
        ]
        
        for data in sample_data:
            cursor.execute('''INSERT OR REPLACE INTO cabinets 
                (cabinet_id, project_name, sales_order_no, total_pages, annotated_pages,
                 total_punches, open_punches, implemented_punches, closed_punches, status,
                 created_date, last_updated) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                data + (datetime.now().isoformat(), datetime.now().isoformat()))
        
        # Sample category data
        categories = [
            ('CAB-001', 'Solar Plant Alpha', 'Wiring', 'Color Code'),
            ('CAB-001', 'Solar Plant Alpha', 'Wiring', 'Termination'),
            ('CAB-002', 'Solar Plant Alpha', 'Label', None),
            ('CAB-003', 'Wind Farm Beta', 'Hardware', 'Missing Bolts'),
        ]
        
        for data in categories:
            cursor.execute('''INSERT INTO category_occurrences 
                (cabinet_id, project_name, category, subcategory, occurrence_date)
                VALUES (?, ?, ?, ?, ?)''',
                data + (datetime.now().isoformat(),))
        
        conn.commit()
        conn.close()
        
        messagebox.showinfo("Success", "Sample data generated!")
        self.show_dashboard()
    
    # ============ ANALYTICS ============
    def show_analytics(self):
        self.set_active_nav('analytics')
        self.clear_content()
        
        header = tk.Frame(self.content, bg='#f8fafc')
        header.pack(fill=tk.X, padx=30, pady=20)
        tk.Label(header, text="Category Analytics", font=('Segoe UI', 16, 'bold'),
                bg='#f8fafc').pack(side=tk.LEFT)
        
        controls = tk.Frame(self.content, bg='white')
        controls.pack(fill=tk.X, padx=30, pady=10)
        
        tk.Label(controls, text="Period:", font=('Segoe UI', 10, 'bold'),
                bg='white').pack(side=tk.LEFT, padx=20, pady=15)
        
        period_var = tk.StringVar(value="all")
        for txt, val in [("All", "all"), ("Week", "weekly"), ("Month", "monthly"), ("Year", "yearly")]:
            tk.Radiobutton(controls, text=txt, variable=period_var, value=val,
                          bg='white', command=lambda: self.update_chart(
                              period_var.get(), project_var.get(), level_var.get())).pack(side=tk.LEFT, padx=3)
        
        tk.Label(controls, text=" | Project:", font=('Segoe UI', 10, 'bold'),
                bg='white').pack(side=tk.LEFT, padx=10)
        
        projects = ["All"] + [p['project_name'] for p in self.db.get_all_projects()]
        project_var = tk.StringVar(value="All")
        project_dd = ttk.Combobox(controls, textvariable=project_var, values=projects,
                                 state='readonly', width=20)
        project_dd.pack(side=tk.LEFT, padx=5)
        project_dd.bind('<<ComboboxSelected>>', lambda e: self.update_chart(
            period_var.get(), project_var.get(), level_var.get()))
        
        tk.Label(controls, text=" | Level:", font=('Segoe UI', 10, 'bold'),
                bg='white').pack(side=tk.LEFT, padx=10)
        
        level_var = tk.StringVar(value="category")
        tk.Radiobutton(controls, text="Category", variable=level_var, value="category",
                      bg='white', command=lambda: self.update_chart(
                          period_var.get(), project_var.get(), level_var.get())).pack(side=tk.LEFT, padx=3)
        tk.Radiobutton(controls, text="Subcategory", variable=level_var, value="subcategory",
                      bg='white', command=lambda: self.update_chart(
                          period_var.get(), project_var.get(), level_var.get())).pack(side=tk.LEFT, padx=3)
        
        tk.Button(controls, text="📥 Export Excel", command=lambda: self.export_excel(
                  period_var.get(), project_var.get()),
                 bg='#10b981', fg='white', font=('Segoe UI', 9, 'bold'),
                 padx=15, pady=8).pack(side=tk.RIGHT, padx=20)
        
        self.chart_frame = tk.Frame(self.content, bg='white')
        self.chart_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)
        
        self.period_var = period_var
        self.project_var = project_var
        self.level_var = level_var
        
        self.update_chart("all", "All", "category")
    
    def update_chart(self, period, project, level):
        for w in self.chart_frame.winfo_children():
            w.destroy()
        
        proj_filter = None if project == "All" else project
        stats = self.db.get_category_stats(period, proj_filter)
        
        if not stats:
            tk.Label(self.chart_frame, text="No data available",
                    font=('Segoe UI', 12), fg='#64748b', bg='white').pack(expand=True)
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
        
        fig = Figure(figsize=(12, 6), facecolor='white')
        ax1 = fig.add_subplot(111)
        ax2 = ax1.twinx()
        
        bars = ax1.bar(range(len(labels)), values, color='#3b82f6', alpha=0.7)
        line = ax2.plot(range(len(labels)), cumulative, color='#ef4444',
                       marker='o', linewidth=2, markersize=6)
        ax2.axhline(y=80, color='#10b981', linestyle='--', linewidth=1.5, alpha=0.7)
        
        ax1.set_xlabel('Category', fontsize=11, fontweight='bold')
        ax1.set_ylabel('Frequency', fontsize=11, fontweight='bold', color='#3b82f6')
        ax2.set_ylabel('Cumulative %', fontsize=11, fontweight='bold', color='#ef4444')
        ax1.set_title(f'Pareto Chart - {level.title()} Analysis', fontsize=14, fontweight='bold')
        
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
    
    def export_excel(self, period, project):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if not file:
            return
        
        proj_filter = None if project == "All" else project
        stats = self.db.get_category_stats(period, proj_filter)
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Category Analysis"
        
        ws['A1'] = "Category"
        ws['B1'] = "Subcategory"
        ws['C1'] = "Count"
        
        for cell in ['A1', 'B1', 'C1']:
            ws[cell].font = Font(bold=True)
            ws[cell].fill = PatternFill(start_color="366092", fill_type="solid")
            ws[cell].font = Font(color="FFFFFF", bold=True)
        
        row = 2
        for item in stats:
            ws[f'A{row}'] = item['category']
            ws[f'B{row}'] = item['subcategory'] or 'N/A'
            ws[f'C{row}'] = item['count']
            row += 1
        
        wb.save(file)
        messagebox.showinfo("Exported", f"Data exported to:\n{file}")
    
    # ============ CATEGORIES ============
    def show_categories(self):
        self.set_active_nav('categories')
        self.clear_content()
        
        header = tk.Frame(self.content, bg='#f8fafc')
        header.pack(fill=tk.X, padx=30, pady=20)
        tk.Label(header, text="Category Management", font=('Segoe UI', 16, 'bold'),
                bg='#f8fafc').pack(side=tk.LEFT)
        
        tk.Button(header, text="➕ Add Category", command=self.add_category,
                 bg='#10b981', fg='white', font=('Segoe UI', 10, 'bold'),
                 padx=15, pady=8).pack(side=tk.RIGHT)
        
        if not self.categories:
            tk.Label(self.content, text="No categories defined.\nClick 'Add Category' to create one.",
                    font=('Segoe UI', 12), fg='#64748b', bg='#f8fafc').pack(expand=True)
            return
        
        canvas = tk.Canvas(self.content, bg='#f8fafc', highlightthickness=0)
        scrollbar = tk.Scrollbar(self.content, command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg='#f8fafc')
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=30)
        scrollbar.pack(side="right", fill="y")
        
        for cat in self.categories:
            self.create_category_card(scroll_frame, cat)
    
    def create_category_card(self, parent, category):
        card = tk.Frame(parent, bg='white', relief=tk.SOLID, borderwidth=1)
        card.pack(fill=tk.X, pady=8)
        
        header = tk.Frame(card, bg='#dbeafe')
        header.pack(fill=tk.X)
        
        tk.Label(header, text=category['name'], font=('Segoe UI', 12, 'bold'),
                bg='#dbeafe', fg='#1e40af').pack(side=tk.LEFT, padx=15, pady=10)
        
        mode_text = "📝 Template" if category.get('mode') == 'template' else "📁 Parent"
        tk.Label(header, text=mode_text, font=('Segoe UI', 9),
                bg='#dbeafe', fg='#64748b').pack(side=tk.LEFT, padx=10)
        
        btn_frame = tk.Frame(header, bg='#dbeafe')
        btn_frame.pack(side=tk.RIGHT, padx=10)
        
        tk.Button(btn_frame, text="✏️", command=lambda: self.edit_category(category),
                 bg='#3b82f6', fg='white', width=3, relief=tk.FLAT).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="🗑️", command=lambda: self.delete_category(category),
                 bg='#ef4444', fg='white', width=3, relief=tk.FLAT).pack(side=tk.LEFT, padx=2)
        
        if category.get('mode') == 'parent' and category.get('subcategories'):
            sub_frame = tk.Frame(card, bg='white')
            sub_frame.pack(fill=tk.X, padx=20, pady=10)
            
            for sub in category['subcategories']:
                sub_row = tk.Frame(sub_frame, bg='#f8fafc')
                sub_row.pack(fill=tk.X, pady=2)
                
                tk.Label(sub_row, text=f"  ↳ {sub['name']}", font=('Segoe UI', 10),
                        bg='#f8fafc', anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=5)
    
    def add_category(self):
        name = simpledialog.askstring("New Category", "Enter category name:")
        if not name:
            return
        
        is_template = messagebox.askyesno("Category Type",
            "Template category (generates punches directly)?\n\n"
            "YES → Template\nNO → Parent (contains subcategories)")
        
        category = {'name': name.strip(), 'mode': 'template' if is_template else 'parent',
                   'inputs': [], 'template': None, 'subcategories': []}
        
        if is_template:
            template = simpledialog.askstring("Template", "Enter punch text template:")
            if template:
                category['template'] = template
        
        self.categories.append(category)
        self.save_categories()
        self.show_categories()
    
    def edit_category(self, category):
        new_name = simpledialog.askstring("Edit", "Enter new name:",
                                         initialvalue=category['name'])
        if new_name:
            category['name'] = new_name
            self.save_categories()
            self.show_categories()
    
    def delete_category(self, category):
        if messagebox.askyesno("Confirm", f"Delete '{category['name']}'?"):
            self.categories.remove(category)
            self.save_categories()
            self.show_categories()


def main():
    root = tk.Tk()
    app = ManagerUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
