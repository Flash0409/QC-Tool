import tkinter as tk
from tkinter import ttk, messagebox
import os, sys, json
from datetime import datetime
import subprocess

# ======================================================
# APP BASE DIR (Portable)
# ======================================================
def get_app_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_app_base_dir()

# JSON files moved into /assets for portability
CRED_FILE = os.path.join(BASE_DIR, "assets", "credentials.json")

# ======================================================
# CREDENTIAL HELPERS
# ======================================================
def load_credentials():
    if not os.path.exists(CRED_FILE):
        messagebox.showerror("Error", "credentials.json not found!")
        sys.exit(1)
    with open(CRED_FILE, "r") as f:
        return json.load(f)

def authenticate_user(username, password, role, credentials):
    cred = credentials.get(role)
    if not cred:
        return False
    return username.strip() == cred["username"] and password == cred["password"]

# ======================================================
# ROUTER (Outside class)
# ======================================================
def route_to_role(username, role, credentials):
    if role == "Quality":
        subprocess.Popen(["py", os.path.join("pages", "quality.py")])
    elif role == "Manager":
        subprocess.Popen(["py", os.path.join("pages", "manager.py")])
    elif role == "Production":
        subprocess.Popen(["py", os.path.join("pages", "production.py")])
    else:
        messagebox.showerror("Routing Error", f"Page for '{role}' not implemented yet!")

# ======================================================
# LOGIN UI
# ======================================================
class LoginPage:
    def __init__(self, root):
        self.root = root
        self.root.title("QC Tool Login")
        self.root.geometry("350x220")
        self.root.resizable(False, False)

        self.credentials = load_credentials()

        tk.Label(root, text="Username:").pack(pady=(15,2))
        self.user_entry = tk.Entry(root, width=30)
        self.user_entry.pack()

        tk.Label(root, text="Password:").pack(pady=2)
        self.pwd_entry = tk.Entry(root, width=30, show="*")
        self.pwd_entry.pack()

        tk.Label(root, text="Login as:").pack(pady=5)
        self.role_var = tk.StringVar(value="Manager")
        self.role_dropdown = ttk.Combobox(root, textvariable=self.role_var, values=list(self.credentials.keys()), state="readonly", width=20)
        self.role_dropdown.pack()

        tk.Button(root, text="Login", command=self.validate_login).pack(pady=10)

    def validate_login(self):
        username = self.user_entry.get()
        password = self.pwd_entry.get()
        role = self.role_var.get()

        if authenticate_user(username, password, role, self.credentials):
            messagebox.showinfo("Success", f"Logged in as {role}")
            self.root.destroy()
            # Route exactly once
            route_to_role(username, role, self.credentials)
        else:
            messagebox.showerror("Login Failed", "Incorrect username or password")

# ======================================================
# RUN APP
# ======================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = LoginPage(root)
    root.mainloop()
