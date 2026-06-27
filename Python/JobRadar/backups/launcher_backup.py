"""
JobRadar Desktop Launcher v2
Polished dashboard launcher for JobRadar.

Run with:
    python jobradar_launcher_v2.py
"""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKEND_SCRIPT = os.path.join(APP_DIR, "job_radar_v2_4.py")
REPORTS_DIR = os.path.join(APP_DIR, "reports")

MAIN_DASHBOARD = os.path.join(REPORTS_DIR, "job_radar.xlsx")
APPLY_NOW = os.path.join(REPORTS_DIR, "apply_now.csv")
RECRUITER_CRM = os.path.join(REPORTS_DIR, "recruiter_crm.xlsx")
JOB_INTELLIGENCE = os.path.join(REPORTS_DIR, "job_intelligence.xlsx")
SKILL_GAP = os.path.join(REPORTS_DIR, "skill_gap_analysis.xlsx")


def count_rows(path):
    if not os.path.exists(path):
        return 0

    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as file:
            lines = file.readlines()
        return max(len(lines) - 1, 0)
    except Exception:
        return 0


def update_dashboard_metrics():
    jobs_found = count_rows(os.path.join(REPORTS_DIR, "job_radar_results.csv"))
    apply_now = count_rows(APPLY_NOW)
    recruiter_leads = count_rows(os.path.join(REPORTS_DIR, "recruiter_leads.csv"))
    skill_gaps = count_rows(os.path.join(REPORTS_DIR, "skill_gap_analysis.csv"))

    jobs_value.config(text=str(jobs_found))
    apply_value.config(text=str(apply_now))
    recruiters_value.config(text=str(recruiter_leads))
    skills_value.config(text=str(skill_gaps))

    if os.path.exists(MAIN_DASHBOARD):
        modified_time = datetime.fromtimestamp(os.path.getmtime(MAIN_DASHBOARD))
        last_scan_value.config(text=modified_time.strftime("%b %d, %Y %I:%M %p"))
    else:
        last_scan_value.config(text="No scan yet")


def open_path(path):
    if not os.path.exists(path):
        messagebox.showwarning(
            "File not found",
            f"I could not find:\n\n{path}\n\nRun JobRadar first to generate this file."
        )
        return

    try:
        os.startfile(path)
    except Exception as exc:
        messagebox.showerror("Open failed", str(exc))


def open_reports_folder():
    os.makedirs(REPORTS_DIR, exist_ok=True)
    os.startfile(REPORTS_DIR)


def run_jobradar():
    if not os.path.exists(BACKEND_SCRIPT):
        messagebox.showerror(
            "Backend not found",
            "I could not find job_radar_v2_4.py.\n\nMake sure it is saved in the same folder as this launcher."
        )
        return

    status_value.config(text="Running scan...", fg="#1f5eff")
    run_button.config(state="disabled", text="Running...")
    root.update_idletasks()

    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        result = subprocess.run(
            [sys.executable, BACKEND_SCRIPT],
            cwd=APP_DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=300,
            env=env,
        )

        if result.returncode == 0:
            update_dashboard_metrics()
            status_value.config(text="Scan complete. Reports updated.", fg="#16803c")
            messagebox.showinfo(
                "JobRadar Complete",
                "Scan complete.\n\nYour reports and dashboard were updated."
            )
        else:
            error_text = result.stderr[-2500:] or result.stdout[-2500:]

            if "PermissionError" in error_text and ".xlsx" in error_text:
                friendly = (
                    "JobRadar could not update one of the Excel files because it is open.\n\n"
                    "Close all Excel workbooks in the reports folder, then run the scan again."
                )
                messagebox.showerror("Excel file is open", friendly)
            else:
                messagebox.showerror("JobRadar Error", error_text)

            status_value.config(text="Scan failed. Check message.", fg="#b00020")

    except subprocess.TimeoutExpired:
        status_value.config(text="Scan timed out.", fg="#b00020")
        messagebox.showerror(
            "Timeout",
            "JobRadar took longer than 5 minutes.\n\nTry running the backend from VS Code to see more details."
        )
    except Exception as exc:
        status_value.config(text="Scan failed.", fg="#b00020")
        messagebox.showerror("Error", str(exc))
    finally:
        run_button.config(state="normal", text="Run JobRadar Scan")


def make_metric(parent, label_text):
    frame = tk.Frame(parent, bg="#ffffff")
    frame.pack(side="left", expand=True, fill="both", padx=6)

    value = tk.Label(frame, text="0", font=("Segoe UI", 22, "bold"), bg="#ffffff", fg="#111111")
    value.pack(pady=(10, 0))

    label = tk.Label(frame, text=label_text, font=("Segoe UI", 9), bg="#ffffff", fg="#555555")
    label.pack(pady=(0, 10))

    return value


def make_button(parent, text, command):
    button = tk.Button(
        parent,
        text=text,
        command=command,
        font=("Segoe UI", 10),
        width=24,
        height=1,
        bg="#f3f3f3",
        activebackground="#e6e6e6",
        relief="groove"
    )
    button.pack(pady=5)
    return button


root = tk.Tk()
root.title("JobRadar")
root.geometry("560x800")
root.resizable(False, False)
root.configure(bg="#f5f6f8")

# Header
header_frame = tk.Frame(root, bg="#f5f6f8")
header_frame.pack(pady=(24, 8))

title = tk.Label(
    header_frame,
    text="JobRadar",
    font=("Segoe UI", 28, "bold"),
    bg="#f5f6f8",
    fg="#111111"
)
title.pack()

subtitle = tk.Label(
    header_frame,
    text="Career Command Center",
    font=("Segoe UI", 12),
    bg="#f5f6f8",
    fg="#555555"
)
subtitle.pack(pady=(2, 0))

# Dashboard card
card = tk.Frame(root, bg="#ffffff", bd=1, relief="solid")
card.pack(padx=24, pady=18, fill="x")

last_scan_label = tk.Label(
    card,
    text="Last Scan",
    font=("Segoe UI", 9, "bold"),
    bg="#ffffff",
    fg="#555555"
)
last_scan_label.pack(pady=(12, 0))

last_scan_value = tk.Label(
    card,
    text="No scan yet",
    font=("Segoe UI", 11),
    bg="#ffffff",
    fg="#111111"
)
last_scan_value.pack(pady=(0, 10))

metrics_row_1 = tk.Frame(card, bg="#ffffff")
metrics_row_1.pack(fill="x", padx=10, pady=(2, 4))

jobs_value = make_metric(metrics_row_1, "Jobs Found")
apply_value = make_metric(metrics_row_1, "Apply Today")

metrics_row_2 = tk.Frame(card, bg="#ffffff")
metrics_row_2.pack(fill="x", padx=10, pady=(2, 12))

recruiters_value = make_metric(metrics_row_2, "Recruiters")
skills_value = make_metric(metrics_row_2, "Skill Gaps")

# Action button
run_button = tk.Button(
    root,
    text="Run JobRadar Scan",
    command=run_jobradar,
    font=("Segoe UI", 13, "bold"),
    width=28,
    height=2,
    bg="#111111",
    fg="#ffffff",
    activebackground="#333333",
    activeforeground="#ffffff"
)
run_button.pack(pady=(4, 16))

# Report buttons
reports_frame = tk.Frame(root, bg="#f5f6f8")
reports_frame.pack()

make_button(reports_frame, "Open Main Dashboard", lambda: open_path(MAIN_DASHBOARD))
make_button(reports_frame, "Open Apply Now", lambda: open_path(APPLY_NOW))
make_button(reports_frame, "Open Recruiter CRM", lambda: open_path(RECRUITER_CRM))
make_button(reports_frame, "Open Job Intelligence", lambda: open_path(JOB_INTELLIGENCE))
make_button(reports_frame, "Open Skill Gap Analysis", lambda: open_path(SKILL_GAP))
make_button(reports_frame, "Open Reports Folder", open_reports_folder)

# Status
status_label = tk.Label(
    root,
    text="Status",
    font=("Segoe UI", 9, "bold"),
    bg="#f5f6f8",
    fg="#555555"
)
status_label.pack(pady=(10, 0))

status_value = tk.Label(
    root,
    text="Ready",
    font=("Segoe UI", 10),
    bg="#f5f6f8",
    fg="#555555"
)
status_value.pack(pady=(2, 4))

footer = tk.Label(
    root,
    text="JobRadar v0.9 Beta",
    font=("Segoe UI", 8),
    bg="#f5f6f8",
    fg="#777777"
)
footer.pack(pady=(2, 4))

exit_button = tk.Button(
    root,
    text="Exit",
    font=("Segoe UI", 9),
    width=14,
    command=root.destroy
)
exit_button.pack(pady=(0, 6))

update_dashboard_metrics()
root.mainloop()