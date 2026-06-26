"""
JobRadar Desktop Launcher

Save this file inside your Python/JobRadar folder next to your main JobRadar script.
Recommended backend file name: job_radar_v2_4.py

Run with:
    python jobradar_launcher.py
"""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

APP_DIR = os.path.dirname(os.path.abspath(__file__))
BACKEND_SCRIPT = os.path.join(APP_DIR, "job_radar_v2_4.py")
REPORTS_DIR = os.path.join(APP_DIR, "reports")

REPORT_PATHS = {
    "Main Dashboard": os.path.join(REPORTS_DIR, "job_radar.xlsx"),
    "Apply Now": os.path.join(REPORTS_DIR, "apply_now.csv"),
    "Recruiter CRM": os.path.join(REPORTS_DIR, "recruiter_crm.xlsx"),
    "Job Intelligence": os.path.join(REPORTS_DIR, "job_intelligence.xlsx"),
    "Skill Gap Analysis": os.path.join(REPORTS_DIR, "skill_gap_analysis.xlsx"),
}


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


def run_jobradar():
    if not os.path.exists(BACKEND_SCRIPT):
        messagebox.showerror(
            "Backend not found",
            "I could not find job_radar_v2_4.py.\n\nSave your latest working JobRadar script with that exact name in this folder."
        )
        return

    status_label.config(text="Running JobRadar... please wait")
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
            status_label.config(text="Scan complete. Reports updated.")
            messagebox.showinfo(
                "JobRadar Complete",
                "JobRadar finished running. Your reports were updated."
            )
        else:
            status_label.config(text="Scan failed. Check error message.")
            messagebox.showerror(
                "JobRadar Error",
                result.stderr[-3000:] or result.stdout[-3000:]
            )

    except subprocess.TimeoutExpired:
        status_label.config(text="Scan timed out.")
        messagebox.showerror(
            "Timeout",
            "JobRadar took longer than 5 minutes. Try running it from VS Code to see details."
        )
    except Exception as exc:
        status_label.config(text="Scan failed.")
        messagebox.showerror("Error", str(exc))


def open_reports_folder():
    os.makedirs(REPORTS_DIR, exist_ok=True)
    os.startfile(REPORTS_DIR)


root = tk.Tk()
root.title("JobRadar")
root.geometry("420x520")
root.resizable(False, False)

header = tk.Label(root, text="JobRadar", font=("Segoe UI", 24, "bold"))
header.pack(pady=(24, 6))

subtitle = tk.Label(root, text="Career command center", font=("Segoe UI", 11))
subtitle.pack(pady=(0, 18))

run_button = tk.Button(
    root,
    text="Run JobRadar Scan",
    font=("Segoe UI", 12, "bold"),
    height=2,
    width=28,
    command=run_jobradar
)
run_button.pack(pady=8)

for label, path in REPORT_PATHS.items():
    button = tk.Button(
        root,
        text=f"Open {label}",
        font=("Segoe UI", 10),
        width=28,
        command=lambda p=path: open_path(p)
    )
    button.pack(pady=5)

folder_button = tk.Button(
    root,
    text="Open Reports Folder",
    font=("Segoe UI", 10),
    width=28,
    command=open_reports_folder
)
folder_button.pack(pady=12)

status_label = tk.Label(root, text="Ready", font=("Segoe UI", 10), fg="gray")
status_label.pack(pady=(12, 6))

exit_button = tk.Button(root, text="Exit", font=("Segoe UI", 10), width=16, command=root.destroy)
exit_button.pack(pady=8)

root.mainloop()