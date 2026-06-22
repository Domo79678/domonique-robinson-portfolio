from pathlib import Path
import re
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from docx import Document
from PyPDF2 import PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

PROJECT_PATH = Path(__file__).parent
REPORT_FILE = PROJECT_PATH / "reports" / "career_report.txt"

STOP_WORDS = {
    "we", "are", "a", "an", "the", "and", "or", "with", "in", "to",
    "of", "for", "on", "is", "be", "as", "this", "that", "you",
    "your", "our", "seeking", "experience", "experienced", "role",
    "candidate", "will", "have", "has", "from", "by", "at", "responsible"
}

IMPORTANT_KEYWORDS = {
    "salesforce", "administrator", "admin", "crm", "flow", "automation",
    "reports", "dashboards", "tableau", "excel", "sql", "soql",
    "business", "analyst", "stakeholder", "communication", "forecasting",
    "pipeline", "revops", "revenue", "operations", "data", "analytics",
    "process", "requirements", "documentation", "testing", "training",
    "integration", "workflow", "optimization", "reporting"
}

IMPORTANT_PHRASES = [
    "salesforce administrator", "business analyst", "stakeholder communication",
    "revenue operations", "sales operations", "pipeline management",
    "flow automation", "salesforce flow", "crm reporting", "data analysis",
    "business requirements", "process improvement", "executive dashboard",
    "user acceptance testing", "data quality", "salesforce reports",
    "salesforce dashboards", "workflow automation"
]

resume_path = None
job_path = None
latest_report = ""
latest_missing_keywords = []
latest_missing_phrases = []
latest_score = 0


def read_file(path):
    path = Path(path)
    file_type = path.suffix.lower()

    if file_type == ".txt":
        return path.read_text(encoding="utf-8")

    if file_type == ".docx":
        document = Document(path)
        return "\n".join(paragraph.text for paragraph in document.paragraphs)

    if file_type == ".pdf":
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return text

    messagebox.showerror("Unsupported File", "Please upload a .txt, .docx, or .pdf file.")
    return ""


def extract_words(text):
    words = set(re.findall(r"\b[a-zA-Z0-9+#.-]+\b", text.lower()))
    return words - STOP_WORDS


def find_phrases(text):
    text = text.lower()
    return sorted([phrase for phrase in IMPORTANT_PHRASES if phrase in text])


def get_rating(score):
    if score >= 90:
        return "Excellent Match", "green"
    if score >= 75:
        return "Good Match", "orange"
    if score >= 50:
        return "Fair Match", "dark orange"
    return "Needs Work", "red"


def get_priority_items(missing_keywords, missing_phrases):
    high_priority_terms = {
        "salesforce", "salesforce administrator", "flow automation",
        "salesforce flow", "business analyst", "stakeholder communication",
        "business requirements", "crm reporting", "data analysis",
        "reports", "dashboards", "sql", "soql"
    }

    missing_items = missing_phrases + missing_keywords
    high_priority = sorted([item for item in missing_items if item in high_priority_terms])
    lower_priority = sorted([item for item in missing_items if item not in high_priority_terms])

    return high_priority, lower_priority


def build_resume_suggestions(missing_keywords, missing_phrases):
    missing_items = missing_phrases + missing_keywords
    suggestions = []

    suggestions.append("Generated Resume Bullet Suggestions")
    suggestions.append("-" * 35)
    suggestions.append("Only use these if they truthfully match your experience.")
    suggestions.append("")

    if "flow automation" in missing_items or "salesforce flow" in missing_items or "automation" in missing_items:
        suggestions.append("• Built Salesforce Flow automations to reduce manual work and improve CRM process consistency.")

    if "stakeholder communication" in missing_items or "stakeholder" in missing_items:
        suggestions.append("• Partnered with stakeholders to gather requirements, clarify business needs, and support solution design.")

    if "forecasting" in missing_items or "pipeline" in missing_items:
        suggestions.append("• Supported sales pipeline visibility through CRM reporting, dashboards, and forecasting-related analysis.")

    if "sql" in missing_items or "soql" in missing_items:
        suggestions.append("• Used SQL/SOQL queries to retrieve, validate, and analyze CRM data for business reporting.")

    if "reports" in missing_items or "dashboards" in missing_items or "crm reporting" in missing_items:
        suggestions.append("• Created reports and dashboards to monitor performance, identify trends, and support decision-making.")

    if len(suggestions) <= 4:
        suggestions.append("• Strengthened CRM documentation, reporting, and process improvement support across business workflows.")

    return "\n".join(suggestions)


def build_bar_chart(score):
    filled = round(score / 10)
    empty = 10 - filled
    return "█" * filled + "░" * empty


def build_report(score, rating, matched_phrases, missing_phrases, matched_keywords, missing_keywords, resume_words, job_words):
    high_priority, lower_priority = get_priority_items(missing_keywords, missing_phrases)
    top_missing = (high_priority + lower_priority)[:5]

    estimated_score = min(score + (len(top_missing) * 4), 100)

    report = []
    report.append("=" * 60)
    report.append("CAREERCOPILOT REPORT")
    report.append("=" * 60)
    report.append(f"Run Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append("")
    report.append("Recruiter Summary")
    report.append("-" * 30)
    report.append(f"Overall Match: {score}%")
    report.append(f"Rating: {rating}")
    report.append(f"ATS Visual: {build_bar_chart(score)}")
    report.append("")
    report.append("Estimated Improvement")
    report.append("-" * 30)
    report.append(f"Current Score: {score}%")
    report.append(f"Estimated Score If Legitimate Missing Items Are Added: {estimated_score}%")
    report.append("")
    report.append("Top Missing Items")
    report.append("-" * 30)

    if top_missing:
        for item in top_missing:
            report.append(f"✘ {item}")
    else:
        report.append("None found.")

    report.append("")
    report.append("ATS Strengths")
    report.append("-" * 30)

    for item in matched_phrases + matched_keywords:
        report.append(f"✔ {item}")

    if not matched_phrases and not matched_keywords:
        report.append("None found.")

    report.append("")
    report.append("Priority Missing Items")
    report.append("-" * 30)
    report.append("HIGH PRIORITY")

    if high_priority:
        for item in high_priority:
            report.append(f"✘ {item}")
    else:
        report.append("None found.")

    report.append("")
    report.append("LOWER PRIORITY")

    if lower_priority:
        for item in lower_priority:
            report.append(f"✘ {item}")
    else:
        report.append("None found.")

    report.append("")
    report.append("Quick Stats")
    report.append("-" * 30)
    report.append(f"Resume keyword count: {len(resume_words)}")
    report.append(f"Job description keyword count: {len(job_words)}")
    report.append(f"Matched keywords: {len(matched_keywords)}")
    report.append(f"Missing keywords: {len(missing_keywords)}")
    report.append(f"Matched phrases: {len(matched_phrases)}")
    report.append(f"Missing phrases: {len(missing_phrases)}")
    report.append("")
    report.append("ATS Score Scale")
    report.append("-" * 30)
    report.append("90-100   Excellent Match")
    report.append("75-89    Strong Match")
    report.append("50-74    Moderate Match")
    report.append("0-49     Needs Work")
    report.append("")
    report.append("Matched Phrases")
    report.append("-" * 30)
    report.extend([f"✔ {item}" for item in matched_phrases] or ["None found."])
    report.append("")
    report.append("Missing Phrases")
    report.append("-" * 30)
    report.extend([f"✘ {item}" for item in missing_phrases] or ["None found."])
    report.append("")
    report.append("Matched Keywords")
    report.append("-" * 30)
    report.extend([f"✔ {item}" for item in matched_keywords] or ["None found."])
    report.append("")
    report.append("Missing Keywords")
    report.append("-" * 30)
    report.extend([f"✘ {item}" for item in missing_keywords] or ["None found."])
    report.append("")
    report.append("Recommendation")
    report.append("-" * 30)
    report.append("Only add missing keywords or phrases if they truthfully match your experience.")
    report.append("")
    report.append(build_resume_suggestions(missing_keywords, missing_phrases))

    return "\n".join(report)


def draw_donut(score, color):
    score_canvas.delete("all")
    score_canvas.create_oval(10, 10, 150, 150, outline="#ddd", width=14)

    extent = int((score / 100) * 359)
    score_canvas.create_arc(10, 10, 150, 150, start=90, extent=-extent, outline=color, width=14, style="arc")
    score_canvas.create_text(80, 72, text=f"{score}%", font=("Arial", 22, "bold"), fill=color)
    score_canvas.create_text(80, 98, text="ATS Match", font=("Arial", 9), fill="black")


def save_pdf_report():
    if not latest_report:
        messagebox.showwarning("No Report", "Please analyze a resume first.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")],
        initialfile="career_report.pdf",
        title="Save PDF Report As"
    )

    if not save_path:
        return

    pdf = canvas.Canvas(save_path, pagesize=letter)
    width, height = letter
    y = height - 40

    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(40, y, "CareerCopilot Report")
    y -= 30

    pdf.setFont("Helvetica", 9)

    for line in latest_report.split("\n"):
        if y < 40:
            pdf.showPage()
            pdf.setFont("Helvetica", 9)
            y = height - 40

        clean_line = line.replace("✔", "[MATCH]").replace("✘", "[MISSING]").replace("█", "#").replace("░", "-")
        pdf.drawString(40, y, clean_line[:100])
        y -= 13

    pdf.save()
    messagebox.showinfo("PDF Exported", f"PDF report saved to:\n{save_path}")


def export_txt_report():
    if not latest_report:
        messagebox.showwarning("No Report", "Please analyze a resume first.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")],
        initialfile="career_report.txt",
        title="Save TXT Report As"
    )

    if not save_path:
        return

    Path(save_path).write_text(latest_report, encoding="utf-8")
    messagebox.showinfo("TXT Exported", f"TXT report saved to:\n{save_path}")


def copy_report():
    if not latest_report:
        messagebox.showwarning("No Report", "Please analyze a resume first.")
        return

    app.clipboard_clear()
    app.clipboard_append(latest_report)
    messagebox.showinfo("Copied", "Full report copied to clipboard.")


def open_reports_folder():
    reports_path = PROJECT_PATH / "reports"
    reports_path.mkdir(exist_ok=True)
    os.startfile(reports_path)


def analyze():
    global resume_path, job_path, latest_missing_keywords, latest_missing_phrases, latest_report, latest_score

    if not resume_path or not job_path:
        messagebox.showwarning("Missing Files", "Please load both a resume and a job description.")
        return

    status_label.config(text="Status: Reading files...")
    progress_bar["value"] = 20
    app.update_idletasks()

    resume_text = read_file(resume_path)
    job_text = read_file(job_path)

    status_label.config(text="Status: Extracting keywords...")
    progress_bar["value"] = 45
    app.update_idletasks()

    resume_words = extract_words(resume_text)
    job_words = extract_words(job_text)

    job_keywords = sorted(job_words & IMPORTANT_KEYWORDS)
    resume_keywords = sorted(resume_words & IMPORTANT_KEYWORDS)

    matched_keywords = sorted(set(job_keywords) & set(resume_keywords))
    missing_keywords = sorted(set(job_keywords) - set(resume_keywords))

    status_label.config(text="Status: Matching phrases...")
    progress_bar["value"] = 70
    app.update_idletasks()

    job_phrases = find_phrases(job_text)
    resume_phrases = find_phrases(resume_text)

    matched_phrases = sorted(set(job_phrases) & set(resume_phrases))
    missing_phrases = sorted(set(job_phrases) - set(resume_phrases))

    latest_missing_keywords = missing_keywords
    latest_missing_phrases = missing_phrases

    total_possible = len(job_keywords) + len(job_phrases)
    total_matched = len(matched_keywords) + len(matched_phrases)
    score = round((total_matched / total_possible) * 100) if total_possible else 0
    latest_score = score

    rating, color = get_rating(score)

    final_report = build_report(
        score,
        rating,
        matched_phrases,
        missing_phrases,
        matched_keywords,
        missing_keywords,
        resume_words,
        job_words,
    )

    latest_report = final_report

    output_box.delete("1.0", tk.END)
    output_box.insert(tk.END, final_report)

    REPORT_FILE.parent.mkdir(exist_ok=True)
    REPORT_FILE.write_text(final_report, encoding="utf-8")

    score_label.config(text=f"{score}%", fg=color)
    rating_label.config(text=rating, fg=color)
    draw_donut(score, color)

    status_label.config(text="Status: Complete")
    progress_bar["value"] = 100

    messagebox.showinfo("Report Saved", f"TXT report automatically saved to:\n{REPORT_FILE}")


def load_resume():
    global resume_path
    resume_path = filedialog.askopenfilename(
        filetypes=[
            ("Supported Files", "*.txt *.docx *.pdf"),
            ("Text Files", "*.txt"),
            ("Word Documents", "*.docx"),
            ("PDF Files", "*.pdf"),
        ]
    )
    if resume_path:
        resume_label.config(text=f"Resume: {Path(resume_path).name}")


def load_job():
    global job_path
    job_path = filedialog.askopenfilename(
        filetypes=[
            ("Supported Files", "*.txt *.docx *.pdf"),
            ("Text Files", "*.txt"),
            ("Word Documents", "*.docx"),
            ("PDF Files", "*.pdf"),
        ]
    )
    if job_path:
        job_label.config(text=f"Job Description: {Path(job_path).name}")


def copy_missing_keywords():
    missing_items = latest_missing_phrases + latest_missing_keywords

    if not missing_items:
        messagebox.showinfo("Nothing to Copy", "No missing keywords or phrases found yet.")
        return

    copied_text = "\n".join(missing_items)
    app.clipboard_clear()
    app.clipboard_append(copied_text)
    messagebox.showinfo("Copied", "Missing keywords and phrases copied to clipboard.")


def improve_resume():
    suggestions = build_resume_suggestions(latest_missing_keywords, latest_missing_phrases)
    output_box.insert(tk.END, "\n\n" + "=" * 60 + "\n")
    output_box.insert(tk.END, suggestions)


def clear_output():
    global latest_report, latest_missing_keywords, latest_missing_phrases, latest_score

    output_box.delete("1.0", tk.END)
    score_label.config(text="--%", fg="black")
    rating_label.config(text="Not analyzed yet", fg="black")
    status_label.config(text="Status: Ready")
    progress_bar["value"] = 0
    latest_report = ""
    latest_missing_keywords = []
    latest_missing_phrases = []
    latest_score = 0
    score_canvas.delete("all")
    score_canvas.create_oval(10, 10, 150, 150, outline="#ddd", width=14)
    score_canvas.create_text(80, 80, text="--%", font=("Arial", 22, "bold"), fill="black")


app = tk.Tk()
app.title("CareerCopilot")
app.geometry("1220x920")

title = tk.Label(app, text="CareerCopilot", font=("Arial", 28, "bold"))
title.pack(pady=4)

subtitle = tk.Label(app, text="Resume + Job Description Match Analyzer", font=("Arial", 13))
subtitle.pack(pady=1)

top_frame = tk.Frame(app)
top_frame.pack(pady=4)

score_canvas = tk.Canvas(top_frame, width=160, height=160, highlightthickness=0)
score_canvas.grid(row=0, column=0, padx=20)

score_canvas.create_oval(10, 10, 150, 150, outline="#ddd", width=14)
score_canvas.create_text(80, 80, text="--%", font=("Arial", 22, "bold"), fill="black")

score_info_frame = tk.Frame(top_frame)
score_info_frame.grid(row=0, column=1, padx=20)

score_label = tk.Label(score_info_frame, text="--%", font=("Arial", 40, "bold"))
score_label.pack(pady=2)

rating_label = tk.Label(score_info_frame, text="Not analyzed yet", font=("Arial", 16, "bold"))
rating_label.pack(pady=2)

scale_label = tk.Label(score_info_frame, text="90+ Excellent | 75+ Strong | 50+ Fair | Below 50 Needs Work", font=("Arial", 10))
scale_label.pack(pady=2)

button_frame = tk.Frame(app)
button_frame.pack(pady=8)

tk.Button(button_frame, text="📄 Load Resume", width=17, command=load_resume).grid(row=0, column=0, padx=3, pady=4)
tk.Button(button_frame, text="💼 Load Job Description", width=23, command=load_job).grid(row=0, column=1, padx=3, pady=4)
tk.Button(button_frame, text="🚀 Analyze", width=13, command=analyze).grid(row=0, column=2, padx=3, pady=4)
tk.Button(button_frame, text="📋 Copy Missing", width=16, command=copy_missing_keywords).grid(row=0, column=3, padx=3, pady=4)
tk.Button(button_frame, text="📋 Copy Report", width=15, command=copy_report).grid(row=0, column=4, padx=3, pady=4)
tk.Button(button_frame, text="✨ Improve Resume", width=17, command=improve_resume).grid(row=0, column=5, padx=3, pady=4)
tk.Button(button_frame, text="📄 Export PDF", width=14, command=save_pdf_report).grid(row=0, column=6, padx=3, pady=4)
tk.Button(button_frame, text="📝 Export TXT", width=14, command=export_txt_report).grid(row=0, column=7, padx=3, pady=4)
tk.Button(button_frame, text="📂 Reports", width=12, command=open_reports_folder).grid(row=0, column=8, padx=3, pady=4)
tk.Button(button_frame, text="🧹 Clear", width=10, command=clear_output).grid(row=0, column=9, padx=3, pady=4)

resume_label = tk.Label(app, text="Resume: Not loaded")
resume_label.pack()

job_label = tk.Label(app, text="Job Description: Not loaded")
job_label.pack()

status_label = tk.Label(app, text="Status: Ready", font=("Arial", 10, "bold"))
status_label.pack(pady=4)

progress_bar = ttk.Progressbar(app, orient="horizontal", length=760, mode="determinate")
progress_bar.pack(pady=4)

output_box = scrolledtext.ScrolledText(app, width=140, height=28)
output_box.pack(pady=8)

app.mainloop()