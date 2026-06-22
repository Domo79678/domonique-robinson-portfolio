# CareerCopilot

CareerCopilot is a Python desktop application that helps job seekers compare a resume against a job description, calculate an ATS-style match score, identify missing keywords, and generate resume improvement suggestions.

## Why I Built This

I built CareerCopilot to support my job search while strengthening my Python, automation, and business analysis skills. The goal was to reduce the manual work of reviewing job descriptions and make resume tailoring more focused, repeatable, and data-informed.

## Features

* Upload resume files in `.txt`, `.docx`, or `.pdf` format
* Upload job descriptions in `.txt`, `.docx`, or `.pdf` format
* Calculate an ATS-style match score
* Display a visual score gauge
* Identify matched keywords and phrases
* Identify missing keywords and phrases
* Prioritize missing skills by importance
* Generate resume improvement suggestions
* Copy missing keywords to clipboard
* Copy the full report to clipboard
* Export reports as `.txt`
* Export reports as `.pdf`
* Open the reports folder directly from the app
* Package the app as a standalone Windows `.exe`

## Tools Used

* Python
* Tkinter
* python-docx
* PyPDF2
* ReportLab
* PyInstaller
* VS Code
* Git/GitHub

## Skills Demonstrated

* Python scripting
* GUI application development
* File handling
* PDF and Word document parsing
* Text analysis
* Keyword matching
* Report generation
* Desktop app packaging
* Career workflow automation
* Business problem solving

## Screenshots

Add screenshots here:

```markdown
![CareerCopilot Home](screenshots/screenshot-1.png)
![CareerCopilot Analysis](screenshots/screenshot-2.png)
![CareerCopilot Report](screenshots/screenshot-3.png)
```

## How It Works

1. Load a resume.
2. Load a job description.
3. Click **Analyze**.
4. CareerCopilot compares the resume to the job description.
5. The app generates a match score, missing skills, strengths, and improvement suggestions.
6. The report can be copied, exported, or saved as a PDF.

## Project Status

Version 1 is complete. Future enhancements may include:

* JobRadar companion app for job alert monitoring
* Gmail job alert parsing
* AI-assisted resume bullet suggestions
* STAR interview story generator
* Cover letter generator
* Job tracker integration
* More advanced scoring logic
* Dark mode UI

## Future Companion App: JobRadar

JobRadar will help collect job alerts from email, remove duplicates, organize roles into one tracker, and send high-fit job descriptions into CareerCopilot for scoring.

## Resume Bullet

Built CareerCopilot, a Python desktop application that analyzes resumes against job descriptions, supports TXT/DOCX/PDF uploads, calculates ATS-style match scores, identifies missing keywords, generates resume improvement suggestions, and exports TXT/PDF reports.
