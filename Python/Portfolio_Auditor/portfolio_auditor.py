from pathlib import Path
from datetime import datetime

PORTFOLIO_ROOT = Path.cwd()

PROJECT_AREAS = [
    "Salesforce",
    "Tableau",
    "SOQL",
    "SQL",
    "Excel",
    "Python",
    "AI",
]

PROJECT_FILE_TYPES = [".py", ".sql", ".twbx", ".xlsx", ".md", ".txt"]
SCREENSHOT_TYPES = [".png", ".jpg", ".jpeg"]

report_lines = []
recommendations = []

total_projects = 0
readme_count = 0
screenshot_count = 0
project_file_count = 0


def add_line(text=""):
    print(text)
    report_lines.append(text)


def has_file_type(folder, extensions):
    for file in folder.rglob("*"):
        if file.is_file() and file.suffix.lower() in extensions:
            return True
    return False


def audit_project(area, project):
    global total_projects, readme_count, screenshot_count, project_file_count

    total_projects += 1

    readme_found = (project / "README.md").exists()
    screenshot_found = has_file_type(project, SCREENSHOT_TYPES)
    project_file_found = has_file_type(project, PROJECT_FILE_TYPES)

    add_line(f"\n🔹 {area}/{project.name}")

    if readme_found:
        readme_count += 1
        add_line("   ✅ README.md found")
    else:
        add_line("   ⚠️ README.md missing")
        recommendations.append(f"Add README.md to {area}/{project.name}")

    if screenshot_found:
        screenshot_count += 1
        add_line("   ✅ Screenshot/image found")
    else:
        add_line("   ⚠️ Screenshot/image missing")
        recommendations.append(f"Add a screenshot to {area}/{project.name}")

    if project_file_found:
        project_file_count += 1
        add_line("   ✅ Project file found")
    else:
        add_line("   ⚠️ Project file missing")
        recommendations.append(f"Add a project file to {area}/{project.name}")


def audit_area(area):
    area_path = PORTFOLIO_ROOT / area

    add_line(f"\n📁 {area}")
    add_line("-" * 60)

    if not area_path.exists():
        add_line("❌ Area folder missing")
        recommendations.append(f"Create the {area} folder")
        return

    project_folders = [item for item in area_path.iterdir() if item.is_dir()]

    if not project_folders:
        add_line("⚠️ No project folders found yet")
        recommendations.append(f"Add at least one project folder inside {area}")
        return

    for project in sorted(project_folders):
        audit_project(area, project)


def calculate_score():
    possible_points = total_projects * 3

    if possible_points == 0:
        return 0

    earned_points = readme_count + screenshot_count + project_file_count
    return round((earned_points / possible_points) * 100)


def save_report():
    reports_folder = PORTFOLIO_ROOT / "Python" / "Portfolio_Auditor" / "reports"
    reports_folder.mkdir(exist_ok=True)

    report_path = reports_folder / "portfolio_report.txt"

    with open(report_path, "w", encoding="utf-8") as file:
        file.write("\n".join(report_lines))

    add_line(f"\nReport saved to: {report_path}")


def run_audit():
    add_line("=" * 60)
    add_line("              PORTFOLIO AUDITOR v2")
    add_line("=" * 60)
    add_line(f"Run Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    for area in PROJECT_AREAS:
        audit_area(area)

    score = calculate_score()

    add_line("\n" + "=" * 60)
    add_line("SUMMARY")
    add_line("=" * 60)
    add_line(f"Projects found: {total_projects}")
    add_line(f"Projects with README: {readme_count}")
    add_line(f"Projects with screenshots/images: {screenshot_count}")
    add_line(f"Projects with project files: {project_file_count}")
    add_line(f"Portfolio Health Score: {score}%")

    if recommendations:
        add_line("\nRecommendations:")
        for recommendation in recommendations:
            add_line(f"- {recommendation}")
    else:
        add_line("\nExcellent! No missing items found.")

    add_line("\nAudit Complete ✅")
    save_report()


run_audit()