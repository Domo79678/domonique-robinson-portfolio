from pathlib import Path

portfolio_path = Path.cwd()

project_areas = [
    "Salesforce",
    "Tableau",
    "SOQL",
    "SQL",
    "Excel",
    "Python",
    "AI",
]

print("=" * 60)
print("              PORTFOLIO AUDITOR v2")
print("=" * 60)

total_projects = 0
projects_with_readme = 0
recommendations = []

for area in project_areas:
    area_path = portfolio_path / area

    print(f"\n📁 {area}")
    print("-" * 60)

    if not area_path.exists():
        print("❌ Area folder missing")
        recommendations.append(f"Create the {area} folder.")
        continue

    project_folders = [item for item in area_path.iterdir() if item.is_dir()]

    if not project_folders:
        print("⚠️ No project folders found yet")
        recommendations.append(f"Add at least one project folder inside {area}.")
        continue

    for project in sorted(project_folders):
        total_projects += 1
        readme_path = project / "README.md"

        if readme_path.exists():
            projects_with_readme += 1
            print(f"✅ {project.name} — README found")
        else:
            print(f"⚠️ {project.name} — README missing")
            recommendations.append(f"Add README.md to {area}/{project.name}.")

print("\n" + "=" * 60)
print("SUMMARY")
print("=" * 60)

print(f"Projects found: {total_projects}")
print(f"Projects with README: {projects_with_readme}")

if total_projects > 0:
    score = round((projects_with_readme / total_projects) * 100)
else:
    score = 0

print(f"Portfolio Documentation Score: {score}%")

if recommendations:
    print("\nRecommendations:")
    for recommendation in recommendations:
        print(f"- {recommendation}")
else:
    print("\nExcellent! All detected projects have README files.")

print("\nAudit Complete ✅")