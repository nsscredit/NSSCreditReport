from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import re
import os
import openpyxl
import glob
import shutil
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter

# Ensure output directory exists
os.makedirs("reports", exist_ok=True)


# Load and clean data from Excel file
def load_and_clean_data(file_path, year):
    file_path = os.path.join("/tmp/uploads", file_path)
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip().str.lower()

    # Rename columns to match expected format
    df.rename(columns={
        's.no': 'S.No',
        'name': 'Name',
        'phone': 'Phone',
        'phone number': 'Phone',
        'contact number': 'Phone',
        'roll no': 'Roll No',
        'roll number': 'Roll No',
        'smail': 'Smail',
        'mail id': 'Smail',
        'total credits earned': 'Credits',
        'total credits': 'Credits',
        'total pending credits': 'Pending Credits',
        'mail sent or not': 'Mail Sent',
        'pass/pending': 'Status'
    }, inplace=True)

    # Check required columns exist
    required_cols = ["Name", "Roll No", "Credits", "Status"]
    if not all(col in df.columns for col in required_cols):
        raise ValueError(f"Missing required columns. Found: {list(df.columns)}")

    # Convert Credits to numeric
    df["Credits"] = pd.to_numeric(df["Credits"], errors="coerce").fillna(0).astype(int)

    # Add Year Column
    df["Year"] = year
    return df


# Process credit data
def process_credit_data(df):
    yearwise_stats = {}

    for year, group in df.groupby("Year"):
        total_volunteers = len(group)
        passed_volunteers = len(group[group["Credits"] >= 55])
        pending_volunteers = len(group[group["Credits"] < 55])
        near_pass_volunteers = len(group[(group["Credits"] >= 40) & (group["Credits"] < 55)])
        active_volunteers = len(group[(group["Credits"] >= 10) & (group["Credits"] < 40)])
        inactive_volunteers = len(group[group["Credits"] < 10])
        average_credits = group["Credits"].mean()
        max_credits = group["Credits"].max()
        min_credits = group["Credits"].min()

        yearwise_stats[year] = {
            "total_volunteers": total_volunteers,
            "passed_volunteers": passed_volunteers,
            "pending_volunteers": pending_volunteers,
            "near_pass_volunteers": near_pass_volunteers,
            "active_volunteers": active_volunteers,
            "inactive_volunteers": inactive_volunteers,
            "average_credits": average_credits,
            "max_credits": max_credits,
            "min_credits": min_credits,
            "pass_percentage": (passed_volunteers / total_volunteers) * 100 if total_volunteers > 0 else 0
        }

    return yearwise_stats


# Generate pie charts
def generate_pie_charts(yearwise_stats):
    pie_chart_paths = {}

    for year, stats in yearwise_stats.items():
        labels = ["Passed", "Pending"]
        sizes = [stats["passed_volunteers"], stats["pending_volunteers"]]
        colors = ["green", "red"]

        plt.figure(figsize=(5, 5))
        plt.pie(sizes, labels=labels, colors=colors, autopct="%1.1f%%", startangle=140)
        plt.title(f"Volunteer Pass Distribution ({year})")

        chart_path = f"reports/pie_chart_{year}.png"
        plt.savefig(chart_path)
        plt.close()
        pie_chart_paths[year] = chart_path

    return pie_chart_paths


# Generate bar charts
def generate_bar_charts(yearwise_stats):
    bar_chart_paths = {}

    for year, stats in yearwise_stats.items():
        categories = ["Active", "Inactive"]
        values = [stats["active_volunteers"], stats["inactive_volunteers"]]
        colors = ["orange", "gray"]

        plt.figure(figsize=(6, 4))
        plt.bar(categories, values, color=colors)
        plt.xlabel("Volunteer Type")
        plt.ylabel("Number of Volunteers")
        plt.title(f"Active vs Inactive Volunteers ({year})")

        chart_path = f"reports/bar_chart_{year}.png"
        plt.savefig(chart_path)
        plt.close()
        bar_chart_paths[year] = chart_path

    return bar_chart_paths


# Generate line graph
def generate_line_graph(yearwise_stats):
    years = list(yearwise_stats.keys())
    participation_rates = [stats["pass_percentage"] for stats in yearwise_stats.values()]

    plt.figure(figsize=(8, 5))
    plt.plot(years, participation_rates, marker='o')
    plt.xlabel("Year")
    plt.ylabel("Pass Percentage (%)")
    plt.title("Volunteer Pass Rate Over Years")
    plt.xticks(years)

    chart_path = f"reports/participation_rate_line_graph.png"
    plt.savefig(chart_path)
    plt.close()

    return chart_path


def extract_credits(column_name):
    """Extract numerical credits from column names, default to 1 if not found."""
    match = re.search(r'(\d+)\s*[Cc]redits?', column_name)
    return int(match.group(1)) if match else 1  # Default to 1 credit


def generate_participation_chart(file_paths, years, output_dir="reports", max_bars_per_chart=20):
    """Generates uniform-sized volunteer participation charts for multiple years and saves them."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    participation_chart_paths = {}
    print("Files in /tmp/uploads:", os.listdir("/tmp/uploads"))

    for file_path, year in zip(file_paths, years):
        full_path = os.path.join("/tmp/uploads", file_path)  # Fix here
        print(f"Reading file: {full_path}")  # Debugging statement
        if not os.path.exists(full_path):  # Check if file exists
            print(f"Error: File not found - {full_path}")
            continue  # Skip if file is missing
        df = pd.read_excel(full_path)

        # Drop unnecessary metadata columns
        metadata_columns = ["S.No", "s. no", "Name", "Roll No", "Smail", "Total Credits Earned",
                            "Total Pending Credits", "Total Credits", "Phone Number", "Contact Number ", "Mail Id",
                            "Event Credits", "Roll Number", "Smail Address", "Mail Sent or Not", "Pass/Pending"]
        df = df.drop(columns=[col for col in metadata_columns if col in df.columns], errors="ignore")

        # Extract valid project columns and their participation values
        participation_data = []
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            total_participants = (df[col] > 0).sum()
            if total_participants >= 0:
                participation_data.append((col, total_participants))

        # Convert to DataFrame
        event_data = pd.DataFrame(participation_data, columns=["Project/Event", "Total Volunteers"])
        if event_data.empty:
            print(f"⚠ No valid participation data found for {year}!")
            continue

        # Sort data by total volunteers
        event_data = event_data.sort_values(by="Total Volunteers", ascending=False)

        # Split into chunks if too many categories
        num_chunks = (len(event_data) + max_bars_per_chart - 1) // max_bars_per_chart

        for chunk_idx in range(num_chunks):
            chunk_data = event_data.iloc[chunk_idx * max_bars_per_chart:(chunk_idx + 1) * max_bars_per_chart]

            # Set a fixed size for consistency
            fig, ax = plt.subplots(figsize=(12, 6))

            max_volunteers = max(chunk_data["Total Volunteers"].max(), 2)
            ax.set_xlim(0, max_volunteers + (max_volunteers * 0.2))

            bars = ax.barh(chunk_data["Project/Event"], chunk_data["Total Volunteers"], color="royalblue")

            for bar in bars:
                label_position = bar.get_width() + 0.2 if bar.get_width() < max_volunteers * 0.1 else bar.get_width() - 0.2
                label_alignment = "left" if bar.get_width() < max_volunteers * 0.1 else "right"
                label_color = "black" if bar.get_width() < max_volunteers * 0.1 else "white"
                ax.text(
                    label_position,
                    bar.get_y() + bar.get_height() / 2,
                    f"{int(bar.get_width())}",
                    va="center",
                    ha=label_alignment,
                    fontsize=10,
                    color=label_color,
                )

            ax.set_xlabel("Number of Volunteers", fontsize=12)
            ax.set_ylabel("Projects / Events", fontsize=12)
            ax.set_title(f"Volunteer Participation ({year}) - Chart {chunk_idx + 1}", fontsize=14)

            ax.tick_params(axis='x', labelsize=10)
            ax.tick_params(axis='y', labelsize=8)
            ax.invert_yaxis()

            chart_path = os.path.join(output_dir, f"participation_chart_{year}part{chunk_idx + 1}.png")
            plt.savefig(chart_path, dpi=300, bbox_inches="tight")
            plt.close()

            participation_chart_paths[f"{year}part{chunk_idx + 1}"] = chart_path

    return participation_chart_paths


def generate_pdf_report(yearwise_stats, output_path, participation_chart_paths):
    doc = SimpleDocTemplate(output_path, pagesize=letter, title="Credit Report")
    styles = getSampleStyleSheet()
    elements = []

    # Title and Date
    elements.append(Paragraph("<b>CREDIT SHEET REPORT</b>", styles["Title"]))
    elements.append(
        Paragraph(f"<i>Generated on: {datetime.now().strftime('%A, %B %d, %Y %H:%M:%S')}</i>", styles["Normal"]))
    elements.append(Spacer(1, 12))

    # Add Year-wise Statistics Table
    for year, stats in yearwise_stats.items():
        year_heading = Paragraph(f"<b>YEAR: {year}</b>", styles["Heading2"])

        data = [
            ["Performance Indicator", "Count"],
            ["Total Volunteers", stats["total_volunteers"]],
            ["Passed Volunteers (>= 55 credits)", stats["passed_volunteers"]],
            ["Pending Volunteers (< 55 credits)", stats["pending_volunteers"]],
            ["Near Pass Volunteers (40-54 credits)", stats["near_pass_volunteers"]],
            ["Active Volunteers (10-39 credits)", stats["active_volunteers"]],
            ["Inactive Volunteers (< 10 credits)", stats["inactive_volunteers"]],
            ["Average Credits Earned by a Volunteer ", f"{stats['average_credits']:.2f}"],
            ["Maximum Credits Earned", stats["max_credits"]],
            ["Minimum Credits Earned", stats["min_credits"]],
            ["Pass Percentage (%)", f"{stats['pass_percentage']:.2f}%"]
        ]

        table = Table(data, colWidths=[250, 150])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
            ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ]))

        elements.append(KeepTogether([year_heading, Spacer(1, 6), table]))
        elements.append(Spacer(1, 12))
        elements.append(
            Paragraph(f"<i>The above table shows the Summary of the Performance Indicators</i>", styles["Normal"]))
    # Generate Charts
    pie_chart_paths = generate_pie_charts(yearwise_stats)
    bar_chart_paths = generate_bar_charts(yearwise_stats)
    line_graph_path = generate_line_graph(yearwise_stats)

    # Add Charts to PDF
    for year in yearwise_stats:
        title = Paragraph(f"<b>Volunteer Statistics ({year})</b>", styles["Heading2"])
        pie_Comment = Paragraph(
            f"<i>Above graph shows Volunteer Pass Distribution.  Volunteers who have scored >= 55 Credits are considered as Passed</i>",
            styles["Normal"])
        bar_Comment = Paragraph(
            f"<i>Above graph shows Volunteer active participation.  Active Volunteers have 10 to 39 credits while inactive volunteers have < 10 credits</i>",
            styles["Normal"])
        pie_chart = Image(pie_chart_paths[year], width=280, height=280)
        bar_chart = Image(bar_chart_paths[year], width=280, height=280)

        # Keep title and pie chart together
        elements.append(KeepTogether([title, pie_chart, pie_Comment, bar_chart, bar_Comment]))
        elements.append(Spacer(1, 3))  # Minimal extra spacing
    if len(years) > 1:
        elements.append(Paragraph("<b>Pass Rate of Volunteers Over Years</b>", styles["Heading2"]))
        elements.append(Image(line_graph_path, width=600, height=300))
        elements.append(
            Paragraph(f"<i>Above graph shows Volunteer Pass Rate trend over Years (Volunteers with Credits >= 55) </i>",
                      styles["Normal"]))
        elements.append(Spacer(1, 20))
    # Add Participation Charts to PDF
    for year in participation_chart_paths.keys():
        chart_title = Paragraph(f"<b>Volunteer Participation in Projects and Events ({year})</b>", styles["Heading2"])
        chart_image = Image(participation_chart_paths[year], width=500, height=300)  # Adjust size as needed
        chart_Comment = Paragraph(
            f"<i> Above graph shows participation Trends for Volunteers across Events & Projects</i>", styles["Normal"])

        # Group title and chart together to prevent splitting across pages
        elements.append(KeepTogether([chart_title, Spacer(1, 12), chart_image, chart_Comment]))
        elements.append(Spacer(1, 24))  # Add spacing between sections

    doc.build(elements)


# Define the folder path
# folder_path = os.getcwd()
# folder_path = os.path.join(os.getcwd(), "uploads")
folder_path = "/tmp/uploads"

# Get all .xls and .xlsx files
excel_files = glob.glob(os.path.join(folder_path, "*.xls*"))

# Extract and print file names only (without the full path)
file_paths = [os.path.basename(file) for file in excel_files]


print("Excel Files Found:", file_paths)

# Extract years using regex
years = [int(re.search(r"\b(20\d{2})\b", file).group(1)) for file in file_paths if re.search(r"\b(20\d{2})\b", file)]

print("Extracted Years:", years)

# Load and process data from each file
all_data = []
for file_path, year in zip(file_paths, years):
    df = load_and_clean_data(file_path, year)
    all_data.append(df)

# Concatenate data from all files
combined_df = pd.concat(all_data, ignore_index=True)

# Process combined data
yearwise_stats = process_credit_data(combined_df)

# Generate participation charts for all years
participation_chart_paths = generate_participation_chart(file_paths, years)

# Generate PDF report
output_pdf_path = "credit_report.pdf"

try:
    generate_pdf_report(yearwise_stats, output_pdf_path, participation_chart_paths)
    print(f"✅ Report successfully generated: {output_pdf_path}")
    folder_path = "/tmp/uploads"
    # Delete all uploaded files after processing
    try:
        shutil.rmtree(folder_path)  # Deletes the entire folder
        os.makedirs(folder_path)  # Recreate the empty folder
        print("All files deleted from /tmp/uploads/")
    except Exception as e:
        print(f"Error deleting folder: {e}")
    # Define the folders
    #ARCHIVE_FOLDER = os.path.join(os.getcwd(), 'archive')
    #os.makedirs(ARCHIVE_FOLDER, exist_ok=True)

    # Move files to the archive after generating the PDF
    #for file_path in glob.glob(os.path.join(os.getcwd(), '*.xls*')):
    #    shutil.move(file_path, os.path.join(ARCHIVE_FOLDER, os.path.basename(file_path)))
    #print("Files moved to archive successfully!")
except Exception as e:
   print(f"❌ Error: {e}")
