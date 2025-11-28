#!/usr/bin/env python3
import typer
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.platypus import (
    PageBreak,
    SimpleDocTemplate,
    Table,
    TableStyle,
    Image,
    Spacer,
    Paragraph,
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from pathlib import Path
from tqdm import tqdm

app = typer.Typer()


def create_envelope_page(story, row, header_text: str, image_path: Path):
    """Create one envelope page from a row of data."""
    # Add image
    if image_path.exists():
        img = Image(str(image_path), width=0.8 * inch, height=0.33 * inch)
        story.append(img)

    # Add header text
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    title_style.fontSize = 11
    header = Paragraph(header_text, title_style)
    story.append(header)
    story.append(Spacer(1, 0.05 * inch))

    # Helper to format cell value
    def fmt(val):
        return (
            ""
            if pd.isna(val) or val == 0
            else int(val) if isinstance(val, float) else val
        )

    # Prepare table data
    table_data = [
        [
            "Envelope\nNumber",
            fmt(row["Envelope Number"]),
            "Save-On-Foods $50",
            fmt(row.get("Save-On-Foods $50")),
        ],
        [
            "Grade",
            f"{row['Grade']}",
            "Save-On-Foods $100",
            fmt(row.get("Save-On-Foods $100")),
        ],
        [
            "Homeroom",
            f"{row['Homeroom']}",
            "Save-On-Foods $200",
            fmt(row.get("Save-On-Foods $200")),
        ],
        [
            "Student #",
            row["Student #"],
            "Choices Market $100",
            fmt(row.get("Choices Market $100")),
        ],
        [
            "Student Name",
            row["Student Name"],
            "Choices Market $200",
            fmt(row.get("Choices Market $200")),
        ],
        ["", "", "Urban Fare $50", fmt(row.get("Urban Fare $50"))],
        ["", "", "Urban Fare $100", fmt(row.get("Urban Fare $100"))],
    ]

    # Create table with appropriate widths
    # Original widths scaled up by factor of (4.05 + 1.18) / 4.05 = 1.2914
    col_widths = [1.098 * inch, 2.195 * inch, 1.485 * inch, 0.452 * inch]
    table = Table(table_data, colWidths=col_widths)

    # Style the table
    style = TableStyle(
        [
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ]
    )
    table.setStyle(style)
    story.append(table)


@app.command()
def main(
    input_file: Path = typer.Argument(..., help="Input Excel file"),
    output_file: Path = typer.Argument(..., help="Output PDF file"),
    header_text: str = typer.Argument(..., help="Header text for each page"),
    image_file: Path = typer.Option(
        "UniversityHill.png", help="Image file for envelope header"
    ),
):
    """Generate envelope PDFs from Excel data."""
    df = pd.read_excel(input_file)

    doc = SimpleDocTemplate(
        str(output_file),
        pagesize=(468, 261),
        leftMargin=0.3 * inch,
        rightMargin=0.3 * inch,
        topMargin=0.3 * inch,
        bottomMargin=0.3 * inch,
    )

    story = []
    for idx, row in tqdm(df.iterrows()):
        if idx > 0:
            story.append(PageBreak())
        create_envelope_page(story, row, header_text, image_file)

    doc.build(story)
    typer.echo(f"Generated {output_file} with {len(df)} envelopes")


if __name__ == "__main__":
    app()
