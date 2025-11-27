# Envelope Mail Merge

Generate envelope PDFs from Excel data for grocery card fundraisers.

## Installation

```bash
uv sync
```

## Usage

```bash
uv run python3 envelope_mail_merge.py <input.xlsx> <output.pdf> "<header text>" [--image-file <image.png>]
```

### Example

```bash
 uv run python envelope_mail_merge.py test_data/sample_envelopes.xlsx /tmp/output.pdf "Elementary PAC Grocery Card Fundraiser – Spring 2025 – Thank You" --image-file UniversityHill.png
 ```

## Testing

```bash
uv run pytest .
```

The test generates a golden output PDF at `tests/golden_output.pdf` that can be checked into version control for comparison.

## Input Format

The input Excel file should contain these columns:
- Envelope Number
- Grade
- Homeroom
- Student #
- Student Name
- Choices Market $100, $200
- Save-On-Foods $50, $100, $200
- Urban Fare $50, $100

## Output

The tool generates a landscape-oriented PDF with one page per student, containing:
- School logo (if provided)
- Header text
- Table with student info and grocery card counts
