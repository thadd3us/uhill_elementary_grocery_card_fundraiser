from pathlib import Path
import subprocess
import pytest

def test_generate_envelope_pdf():
    """Test generating envelope PDF from sample Excel file."""
    input_file = Path("test_data/sample_envelopes.xlsx")
    output_file = Path("test_data/golden_output.pdf")
    header_text = "Elementary PAC Grocery Card Fundraiser – Spring 2025 – Thank You!"
    image_file = Path("UniversityHill.png")

    assert input_file.exists(), f"Input file {input_file} not found"
    assert image_file.exists(), f"Image file {image_file} not found"

    output_file.parent.mkdir(parents=True, exist_ok=True)

    result = subprocess.run(
        [
            "uv", "run", "python3", "envelope_mail_merge.py",
            str(input_file),
            str(output_file),
            header_text,
            "--image-file", str(image_file)
        ],
        capture_output=True,
        text=True
    )

    assert result.returncode == 0, f"Command failed: {result.stderr}"
    assert output_file.exists(), f"Output file {output_file} was not created"
    assert output_file.stat().st_size > 0, "Output file is empty"

    print(f"\nGenerated golden output at: {output_file}")
    print(f"Output: {result.stdout}")
