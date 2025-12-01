from pathlib import Path
import subprocess
import pytest
import pandas as pd


def test_assign_serial_numbers():
    """Test assigning serial numbers to envelopes from sample data."""
    envelopes_file = Path("test_data/sample_envelopes.xlsx")
    serial_numbers_file = Path("test_data/sample_serial_numbers.xlsx")
    output_file = Path("test_data/assigned_envelopes.xlsx")

    assert envelopes_file.exists(), f"Envelopes file {envelopes_file} not found"
    assert serial_numbers_file.exists(), f"Serial numbers file {serial_numbers_file} not found"

    output_file.parent.mkdir(parents=True, exist_ok=True)

    result = subprocess.run(
        [
            "uv", "run", "python3", "assign_serial_numbers_envelopes.py",
            str(envelopes_file),
            str(serial_numbers_file),
            str(output_file),
        ],
        capture_output=True,
        text=True
    )

    assert result.returncode == 0, f"Command failed: {result.stderr}"
    assert output_file.exists(), f"Output file {output_file} was not created"
    assert output_file.stat().st_size > 0, "Output file is empty"

    # Verify the output has the expected structure
    output_df = pd.read_excel(output_file)
    assert "Envelope Number" in output_df.columns
    assert "Student Name" in output_df.columns
    assert "Assigned Serial Numbers" in output_df.columns

    # Verify that serial numbers were assigned (check that the column has content)
    assigned_serials = output_df["Assigned Serial Numbers"].dropna()
    assert len(assigned_serials) > 0, "No serial numbers were assigned"

    print(f"\nGenerated output at: {output_file}")
    print(f"Output: {result.stdout}")
    print(f"\nFirst few rows:\n{output_df.head()}")
    print(f"\nSample assigned serial numbers:\n{assigned_serials.iloc[0] if len(assigned_serials) > 0 else 'None'}")
