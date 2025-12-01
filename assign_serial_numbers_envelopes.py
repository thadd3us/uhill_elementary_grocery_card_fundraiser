#!/usr/bin/env python3
import typer
import pandas as pd
from pathlib import Path
from tqdm import tqdm

app = typer.Typer()


@app.command()
def main(
    envelopes_file: Path = typer.Argument(..., help="Input Excel file with envelope data"),
    serial_numbers: Path = typer.Argument(..., help="Excel file with serial numbers by card type"),
    output_file: Path = typer.Argument(..., help="Output Excel file with assigned serial numbers"),
):
    """Assign serial numbers to envelopes based on card orders."""

    # Read input files
    envelopes_df = pd.read_excel(envelopes_file)
    # Read serial numbers as strings to preserve full precision
    serial_numbers_df = pd.read_excel(serial_numbers, dtype=str)

    # Get card type columns (those that end with a dollar amount)
    card_columns = [col for col in envelopes_df.columns
                   if any(col.endswith(f'${amt}') for amt in ['50', '100', '200'])]

    # Verify that serial_numbers_df has matching columns
    missing_cols = set(card_columns) - set(serial_numbers_df.columns)
    if missing_cols:
        typer.echo(f"Error: Serial numbers file is missing columns: {missing_cols}", err=True)
        raise typer.Exit(1)

    # Build the serial number pools
    serial_pools = {}
    for col in card_columns:
        # Get non-null serial numbers from the column (already strings from dtype=str)
        serials = serial_numbers_df[col].dropna().tolist()

        # Special handling for Save-On $200 - create synthetic cards from pairs of $100
        if col == "Save-On-Foods $200":
            # Count how many $200 cards we need
            total_200_needed = envelopes_df[col].fillna(0).astype(int).sum()

            # Get $100 cards to pair up
            save_on_100_col = "Save-On-Foods $100"
            if save_on_100_col in serial_numbers_df.columns:
                all_100_serials = serial_numbers_df[save_on_100_col].dropna().tolist()

                # Create synthetic $200 cards from pairs of $100 cards
                synthetic_200 = []
                cards_needed = total_200_needed * 2  # Need 2x $100 cards
                for i in range(0, cards_needed, 2):
                    if i + 1 < len(all_100_serials):
                        synthetic_200.append(f"{all_100_serials[i]},{all_100_serials[i+1]}")

                # Remove the used $100 cards from the pool
                serial_pools["Save-On-Foods $100"] = all_100_serials[cards_needed:]
                serial_pools[col] = synthetic_200
            else:
                serial_pools[col] = serials
        elif col != "Save-On-Foods $100":
            # For non-Save-On-100 columns, just use the serials as-is
            serial_pools[col] = serials
        # Save-On-Foods $100 is handled in the $200 block above

    # Verify we have enough serial numbers
    for col in card_columns:
        if col not in serial_pools:
            continue
        total_needed = envelopes_df[col].fillna(0).astype(int).sum()
        available = len(serial_pools[col])
        if available < total_needed:
            typer.echo(
                f"Error: Not enough serial numbers for {col}. "
                f"Need {total_needed}, have {available}",
                err=True
            )
            raise typer.Exit(1)

    # Create output dataframe with envelope number and student name
    output_rows = []

    for idx, row in tqdm(envelopes_df.iterrows(), total=len(envelopes_df)):
        envelope_num = row.get("Envelope Number", idx)
        student_name = row.get("Student Name", "")

        result_row = {
            "Envelope Number": envelope_num,
            "Student Name": student_name,
        }

        # For each card type, assign serial numbers
        for col in card_columns:
            if col not in serial_pools:
                continue

            count = int(row.get(col, 0)) if pd.notna(row.get(col)) and row.get(col) != "" else 0

            if count > 0:
                # Pop the required number of serial numbers
                assigned_serials = []
                for _ in range(count):
                    if serial_pools[col]:
                        assigned_serials.append(serial_pools[col].pop(0))

                # Join multiple serials with space
                result_row[col] = " ".join(assigned_serials) if assigned_serials else ""
            else:
                result_row[col] = ""

        output_rows.append(result_row)

    # Create output dataframe
    output_df = pd.DataFrame(output_rows)

    # Verify all serial numbers were used
    unused = {col: len(serials) for col, serials in serial_pools.items() if serials}
    if unused:
        typer.echo(f"Warning: Unused serial numbers remaining: {unused}", err=True)

    # Write output - use ExcelWriter to ensure strings are preserved
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False)
        # Format all card columns as text to prevent scientific notation
        worksheet = writer.sheets['Sheet1']
        for col_idx, col_name in enumerate(output_df.columns, start=1):
            if any(col_name.endswith(f'${amt}') for amt in ['50', '100', '200']):
                for row_idx in range(2, len(output_df) + 2):  # Start from row 2 (after header)
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    if cell.value:
                        cell.number_format = '@'  # Text format

    typer.echo(f"Generated {output_file} with {len(output_df)} envelopes")


if __name__ == "__main__":
    app()
