import tabula
import pandas as pd
import os

def pdf_to_excel(pdf_path, output_path):
    # Extract all tables from PDF
    dfs = tabula.read_pdf(
        pdf_path,
        pages="all",
        multiple_tables=True
    )

    if not dfs or len(dfs) == 0:
        raise ValueError("No tables found in PDF")

    # Save all tables into one Excel file (multiple sheets)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for i, df in enumerate(dfs):
            df.to_excel(
                writer,
                sheet_name=f"Table_{i+1}",
                index=False
            )

    return output_path
