import os

import camelot
import pandas as pd

base_dir = os.path.abspath(os.path.join(__file__, "..", ".."))

def pdf_to_excel(pdf_file_path):
    path_to_pdf = os.path.join(base_dir, pdf_file_path)
    tables = camelot.read_pdf(path_to_pdf, pages="all")

    combined_df = pd.concat([table.df for table in tables], ignore_index=True)

    combined_df.to_excel(
        pdf_file_path.replace(".pdf", ".xlsx"), sheet_name="Sheet1", index=False
    )


if __name__ == "__main__":
    for pdf_file_path in os.listdir(base_dir):
        if pdf_file_path.endswith(".pdf"):
            pdf_to_excel(pdf_file_path)
