import os
import asyncio

import camelot
import pandas as pd

base_dir = os.path.abspath(os.path.join(__file__, "..", ".."))


"""
Asynchronously extracts tabular data from a PDF file into Excel spreadsheets. 

Splits the PDF pages into batches of 100 pages, and processes each batch concurrently using asyncio.gather. 
Merges the output Excel files if there are multiple batches.

Args:
    pdf_file_path (str): Path to the input PDF file.

Returns:
    None. Outputs Excel files to the same directory as the input PDF.
"""


async def get_pdf_data(pdf_file_path):
    path_to_pdf = os.path.join(base_dir, pdf_file_path)

    handler = camelot.handlers.PDFHandler(path_to_pdf)
    page_list = handler._get_pages(pages="all")

    tasks = []

    for i in range(0, len(page_list), 100):
        tasks.append(pdf_to_excel(path_to_pdf, pdf_file_path, page_list, i))

    await asyncio.gather(*tasks)

    if len(page_list) > 100:
        await merge_xlsx_files(pdf_file_path)


"""Converts a range of pages from a PDF file to Excel format.

Args:
  path_to_pdf: The file path to the PDF file.
  pdf_file_path: The file name of the PDF file.
  page_list: The list of page numbers to convert.
  i: The start index in page_list to convert pages from.

Returns:
  None. Saves the converted Excel file to disk.
"""


async def pdf_to_excel(path_to_pdf, pdf_file_path, page_list, i):
    pages = ", ".join([str(page) for page in page_list[i : i + 100]])
    tables = camelot.read_pdf(path_to_pdf, pages=pages)

    combined_df = pd.concat([table.df for table in tables], ignore_index=True)

    combined_df.to_excel(
        os.path.join(base_dir, pdf_file_path.replace(".pdf", f"_{i//100+1}.xlsx")),
        sheet_name="Sheet1",
        index=False,
    )
    print(
        f"Partially converted in {pdf_file_path.replace('.pdf', f'_{i//100+1}.xlsx')} to Excel format."
    )


"""Merges multiple partially converted Excel files into one Excel file.

This merges the Excel files that were generated from batches of PDF pages into a single Excel file for the full PDF. It assumes the partial files were named like {original_pdf_name}_{n}.xlsx and sorts them before concatenating into a single dataframe.

Args:
  pdf_file_path: The path to the original PDF file. Used to construct the merged file name.
"""


async def merge_xlsx_files(pdf_file_path):
    dfs = []
    for xlsx_file_path in sorted(os.listdir(base_dir)):
        if xlsx_file_path.startswith(
            pdf_file_path.replace(".pdf", "_")
        ) and xlsx_file_path.endswith(".xlsx"):
            df = pd.read_excel(
                os.path.join(base_dir, xlsx_file_path),
            )
            dfs.append(df)

            os.remove(os.path.join(base_dir, xlsx_file_path))

    combined_df = pd.concat(dfs, ignore_index=True)
    combined_df.to_excel(
        os.path.join(base_dir, pdf_file_path.replace(".pdf", ".xlsx")),
        sheet_name="Sheet1",
        index=False,
    )
    print("Merged all xlsx files.")


"""Main function to convert PDFs to Excel.

Iterates through all PDF files in the provided directory, converts each one to 
Excel format using the get_pdf_data() function, and gathers all the async tasks.
Waits for all conversions to complete before exiting.
"""


async def main():
    tasks = []
    for pdf_file_path in os.listdir(base_dir):
        if pdf_file_path.endswith(".pdf"):
            print(f"Converting {pdf_file_path} to Excel format...")
            tasks.append(get_pdf_data(pdf_file_path))

    await asyncio.gather(*tasks)

    print("All tasks completed.")


if __name__ == "__main__":
    asyncio.run(main())
