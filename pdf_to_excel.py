import os
import asyncio

import camelot
import pandas as pd

base_dir = os.path.abspath(os.path.join(__file__, "..", ".."))


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
