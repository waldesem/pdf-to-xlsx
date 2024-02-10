import os
import asyncio

import camelot
import pandas as pd

base_dir = os.path.abspath(os.path.join(__file__, "..", ".."))


async def pdf_to_excel(pdf_file_path):
    path_to_pdf = os.path.join(base_dir, pdf_file_path)
        
    # pages_num = int(input("Enter number of pages: "))

    # if pages_num < 100:
    #     tables = camelot.read_pdf(path_to_pdf, pages="all")
    # else:
    #     step = 100
    #     count = pages_num
    #     for i in range(1, pages_num, step):
    #         tables = camelot.read_pdf(path_to_pdf, pages=f"{i}-{count}")
    #         count -= 100
    #         print(count)
    #         if count < 0:
    #             step = -count
    
    tables = camelot.read_pdf(path_to_pdf, pages="all")

    combined_df = pd.concat([table.df for table in tables], ignore_index=True)

    combined_df.to_excel(
        os.path.join(base_dir, pdf_file_path.replace(".pdf", ".xlsx")),
        sheet_name="Sheet1",
        index=False,
    )


async def main():
    tasks = []
    for pdf_file_path in os.listdir(base_dir):
        if pdf_file_path.endswith(".pdf"):
            print(f"Converting {pdf_file_path} to Excel format...")
            tasks.append(pdf_to_excel(pdf_file_path))

    await asyncio.gather(*tasks)

    print("All tasks completed.")


if __name__ == "__main__":
    asyncio.run(main())
