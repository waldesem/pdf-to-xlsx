"""Pdf recognition."""

import io
import os
import tkinter as tk
from pathlib import Path
from threading import Thread
from tkinter import filedialog, messagebox, ttk

import camelot
import cv2
import fitz
import numpy as np
import pandas as pd
import pytesseract
from PIL import Image


class PdfConverter:
    """GUI for pdf conversion."""

    def __init__(self, root: tk.Tk) -> None:
        """Initialize the GUI."""
        self.root = root
        self.root.title("Pdf-Converter")
        self.root.geometry("540x260")

        # Variables
        self.selected_file = None
        self.output_file = None

        self.setup_ui()

    def setup_ui(self) -> None:
        """Set up the UI components."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Pdf converter",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # File selection
        ttk.Label(main_frame, text="Select pdf file:").grid(
            row=1,
            column=0,
            sticky=tk.W,
            pady=5,
        )

        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.file_label = ttk.Label(
            file_frame,
            text="No file selected",
            foreground="gray",
            width=50,
        )
        self.file_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))

        self.upload_btn = ttk.Button(
            file_frame,
            text="Browse",
            command=self.browse_file,
        )
        self.upload_btn.grid(row=0, column=1)

        # Checkbox
        self.checkbox_var = tk.BooleanVar()
        self.check_button = ttk.Checkbutton(
            main_frame, text="Pdf as images", variable=self.checkbox_var,
        )
        self.check_button.grid(
            row=3, column=0, sticky=tk.W, pady=(10, 0), padx=(0, 10), columnspan=2,
        )
        self.checkbox_var.set(True)

        # Progress bar
        self.progress_label = ttk.Label(main_frame, text="")
        self.progress_label.grid(row=5, column=0, columnspan=2, pady=(20, 5))

        self.progress_bar = ttk.Progressbar(main_frame, mode="determinate")
        self.progress_bar.grid(
            row=6,
            column=0,
            columnspan=2,
            sticky=(tk.W, tk.E),
            pady=5,
        )

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=(20, 0))

        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        file_frame.columnconfigure(0, weight=1)

    def browse_file(self) -> None:
        """Open file dialog to select pdf file."""
        file_types = [
            ("Pdf file", "*.PDF *.pdf"),
            ("All files", "*.*"),
        ]

        filename = filedialog.askopenfilename(
            title="Select pdf",
            filetypes=file_types,
        )

        if filename:
            self.selected_file = Path(filename)
            self.file_label.config(text=self.selected_file.name, foreground="black")
            self.output_file = None
            self.start_conversion()

    def start_conversion(self) -> None:
        """Start conversion process."""
        thread = Thread(target=self.get_pdf_data)
        thread.daemon = True
        thread.start()

    def conversion_completed(self) -> None:
        """Handle successful conversion completion."""
        self.progress_bar.stop()
        self.progress_label.config(text="Conversion completed successfully!")

        # Re-enable button
        self.upload_btn.config(state="normal")
        self.file_label.config(text="No file selected", foreground="gray")

        messagebox.showinfo("Success", f"Text file saved as:\n{self.output_file.name}")
        self.open_folder()
        self.root.after(
            3000,
            lambda: self.progress_label.config(text=""),
        )

    def conversion_failed(self, error_message: str) -> None:
        """Handle conversion failure."""
        self.progress_bar.stop()
        self.progress_label.config(text="Conversion failed!")

        # Re-enable button
        self.upload_btn.config(state="normal")
        self.file_label.config(text="No file selected", foreground="gray")

        messagebox.showerror("Error", f"Conversion failed:\n{error_message}")
        self.root.after(
            3000,
            lambda: self.progress_label.config(text=""),
        )

    def open_folder(self) -> None:
        """Open the folder containing the generated text file."""
        if self.output_file and self.output_file.exists():
            try:
                # Get the parent directory of the output file
                folder_path = self.output_file.parent

                # Try to open folder with default system application
                if os.name == "nt":  # Windows
                    os.startfile(folder_path)  # noqa: S606
                elif os.name == "darwin":  # macOS
                    os.system(f'open "{folder_path}"')  # noqa: S605
                elif os.name == "posix":  # Linux
                    os.system(f'xdg-open "{folder_path}"')  # noqa: S605
            except OSError as e:
                messagebox.showerror("Error", f"Could not open folder:\n{e}")
        else:
            messagebox.showwarning("Warning", "No file available to open.")

    def get_pdf_data(self) -> None:
        """Extract tabular data from a PDF file."""
        if self.checkbox_var:
            result = self.pdf_image_to_table()
            if result:
                self.output_file = Path(
                        self.selected_file.parent,
                        f"{self.selected_file.stem}.txt",
                )
                with self.output_file.open("w") as f:
                    for table in result:
                        f.write(table)
                self.root.after(0, self.conversion_completed)
            else:
                self.conversion_failed("Failed to extract data from PDF file")
        else:
            handler = camelot.handlers.PDFHandler(self.selected_file)
            page_list = handler._get_pages(pages="all")  # noqa: SLF001
            for i in range(0, len(page_list), 100):
                self.pdf_to_excel(page_list, i)
                self.merge_xlsx_files(self.selected_file.parent)

            self.root.after(0, self.conversion_completed)

    def pdf_image_to_table(self) -> list:
        """Extract tabular data from a PDF file."""
        images = self.extract_images_from_pdf()
        self.progress_bar["maximum"] = len(images)
        tables = []
        for page_num, image_bytes in images:
            processed_image = self.preprocess_image(image_bytes)
            text: str = pytesseract.image_to_string(
                processed_image, lang="rus+eng'", config="--psm 6",
            )
            tables.append(text)
            self.progress_bar.step(page_num)
        return tables

    def extract_images_from_pdf(self) -> list[tuple[int, bytes]]:
        """Extract images from a PDF file."""
        doc = fitz.open(self.selected_file)
        images = []
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            image_list = page.get_images(full=True)
            for img in image_list:
                base_image = doc.extract_image(img[0])
                image_bytes = base_image["image"]
                images.append((page_num, image_bytes))
        return images

    @staticmethod
    def preprocess_image(image_bytes: bytes) -> bytes:
        """Preprocess the image for OCR."""
        image = Image.open(io.BytesIO(image_bytes))
        image = np.array(image.convert("L"))  # Перевод в grayscale
        _, binary = cv2.threshold(image, 150, 255, cv2.THRESH_BINARY)  # Бинаризация
        return binary

    def pdf_to_excel(self, page_list: list, i: int) -> None:
        """Convert a range of pages from a PDF file to Excel format."""
        pages = ", ".join([str(page) for page in page_list[i : i + 100]])
        tables = camelot.read_pdf(self.selected_file, pages=pages)

        df = pd.concat([table.df for table in tables], ignore_index=True)

        df.to_excel(
            Path(
                self.selected_file.parent,
                self.selected_file.replace(".pdf", f"_{i // 100 + 1}.xlsx"),
            ),
            sheet_name="Sheet1",
            index=False,
        )

    def merge_xlsx_files(self, pdf_file_path: Path) -> None:
        """Merge multiple partially converted Excel files into one Excel file."""
        dfs = []
        for xlsx_file_path in sorted(Path(self.selected_file.parent).iterdir()):
            if xlsx_file_path.match(
                pdf_file_path.replace(".pdf", "_"),
            ) and xlsx_file_path.suffix(".xlsx"):
                df = pd.read_excel(
                    Path(self.selected_file.parent, xlsx_file_path),
                )
                dfs.append(df)

                Path.unlink(Path(self.selected_file.parent, xlsx_file_path))

        combined_df = pd.concat(dfs, ignore_index=True)
        headers = [f"column_{index}" for index, _ in enumerate(combined_df.columns)]
        combined_df.columns = headers

        final_df = combined_df.replace("\n", " ", regex=True)
        final_df.to_excel(
            Path(self.selected_file.parent, pdf_file_path.replace(".pdf", ".xlsx")),
            sheet_name="Sheet1",
            index=False,
        )
        messagebox.showinfo("Merged in Excel format.")


def main() -> None:
    """Run the application."""
    root = tk.Tk()
    PdfConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
