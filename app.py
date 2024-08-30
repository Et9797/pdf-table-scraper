import re
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pandastable import Table
from pathlib import Path
from loguru import logger


PROJECT_ROOT_DIR = Path.cwd().resolve()
PROJECT_ROOT_DIR


logger.add(PROJECT_ROOT_DIR / "pdf-scraper-log.txt", level="DEBUG", format='{time:HH:mm:ss} ({name}:{function}:{line}) - {message}')


class PDFScraperApp:
    def __init__(self, root) -> None:
        self.root = root
        self.root.title("PDF Scraper")
        self.center_window(660, 640)  # Center the window with 660x640 size

        # Create a canvas widget to hold the frames and add a scrollbar
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        # Configure the canvas and scrollbar
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack the canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Frame to hold the content
        self.frame1 = tk.Frame(self.scrollable_frame)
        self.frame1.pack(fill=tk.BOTH, expand=True)  # Adjust fill and expand to make sure it takes available space
        self.frame2 = tk.Frame(self.scrollable_frame)
        self.frame2.pack(fill=tk.BOTH, expand=True)

        # Create the buttons
        self.upload_button = tk.Button(self.frame2, text="Upload PDF", command=self.upload_pdf)
        self.export_button = tk.Button(self.frame2, text="Export to Excel", command=self.export_to_excel)

        # Configure the grid to center the buttons
        self.frame2.grid_columnconfigure(0, weight=1)
        self.frame2.grid_columnconfigure(1, weight=1)
        self.frame2.grid_rowconfigure(0, weight=1)

        # Place the buttons in the grid
        self.upload_button.grid(row=0, column=0, padx=260, pady=5, sticky="w")  # Align to the left
        self.export_button.grid(row=1, column=0, padx=250, pady=5, sticky="e")  # Align to the right

        self.pdf_file_path: Path = None

    def center_window(self, width: int, height: int) -> None:
        # Get screen width and height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Calculate x and y coordinates to center the window
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2 - 100

        # Set the window position and size
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def upload_pdf(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

        if file_path:
            if file_path.lower().endswith('.pdf'):
                messagebox.showinfo(title="Success", message=f"Uploaded:\n{file_path}")
                self.pdf_file_path = Path(file_path)
            else:
                messagebox.showerror(title="Error", message="Please upload a PDF file.")
    
    def scrape_pdf(self) -> pd.DataFrame:
        with pdfplumber.open(self.pdf_file_path) as pdf:
            table = []
            for page in pdf.pages:
                t = page.extract_table()
                if t:
                    table.extend(t)

        assert table, logger.debug("Unable to extract any table from PDF file.")

        data = [self.split_string(row[0]) for row in table]
        df = pd.DataFrame(
            data,
            columns=[
                "Artikelnr",
                "Aantal",
                "Omschrijving",
                "Prijs per stuk",
                "Korting",
                "Regeltotaal"
            ]
        )
        return df

    def split_string(self, row: str) -> tuple[str]:
        pattern = re.compile(r'(\S+)\s+(\d+,\d{2})\s+(.+?)\s+(\d+,\d{2})\s+(\d+ %)\s+(\d+,\d{2})')
        match = re.match(pattern, row)
        assert len(match.groups()) == 6, logger.debug(f"{row}\nTuple length is not equal to 6.")
        return match.groups()

    def export_to_excel(self) -> None:
        if not self.pdf_file_path:
            messagebox.showerror(title="Error", message="No PDF file uploaded.")
            return
        
        df = self.scrape_pdf()
        df.to_excel(
            f"{self.pdf_file_path.parent / self.pdf_file_path.stem}.xlsx", 
            index=False, 
            engine="openpyxl"
        )

        pandas_table = Table(self.frame1, dataframe=df, showstatusbar=False, showtoolbar=True)
        pandas_table.show()

        messagebox.showinfo(
            title="Success",
            message=(
                "Successfully converted PDF to Excel.\n\nFile location:\n"
                f"{self.pdf_file_path.parent / self.pdf_file_path.stem}.xlsx"
            )
        )


if __name__ == "__main__":
    # Create the main application window
    root = tk.Tk()
    app = PDFScraperApp(root)

    # Run the application
    root.mainloop()
