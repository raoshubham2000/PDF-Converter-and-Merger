import shutil
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import IntVar
from PIL import Image
from fpdf import FPDF
import subprocess
import os
from threading import Thread
import PyPDF2
from concurrent.futures import ThreadPoolExecutor


class FileConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("File Conversion and PDF Merge")

        self.input_files = []  # Initialize as a list
        self.output_directory = ""
        self.merge_var = IntVar()
        self.merge_var.set(0)  # By default, do not merge files

        # Create and configure a listbox to display selected files
        self.file_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=40, height=10)
        self.file_listbox.pack(pady=10)

        # Create and configure buttons for file rearrangement
        self.move_up_button = tk.Button(root, text="Move Up", command=self.move_up)
        self.move_down_button = tk.Button(root, text="Move Down", command=self.move_down)
        self.move_up_button.pack(side=tk.LEFT)
        self.move_down_button.pack(side=tk.LEFT)

        # Create and configure a button to select input files
        self.choose_input_button = tk.Button(root, text="Choose Input Files", command=self.choose_input)
        self.choose_input_button.pack(pady=5)

        # Create and configure a button to select the output directory
        self.choose_output_button = tk.Button(root, text="Choose Output Location", command=self.choose_output_directory)
        self.choose_output_button.pack(pady=5)

        # Create a checkbutton to toggle the merge option
        self.merge_checkbox = tk.Checkbutton(root, text="Merge into a single PDF", variable=self.merge_var)
        self.merge_checkbox.pack()

        # Create and configure a button to start the conversion and merge process
        self.convert_button = tk.Button(root, text="Convert to PDF", command=self.start_conversion)
        self.convert_button.pack(pady=5)

        # Create a progress bar
        self.progress_bar = ttk.Progressbar(root, length=200, mode="determinate")
        self.progress_bar.pack()

        # Create a label to display the result
        self.result_label = tk.Label(root, text="")
        self.result_label.pack()

    def choose_input(self):
        self.input_files = list(filedialog.askopenfilenames(filetypes=[("All Files", "*.*")]))
        self.update_file_listbox()

    def choose_output_directory(self):
        self.output_directory = filedialog.askdirectory()

    def update_file_listbox(self):
        self.file_listbox.delete(0, tk.END)
        for file in self.input_files:
            self.file_listbox.insert(tk.END, os.path.basename(file))

    def move_up(self):
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            if index > 0:
                item = self.file_listbox.get(index)
                self.file_listbox.delete(index)
                self.file_listbox.insert(index - 1, item)
                self.input_files.insert(index - 1, self.input_files.pop(index))

    def move_down(self):
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            if index < self.file_listbox.size() - 1:
                item = self.file_listbox.get(index)
                self.file_listbox.delete(index)
                self.file_listbox.insert(index + 1, item)
                self.input_files.insert(index + 1, self.input_files.pop(index))

    def update_progress_bar(self, value):
        self.progress_bar["value"] = value
        self.root.update_idletasks()

    def convert_files_parallel(self):
        with ThreadPoolExecutor() as executor:
            total_files = len(self.input_files)

            def convert_file(input_file):
                i = self.input_files.index(input_file)
                output_pdf_path = os.path.join(self.output_directory, f"file_{i + 1}.pdf")
                self.convert_file_to_pdf(input_file, output_pdf_path)
                self.update_progress_bar((i + 1) * 100 // total_files)

            executor.map(convert_file, self.input_files)
            self.result_label.config(text=f"Files converted to PDF in the output directory: {self.output_directory}")
            self.update_progress_bar(100)

    def start_conversion(self):
        if not self.input_files:
            self.result_label.config(text="Please choose input files.")
            return

        if not self.output_directory:
            self.result_label.config(text="Please choose an output directory.")
            return

        if self.merge_var.get() == 1:
            self.merge_and_convert()
        else:
            self.convert_files_parallel()

    def merge_and_convert(self):
        merged_pdf_path = os.path.join(self.output_directory, "merged.pdf")

        def conversion_task():
            pdf_merger = PyPDF2.PdfFileMerger()
            total_files = len(self.input_files)

            for i, input_file in enumerate(self.input_files):
                output_pdf_path = os.path.join(self.output_directory, f"file_{i + 1}.pdf")
                self.convert_file_to_pdf(input_file, output_pdf_path)
                pdf_merger.append(output_pdf_path)
                self.update_progress_bar((i + 1) * 100 // total_files)

            with open(merged_pdf_path, "wb") as merged_pdf_file:
                pdf_merger.write(merged_pdf_file)

            self.result_label.config(text=f"Files converted and merged to PDF: {merged_pdf_path}")
            self.update_progress_bar(100)

        thread = Thread(target=conversion_task)
        thread.start()

        self.update_progress_bar(0)

    def convert_files_individual(self):
        total_files = len(self.input_files)

        def conversion_task():
            for i, input_file in enumerate(self.input_files):
                output_pdf_path = os.path.join(self.output_directory, f"file_{i + 1}.pdf")
                self.convert_file_to_pdf(input_file, output_pdf_path)
                self.update_progress_bar((i + 1) * 100 // total_files)

            self.result_label.config(text=f"Files converted to PDF in the output directory: {self.output_directory}")
            self.update_progress_bar(100)

        thread = Thread(target=conversion_task)
        thread.start()

        self.update_progress_bar(0)

    def convert_file_to_pdf(self, input_file, output_pdf_path):
        input_file_name = os.path.basename(input_file)
        if input_file_name.lower().endswith(('.jpg', '.jpeg', '.png', '.gif')):
            self.convert_image_to_pdf(input_file, output_pdf_path)
        elif input_file_name.lower().endswith('.txt'):
            self.convert_text_to_pdf(input_file, output_pdf_path)
        elif input_file_name.lower().endswith(('.docx', '.xlsx', '.pptx')):
            self.convert_office_to_pdf(input_file, output_pdf_path)
        elif input_file_name.lower().endswith('.pdf'):
            shutil.copy(input_file, output_pdf_path)
        else:
            self.result_label.config(text="Unsupported file type. Please select image, text, office, or PDF files.")

    def convert_image_to_pdf(self, image_path, pdf_path):
        img = Image.open(image_path)
        pdf = FPDF()
        pdf.add_page()
        pdf.image(image_path, x=10, y=10, w=180)
        pdf.output(pdf_path)

    def convert_text_to_pdf(self, text_path, pdf_path):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        with open(text_path, "r") as file:
            for line in file:
                pdf.cell(200, 10, txt=line, ln=True)

        pdf.output(pdf_path)

    def convert_office_to_pdf(self, office_path, pdf_path):
        # Convert office files to PDF using LibreOffice (you need to have LibreOffice installed)
        cmd = ['libreoffice', '--headless', '--convert-to', 'pdf', office_path, '--outdir', self.output_directory]
        subprocess.run(cmd)


if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverter(root)
    root.mainloop()
