# PDF-Converter-and-Merger
A versatile tool for converting various file types to PDF and merging them into a single document. Simplify file management and PDF creation in one place!

This Python script provides a versatile graphical user interface (GUI) application for file conversion and PDF merging. With this tool, you can easily convert various types of files, including images, text documents, office files, and PDFs, into PDF format. It also offers an option to merge the converted PDFs into a single PDF file. The script utilizes popular libraries like Tkinter for the GUI, PIL (Pillow) for image handling, fpdf for PDF generation, PyPDF2 for PDF merging, and more.

Key Features:

>Select Input Files: You can choose multiple input files of various types, such as images (JPEG, PNG, GIF), text files (TXT), office documents (DOCX, XLSX, PPTX), and existing PDFs.
>File Reordering: The GUI allows you to rearrange the order of selected files by moving them up or down in the list.
>Output Directory: Specify the directory where the converted and merged PDFs will be saved.
>Merge Option: You can choose to merge the selected files into a single PDF or convert them individually to separate PDFs.
>Parallel Processing: The script takes advantage of parallel processing using Python's ThreadPoolExecutor for efficient conversion of multiple files.
>PDF Conversion: Supported file types are automatically converted to PDF, and unsupported types are flagged.
>Progress Tracking: A progress bar keeps you informed of the conversion and merging progress.

Feel free to modify and enhance this tool, contribute to its development, or adapt it for your specific needs.
