# DocxMergeTOC

This script takes a folder of .docx Word documents, converts each one to PDF, merges them into a single PDF, and prepends a Table of Contents (TOC) page based on the original filenames. Each TOC entry links to the corresponding section of the merged document.

## Process Breakdown
1. Imports & Setup:
    - Uses win32com.client to automate Microsoft Word for DOCX to PDF conversion.

    - Uses fitz (PyMuPDF) to merge PDF files and create clickable links.

    - Uses reportlab to generate the TOC page dynamically.

    - Uses a temporary directory to store intermediate PDF files.

2.  Font Registration:

    - Loads a custom Hebrew TrueType font (.ttf) so RTL (Right-To-Left) Hebrew text can be rendered correctly in the TOC page.

3. DOCX to PDF Conversion:

    - Iterates over all .docx files in the specified folder.

    - Each file is converted to PDF using Word automation (SaveAs(..., FileFormat=17)).

    - Skips other file types or temporary lock files.

4. Merging PDFs:

    - Each converted PDF is opened and inserted into a single fitz PDF document (merged_pdf).

    - While inserting, the code keeps track of each file’s starting page number in the merged document.

5. TOC Page Generation (via reportlab):

    - A new PDF page is created in memory with a TOC heading.

    - For each original filename, a line like filename .......... page_number is added (in RTL Hebrew).

    - The x/y coordinates of each line are saved for hyperlinking.

6. Merge TOC into Final Document:

    - The TOC PDF is inserted at the beginning of the final PDF.

    - The original content (merged PDFs) is appended after that.

7. Add Hyperlinks to TOC:

    - For each TOC line, a clickable fitz.LINK_GOTO is added over the text area.

    - Clicking a TOC entry jumps to the corresponding page in the document.

    - A blue underline is rendered to indicate the link visually.

8. Final Output:

    - The final PDF (with TOC + content + links) is saved to disk.

