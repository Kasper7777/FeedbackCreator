# Assignment Feedback Form Generator

This Python script is a GUI-based application that automates the process of generating assignment feedback forms for students based on their marks and feedback stored in an Excel file. It supports the optional inclusion of a rubric or table from a Microsoft Word document. The feedback forms are saved as `.docx` files, each named after the respective student.

## Features

- **Automated Feedback Generation:** Reads an Excel file containing student names, marks, and feedback and generates individual Word documents for each student.
- **Customisable Configurations:** Update details such as tutor name, module title, assignment name, and the percentage the assignment contributes to the module through the GUI.
- **Optional Rubric:** Includes a table from a Word document (e.g., rubric) if selected, appending it to each student's feedback form.
- **Styled Word Documents:** Feedback forms are formatted with borders, centre-aligned headings, and footer messages.
- **Document Margins:** Word documents are set with 1-inch margins for a clean layout.

## Requirements

- Python 3.x
- `python-docx`
- `openpyxl`
- `tkinter`
  
Install the required packages using:

```bash
pip install python-docx openpyxl
