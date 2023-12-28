# offer Letter Generator

## Overview

The offer Letter Generator is a Python application built using the Tkinter GUI toolkit. It reads employee information from an Excel file and generates offer letters for each employee. The offer letters are created in both DOCX and PDF formats.

## Author

- **Author:** Surya Yadav

## Prerequisites

- Python 3.x
- Required Python packages (install via `pip install <package_name>`):
  - tkinter
  - pandas
  - python-docx
  - docx2pdf

## How to Use

1. Run the application by executing the script.
2. Browse and select the Excel file containing employee information.
3. Click the "Generate Leave Letters" button to create offer letters.
4. The generated offer letters will be displayed in the text area.

## File Descriptions

- `leave_letter_generator.py`: The main Python script containing the application code.
- `README.md`: This documentation file.

## Additional Notes

- Ensure that the Excel file follows the required format with columns such as `employee_name`, `company_name`, `job_name`, `start_date`, `salary`, and `hr_name`.

## Acknowledgments

The Leave Letter Generator uses the following Python libraries:
- Tkinter: For creating the graphical user interface.
- pandas: For reading and processing Excel files.
- python-docx: For creating and modifying Word documents.
- docx2pdf: For converting Word documents to PDF.

## License

This project is licensed under the [MIT License](LICENSE).

