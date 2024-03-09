# xlsx2docx4budget
creating an xlsx (excel file) converter which converts to word documents for ease of reading when attending budget meetings

This script is designed to convert Excel files (xlsx) into Word documents (docx). The primary use case is to facilitate the reading of budget data during meetings. Instead of scrolling through rows and columns of numbers, you can read a well-formatted Word document that presents the data in a more digestible format.

### How to Use
1. **Installation**: Clone this repository to your local machine. Ensure you have Python installed, along with the required libraries. You can install the necessary libraries by running `pip install -r requirements.txt` in your terminal.
2. **Prepare Your Excel File**: The script expects an Excel file with specific column headers. Make sure your Excel file is formatted correctly.
3. **Run the Script**: Navigate to the directory containing the script in your terminal. Run the script by typing `python xlsx2docx4budget.py` followed by the name of your Excel file. For example: `python xlsx2docx4budget.py budget.xlsx`.
4. **Check the Output**: The script will generate a Word document in the same directory. The name of the Word document will be the same as the Excel file, but with a .docx extension.

### Requirements
- Python 3.6 or higher
- pandas
- openpyxl
- python-docx

### Required Headers in xlsx file
The Excel file should have the following column headers, in this exact order:

1. Project title
2. Project description
3. Business driver
4. Business value
5. Business risk
6. Budget expense
7. Internal hours
8. External hours
9. Solutions involvement
10. Solutions hours
11. PMO involvement
12. PMO hours
13. Total expected hours

### Contributing
We welcome contributions to this project. Please feel free to submit issues and pull requests.

### License
This project is licensed under the MIT License. See the LICENSE file for details.