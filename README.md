Excel Automation with Pywinauto
This Python script demonstrates how to automate tasks in Microsoft Excel using pywinauto. It opens Excel, creates a new workbook, writes "Cloudcx.ai" in cell A1, saves the file, and then closes Excel.

Prerequisites
Python 3.x installed on your system.
pywinauto library installed. You can install it using pip:
Copy code
pip install pywinauto
Setup
Excel Path: Modify the excel_path variable in the script to point to the location of your Excel executable.

python
Copy code
excel_path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"  # Adjust this path
Coordinates: Adjust the coordinates in the script if needed to interact with specific elements in Excel GUI.

python
Copy code
main_window.click_input(coords=(200, 150))  # Adjust coordinates as needed
Usage
Run the script:

sh
Copy code
python automate_excel.py
Observe Excel opening, creating a new workbook, typing "Cloudcx.ai" in cell A1, saving the file as "test.xlsx", and closing Excel.
