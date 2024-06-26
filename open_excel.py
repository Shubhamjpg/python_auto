from pywinauto import application
from pywinauto.keyboard import send_keys
import time


excel_path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"  # Adjust this path

app = application.Application().start(excel_path)

time.sleep(5)

app.connect(path=excel_path)

main_window = app.top_window()

# Wait for the main window to be ready
main_window.wait('ready')

# Maximize the window to ensure all elements are visible
main_window.maximize()

# Create a new workbook (Ctrl+N)
send_keys("^n")
time.sleep(2)

# Click on the first cell (A1)
main_window.click_input(coords=(200, 150))  # Adjust coordinates as needed

# Write "Hello World" in cell A1
send_keys("Cloudcx.ai{ENTER}")

# Save the file (Ctrl+S)
send_keys("^s")
time.sleep(2)

send_keys(r"test{ENTER}")

# Wait for the file to save
time.sleep(2)

# Close Excel (Alt+F4)
send_keys("%{F4}")
