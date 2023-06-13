from ox_script import *

print_num = 1  # Default print quantity
starting_index = 0  # Default starting index
workbook = None  # Placeholder for the Excel workbook
# Path to the Excel file, you can feel free to change this as well to support 
# multiple Excel files
excel_path = "Fixed Assets Data.xlsx"  

def print_label():
    # Loop to print the desired number of labels
    for i in range(0, int(print_num_controller.value)):
        generate_label(i, starting_index)

def generate_label(current_index, start_at=0):
    global workbook
    if type(workbook) != type(None):  # Check if the workbook has been loaded
        current_row = workbook.active[current_index + 2 + start_at]  # Retrieve the current row
        # Generate the command to update form variables and print the label
        cmd = PTK_UpdateAllFormVariables(
            "command4-en.txt",
            Input1=str(current_row[0].value),
            Input2=str(current_row[1].value),
            Input3=str(current_row[2].value),
            Input4=str(current_row[3].value),
            Input5=str(current_row[4].value),
            Input6=str(current_row[5].value),
        )
        PTK_SendCmdToPrinter(cmd)  # Send the command to the printer
        current_row[6].value = datetime.datetime.now()  # Update printed date in the Excel file
        current_row[7].value = "Printed"  # Update status in the Excel file
        workbook.save(excel_path)  # Save the changes to the Excel file

def update_print_num(value):
    global print_num
    print_num = int(print_num_controller.value)  # Update the print quantity

def update_starting_index(value):
    global starting_index
    starting_index = int(starting_index_controller.value)  # Update the starting index

def load_excel(path):
    from openpyxl import load_workbook
    global workbook
    workbook = load_workbook(path)  # Load the Excel workbook

if __name__ == "__main__":
    controller = PTK_UIInit(
        PTK_UIPage(
            PTK_UIText(title="File Name: Fixed Assets Data.xlsx"),  # Display the file name
            print_num_controller := PTK_UIInput(title="Print Quantity", value="1",  Onsubmit=update_print_num),
            starting_index_controller := PTK_UIInput(title="Starting Index", value="0",  Onsubmit=update_starting_index),
            PTK_UIButton(title="Print", Onpressed=print_label),
        ),
        require_execute_confirmation=False,  # Allow the script to run without execute confirmation
    )
    load_excel(excel_path)  # Load the Excel file into the workbook
