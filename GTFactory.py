import xlsxwriter as xl
from xlsxwriter import exceptions
import os
import time
import tkinter as tk
from tkinter import messagebox
import tkinter.filedialog as fd
import configparser as cp

error = False  # For errors.


def SetPathLogs():  # Set the path of the logs.
    path = fd.askdirectory()  # Get the directory.
    if path != "":  # If path is selected -> Do, if not selected -> don't deleted already text in the text box.
        TextLog.delete(1.0, "end-1c")  # Clears the text box.
        TextLog.insert('insert', path)  # Insert the path.


def SetPathSave():  # Set the path of the save.
    path = fd.askdirectory()  # Get the directory.
    if path != "":  # If path is selected -> Do, if not selected -> don't deleted already text in the text box.
        TextSave.delete(1.0, "end-1c")  # Clears the text box.
        TextSave.insert('insert', path)  # Insert the path.


def getTextLog():  # Get the text from the log's text box.
    config.set('directory', 'logs', TextLog.get(1.0, "end-1c"))  # Save the input from the text box to the config for the next time.
    return TextLog.get(1.0, "end-1c")  # Return the text from the text box of the logs.


def getTextSave():  # Get the text from the save's text box.
    global error  # Take the error from outside of the function.
    if TextSave.get(1.0, "end-1c") is "":  # Check if the text is empty.
        error = True  # Set the error to True.
        messagebox.showerror("Error", "Please enter a save path.")  # Create a message box.
    else:  # If the text is not empty.
        error = False  # Set the error to False.
    config['directory']['save'] = str(TextSave.get(1.0, "end-1c"))  # Save the input from the text box to the config for the next time.
    return TextSave.get(1.0, "end-1c")  # Return the text from the text box of the save.


def Done():  # Create a "Done" message box after completing the program.
    global error  # Take the error from outside of the function.
    if error is False:  # Check if there is no errors.
        messagebox.showinfo("Status", "Done!")  # Set message box.


def main():  # The main program.
    global error  # Take the error from outside of the function.

    if error is True:  # Check if there is an error.
        return True  # Return true.

    date_dic = {  # Dictionary for months.
        "Jan": "1",
        "Feb": "2",
        "Mar": "3",
        "Apr": "4",
        "May": "5",
        "Jun": "6",
        "Jul": "7",
        "Aug": "8",
        "Sep": "9",
        "Oct": "10",
        "Nov": "11",
        "Dec": "12",
    }

    workbook = xl.Workbook(f'{getTextSave()}\\Shelves.xlsx')  # Create an excel file.
    worksheet = workbook.add_worksheet('דו"ח מדפים')  # Create a worksheet.

    worksheet.ignore_errors({'number_stored_as_text': 'A1:E10000'})  # Ignore errors.

    cell_title = workbook.add_format({'bold': True, 'border': True})  # Make bold and border.
    cell_pass = workbook.add_format({'bg_color': '#C6EFCE', 'border': True})  # Make background color and border.
    cell_fail = workbook.add_format({'bg_color': '#FFCCCB', 'border': True})  # Make background color and border.
    cell_other = workbook.add_format({'border': True})  # Make border.

    row = 1  # Start after the titles in the excel.
    worksheet.write(f'A{row}', "Serial", cell_title)  # Set the title "serial" to A1.
    worksheet.write(f'B{row}', "Result", cell_title)  # Set the title "result" to B1.
    worksheet.write(f'C{row}', "Convertor", cell_title)  # Set the title "convertor" to C1.
    worksheet.write(f'D{row}', "Error", cell_title)  # Set the title "error" to D1.
    worksheet.write(f'E{row}', "Date", cell_title)  # Set the title "dates" to E1.

    try:  # Check if there is such a directory
        files = os.listdir(getTextLog())  # Get list of files from the "logs" folder.
    except FileNotFoundError as e:  # Set the error to "e".
        error = True  # Set that there is an error.
        messagebox.showerror("Error", str(e)[13:-1].replace("\\\\", "\\"))  # Remove useless stuff and print to the message box.
    except OSError as e:  # Set the error to "e".
        error = True  # Set the error to true.
        messagebox.showerror("Error", 'Path Error - No such path at "Logs directory".')  # Create a message box.
    else:
        error = False  # Set that there is no error.
        for file in files:  # Enter for each file.
            data = {"serial": "", "result": "", "convertor": "", "error": "", "date": ""}  # All the data.
            try:  # Check if there is a logs in the specified path.
                open(f"{path_logs}\\{file}", "r")  # Read the file.
            except PermissionError as e:
                messagebox.showerror("Error", "Permission error. please run as administrator")  # Create a message box.
                error = True  # Set the error to true.
                break
            except FileNotFoundError as e:  # Set the error to "e".
                messagebox.showerror("Error", "There is no logs in the path you specified")  # Create message box.
                error = True  # Set the error to true.
                break
            except OSError as e:  # Set the error to e.
                messagebox.showerror("Error", 'Please re-open the program')  # Create a message box.
                error = True  # Set the error to true.
                break
            else:
                error = False  # Set the error to false.
                with open(f"{path_logs}\\{file}", "r") as f:  # Read the file.
                    if ".txt" in file:  # Skip other file that are not ".txt".
                        lines = f.readlines()   # Get a list of lines from the file.
                        if len(lines) > 20:  # Check if the file does validation correctly.
                            if "PASS" in lines[-1]:  # Check if there is a "PASS" in the line.
                                result = "Pass"  # Set the result to "Pass".
                            elif "FAIL" in lines[-1]:    # Check if there is a "FAIL" in the line.
                                result = "Fail"  # Set the result to "Fail".
                            else:  # Skip a manual exit {
                                data["error"] = "not finished"  # -
                            if data["error"] != "not finished":  # }
                                data["serial"] = file[0:13]  # Set the serial number.
                                data["result"] = result  # Set the result
                                if "30000" in lines[12]:  # Check if the convertor is 30k.
                                    data["convertor"] = "30000"  # Set the convertor to 30k.
                                elif "7500" in lines[12]:  # Check if the convertor is 7.5k.
                                    data["convertor"] = "7500"  # Set the convertor to 7.5k.
                                if "average" in lines[-3]:  # Check if there is an average error.
                                    data["error"] = lines[-3][15:20]  # Set the numbers after the "average error".
                                else:
                                    if "FAIL #8 - points dump hx711 sens avg" in lines[-2]:  # Check if fail because of dump fail.
                                        data["error"] = "dump fail"  # Set "dump fail" to "error".
                                    elif "FAIL #8 - points" in lines[-2] and data["error"] != "dump fail":  # Check if the program was stopped manually.
                                        data["error"] = "program was stopped manually"  # Set "program was stopped manually" to "error".
                                date = time.ctime(os.path.getmtime(f"{path_logs}\\{file}"))  # Get the date of the file modification.
                                data["date"] = (date[8:10] + "/" + date_dic[date[4:7]]).replace(" ", "")  # Set the date to "date" at data.
                                # print(data)  # print the data
                                row += 1  # Add a new row.
                                worksheet.write(f'A{row}', data["serial"], cell_other)  # Set serials to the A column.
                                worksheet.write(f'B{row}', data["result"], cell_pass) if result is "Pass" else worksheet.write(f'B{row}', data["result"], cell_fail)  # Set results to the B column.
                                worksheet.write(f'C{row}', data["convertor"], cell_other)  # Set convertors to the C column.
                                worksheet.write(f'D{row}', data["error"], cell_other)  # Set errors to the D column.
                                worksheet.write(f'E{row}', data["date"], cell_other)  # Set dates to the E column.

        try:
            workbook.close()  # Close the excel to save.
        except exceptions.FileCreateError as e:  # Set the error to "e".
            error = True  # Set that there is an error.
            messagebox.showerror("Error", str(e)[10:-15].replace("\\\\", "\\"))  # Remove useless stuff and print to the message box.

        config.write(open('config.ini', 'w'))  # Must for update to config.ini.


def GetPass():
    if EntryPassword.get() == "here comes the money":  # Check if the password is correct.
        frame.destroy()  # close the window.
        MainProgram()  # Start the main program "main()".
    else:
        time.sleep(1)  # Sleep for 1 second to avoid brute force hack.


frame = tk.Tk()  # Set tkinter command.
frame.title("Password")  # Set error frame if didn't start correctly. {
frame.geometry('200x100+900+500')  # Set the size of the window. }
frame.resizable(width=False, height=False)  # Disable resize of the window.
if os.path.exists('trigo_icon.ico'):  # Check if the icon is exist.
    frame.iconbitmap('trigo_icon.ico')  # Set the program icon.
LabelPassword = tk.Label(frame, text="Please enter the password:")  # Create the label of the "Password"
LabelPassword.pack()  # Set the label.
LabelPassword.place(x=25, y=10)  # Set position of the label.
EntryPassword = tk.Entry(frame, show="*", width=15)  # Create the entry of the "Password"
EntryPassword.pack()  # Set the entry.
EntryPassword.place(x=50, y=35)  # Set position of the entry.
BtnPassword = tk.Button(frame, height=1, width=8, text="Continue", command=lambda: GetPass())  # Create the button to continue.
BtnPassword.pack()  # Set the button.
BtnPassword.place(x=67, y=60)  # Set position of the entry.


def MainProgram():
    global config
    global error
    config = cp.ConfigParser()  # Set config command.
    config.read('config.ini')  # Read the config.
    try:
        global path_logs
        path_logs = config['directory']['logs']  # Set the path of logs.
        path_save = config['directory']['save']  # Set the path of Save.
    except KeyError as e:  # Set the error to "e".
        with open("config.ini", "w") as configfile:
            config['directory'] = {'logs': "", "save": ""}
            config.write(configfile)
        error = True  # Set that there is an error.
        frame = tk.Tk()  # Set tkinter command.
        frame.title("Error")  # Set error frame if didn't start correctly. {
        frame.geometry('1x1+900+500')  # Set the size of the window. }
        messagebox.showerror("Error", 'Please re-open the program so the "config.ini" file will be created')  # Create a message box.
    else:
        frame = tk.Tk()  # Set tkinter command.
        frame.title("Shelves")  # Set title name of the window.
        frame.geometry('400x200+700+400')  # Set the size of the window.
        frame.resizable(width=False, height=False)

        LabelLog = tk.Label(frame, text="Logs directory:")  # Create the label of the "Logs directory:"
        LabelLog.pack()  # Set the label.
        LabelLog.place(x=30, y=20)  # Set position for the label.
        global TextLog
        TextLog = tk.Text(frame, height=1, width=40)  # Create the Text box for the logs directory.
        TextLog.insert('insert', path_logs)  # Insert the path to the text box.
        TextLog.pack()  # Set the text box.
        TextLog.place(x=30, y=40)  # Set position for the text box.
        LabelSave = tk.Label(frame, text="Save directory:")  # Create the label of the "Save directory:"
        LabelSave.pack()  # Set the label.
        LabelSave.place(x=30, y=80)  # Set position for the label.
        global TextSave
        TextSave = tk.Text(frame, height=1, width=40)  # Create the Text box for the save directory.
        TextSave.insert('insert', path_save)  # Insert the path to the text box.
        TextSave.pack()  # Set the text box.
        TextSave.place(x=30, y=100)  # Set position for the text box.
        BtnStart = tk.Button(frame, height=1, width=8, text="Start", command=lambda: [getTextLog(), getTextSave(), main(), Done()])  # Create a button that do 4 commands with lambda.
        BtnStart.pack()  # Set the button.
        BtnStart.place(x=290, y=150)  # Set position for the button.
        BtnPathLogs = tk.Button(frame, height=0, width=1, bg="yellow", command=lambda: [SetPathLogs()])  # Create a button to set the path of logs. lambda is necesery to make a "When press command".
        BtnPathLogs.pack()  # Set the button.
        BtnPathLogs.place(x=350, y=40)  # Set position for the button.
        BtnPathSave = tk.Button(frame, height=0, width=1, bg="yellow", command=lambda: [SetPathSave()])  # Create a button to set the path of logs. lambda is necesery to make a "When press command".
        BtnPathSave.pack()  # Set the button.
        BtnPathSave.place(x=350, y=100)  # Set position for the button.
        LabelComment = tk.Label(frame, justify="left", text="""© 2022 Reuven Itzhakov
Instagram: reuven.itz
Gmail: itzhakovreuven@gmail.com""")  # Create a label of copyright.
        LabelComment.pack(fill='both')  # Set the text label from left to right side.
        LabelComment.place(x=30, y=140)  # Set the label position.

        if os.path.exists('trigo_icon.ico'):  # Check if the icon is exist.
            frame.iconbitmap('trigo_icon.ico')  # Set the program icon.

        frame.mainloop()  # Loop the tkinter program


frame.mainloop()  # Loop the tkinter program
