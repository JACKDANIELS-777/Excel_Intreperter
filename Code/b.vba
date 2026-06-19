Sub show_custom_msgbox(Optional ByVal title As String = "Default Title", Optional ByVal msg As String = "Default Message")
    Dim filePath As String
    Dim pythonCode As String
    Dim pythonExe As String
    Dim wsh As Object
    Dim command As String

    ' Path to the Python executable (adjust as necessary)
    pythonExe = "C:\Users\elrai\PycharmProjects\open_word_py\.venv\Scripts\python.exe"

    ' Temporary file path for Python script
    filePath = "C:\Users\elrai\vba_interperter_code\a.py" ' Adjust as necessary

    ' Build Python code
    pythonCode = "import tkinter as tk" & vbCrLf
    pythonCode = pythonCode & "def create_custom_msgbox(title, msg):" & vbCrLf
    pythonCode = pythonCode & "    root = tk.Tk()" & vbCrLf
    pythonCode = pythonCode & "    root.withdraw()  # Hide the root window" & vbCrLf
    pythonCode = pythonCode & "    msgbox = tk.Toplevel()" & vbCrLf
    pythonCode = pythonCode & "    msgbox.title(title)" & vbCrLf
    pythonCode = pythonCode & "    msgbox.geometry('300x150')" & vbCrLf
    pythonCode = pythonCode & "    label = tk.Label(msgbox, text=msg, wraplength=250)" & vbCrLf
    pythonCode = pythonCode & "    label.pack(pady=20)" & vbCrLf
    pythonCode = pythonCode & "    button = tk.Button(msgbox, text='OK', command=msgbox.destroy)" & vbCrLf
    pythonCode = pythonCode & "    button.pack(pady=10)" & vbCrLf
    pythonCode = pythonCode & "    msgbox.transient()" & vbCrLf
    pythonCode = pythonCode & "    msgbox.grab_set()" & vbCrLf
    pythonCode = pythonCode & "    msgbox.resizable(False, False)" & vbCrLf

    pythonCode = pythonCode & "def create_custom_msgbox_yes_no(title, msg):" & vbCrLf
    pythonCode = pythonCode & "    import tkinter as tk" & vbCrLf
    pythonCode = pythonCode & "    root = tk.Tk()" & vbCrLf
    pythonCode = pythonCode & "    root.withdraw()  # Hide the root window" & vbCrLf
    pythonCode = pythonCode & "    msgbox = tk.Toplevel()" & vbCrLf
    pythonCode = pythonCode & "    msgbox.title(title)" & vbCrLf
    pythonCode = pythonCode & "    msgbox.geometry('300x150')" & vbCrLf
    pythonCode = pythonCode & "    label = tk.Label(msgbox, text=msg, wraplength=250)" & vbCrLf
    pythonCode = pythonCode & "    label.pack(pady=20)" & vbCrLf
    pythonCode = pythonCode & "    def on_yes(): " & vbCrLf
    pythonCode = pythonCode & "        button_yes.config(bg=input('Enter btn color: '))" & vbCrLf
    'pythonCode = pythonCode & "        button_yes.config(bg='red')" & vbCrLf
    pythonCode = pythonCode & "        print(input('ok'))" & vbCrLf
    pythonCode = pythonCode & "        with open('C:\\Users\elrai\\vba_interperter_code\\output.txt', 'w') as f: f.write('Yes')" & vbCrLf
    pythonCode = pythonCode & "        #msgbox.destroy()" & vbCrLf
    pythonCode = pythonCode & "    button_yes = tk.Button(msgbox, text='Yes', command=on_yes)" & vbCrLf
    pythonCode = pythonCode & "    button_yes.pack(side=tk.LEFT, padx=20, pady=10)" & vbCrLf
    pythonCode = pythonCode & "    def on_no(): " & vbCrLf
    pythonCode = pythonCode & "        with open('C:\\Users\\elrai\\vba_interperter_code\\output.txt', 'w') as f: f.write('No')" & vbCrLf
    pythonCode = pythonCode & "        msgbox.destroy()" & vbCrLf
    pythonCode = pythonCode & "    button_no = tk.Button(msgbox, text='No', command=on_no)" & vbCrLf
    pythonCode = pythonCode & "    button_no.pack(side=tk.LEFT, padx=20, pady=10)" & vbCrLf
    pythonCode = pythonCode & "    msgbox.transient()" & vbCrLf
    pythonCode = pythonCode & "    msgbox.grab_set()" & vbCrLf
    pythonCode = pythonCode & "    msgbox.resizable(False, False)" & vbCrLf
    pythonCode = pythonCode & "    root.mainloop()" & vbCrLf
 'ythonCode = pythonCode & "    print()" & vbCrLf
   'pythonCode = pythonCode & "        #root.mainloop()"
        
    ' Call the function with parameters
    pythonCode = pythonCode & "create_custom_msgbox_yes_no('" & title & "', '" & msg & "')" & vbCrLf

    ' Save the Python code to the file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, pythonCode
    Close #fileNum
    
    ' Create WScript.Shell object
    Set wsh = CreateObject("WScript.Shell")

    ' Construct the command to execute the Python script
    command = pythonExe & " """ & filePath & """"

    ' Execute the Python script and wait for it to finish
    wsh.Run command, 1, True ' 1 = Show window, True = Wait until it finishes
End Sub

