Attribute VB_Name = "UTIL_FileExplorer"
'############################################################################################
' Open a File Explorer
' Returns the File Path that you selected
' ############################################################################################
Public Function OpenFileExplorer(Optional title As String = "Open a file", Optional default As String = vbNullString) As String

On Error GoTo Errhandling:

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    ' Only one file can be selected
    .AllowMultiSelect = False

    ' Set the title of the dialog box.
    .title = title
    
    ' Default Path
    If Trim(default) <> vbNullString Then
        .InitialFileName = default
    Else
        .InitialFileName = ThisWorkbook.path & "\"
    End If

    ' Clear out the current filters, and add our own.
    .Filters.clear
    .Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb; *.csv"
    .Filters.Add "All Files", "*.*"

    ' Show the dialog box. If the .Show method returns True, the
    ' user picked at least one file. If the .Show method returns
    ' False, the user clicked Cancel.
    If .Show = True Then
        OpenFileExplorer = .SelectedItems(1) 'replace txtFileName with your textbox
'        Debug.Print "[Selected] " & openFileExplorer
    Else
        OpenFileExplorer = vbNullString
    End If

End With

Exit Function

Errhandling:
    MsgBox "Can not open File Explorer", vbCritical, "Error"
    
End Function


'############################################################################################
' Check if the File by given name is open
' Returns TRUE or FALSE
' ############################################################################################
Function IsFileOpen(filename As String) As Boolean
       Dim filenum As Integer, errnum As Integer

       On Error Resume Next   ' Turn error checking off.
       filenum = FreeFile()   ' Get a free file number.
       ' Attempt to open the file and lock it.
       Open filename For Input Lock Read As #filenum
       Close filenum          ' Close the file.
       errnum = Err           ' Save the error number that occurred.
       On Error GoTo 0        ' Turn error checking back on.

       ' Check to see which error occurred.
       Select Case errnum

           ' No error occurred.
           ' File is NOT already open by another user.
           Case 0
               IsFileOpen = False

           ' Error number for "Permission Denied."
           ' File is already opened by another user.
           Case 70
               IsFileOpen = True

           ' Another error occurred.
           Case Else
               Error errnum
       End Select
       
End Function
