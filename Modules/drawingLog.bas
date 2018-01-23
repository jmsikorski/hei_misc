Attribute VB_Name = "genLog"
Public Function getFileList(folder As String) As String()
Attribute getFileList.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Integer
    Dim sheet() As String
    Dim sheets() As String
    
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
'    With fldr
'        .Title = "Select a Folder"
'        .AllowMultiSelect = False
'        .InitialFileName = Application.DefaultFilePath
'        If .Show <> -1 Then GoTo NextCode
'        sItem = .SelectedItems(1)
'    End With
'NextCode:
'    folder = sItem
'    Set fldr = Nothing
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.getFolder(folder)
    i = 2
    'loops through each file in the directory and prints their names and path
    For Each objFile In objFolder.Files
     'print file name
        i = i + 1
    Next objFile
    ReDim sheets(i)
    i = 0
    For Each objFile In objFolder.Files
        sheet = Split(objFile.Name, "-")
        sheets(i) = sheet(0)
        i = i + 1
    Next objFile
    getFileList = sheets

End Function

Public Sub buildSheets()
Attribute buildSheets.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim sheetTitles() As String
    Dim drawingSets() As String
    Dim drawingSheets() As String
    Dim vArray() As String
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim objSubSubFolder As Object
    Dim objFile As Object
    Dim i As Integer
    Dim x As Integer
    Dim folder As String
    Dim file As String
    
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    folder = sItem
    Set fldr = Nothing
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.getFolder(folder)
    i = 0
    x = 0
    'loops through each file in the directory and prints their names and path
    For Each objFolder In objFolder.SubFolders
        Set objSubFolder = objFSO.getFolder(objFolder)
        For Each objSubFolder In objSubFolder.SubFolders
            x = x + 1
        Next objSubFolder
         i = i + 1
    Next objFolder
    ReDim drawingSets(i)
    ReDim sheetTitles(x)
    x = 0
    i = 0
    Set objFolder = objFSO.getFolder(folder)
    For Each objFolder In objFolder.SubFolders
     'print file name
        drawingSets(i) = objFolder.Name
        Set objSubFolder = objFSO.getFolder(objFolder)
        For Each objSubFolder In objSubFolder.SubFolders
            sheetTitles(x) = objSubFolder.Name
            drawingSheets = getFileList(objFSO.getFolder(objSubFolder))
            On Error GoTo 10
            ThisWorkbook.Worksheets(objSubFolder.Name).Activate
            Cells(1, i + 1).Value = drawingSets(i)
            For sht = 0 To UBound(drawingSheets) - 1
                Cells(sht + 2, i + 1).Value = drawingSheets(sht)
            Next sht
            Columns(i + 1).EntireColumn.AutoFit
            x = x + 1
        Next objSubFolder
        i = i + 1
    Next objFolder
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Sheet1").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    vArray = Split(folder, "\")
    Size = UBound(vArray)
    file = vArray(Size - 1)
    file = fodlder & "\" & file & "_" & Format(Date, "mm.dd.yy") & ".xlsm"
    ThisWorkbook.SaveAs (file)
    Exit Sub
10
    ThisWorkbook.Worksheets.Add
    ActiveSheet.Name = objSubFolder.Name
    Resume Next
End Sub
