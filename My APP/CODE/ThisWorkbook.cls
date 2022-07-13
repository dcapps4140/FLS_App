VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Print_Report_Example()
    Call DeleteDataFromDatabase
    Call AddRecordsIntoAccessTable
    Call Print_Report("FLS_Report1")
End Sub
Sub AddRecordsIntoAccessTable()

    '-----------------------------------------------------------------------------
    'The macro opens the Sample.accdb database and adds the 7 rows from the sheet
    '"Excel Data" in the "Customers" table of the database.
    'The code uses late binding, so no reference to external library is required.
    'Written By:    Christos Samaras
    'Date:          27/06/2020
    'E-mail:        xristos.samaras@gmail.com
    'Site:          https://myengineeringworld.netY
    'MODIFIED FOR THIS USE BY: DACID CAPPS, 22JUN2022
    '-----------------------------------------------------------------------------
    'Declaring the necessary variables.
    Dim accessFile  As String
    Dim accessTable As String
    Dim sht         As Worksheet
    Dim lastRow     As Long
    Dim lastColumn  As Integer
    Dim cn         As Object
    Dim rs          As Object
    Dim sql         As String
    Dim i           As Long
    Dim j           As Integer
    Dim errorcount  As Integer
    Dim MAX_RETRIES As Integer
    
    'Disable the screen flickering.
    Application.ScreenUpdating = False
    
    'Specify the file path of the accdb file. You can also use the full path of the file like this:
    'AccessFile = "C:\Users\Christos\Desktop\Sample.accdb"
    accessFile = "C:\Users\cappsdax\OneDrive - Intel Corporation\Documents\FLS_Database1.accdb"
    'accessFile = ThisWorkbook.Path & "\" & "FLS_Database1.accdb"
    
    'Ensure that the Access file exists.
    If FileExists(accessFile) = False Then
        MsgBox "The Access file doesn't exist!", vbCritical, "Invalid Access file path"
        Exit Sub
    End If
    
    'Set the name of the table you want to add the data.
    accessTable = "FLS_Database"
    
    'Set the worksheet that contains the data.
    On Error Resume Next
    Set sht = ThisWorkbook.Sheets("Database")
    If Err.Number <> 0 Then
        MsgBox "The given worksheet does not exist!", vbExclamation, "Invalid Sheet Name"
        Exit Sub
    End If
    Err.Clear
    
    'Find the last row and last column in the given worksheet.
    With sht
        'lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        lastRow = Sheets("Database").Range("B2668").Value + 1
        lastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    
    'Check if there are data in the worksheet.
    If lastRow < 2 Or lastColumn < 1 Then
        MsgBox "There are no data in the given worksheet!", vbCritical, "Empty Data"
        Exit Sub
    End If
    
    'Create the ADODB connection object.
    Set cn = CreateObject("ADODB.connection")
    
    'Check if the object was created.
    If Err.Number <> 0 Then
        MsgBox "The connection was not created!", vbCritical, "Connection Error"
        Exit Sub
    End If
    Err.Clear
    
    'Open the connection.
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessFile
    
    'Create the SQL statement to retrieve the table data (the entire table).
    sql = "SELECT * FROM " & accessTable
    
    'Create the ADODB recordset object.
    Set rs = CreateObject("ADODB.Recordset")
    
    'Check if the object was created.
    If Err.Number <> 0 Then
        'GoTo RetryHandler
        MsgBox "The recordset was not created!", vbCritical, "Recordset Error" & Err.Number
        Set rs = Nothing
        Set cn = Nothing
        
        Exit Sub
    End If
    Err.Clear
    
    'Set the necessary recordset properties.
    rs.CursorType = 1   'adOpenKeyset on early binding
    rs.LockType = 3     'adLockOptimistic on early binding
    
    'Open the recordset.
    rs.Open sql, cn
    
    'Add the records from Excel to Access by looping through the rows and columns of the given worksheet.
    'Here the headers are in the row 1 and they are identical to the Access table headers.
    'This is the reason why, for example, there are no spaces in the headers of the sample worksheet.
    For i = 2 To lastRow
        rs.AddNew
        For j = 1 To lastColumn
            'This is how it will look like the first time (i = 2, j = 1):
            'rs("FirstName") = "Bob"
            rs(sht.Cells(1, j).Value) = sht.Cells(i, j).Value
        Next j
        rs.Update
    Next i
    
    rs.Resync
    
    'Close the recordet and the connection.
    rs.Close
    cn.Close
    
    'Release the objects.
    Set rs = Nothing
    Set cn = Nothing
    
    'Re-enable the screen.
    Application.ScreenUpdating = True
    
    'Inform the user that the macro was executed successfully.
    MsgBox lastRow - 1 & " rows were successfully added into the '" & accessTable & "' table!", vbInformation, "Done"
    
'RetryHandler:
' Retry until MAX_RETRIES are hit to increment your error count.
'errorcount = errorcount + 1
'If errorcount < MAX_RETRIES Then

' Sleep a random amount of time, and retry the same operation.
'    Application.Wait "00:00:01"
'    Resume
'Else
' Retries did not help. Show the error, and fall out.
'    MsgBox Err.Number & " " & Err.Description
'    Exit Sub
'End If

End Sub

Sub DeleteDataFromDatabase()
    'Declaring the necessary variables.
    Dim accessFile  As String
    Dim accessTable As String
    Dim con As ADODB.Connection, rs As ADODB.Recordset
    Dim i As Long, lr As Long, effRecsCnt As Long
    Dim xSql As String
    
    'Specify the file path of the accdb file. You can also use the full path of the file like this:
    'AccessFile = "C:\Users\Christos\Desktop\Sample.accdb"
    accessFile = "C:\Users\cappsdax\OneDrive - Intel Corporation\Documents\FLS_Database1.accdb"
    
    'Set the name of the table you want to modify the data.
    accessTable = "FLS_Database"
    
    On Error GoTo Errhandler
    
    'Create the ADODB connection object.
    Set con = CreateObject("ADODB.connection")
    
    'Check if the object was created.
    If Err.Number <> 0 Then
        MsgBox "The connection was not created!", vbCritical, "Connection Error"
        Exit Sub
    End If
    Err.Clear
    
    'Open the connection.
    con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessFile
    
    'Create the SQL statement to retrieve the table data (the entire table).
    xSql = "Delete * FROM " & accessTable
    
    'Create the ADODB recordset object.
    Set rs = CreateObject("ADODB.Recordset")
    
    'Check if the object was created.
    If Err.Number <> 0 Then
        'GoTo RetryHandler
        MsgBox "The DEL recordset was not created!", vbCritical, "Recordset Error" & Err.Number
        Set rs = Nothing
        Set con = Nothing
        
        Exit Sub
    End If
    Err.Clear
    
    'Set the necessary recordset properties.
    rs.CursorType = 1   'adOpenKeyset on early binding
    rs.LockType = 3     'adLockOptimistic on early binding
    
    'Open the recordset.
    rs.Open xSql, con
    
    'Check if the ERROR occured.
    If Err.Number = 0 Then
        'GoTo RetryHandler
        MsgBox "The recordset was deleted!", vbOK, "Recordset Deleted"
    End If
    Err.Clear
    
    'Close the recordet and the connection.
    rs.Close
    con.Close
    
    'Release the objects.
    Set rs = Nothing
    Set con = Nothing
    Exit Sub
    
Errhandler:
    Debug.Print Err.Number & ":" & Err.Description

End Sub
Function FileExists(FilePath As String) As Boolean
    '--------------------------------------------------
    'Checks if a file exists (using the Dir function).
    '--------------------------------------------------
    On Error Resume Next
    If Len(FilePath) > 0 Then
        If Not Dir(FilePath, vbDirectory) = vbNullString Then FileExists = True
    End If
    On Error GoTo 0
End Function

Public Function Print_Report(ReportName As String)
    On Error GoTo SubError
    
    Dim MsAxs As Access.Application
    Dim accessFile  As String
    
    accessFile = "C:\Users\cappsdax\OneDrive - Intel Corporation\Documents\FLS_Database1.accdb"
    
    Set MsAxs = New Access.Application
    MsAxs.OpenCurrentDatabase (accessFile)
    

    MsAxs.DoCmd.OpenReport ReportName, acViewNormal, , , , acWindowNormal
    'MsAxs.DoCmd.SelectObject acReport, ReportName
    'MsAxs.DoCmd.RunCommand acCmdPrint
    
    MsAxs.CloseCurrentDatabase
    
SubExit:
    Exit Function
SubError:
    MsgBox "Print_Report error: " & vbCrLf & Err.Number & ": " & Err.Description
End Function

