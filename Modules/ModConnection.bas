Attribute VB_Name = "ModConnection"
Public dbPath As String

Public Sub ConnectDB()
'On Error Resume Next
    'Get the path of the database
    dbPath = App.Path & "\Database\MasterFile.mdb"
    'Open the database
    With CN
        .CommandTimeout = 5
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & dbPath & ";Persist Security Info=False"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub

Sub Main()
If App.PrevInstance = True Then MsgBox "Application is already running", vbExclamation: Exit Sub    'Check if the application is running
MDIForm1.Show   'Load the MainForm
frmLogin.Show vbModal   'Load the login form
End Sub
