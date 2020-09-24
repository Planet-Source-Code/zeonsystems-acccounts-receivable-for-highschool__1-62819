Attribute VB_Name = "ModVariables"
Public CN As New ADODB.Connection
Public rs_stud As New ADODB.Recordset
Public rs_level As New ADODB.Recordset
Public rs_fees As New ADODB.Recordset
Public rs_discount As New ADODB.Recordset
Public rs_user As New ADODB.Recordset
Public rs_payment As New ADODB.Recordset
Public rs_school_yr As New ADODB.Recordset
Public rs_sections As New ADODB.Recordset

Public Const AppTitle = "FVHS Accounts Receivable"
Public SysSY As CURRENT_SY
Public SysUser As CURRENT_USER
Public Type CURRENT_USER
    UN As String
    UP As String
    UA As Boolean
End Type

Public Type CURRENT_SY
    SY As String
End Type
