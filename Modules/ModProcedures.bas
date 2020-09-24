Attribute VB_Name = "ModProcedures"
'check the id number if it is in the record or not
Public Function check_id(ByRef sText As TextBox) As Boolean
Dim rsID As New ADODB.Recordset

rsID.Open "Select * From tblStudent", CN, adOpenStatic, adLockReadOnly
rsID.Requery
rsID.Find "ID='" & sText & "'"

If rsID.EOF Then
    check_id = False
Else
    check_id = True
    MsgBox "ID Number is alredy existing. Check the ID Number.", vbExclamation, AppTitle
    Call highlight_focus(sText)
End If
End Function
'Highlight the whole text in the textbox
Public Sub highlight_focus(ByRef sText As TextBox)
sText.SetFocus
With sText
    .SelStart = 0
    .SelLength = Len(sText.Text)
End With
End Sub
'check if a control is empty
Public Function is_empty(ByRef sText As Variant) As Boolean
If sText.Text = "" Then
    is_empty = True
    MsgBox "The field is required.Please check it!", vbExclamation, AppTitle
    sText.SetFocus
Else
    is_empty = False
End If
End Function
'check if the value inputed is a number
Public Function not_curr(ByRef sText As Variant) As Boolean
If Not IsNumeric(sText.Text) Then
    not_curr = True
    MsgBox "Cannot accept non-numeric input!", vbOKOnly + vbExclamation, AppTitle
    sText.SetFocus
    SendKeys "{Home}+{End}"
Else
    not_curr = False
End If
End Function
'This will fill the combo box
Public Sub fill_yr(ByVal sCombo As ComboBox)
Dim rs As New ADODB.Recordset

rs.Open "Select * From tblLevel Order By YR Asc", CN, adOpenStatic, adLockReadOnly

rs.Requery
With sCombo
    .Clear
    While Not rs.EOF
        .AddItem rs!yr
        rs.MoveNext
    Wend
End With

Set rs = Nothing
End Sub
'This will fill the combo box according to discount field
Public Function fill_discount(ByVal sCombo As ComboBox)
Dim rs As New ADODB.Recordset

rs.Open "Select * From tblDiscountType Order By DiscountName Asc", CN, adOpenStatic, adLockReadOnly

rs.Requery
With sCombo
    .Clear
    .AddItem "Brothers/Sisters"
    While Not rs.EOF
        .AddItem rs!DiscountName
        rs.MoveNext
    Wend
End With

Set rs = Nothing

End Function
'This will fill the combo box according to payable fees
Public Function fill_fee(ByVal sCombo As ComboBox)
Dim rs As New ADODB.Recordset

rs.Open "Select * From tblFeesType Order By FeeName Asc", CN, adOpenStatic, adLockReadOnly

rs.Requery
With sCombo
    .Clear
    While Not rs.EOF
        .AddItem rs!FeeName
        rs.MoveNext
    Wend
End With

Set rs = Nothing

End Function
'get the total discount of a certain record
Public Function Acct_Disc(ByVal acct As String, ByRef sText As TextBox)
Dim rs_disc As New ADODB.Recordset

rs_disc.Open "Select * From qryDiscount Where AccountNo = '" & acct & "'", CN, adOpenStatic, adLockReadOnly
sText = "0.00"
If rs_disc.EOF Then
    sText = "0.00"
Else
    sText = Format$(rs_disc!SumOfAmount, "#,##0.00")
End If

Set rs_disc = Nothing
End Function

Public Function Acct_Fees(ByVal acct As String, ByRef sText As TextBox)
On Error Resume Next
Dim rs_fee As New ADODB.Recordset


rs_fee.Open "Select * From qryFees Where AccountNo = '" & acct & "'", CN, adOpenStatic, adLockReadOnly
sText = "0.00"
If rs_fee.EOF Then
    sText = "0.00"
Else
    sText = Format$(rs_fee!SumOfAmount, "#,##0.00")
End If

Set rs_fee = Nothing
End Function

Public Function add_tuition(ByVal AcctNo As String, ByRef sYR As String)
Dim rsLevelFee As New ADODB.Recordset
Dim rs_tuition As New ADODB.Recordset

rs_tuition.Open "Select * From tblFees Where AccountNo ='" & AcctNo & "'", CN, adOpenStatic, adLockOptimistic
rsLevelFee.Open "Select * From tblLevel Where YR ='" & sYR & "'", CN, adOpenStatic, adLockReadOnly
With rs_tuition
    .AddNew
    .Fields("AccountNo") = AcctNo
    .Fields("FeeName") = "Tuition Fee"
    .Fields("Amount") = rsLevelFee!Amount
    .Update
    .Requery
End With

Set rsLevelFee = Nothing
Set rs_tuition = Nothing
End Function

Public Sub add_fees(ByVal AcctNo As String, ByRef SY As String, ByVal yr As Label)
Dim rs As New ADODB.Recordset
Dim rs_fees_type As New ADODB.Recordset

rs_fees_type.Open "Select * From tblFeesType", CN, adOpenStatic, adLockReadOnly
rs.Open "Select * From tblFees", CN, adOpenKeyset, adLockOptimistic
rs_fees_type.Requery
Do While Not rs_fees_type.EOF
    With rs
        .AddNew
        .Fields("AccountNo") = AcctNo
        .Fields("FeeName") = rs_fees_type!FeeName
        .Fields("Amount") = rs_fees_type!Amount
        .Update
        .Requery
    End With
    rs_fees_type.MoveNext
    If rs_fees_type.Fields("FeeName") = "Graduation Fee" And Mid(yr, 1, 1) < 4 Then
        rs_fees_type.MoveNext
    End If
Loop
Set rs = Nothing
Set rs_fees_type = Nothing
End Sub


Public Function get_Balance(ByRef acct As String, ByRef sBal As TextBox, ByVal sTotal As String)

Dim rs_bal As New ADODB.Recordset
Dim rs_qrybal As New ADODB.Recordset

rs_bal.Open "Select * From qryPayment Where AccountNo = '" & acct & "'", CN, adOpenStatic, adLockReadOnly
rs_bal.Requery
sText = "0.00"
If rs_bal.EOF Then
    sBal = sTotal
Else
    sBal = Format(sTotal - rs_bal!SumOfAmount, "#,##0.00")
End If

Set rs_bal = Nothing
End Function

Public Function check_data(ByVal sText As Variant) As Boolean
Dim rs_year_level As New ADODB.Recordset

rs_year_level.Open "Select * From tblStudent", CN, adOpenStatic, adLockReadOnly

rs_year_level.Requery

rs_year_level.Find "YR='" & sText & "'"

If rs_year_level.EOF Then
    check_data = False
Else
    check_data = True
End If

Set rs_year_level = Nothing
End Function

Public Function chk_data(ByVal sText As Variant) As Boolean
Dim rs_year_level As New ADODB.Recordset

rs_year_level.Open "Select * From tblDiscount", CN, adOpenStatic, adLockReadOnly

rs_year_level.Requery

rs_year_level.Find "DiscountName='" & sText & "'"

If rs_year_level.EOF Then
    chk_data = False
Else
    chk_data = True
End If

Set rs_year_level = Nothing
End Function

Private Sub CopyFeesType()
Dim rs_fees As New ADODB.Recordset
rs_FeesType.Open "SELECT * FROM tblFeesType", CN, adOpenStatic, adLockReadOnly
rs_fees.Open "SELECT * FROM tblFees", CN, adOpenStatic, adLockReadOnly
With rs_fees
        .Fields("AccountNo") = AcctNo
        .Fields("FeeName") = rs_FeesType!FeeName
        .Fields("Amount") = rs_FeesType!Amount
        .Fields("Standsard") = rs_FeesType!Standard
        .Update
        .Requery
    End With
End Sub

Public Sub add_discount(ByVal AccountNo As String)

Dim rs As New ADODB.Recordset

rs.Open "Select * From tblDiscount", CN, adOpenKeyset, adLockOptimistic

With rs
    .AddNew
    .Fields("AccountNo") = AccountNo
    .Fields("DiscountName") = "None"
    .Fields("Amount") = "0.00"
    .Update
    .Requery
End With

Set rs = Nothing
End Sub

Public Sub ToUpper(ByRef sText As TextBox)
    sText = UCase(sText)
    'sText = UCase(Mid$(sText, 1, 1)) + Mid$(sText, 2, Len(sText))
                         
End Sub

Public Sub Message()
    MsgBox "No record found.", vbInformation
End Sub

Public Sub UnEnrolled()
Dim rs As New ADODB.Recordset

rs.Open "Select * From tblStudent", CN, adOpenKeyset, adLockOptimistic
With rs
    .Requery
    Do While Not .EOF
        .Fields("Enrolled") = False
        .Update
        .MoveNext
    Loop
End With

Set rs = Nothing
End Sub


'Public Sub fill_sec_yr(ByVal sYR As ComboBox, ByRef sSection As ComboBox)
'Dim rs As New ADODB.Recordset

'rs.Open "Select * From tblSections Where YR ='" & sYR.Text & "'", CN, adOpenStatic, adLockReadOnly

'With sSection
'    .Clear
'    Do While Not rs.EOF
'        .AddItem rs!Section
'        rs.MoveNext
'    Loop
'End With

'Set rs = Nothing
'End Sub

Public Sub del_record(ByVal ID As TextBox, ByVal acct As String)
Dim rs As New ADODB.Recordset
Dim rs_fee As New ADODB.Recordset

rs.Open "Select * From tblStudent Where ID = '" & ID.Text & "'", CN, adOpenStatic, adLockReadOnly
rs_fee.Open "Select * From tblFee Where AccountNo ='" & acct & "'", CN, adOpenKeyset, adLockOptimistic

If rs.Fields("YR") < "4" Then
    rs_fee.Requery
    rs_fee.Find "FeeName = " & "Graduation Fee" & ""
    rs_fee.Delete
    rs_fee.Requery
End If


Set rs = Nothing
Set rs_fee = Nothing
    
End Sub

Public Sub old_account(ByVal ID As TextBox, ByVal SY As String, ByVal acct As String)
Dim rs As New ADODB.Recordset
Dim rs_acct As New ADODB.Recordset
rs.Open "Select * From tblAccount Where ID = '" & ID.Text & "' And SY <> '" & SysSY.SY & "' And Status = " & False & "", CN, adOpenKeyset, adLockOptimistic
rs_acct.Open "Select * From tblFees", CN, adOpenKeyset, adLockOptimistic
If rs.RecordCount >= 1 Then
        MsgBox ("You have an unpaid old account. This will be added to your current account."), vbExclamation, AppTitle
        With rs_acct
            .AddNew
            .Fields("AccountNo") = acct
            .Fields("FeeName") = "Old Account"
            .Fields("Amount") = rs!Balance
            .Update
            .Requery
        End With
        rs.Fields("Status") = True
        rs.Update
        rs.Requery
End If

End Sub
