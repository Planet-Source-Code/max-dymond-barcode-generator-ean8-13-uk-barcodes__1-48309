Attribute VB_Name = "modFunctions"
Function MakeBarcode(ByVal v_BarcodeRoot As String) As String
MakeBarcode = v_BarcodeRoot & CheckDigit(v_BarcodeRoot)
End Function

Function CheckDigit(ByVal v_Barcode As String) As Byte
Dim bytDigit As Byte, bytCalcTotal As Byte, sTempCode As String, bytToggle As Byte, bytCount As Byte

'This function iterates through each digit of the barcode,
'assigning alternate values of digit*3 and digit*1.
'The checkdigit is 10-(final value) Mod 10, unless (final value) Mod 10
'is 0, in which case, the check digit is 0

Select Case Len(v_Barcode)
Case 7, 12
    sTempCode = Right$("0000000000000000" & v_Barcode, 17)
    bytToggle = 3
    For bytCount = 1 To 17
        bytCalcTotal = bytCalcTotal + Val(Mid$(sTempCode, bytCount, 1)) * bytToggle
        bytToggle = 4 - bytToggle
    Next
    bytDigit = bytCalcTotal Mod 10
    bytDigit = IIf(bytDigit = 0, 0, 10 - bytDigit)
End Select
CheckDigit = bytDigit
End Function

Sub Alert(ByVal v_MessageString As String)
MsgBox v_MessageString, vbExclamation, App.Title
End Sub

Function MidInt(ByVal v_TempStr, ByVal v_Position)
MidInt = CInt(Mid(v_TempStr, v_Position, 1))
End Function
