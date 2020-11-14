' ***************************************************************************
' *   This program edited by                                                *
' *   SM Mehdi Akram                                                        *
' *   CEO, Royal Technologies                                               *
' *   Cell: +8801973245450                                                  *
' *   E-Mail: mehdi.akram@gmail.com                                         *
' *   Website: http://www.royaltechbd.com                                       *
' ***************************************************************************


' How to use this function.
' ****************************************************************************************
' *  1. Start Microsoft Excel.                                                           *
' *  2. Press ALT+F11 to start the Visual Basic Editor.                                  *
' *  3. On the Insert menu, click Module.                                                *
' *  4. Type the following code into the module sheet.                                   *
' *  5. Save Module                                                                      *
' *  6. Type SpellNumber in any cell and give you desired cell name as function argument.*
' ****************************************************************************************

'Limitation: It can not convert number greater than 2147483647.

Function SpellNumber(ByVal MyNumber)

Dim Taka, Paisa, Temp
Dim DecimalPlace, Count
ReDim Place(9) As String
Place(2) = "Thousand "
Place(3) = "Lac "
Place(4) = "Crore "
Place(5) = "Arab " ' String representation of amount
MyNumber = Trim(Str(MyNumber)) ' Position of decimal place 0 if none
DecimalPlace = InStr(MyNumber, ".")
' Convert Paisa and set MyNumber to Taka amount
If DecimalPlace > 0 Then
Paisa = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
End If
Count = 1
Do While MyNumber <> ""
If Count = 1 Then Temp = GetHundreds(Right(MyNumber, 3))
If Count > 1 Then Temp = GetHundreds(Right(MyNumber, 2))
If Temp <> "" Then Taka = Temp & Place(Count) & Taka
If Count = 1 And Len(MyNumber) > 3 Then
MyNumber = Left(MyNumber, Len(MyNumber) - 3)
Else
If Count > 1 And Len(MyNumber) > 2 Then
MyNumber = Left(MyNumber, Len(MyNumber) - 2)
Else
MyNumber = ""
End If
End If
Count = Count + 1
Loop
Select Case Taka
Case ""
Taka = "No Taka "
Case "One"
Taka = "One Taka "
Case Else
'****************************************************************
'modified the following two lines to display "Taka" to precede
' rem'd the first line and added the second line
'****************************************************************
'Taka = Taka & " Taka"
Taka = "Taka " & Taka

End Select
Select Case Paisa
Case ""
'****************************************************************
'modified the following two lines to display nothing for no Paisa
' rem'd the first line and added the second line
'****************************************************************

'Paisa = " and No Paisa"
'****************************************************************
'modified the following line to display " Only" for no Paisa
' rem'd the first line and added the second line
'****************************************************************
'Paisa = ""
Paisa = "Only"
Case "One"
Paisa = "and One Paisa"
Case Else
Paisa = "and " & Paisa & "Paisa"

End Select
SpellNumber = Taka & Paisa
End Function
'*******************************************
' Converts a number from 100-999 into text *
'*******************************************
Function GetHundreds(ByVal MyNumber)
Dim Result As String
If Val(MyNumber) = 0 Then Exit Function
MyNumber = Right("000" & MyNumber, 3) 'Convert the hundreds place
If Mid(MyNumber, 1, 1) <> "0" Then
Result = GetDigit(Mid(MyNumber, 1, 1)) & "Hundred "
End If
'Convert the tens and ones place
If Mid(MyNumber, 2, 1) <> "0" Then
Result = Result & GetTens(Mid(MyNumber, 2))
Else
Result = Result & GetDigit(Mid(MyNumber, 3))
End If
GetHundreds = Result
End Function
'*********************************************
' Converts a number from 10 to 99 into text. *
'*********************************************
Function GetTens(TensText)
Dim Result As String
Result = "" ' null out the temporary function value

If Val(Left(TensText, 1)) = 1 Then ' If value between 10-19
Select Case Val(TensText)
Case 10: Result = "Ten "
Case 11: Result = "Eleven "
Case 12: Result = "Twelve "
Case 13: Result = "Thirteen "
Case 14: Result = "Fourteen "
Case 15: Result = "Fifteen "
Case 16: Result = "Sixteen "
Case 17: Result = "Seventeen "
Case 18: Result = "Eighteen "
Case 19: Result = "Nineteen "
Case Else
End Select
Else ' If value between 20-99
Select Case Val(Left(TensText, 1))
Case 2: Result = "Twenty "
Case 3: Result = "Thirty "
Case 4: Result = "Forty "
Case 5: Result = "Fifty "
Case 6: Result = "Sixty "
Case 7: Result = "Seventy "
Case 8: Result = "Eighty "
Case 9: Result = "Ninety "
Case Else
End Select
Result = Result & GetDigit(Right(TensText, 1))  'Retrieve ones place
End If
GetTens = Result
End Function
'*******************************************
' Converts a number from 1 to 9 into text. *
'*******************************************
Function GetDigit(Digit)
Select Case Val(Digit)
Case 1: GetDigit = "One "
Case 2: GetDigit = "Two "
Case 3: GetDigit = "Three "
Case 4: GetDigit = "Four "
Case 5: GetDigit = "Five "
Case 6: GetDigit = "Six "
Case 7: GetDigit = "Seven "
Case 8: GetDigit = "Eight "
Case 9: GetDigit = "Nine "
Case Else: GetDigit = ""
End Select
End Function


