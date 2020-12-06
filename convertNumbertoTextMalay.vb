Option Explicit
'Main Function
Function EjaNombor(ByVal MyNumber)
    Dim Dollars, Cents, Temp, Only
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = "Ribu "
    Place(3) = "Juta "
    Place(4) = "Bilion "
    Place(5) = "Trilion "
    Only = "Sahaja "
 
    MyNumber = Trim(Str(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & _
                  "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(Right(MyNumber, 3))
        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Dollars
        Case ""
            Dollars = ""
        Case "One"
            Dollars = "Satu Ringgit "
         Case Else
            Dollars = Dollars & "Ringgit "
    End Select
    Select Case Cents
        Case ""
            Cents = ""
        Case "One"
            Cents = "Satu Sen"
              Case Else
            Cents = "" & Cents & "Sen "
    End Select
    EjaNombor = Dollars & Cents & Only
End Function
 
Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & "Ratus "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function
 
Function GetTens(TensText)
    Dim Result As String
    Result = "" ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "Sepuluh "
            Case 11: Result = "Sebelas "
            Case 12: Result = "Dua Belas "
            Case 13: Result = "Tiga Belas "
            Case 14: Result = "Empat Belas "
            Case 15: Result = "Lima Belas "
            Case 16: Result = "Enam Belas "
            Case 17: Result = "Tujuh Belas "
            Case 18: Result = "Lapan Belas "
            Case 19: Result = "Sembilan Belas "
            Case Else
        End Select
    Else ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "Dua Puluh "
            Case 3: Result = "Tiga Puluh "
            Case 4: Result = "Empat Puluh "
            Case 5: Result = "Lima Puluh "
            Case 6: Result = "Enam Puluh "
            Case 7: Result = "Tujuh Puluh "
            Case 8: Result = "Lapan Puluh "
            Case 9: Result = "Sembilan Puluh "
            Case Else
        End Select
        Result = Result & GetDigit _
            (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = Result
End Function
 
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "Satu "
        Case 2: GetDigit = "Dua "
        Case 3: GetDigit = "Tiga "
        Case 4: GetDigit = "Empat "
        Case 5: GetDigit = "Lima "
        Case 6: GetDigit = "Enam "
        Case 7: GetDigit = "Tujuh "
        Case 8: GetDigit = "Lapan "
        Case 9: GetDigit = "Sembilan "
        Case Else: GetDigit = ""
    End Select
End Function
