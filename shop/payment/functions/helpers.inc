<%

Function FormatDt(value)
    d = Pad(Day(value),"start","0",2)
    m = Pad(Month(value),"start","0",2)
    y = Year(value)
    FormatDt= y & "-" & m & "-" & d
End Function

Function FormatDtTime(value)
    d = Pad(Day(value),"start","0",2)
    m = Pad(Month(value),"start","0",2)
    y = Year(value)
    h = Pad(Hour(value),"start","0",2)
    mi = Pad(Minute(value),"start","0",2)
    s = Pad(Second(value),"start","0",2)
    output = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
    FormatDtTime = output
End Function

Function LogTime(value)
    h = Pad(Hour(value),"start","0",2)
    mi = Pad(Minute(value),"start","0",2)
    s = Pad(Second(value),"start","0",2)
    output = h & ":" & mi & ":" & s
    LogTime = output
End Function

Function LogDt(value)
    d = Pad(Day(value),"start","0",2)
    m = Pad(Month(value),"start","0",2)
    y = Year(value)
    h = Pad(Hour(value),"start","0",2)
    mi = Pad(Minute(value),"start","0",2)
    s = Pad(Second(value),"start","0",2)
    output = y & m & d & "_" & h & "00"
    LogDt = output
End Function

' Pad strings with specified character to specified length
Function Pad(value, where, char, length)
  Select case where
    Case "start","left"
      Do While Len(value) < length
        value = char & value
      Loop
    Case "end","right"
      Do While Len(value) < length
        value = value & char
      Loop
    Case "both"
      Do While Len(value) < length
        value = char & value & char
      Loop
  End Select
  Pad = value
End Function

Function RW(value)
  value = "<pre style=""color:red;font-weight:600;font-size:14px;background:white;padding: 4px;"">" & value & "</pre>"
  Response.Write value
  RW = value
End Function

Function PluralTest(input, one, others)
  If input = 1 Then
    PluralTest = one
  Else
    PluralTest = others
  End If

End Function

Function ValidateBarcode(barcode)
    validBarcode = false
    barcodeType = "unknown"
    newBarcode = ""
    Select Case Len(barcode)
        Case 14
            validBarcode = ValidateUPC(barcode)
            if validBarcode = true then
                barcodeType = "GTIN-14"
            end if
        Case 13
            validBarcode = ValidateUPC(barcode)
            if validBarcode = true then
                barcodeType = "GTIN-13"
            end if
        Case 12
            validBarcode = ValidateUPC(barcode)
            if validBarcode = true then
                barcodeType = "GTIN-12"
                'eanBarcode = "0" & barcode    
                ' validBarcode = ValidateUPC(eanBarcode)
                ' if validBarcode = true then
                '     barcodeType = "GTIN-13"
                '     barcode = eanBarcode
                ' end if
            end if
        Case 8
            validBarcode = ValidateUPC(barcode)
            if validBarcode = true then
                barcodeType = "GTIN-8"
            end if
        Case Else
            Do While Len(barcode) < 13
                barcode = "0" & barcode
            Loop
            validBarcode = ValidateUPC(barcode)
            if validBarcode = true then
                barcodeType = "GTIN-13"
            end if
    End Select
    newBarcode = barcode
    result = "{""validBarcode"":" & validBarcode & ",""barcodeType"":""" & barcodeType & """,""barcode"":""" & newBarcode & """}"
    ValidateBarcode = result
End Function
'   ---------------------------------------------------------------
Function ValidateEAN(code)
    validEAN = true
    If Len(code) <> 8 And Len(code) <> 13 Then : validEAN = false
    ValidChars = "0123456789"
    For i = 1 to Len(code)
        If InStr(1, ValidChars, Mid(code, i, 1)) < 1 Then : validEAN = false
    Next
    code13 = code
    While Len(code13) < 13
        code13 = "0" & code13
    Wend
    origCheck = CInt(Mid(code13, Len(code13), 1))
    testCode = Left(code13, Len(code13) - 1)
    even = CInt(Mid(testCode, 2, 1)) + _
            CInt(Mid(testCode, 4, 1)) + _
            CInt(Mid(testCode, 6, 1)) + _
            CInt(Mid(testCode, 8, 1)) + _
            CInt(Mid(testCode, 10, 1)) + _
            CInt(Mid(testCode, 12, 1))
    even = even * 3
    odd =  CInt(Mid(testCode, 1, 1)) + _
            CInt(Mid(testCode, 3, 1)) + _
            CInt(Mid(testCode, 5, 1)) + _
            CInt(Mid(testCode, 7, 1)) + _
            CInt(Mid(testCode, 9, 1)) + _
            CInt(Mid(testCode, 11, 1))
    total = odd + even
    checkSum = total mod 10
    if checkSum <> 0 Then : checkSum = 10 - checkSum
    if checkSum <> origCheck Then : validEAN = false

    ValidateEAN = validEAN
End Function
'   ---------------------------------------------------------------
Function ValidateUPC(code)
    validUPC = true
    If Len(code) <> 8 And Len(code) <> 12 And Len(code) <> 13 _
        And Len(code) <> 14 Then : validUPC = false
    '//-------------------------------------------------------
    '// Not sure whether to do this as 8 digit should be EAN-8
    '//-------------------------------------------------------
    'If Len(code) = 8 Then
    '    newCode = UPCEtoUPCA(code)
    '    If Len(newCode) = 12 Then : code = newCode
    'End If
    '//-------------------------------------------------------
    ValidChars = "0123456789"
    For i = 1 to Len(code)
        If InStr(1, ValidChars, Mid(code, i, 1)) < 1 Then : validUPC = false
    Next
    origCheck = CInt(Mid(code, Len(code), 1))

    '//Reverse the remain barcode string
    If validUPC = true Then
        testCode = Mid(code, 1, Len(code) - 1)
        Dim arrDigits()
        ReDim arrDigits(Len(testCode) - 1)
        For i = 0 to Len(testCode) - 1
            arrDigits(i) = CInt(Mid(testCode, Len(testCode) - i, 1))
        Next
        oddTotal = 0
        evenTotal = 0
        For i = 0 to Ubound(arrDigits)
            If i mod 2 = 0 Then
                evenTotal = evenTotal + (arrDigits(i) * 3)
            Else
                oddTotal = oddTotal + arrDigits(i)
            End If
        Next
        checkSum = (10 - ((evenTotal + oddTotal) mod 10)) mod 10
        If checkSum <> origCheck Then : validUPC = false
    End If

    ValidateUPC = validUPC
End Function
'   ---------------------------------------------------------------
Function UPCEtoUPCA(code)
    firstDigit = left(code, 1)
    checkDigit = Mid(code, Len(code), 1)
    upcCode = Mid(code, 2, Len(code) -2)

    lastDigit = Mid(upcCode, 6, 1)
    Select Case lastDigit
        Case "0", "1", "2"
            manuf_code = Mid(upcCode, 1, 2) & lastDigit & "00"
            prod_code = "00" & Mid(upcCode, 3, 3)
        Case "3"
            manuf_code = Mid(upcCode, 1, 3) & "00"
            prod_code = "000" & id(upcCode, 4, 2)
        Case "4"
            manuf_code = Mid(upcCode, 1, 4) & "0"
            prod_code = "0000" & Mid(upcCode, 5, 1)
        Case Else
            manuf_code = Mid(upcCode, 1, 5)
            prod_code = "0000" & Mid(upcCode, 6, 1)
    End Select                
    UPCA = firstDigit & manuf_code & prod_code & checkDigit
    UPCEtoUPCA = UPCA
End Function
'   ---------------------------------------------------------------
Function Bool2Int(input)
    If input = True Then : output = 1 : Else : output = 0 :End If
    Bool2Int = output
End Function

Function BytesToStr(bytes)
		Dim Stream
		Set Stream = Server.CreateObject("Adodb.Stream")
				Stream.Type = 1 'adTypeBinary
				Stream.Open
				Stream.Write bytes
				Stream.Position = 0
				Stream.Type = 2 'adTypeText
				'Stream.Charset = "iso-8859-1"
				Stream.Charset = "utf-8"
				BytesToStr = Stream.ReadText
				Stream.Close
		Set Stream = Nothing
End Function

Function ConvertUnixTimeStampToDateTime(input_unix_timestamp)
	ConvertUnixTimeStampToDateTime = CStr(DateAdd("s", input_unix_timestamp, "01/01/1970 00:00:00"))
End Function

Function SQLTC(stmt)
    sql_stmt = Replace(stmt,";","")
	sql = "BEGIN TRY " & vbCrLf
	sql = sql & "   " & sql_stmt & " " & vbCrLf
	sql = sql & "END TRY " & vbCrLf
	sql = sql & "BEGIN CATCH " & vbCrLf
	sql = sql & "END CATCH "
	SQLTC = sql
End Function

Function GetSlug(title)
    Set regEx = New RegExp
    regEx.Global = True
    regEx.Pattern = "[^a-zA-Z0-9\-\s]"
    Result = regEx.Replace(title, " ")
    Result = LCase(Replace(Trim(Result), "  "," "))
    Result = Replace(Result, " ","-")
    GetSlug = Result
End Function

%>