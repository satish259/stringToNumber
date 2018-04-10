Attribute VB_Name = "StringToNumbers"
Option Explicit
' Please reference Microsoft Scripting Runtime
' Code from https://bettersolutions.com/excel/functions/function-stringtolong.htm
Public Function StringToLong(ByVal strSource, ByRef lngRes, ByRef strError) As Boolean
Dim odictionary As Scripting.Dictionary
Dim arwords As Variant
Dim slastword As String
Dim lmultiple As Long

    On Error GoTo ErrorHandler
    Set odictionary = StringToLong_Dictionary
    lngRes = Empty
    
    
    If IsNumeric(strSource) = True Then
        lngRes = CLng(strSource)
    Else
        lmultiple = 1
        strSource = LCase(strSource)
        
        If (StringToLong_Validation(odictionary, strSource, lngRes, strError) = False) Then
            StringToLong = False
            Exit Function
        End If
        
        If (odictionary.Exists(strSource) = True) Then
            lngRes = odictionary.Item(strSource)
        Else
            arwords = Split(strSource, " ")
            Do While Len(strSource) > 0
                slastword = arwords(UBound(arwords))
                Select Case slastword
                    Case "and":
                    Case "hundred":
                                     If (lmultiple = 1000) Then
                                         lmultiple = 100000
                                     Else: lmultiple = 100
                                     End If
                    Case "thousand": lmultiple = 1000
                    Case Else:
                        If (odictionary.Exists(slastword) = True) Then
                            lngRes = lngRes + (odictionary.Item(slastword) * lmultiple)
                        End If
                End Select
                strSource = Trim(Left(strSource, InStrRev(strSource, " ")))
                arwords = Split(strSource, " ")
            Loop
        End If
    End If

    strError = Empty
    StringToLong = True
    Exit Function
    
ErrorHandler:
    lngRes = Empty
    strError = "Error"
    StringToLong = False
End Function

Private Function StringToLong_Validation(ByVal objDictionary As Scripting.Dictionary, ByVal strSource As Variant, ByRef lngRes As Variant, ByRef strError As Variant) As Boolean
    
Dim arwords As Variant
Dim lcount As Long

    On Error GoTo ErrorHandler
    StringToLong_Validation = False
    
    If (strSource Like "*,*") Then
        strError = "Punctuation characters are not allowed"
        Exit Function
    End If
    
    arwords = Split(strSource, " ")
    For lcount = 0 To UBound(arwords)
        If objDictionary.Exists(arwords(lcount)) = False Then
            strError = "Spelling mistake"
            Exit Function
        End If
    Next lcount
        
    If (InStr(1, strSource, "thousand") > 0) Then
        If (Right(strSource, 8) <> "thousand") Then
        
            If (InStr(InStr(1, strSource, "thousand"), strSource, "hundred") > 0) Then
                If (InStr(1, strSource, "thousand and") > 0) Then
                    strError = "Invalid 'and' after the thousand"
                    Exit Function
                End If
            Else
                If (InStr(1, strSource, "thousand and") = 0) Then
                    strError = "Missing 'and' after the thousand"
                    Exit Function
                End If
            End If
        End If
    End If
    
    If (InStr(1, strSource, "hundred") > 0) Then
        If (Right(strSource, 7) <> "hundred") Then
            If ((InStr(1, strSource, "hundred and") = 0) And _
                (InStr(1, strSource, "hundred thousand") = 0)) Then
                strError = "Missing 'and' after the hundred"
                Exit Function
            End If
        End If
        
        If (InStr(1, strSource, "thousand") > 0) Then
            strSource = Mid(strSource, InStr(1, strSource, "thousand") + 9)
        End If
        If (InStr(1, strSource, "hundred") > 0) Then
            strSource = Left(strSource, InStr(1, strSource, "hundred") + 6)
            If ((strSource <> "one hundred") And _
                (strSource <> "two hundred") And _
                (strSource <> "three hundred") And _
                (strSource <> "four hundred") And _
                (strSource <> "five hundred") And _
                (strSource <> "six hundred") And _
                (strSource <> "seven hundred") And _
                (strSource <> "eight hundred") And _
                (strSource <> "nine hundred")) Then
                strError = "You cannot have more than 9 hundreds"
                Exit Function
            End If
        End If
    End If
    StringToLong_Validation = True
    Exit Function
    
ErrorHandler:
    lngRes = Empty
    strError = "Error"
    StringToLong_Validation = False
End Function

Private Function StringToLong_Dictionary() As Scripting.Dictionary
Dim objDictionary As Scripting.Dictionary
    Set objDictionary = New Scripting.Dictionary
    objDictionary.Add "one", 1
    objDictionary.Add "two", 2
    objDictionary.Add "three", 3
    objDictionary.Add "four", 4
    objDictionary.Add "five", 5
    objDictionary.Add "six", 6
    objDictionary.Add "seven", 7
    objDictionary.Add "eight", 8
    objDictionary.Add "nine", 9
    objDictionary.Add "ten", 10
    objDictionary.Add "eleven", 11
    objDictionary.Add "twelve", 12
    objDictionary.Add "thirteen", 13
    objDictionary.Add "fourteen", 14
    objDictionary.Add "fifteen", 15
    objDictionary.Add "sixteen", 16
    objDictionary.Add "seventeen", 17
    objDictionary.Add "eighteen", 18
    objDictionary.Add "nineteen", 19
    objDictionary.Add "twenty", 20
    objDictionary.Add "thirty", 30
    objDictionary.Add "forty", 40
    objDictionary.Add "fifty", 50
    objDictionary.Add "sixty", 60
    objDictionary.Add "seventy", 70
    objDictionary.Add "eighty", 80
    objDictionary.Add "ninety", 90
    
    objDictionary.Add "hundred", -1
    objDictionary.Add "thousand", -1
    objDictionary.Add "and", -1
    Set StringToLong_Dictionary = objDictionary
End Function
Public Function StringToRoman(ByVal strSource, ByRef strRes, ByRef strError) As Boolean
Dim arRomans As Variant
Dim lposition As Long
Dim sroman As String
Dim strResOrg As String
Dim strErrorOrg As String

    On Error GoTo ErrorHandler
    
    strRes = Empty
    strResOrg = strRes
    strErrorOrg = strError
    
    If Not IsNumeric(strSource) Then
        If StringToLong(strSource, strRes, strError) = True Then
            strSource = strRes
            strRes = strResOrg
            strError = strErrorOrg
        Else
            
            strRes = strRes
            strError = strError
            GoTo ErrorHandler
        End If
    End If
    
    If (NumberToRoman_Validation(strSource, strRes, strError) = False) Then
        StringToRoman = False
        Exit Function
    End If
        
    arRomans = Array("", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX")
    Do While (Len(strSource) > 0)
        sroman = arRomans(Left(strSource, 1))
        
        Select Case Len(strSource)
            Case 2: sroman = Replace(Replace(Replace(sroman, "X", "C"), "V", "L"), "I", "X")
            Case 3: sroman = Replace(Replace(Replace(sroman, "X", "M"), "V", "D"), "I", "C")
            Case 4: sroman = Replace(sroman, "I", "M")
            Case Else:
        End Select
        
        strRes = strRes & sroman
        strSource = Right(strSource, Len(strSource) - 1)
    Loop

    strError = Empty
    StringToRoman = True
    Exit Function
    
ErrorHandler:
    strRes = Empty
    If IsEmpty(strError) Then strError = "Error"
    StringToRoman = False
End Function

Private Function NumberToRoman_Validation(ByVal strSource As Variant, ByRef strRes As Variant, ByRef strError As Variant) As Boolean

    On Error GoTo ErrorHandler
    NumberToRoman_Validation = False
    
    If (IsNumeric(strSource) = False) Then
        strError = "Numerical values only, no text allowed"
        Exit Function
    End If
    
    If (Val(strSource) < 0) Then
        strError = "Negative values are not allowed"
        Exit Function
    End If
    
    If (Val(strSource) < 1) Or (Val(strSource) > 3000) Then
        strError = "Numbers must be between 1 and 3000 inclusive"
        Exit Function
    End If
    
    NumberToRoman_Validation = True
    Exit Function
ErrorHandler:
    NumberToRoman_Validation = False
End Function
