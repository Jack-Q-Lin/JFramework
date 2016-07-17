Module VBtoSQLScripts
    Sub VBtoSQL_Quoted_Strings_Get(ByRef cLine As String, ByRef aSubStr As String, ByRef nSubStrCount As Integer)
        Dim nP1, nP2 As Integer
        Dim cLineStringTmp As String

        nP1 = 0     ' Start position of substring
        nP2 = 1     ' End position of substring
        cLineStringTmp = cLine
        cLine = ""
        Do While nP2 <= Len(cLineStringTmp)
            If Mid(cLineStringTmp, nP2, 1) = """" Then
                If nP1 = 0 Then
                    nP1 = nP2
                Else
                    aSubStr = aSubStr + Mid(cLineStringTmp, nP1, nP2 - nP1 + 1) + "<Q" + Right("00" + CStr(nSubStrCount), 2) + "/>"
                    cLine = cLine + "<Q" + Right("00" + CStr(nSubStrCount), 2) + "/>"
                    nSubStrCount = nSubStrCount + 1
                    nP1 = 0
                End If
            ElseIf nP1 = 0 Then
                cLine = cLine + Mid(cLineStringTmp, nP2, 1)
            End If
            nP2 = nP2 + 1
        Loop

    End Sub

    Sub VBtoSQL_Quoted_Strings_Set(ByRef cLine As String, ByRef aSubStr As String)
        Dim nP1, nP2 As Integer
        Dim cID As String

        If InStr(cLine, "<Q") = 0 And InStr(cLine, "/>") = 0 Then Exit Sub

        nP1 = 1     ' Start position of substring
        nP2 = 1
        Do While Len(aSubStr) - nP1 > 5
            nP2 = InStr(nP1, aSubStr, """<Q")   ' End position of substring
            cID = Mid(aSubStr, nP2 + 1, 6)
            If InStr(cLine, cID) > 0 Then
                cLine = Replace(cLine, cID, Mid(aSubStr, nP1, nP2 - nP1 + 1))
                aSubStr = Left(aSubStr, nP1 - 1) + Mid(aSubStr, nP2 + 7)
            Else
                nP1 = nP2 + 7
            End If
        Loop

    End Sub

    Function VBtoSQL_Type(ByVal sType) As String
        Dim aVBTypes(5) As String
        Dim nP As Integer

        aVBTypes(0) = "Boolean</>bool"
        aVBTypes(1) = "Integer</>int"
        aVBTypes(2) = "UInteger</>uint"
        aVBTypes(3) = "Single</>float"
        aVBTypes(4) = "Date</>DateTime"

        sType = LCase(Trim(sType))
        For nP = 0 To UBound(aVBTypes) - 1
            If LCase(aVBTypes(nP)(0)) = sType Then
                sType = aVBTypes(nP)(1)
                Exit For
            End If
        Next nP

        VBtoSQL_Type = sType
    End Function

    Sub VBtoSQL_VariableDefine(ByVal cLeftBlank As String, ByRef cVarDefs As String, ByRef aSubStrInQuotes As String, ByRef aArrayDefines As String, ByRef outCSFile As Object)
        ' cVarDefs can be:
        '   Dim lTotal As Long = 123456L
        '   Dim a, b, c As String, i, j As Integer
        '   Dim nums = New Integer() {1, 2, 3}
        '   Const MAX_STUDENTS As Integer = 25
        '   ReadOnly MIN_DIAMETER As Single = 4.93
        '   ByVal cLeftBlank As String, ByRef cVarDefs As String

        Dim nP As Integer
        Dim sType As String
        Dim aVars As String     ' a substring of cVarDefs for each "As"
        Dim aVar As String     ' a substring of aVars
        Dim cValues As String
        Dim aInBrackets(5) As Integer
        Dim cKey, cKeyMem As String
        Dim cSeperator As String
        Dim cDemension1, cDemension2 As String

        Dim a, b, c As String, i, j As Integer

        cVarDefs = cVarDefs + ","
        Do While Len(cVarDefs) > 1
            nP = InStr(cVarDefs, " As ")

            If nP = 0 Then
                aVars = Trim(Left(cVarDefs, Len(cVarDefs) - 1))
                cVarDefs = ""

                sType = "var"
                cValues = ""
            Else
                aVars = Trim(Left(cVarDefs, nP - 1))
                cVarDefs = Mid(cVarDefs, nP + 4)

                ' Get the position of the common after " As ". the end of this phase
                For nP = 0 To 4
                    aInBrackets(nP) = 0
                Next nP
                nP = 1
                Do While (aInBrackets(0) > 0 Or aInBrackets(1) > 0 Or aInBrackets(2) > 0) Or Mid(cVarDefs, nP, 1) <> ","
                    If Mid(cVarDefs, nP, 1) = "(" Then aInBrackets(0) = aInBrackets(0) + 1
                    If Mid(cVarDefs, nP, 1) = ")" Then aInBrackets(0) = aInBrackets(0) - 1
                    If Mid(cVarDefs, nP, 1) = "[" Then aInBrackets(1) = aInBrackets(1) + 1
                    If Mid(cVarDefs, nP, 1) = "]" Then aInBrackets(1) = aInBrackets(1) - 1
                    If Mid(cVarDefs, nP, 1) = "{" Then aInBrackets(2) = aInBrackets(2) + 1
                    If Mid(cVarDefs, nP, 1) = "}" Then aInBrackets(2) = aInBrackets(2) - 1
                    nP = nP + 1
                Loop
                sType = Trim(Left(cVarDefs, nP - 1)) + " "
                cVarDefs = Mid(cVarDefs, nP + 1)

                nP = 1
                Do While Mid(sType, nP, 1) <> " "
                    nP = nP + 1
                Loop
                cValues = Trim(Mid(sType, nP))
                sType = Left(sType, nP - 1)
            End If

            nP = InStr(aVars + " ", " ")
            cKey = Trim(Left(aVars, nP - 1))
            cKeyMem = cKey              ' Memery the key for next " As " phase in same line.
            If cKey = "Optional" Then
                aVars = Mid(aVars, 10)
                nP = InStr(aVars + " ", " ")
                cKey = Trim(Left(aVars, nP - 1))
            End If

            cSeperator = ", "
            If cKey = "Dim" Then
                cKey = ""
                cSeperator = "; "
                aVars = Mid(aVars, 5)
            ElseIf cKey = "Const" Then
                cKey = "const"
                cSeperator = "; "
                aVars = Mid(aVars, 7)
            ElseIf cKey = "ReadOnly" Then
                cKey = "readonly"
                cSeperator = "; "
                aVars = Mid(aVars, 10)
            ElseIf cKey = "ByVal" Then
                cKey = ""
                aVars = Mid(aVars, 7)
            ElseIf cKey = "ByRef" Then
                cKey = "ref"
                aVars = Mid(aVars, 7)
            End If

            If Len(cVarDefs) <= 1 And cSeperator = ", " Then cSeperator = "" 'for parameters in a routine

            sType = VBtoSQL_Type(sType)

            If InStr(aVars, "(") Then      ' Array
                aVars = aVars + ","
                Do While Len(aVars) > 1
                    nP = 1
                    cDemension1 = "not in"
                    Do While cDemension1 = "in" Or Mid(aVars, nP, 1) <> ","
                        If Mid(aVars, nP, 1) = "(" Then
                            cDemension1 = "in"
                        ElseIf Mid(aVars, nP, 1) = ")" Then
                            cDemension1 = "not in"
                        End If
                        nP = nP + 1
                    Loop
                    aVar = Trim(Left(aVars, nP - 1))
                    aVars = Mid(aVars, nP + 1)

                    If InStr(aVar, "(") Then      ' Array
                        nP = InStr(aVar, "(")
                        cDemension1 = ""        ' for Type
                        cDemension2 = ""        ' for New
                        Do While InStr("(0123456789, )", Mid(aVar, nP, 1)) > 0
                            If Mid(aVar, nP, 1) = "(" Or Mid(aVar, nP, 1) = ")" Or Mid(aVar, nP, 1) = "," Then cDemension1 = cDemension1 + Mid(aVar, nP, 1)
                            cDemension2 = cDemension2 + Mid(aVar, nP, 1)
                            nP = nP + 1
                            If nP > Len(aVar) Then Exit Do
                        Loop
                        cDemension1 = Replace(Replace(cDemension1, "(", "["), ")", "]")
                        cDemension2 = Replace(Replace(cDemension2, "(", "["), ")", "]")

                        ' Remember this Array name
                        aArrayDefines = "<" + Left(aVar, InStr(aVar, "(") - 1) + cDemension1 + "/>" + aArrayDefines

                        aVar = sType + cDemension1 + " " + Left(aVar, InStr(aVar, "(") - 1)
                        aVar = aVar + " = new " + sType + cDemension2

                        'aVar = aVar + " {" + cValues + "}" + cSeperator

                        Call VBtoSQL_Quoted_Strings_Set(aVar, aSubStrInQuotes)
                        outCSFile.WriteLine(cLeftBlank + aVar)
                    Else
                        If cKey <> "" Then
                            sType = cKey + " " + sType
                        End If

                        If cValues <> "" Then
                            aVar = sType + " " + Trim(aVar) + " " + cValues + cSeperator
                        Else
                            aVar = sType + " " + Trim(aVar) + cSeperator
                        End If

                        Call VBtoSQL_Quoted_Strings_Set(aVar, aSubStrInQuotes)
                        outCSFile.WriteLine(cLeftBlank + aVar)
                    End If
                Loop
            Else
                If cKey <> "" Then
                    sType = cKey + " " + sType
                End If

                If cValues <> "" Then
                    aVars = sType + " " + Trim(aVars) + " " + cValues + cSeperator
                Else
                    aVars = sType + " " + Trim(aVars) + cSeperator
                End If

                Call VBtoSQL_Quoted_Strings_Set(aVars, aSubStrInQuotes)
                outCSFile.WriteLine(cLeftBlank + aVars)
            End If

            If Len(cVarDefs) > 1 Then cVarDefs = cKeyMem + " " + cVarDefs 'prepare for next " As " phase in same line.
        Loop


    End Sub

    Function VBtoSQL_ArrayReQuote(ByVal cExpression As String, ByVal aArrayDefines As String) As String
        Dim cVName As String
        Dim nP, i, j, nQuotes As Integer

        ' Array
        nP = 1
        Do While InStr(nP, cExpression, "(") > 0
            nP = InStr(nP, cExpression, "(") + 1
            cVName = ""
            i = nP - 2
            Do While i > 0
                If InStr(" ([", Mid(cExpression, i, 1)) > 0 Then Exit Do
                cVName = Mid(cExpression, i, 1) + cVName
                i = i - 1
                If i = 0 Then Exit Do
            Loop
            If InStr(aArrayDefines, "<" + cVName + "[") > 0 Then
                i = InStr(aArrayDefines, "<" + cVName + "[") + Len(cVName) + 1
                j = nP - 1
                Do While Mid(aArrayDefines, i, 1) <> "/"
                    If Mid(aArrayDefines, i, 1) = "[" Then
                        Do While Mid(cExpression, j, 1) <> "("
                            j = j + 1
                        Loop
                        cExpression = Left(cExpression, j - 1) + "[" + Mid(cExpression, j + 1)
                        nQuotes = 0
                        Do While nQuotes > 0 Or Mid(cExpression, j, 1) <> ")"
                            If Mid(cExpression, j, 1) = "(" Then
                                nQuotes = nQuotes + 1
                            ElseIf Mid(cExpression, j, 1) = ")" Then
                                nQuotes = nQuotes - 1
                            End If
                            j = j + 1
                        Loop
                        cExpression = Left(cExpression, j - 1) + "]" + Mid(cExpression, j + 1)
                    End If
                    i = i + 1
                Loop
            End If

        Loop

        VBtoSQL_ArrayReQuote = cExpression

    End Function

    Function VBtoSQL_ConvertExpresion(ByVal cExpression As String, ByVal aArrayDefines As String, ByVal cWith As String) As String
        Dim cVName As String
        Dim nP, i, j, nQuotes As Integer

        cExpression = Replace(cExpression, " = ", " == ")
        cExpression = Replace(cExpression, " <> ", " != ")
        cExpression = Replace(cExpression, " Mod ", " % ")
        cExpression = Replace(cExpression, " \ ", " / ")
        cExpression = Replace(cExpression, " And ", " & ")
        cExpression = Replace(cExpression, " Or ", " | ")
        cExpression = Replace(cExpression, " AndAlso ", " && ")
        cExpression = Replace(cExpression, " OrElse ", " || ")
        cExpression = Replace(cExpression, " Xor ", " ^ ")
        cExpression = Replace(cExpression, " Not ", " ! ")

        cExpression = Replace(cExpression, " & ", " + ")

        cExpression = Replace(cExpression, " vbCrLf ", " ""\r\n"" ")
        cExpression = Replace(cExpression, " vbCr ", " ""\r"" ")
        cExpression = Replace(cExpression, " vbLf ", " ""\n"" ")
        cExpression = Replace(cExpression, " vbNewLine ", " ""\r\n"" ")
        cExpression = Replace(cExpression, " vbNullString ", " null ")
        cExpression = Replace(cExpression, " vbTab ", " \t ")

        ' With
        cExpression = Replace(cExpression, "(.", "( .")
        cExpression = Replace(" " + cExpression, " .", " " + cWith + ".")
        cExpression = Trim(Replace(cExpression, "( ", "("))

        VBtoSQL_ConvertExpresion = VBtoSQL_ArrayReQuote(cExpression, aArrayDefines)

    End Function

    Sub VBtoSQL_CommonStatements(ByVal cLeftBlank As String, ByRef cExpression As String, ByRef aSubStr As String, ByVal aArrayDefines As String, ByVal cWith As String, ByRef outCSFile As Object)
        cExpression = Replace(" " + cExpression, " .", " " + cWith + ".")
        cExpression = VBtoSQL_ArrayReQuote(cExpression, aArrayDefines)

        Call VBtoSQL_Quoted_Strings_Set(cExpression, aSubStr)
        outCSFile.WriteLine(cLeftBlank + cExpression + ";")

    End Sub

    Sub VBtoSQL_TextFile_To_TextFile()
        Dim nM, nLine, nP1, nP2, nGlobleVariableCount, nVariableCount As Integer
        Dim fs As Object
        Dim inVBAFile, outCSFile As Object
        Dim cLineString As String
        Dim cContinueLineString As String
        Dim aSubStrInQuotes As String
        Dim cLineStringCS As String
        Dim cParameterDefines As String
        Dim nP As Integer
        Dim aWith(20) As String
        Dim nWith As Integer
        Dim cThisLine, cLastLine, cLeftSpaces As String
        Dim cRoutineName, cRoutineType, cReturnType As String
        Dim bYes, lContinue, bContinueLine As Boolean
        Dim aArrayDefines As String

        Const cInVBAFileName = "D:\JackWorks\Programming\Code128Generator\VB-Code.txt"
        Const cOutCSFileName = "D:\JackWorks\Programming\Code128Generator\VC-Code.txt"

        '    Const cInVBAFileName = "E:\VB - Codes.txt"
        '    Const cOutCSFileName = "E:\CS - Codes.txt"

        'On Error Resume Next

        Const ForReading = 1, ForWriting = 2, ForAppending = 8
        ' The following line contains constants for the OpenTextFile
        Dim TextLine As String

        'Create out file
        fs = CreateObject("Scripting.FileSystemObject")

        inVBAFile = fs.OpenTextFile(cInVBAFileName, ForReading)
        outCSFile = fs.CreateTextFile(cOutCSFileName, ForWriting, True)
        aArrayDefines = ""
        aWith(0) = ""
        nWith = 1
        Do While inVBAFile.AtEndOfStream <> True
            cLineString = inVBAFile.ReadLine
            nVariableCount = 1

            'For testing ======================================================================================================
            If InStr(cLineString, "b((3+4)-") > 0 Then
                nP = 1
            End If
            'For testing ======================================================================================================

            ' Cut out the substrings include " and "
            aSubStrInQuotes = ""
            nP = 1      ' Index of substring
            Call VBtoSQL_Quoted_Strings_Get(cLineString, aSubStrInQuotes, nP)

            ' === Remove comments
            nP1 = InStr(cLineString, "'")
            If nP1 > 0 Then
                cLineStringCS = Mid(cLineString, nP1 + 1)
                Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                outCSFile.WriteLine(cLeftSpaces + "//" + cLineStringCS)
                cLineString = RTrim(Left(cLineString, nP1 - 1))
            Else
                cLineString = RTrim(cLineString)
            End If

            ' === Continue lines
            Do While Right(cLineString, 1) = "_" And inVBAFile.AtEndOfStream <> True
                cContinueLineString = inVBAFile.ReadLine

                Call VBtoSQL_Quoted_Strings_Get(cContinueLineString, aSubStrInQuotes, nP)

                ' === Remove comments
                nP1 = InStr(cContinueLineString, "'")
                If nP1 > 0 Then
                    cLineStringCS = Mid(cContinueLineString, nP1 + 1)
                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                    outCSFile.WriteLine(cLeftSpaces + "//" + cLineStringCS)
                    cLineString = cLineString + " " + RTrim(Left(cContinueLineString, nP1 - 1))
                End If

                cLineString = Left(cLineString, Len(cLineString) - 1) + Trim(cContinueLineString)
            Loop

            ' Get the left spaces
            cLineString = Replace(cLineString, vbTab, "   ")
            nP = 1
            Do While (Mid(cLineString, nP, 1) = " ") And (nP <= Len(cLineString))
                nP = nP + 1
            Loop
            cLeftSpaces = Left(cLineString, nP - 1)
            cLineString = Trim(cLineString)
            cThisLine = cLineString

            If (cLineString = "") And (cLastLine = "") Then GoTo lNextLine

            cLineStringCS = ""
            lContinue = True

            ' === Empty line
            If cLineString = "" Then
                If cLastLine <> "" Then outCSFile.WriteLine(cLeftSpaces + "")
                lContinue = False
            End If

            ' === Key: Sub
            If lContinue Then
                nP = InStr(" " + cLineString, " Sub ")
                If nP > 0 Then
                    cRoutineType = Trim(Left(cLineString, nP - 1))
                    cLineString = Trim(Mid(cLineString, nP + Len(" Sub ") - 1))

                    nP = InStr(cLineString, "(")
                    cRoutineName = Trim(Left(cLineString, nP - 1))
                    cLineString = Trim(Mid(cLineString, nP + 1))

                    cParameterDefines = Trim(Left(cLineString, Len(cLineString) - 1))
                    cReturnType = "void"

                    If cParameterDefines = "" Then
                        cLineStringCS = cRoutineType + " " + cReturnType + " " + cRoutineName + "() {"
                        Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                        outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    Else
                        cLineStringCS = cRoutineType + " " + cReturnType + " " + cRoutineName + "("
                        Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                        outCSFile.WriteLine(cLeftSpaces + cLineStringCS)

                        Call VBtoSQL_VariableDefine(cLeftSpaces + "      ", cParameterDefines, aSubStrInQuotes, aArrayDefines, outCSFile)

                        outCSFile.WriteLine(cLeftSpaces + "      " + ") {")

                    End If
                    aArrayDefines = "<Procedure/>" + aArrayDefines
                    lContinue = False
                End If
            End If

            ' === Key: End Sub
            If cLineString = "End Sub" Then
                outCSFile.WriteLine(cLeftSpaces + "}")
                aArrayDefines = Mid(aArrayDefines, InStr(aArrayDefines, "<Procedure/>") + Len("<Procedure/>"))
                lContinue = False
            End If

            ' === Key: Function
            If lContinue Then
                nP = InStr(" " + cLineString, " Function ")
                If nP > 0 Then
                    cRoutineType = Trim(Left(cLineString, nP - 1))
                    cLineString = Trim(Mid(cLineString, nP + Len(" Function ") - 1))

                    nP = InStr(cLineString, "(")
                    cRoutineName = Trim(Left(cLineString, nP - 1))
                    cLineString = Trim(Mid(cLineString, nP + 1))

                    nP = InStr(cLineString, ") As ")
                    cParameterDefines = Trim(Left(cLineString, nP - 1))
                    cReturnType = VBtoSQL_Type(Trim(Mid(cLineString, nP + 5)))

                    If cParameterDefines = "" Then
                        cLineStringCS = cRoutineType + " " + cReturnType + " " + cRoutineName + "() {"
                        Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                        outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    Else
                        cLineStringCS = cRoutineType + " " + cReturnType + " " + cRoutineName + "("
                        Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                        outCSFile.WriteLine(cLeftSpaces + cLineStringCS)

                        Call VBtoSQL_VariableDefine(cLeftSpaces + "      ", cParameterDefines, aSubStrInQuotes, aArrayDefines, outCSFile)

                        outCSFile.WriteLine(cLeftSpaces + "      " + ") {")
                    End If
                    aArrayDefines = "<Procedure/>" + aArrayDefines
                    lContinue = False
                End If
            End If

            ' === Key: End Fuction
            If cLineString = "End Function" Then
                outCSFile.WriteLine(cLeftSpaces + "}")
                aArrayDefines = Mid(aArrayDefines, InStr(aArrayDefines, "<Procedure/>") + Len("<Procedure/>"))
                lContinue = False
            End If

            ' === Key: Dim or Const or ReadOnly
            If lContinue Then
                If Left(cLineString, 4) = "Dim " Or Left(cLineString, 6) = "Const " Or Left(cLineString, 9) = "ReadOnly " Then
                    Call VBtoSQL_VariableDefine(cLeftSpaces, cLineString, aSubStrInQuotes, aArrayDefines, outCSFile)

                    lContinue = False
                End If
            End If

            ' === Key: If
            If lContinue Then
                If Left(cLineString, 3) = "If " Then
                    nP = InStr(cLineString, " Then")
                    cLineStringCS = "if (" + VBtoSQL_ConvertExpresion(Mid(cLineString, 4, nP - 4), aArrayDefines, aWith(nWith - 1)) + ") "
                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    outCSFile.WriteLine(cLeftSpaces + "begin")

                    If Right(cLineString, 5) = " Then" Then
                    Else
                        cLineString = Mid(cLineString, nP + 5)
                        nP = InStr(cLineString, " Else ")
                        If nP > 0 Then
                            cLineStringCS = Left(cLineString, nP - 1)
                            Call VBtoSQL_CommonStatements(cLeftSpaces, cLineStringCS, aSubStrInQuotes, aArrayDefines, aWith(nWith - 1), outCSFile)

                            outCSFile.WriteLine(cLeftSpaces + "end")
                            outCSFile.WriteLine(cLeftSpaces + "else")
                            outCSFile.WriteLine(cLeftSpaces + "begin")

                            cLineStringCS = Mid(cLineString, nP + 6)
                            Call VBtoSQL_CommonStatements(cLeftSpaces, cLineStringCS, aSubStrInQuotes, aArrayDefines, aWith(nWith - 1), outCSFile)
                        Else
                            nP = InStr(cLineString, " : ")
                            If nP > 0 Then
                                cLineStringCS = Left(cLineString, nP - 1)
                                Call VBtoSQL_CommonStatements(cLeftSpaces, cLineStringCS, aSubStrInQuotes, aArrayDefines, aWith(nWith - 1), outCSFile)

                                cLineStringCS = Mid(cLineString, nP + 3)
                                Call VBtoSQL_CommonStatements(cLeftSpaces, cLineStringCS, aSubStrInQuotes, aArrayDefines, aWith(nWith - 1), outCSFile)
                            Else
                                cLineStringCS = cLineString
                                Call VBtoSQL_CommonStatements(cLeftSpaces, cLineStringCS, aSubStrInQuotes, aArrayDefines, aWith(nWith - 1), outCSFile)
                            End If
                        End If
                        outCSFile.WriteLine(cLeftSpaces + "end")
                    End If

                    lContinue = False
                End If
            End If

            ' === Key: ElseIf
            If lContinue Then
                If Left(cLineString, 7) = "ElseIf " Then
                    outCSFile.WriteLine(cLeftSpaces + "end")

                    nP = InStr(cLineString, " Then")
                    cLineStringCS = " else if (" + VBtoSQL_ConvertExpresion(Mid(cLineString, 8, nP - 8), aArrayDefines, aWith(nWith - 1)) + ") "

                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)

                    outCSFile.WriteLine(cLeftSpaces + "begin")

                    If nP <> Len(cLineString) - 4 Then
                        cLineStringCS = Mid(cLineString, nP + 5) + ";"
                        Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                        outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                        outCSFile.WriteLine(cLeftSpaces + "end")
                    End If


                    lContinue = False
                End If
            End If

            ' === Key: Else
            If lContinue Then
                If cLineString = "Else" Then
                    outCSFile.WriteLine(cLeftSpaces + "} else {")
                    lContinue = False
                End If
            End If

            ' === Key: End If
            If lContinue Then
                If cLineString = "End If" Then
                    outCSFile.WriteLine(cLeftSpaces + "}")
                    lContinue = False
                End If
            End If

            ' === Key: Select Case
            If lContinue Then
                If Left(cLineString, 12) = "Select Case " Then
                    cLineStringCS = "switch (" + VBtoSQL_ConvertExpresion(Mid(cLineString, 13), aArrayDefines, aWith(nWith - 1)) + ") {"

                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    lContinue = False
                End If
            End If

            ' === Key: Case
            If lContinue Then
                If Left(cLineString, 5) = "Case " Then
                    If Left(cLastLine, 12) <> "Select Case " Then
                        outCSFile.WriteLine(cLeftSpaces + "    break;")
                    End If
                    cLineString = Mid(cLineString, 6)
                    nP = InStr(cLineString, ",")
                    Do While nP > 0
                        cLineStringCS = "case " + VBtoSQL_ConvertExpresion(Left(cLineString, nP - 1), aArrayDefines, aWith(nWith - 1)) + ":"

                        Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                        outCSFile.WriteLine(cLeftSpaces + cLineStringCS)

                        cLineString = Mid(cLineString, nP + 1)
                        nP = InStr(cLineString, ",")
                    Loop

                    cLineStringCS = "case " + VBtoSQL_ConvertExpresion(cLineString, aArrayDefines, aWith(nWith - 1)) + ":"

                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    lContinue = False
                End If
            End If

            ' === Key: Case Else
            If lContinue Then
                If cLineString = "Case Else" Then
                    outCSFile.WriteLine(cLeftSpaces + "    break;")
                    outCSFile.WriteLine(cLeftSpaces + "default:")
                    lContinue = False
                End If
            End If

            ' === Key: End Select
            If lContinue Then
                If cLineString = "End Select" Then
                    outCSFile.WriteLine(cLeftSpaces + "    break;")
                    outCSFile.WriteLine(cLeftSpaces + "}")
                    lContinue = False
                End If
            End If

            ' === Key: Do
            If lContinue Then
                If cLineString = "Do" Then
                    cLineStringCS = "do {"

                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    lContinue = False
                End If
            End If

            ' === Key: Loop or End While or Next
            If lContinue Then
                If cLineString = "Loop" Or cLineString = "End While" Or cLineString = "Next" Or Left(cLineString, 5) = "Next " Then
                    cLineStringCS = "}"

                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    lContinue = False
                End If
            End If

            ' === Key: While or Do While or Do Until
            If lContinue Then
                cLineStringCS = ""
                If Left(cLineString, 6) = "While " Then
                    cLineStringCS = "while (" + VBtoSQL_ConvertExpresion(Mid(cLineString, 7), aArrayDefines, aWith(nWith - 1)) + ") {"
                ElseIf Left(cLineString, 9) = "Do While " Then
                    cLineStringCS = "while (" + VBtoSQL_ConvertExpresion(Mid(cLineString, 10), aArrayDefines, aWith(nWith - 1)) + ") {"
                ElseIf Left(cLineString, 9) = "Do Until " Then
                    cLineStringCS = "while (!(" + VBtoSQL_ConvertExpresion(Mid(cLineString, 10), aArrayDefines, aWith(nWith - 1)) + ")) {"
                End If
                If cLineStringCS <> "" Then
                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    lContinue = False
                End If
            End If

            ' === Key: Loop While or Loop Until
            If lContinue Then
                cLineStringCS = ""
                If Left(cLineString, 11) = "Loop While " Then
                    cLineStringCS = "while (" + VBtoSQL_ConvertExpresion(Mid(cLineString, 12), aArrayDefines, aWith(nWith - 1)) + ");"
                ElseIf Left(cLineString, 11) = "Loop Until " Then
                    cLineStringCS = "while (!(" + VBtoSQL_ConvertExpresion(Mid(cLineString, 12), aArrayDefines, aWith(nWith - 1)) + "));"
                End If
                If cLineStringCS <> "" Then
                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    lContinue = False
                End If
            End If

            ' === Key: For Each or For
            If lContinue Then
                cLineStringCS = ""
                If Left(cLineString, 9) = "For Each " Then
                    cLineStringCS = "foreach (" + Mid(cLineString, 10) + ") {"
                ElseIf Left(cLineString, 4) = "For " Then
                    nP = InStr(cLineString, " = ")
                    nP1 = InStr(cLineString, " To ")
                    nP2 = InStr(cLineString, " Step ")
                    If nP2 > 0 Then
                        cLineStringCS = VBtoSQL_ConvertExpresion(Mid(cLineString, nP1 + 4, nP2 - nP1 - 4), aArrayDefines, aWith(nWith - 1))
                        cLineStringCS = "for (" + Mid(cLineString, 5, nP1 - 5) + "; " + Mid(cLineString, 5, nP - 5) + " <= " + cLineStringCS + "; " + Mid(cLineString, 5, nP - 5) + " += " + Mid(cLineString, nP2 + 6) + ") {"
                    Else
                        cLineStringCS = VBtoSQL_ConvertExpresion(Mid(cLineString, nP1 + 4), aArrayDefines, aWith(nWith - 1))
                        cLineStringCS = "for (" + Mid(cLineString, 5, nP1 - 5) + "; " + Mid(cLineString, 5, nP - 5) + " <= " + cLineStringCS + "; " + Mid(cLineString, 5, nP - 5) + " += 1 ) {"
                    End If
                End If
                If cLineStringCS <> "" Then
                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)
                    outCSFile.WriteLine(cLeftSpaces + cLineStringCS)
                    lContinue = False
                End If
            End If

            ' === Key: Exit Do or Exit While or Exit For
            If lContinue Then
                If cLineString = "Exit Do" Or cLineString = "Exit While" Or cLineString = "Exit For" Then
                    outCSFile.WriteLine(cLeftSpaces + "break;")
                    lContinue = False
                End If
            End If

            ' === Key: Continue For
            If lContinue Then
                If cLineString = "Continue For" Then
                    outCSFile.WriteLine(cLeftSpaces + "continue;")
                    lContinue = False
                End If
            End If

            ' === Key: With
            If lContinue Then
                If Left(cLineString, 5) = "With " Then
                    cLineStringCS = VBtoSQL_ConvertExpresion(Trim(Mid(cLineString, 6)), aArrayDefines, aWith(nWith - 1))
                    Call VBtoSQL_Quoted_Strings_Set(cLineStringCS, aSubStrInQuotes)

                    aWith(nWith) = cLineStringCS
                    nWith = nWith + 1

                    lContinue = False
                End If
            End If

            ' === Key: End With
            If lContinue Then
                If cLineString = "End With" Then
                    nWith = nWith - 1
                    lContinue = False
                End If
            End If

            If lContinue Then

                cLineStringCS = cLineString
                Call VBtoSQL_CommonStatements(cLeftSpaces, cLineStringCS, aSubStrInQuotes, aArrayDefines, aWith(nWith - 1), outCSFile)

                lContinue = False
            End If

lNextLine:
            cLastLine = cThisLine

        Loop

        inVBAFile.Close()
        outCSFile.Close()

        inVBAFile = Nothing
        outCSFile = Nothing
        fs = Nothing

        MsgBox("Done. ")

    End Sub

End Module
