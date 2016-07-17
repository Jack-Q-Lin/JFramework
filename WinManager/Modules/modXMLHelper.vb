Imports System.IO
Imports System.Xml

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' XMLSimple: 
'
' Author....: Zachary Lewis
' Email.....: freeware@zcstr.com
'
' Disclaimer: Use this DEMO code entirely at your own risk
'             and accept all responsibility for your actions.
'
' Freeware..: Cost is Free
'
' Purpose...: This class XMLSimple Supports simple XML Element retreival.
'

Public Class XMLFile

    Private myXMLFileName As String
    Private myXMLReader As XmlTextReader
    Public hasErrors As Boolean
    Public lastErrorMessage As String

    Private isInitialized As Boolean


    '
    ' Internal reset of XML file Stream
    '
    Private Sub resetXMLFile()
        Dim tmpTextStream As StreamReader
        Dim returnResult As Boolean

        returnResult = False

        Try
            '
            ' My Warm in fuzz.  Not about performance, simple.
            myXMLReader = Nothing
            System.GC.Collect()

            tmpTextStream = New StreamReader(myXMLFileName)
            myXMLReader = New XmlTextReader(tmpTextStream)

            myXMLFileName = myXMLFileName
            returnResult = True

        Catch exIO As FileNotFoundException


            ' Handle File not Found issue here
            ' find Message in message file based on locale
            ' 
            lastErrorMessage = exIO.ToString
            hasErrors = True

        Catch exXML As XmlException
            '
            ' Handle XML issue here

            lastErrorMessage = exXML.ToString
            hasErrors = True

        Catch ex As Exception
            '            '
            '

            lastErrorMessage = ex.ToString
            hasErrors = True

        End Try

    End Sub

    '
    ' Initialize the XML File that you will be using
    '
    Public Function setXMLFile(ByVal pFileName As String) As Boolean

        Dim myTextStream As StreamReader
        Dim returnResult As Boolean

        returnResult = False

        Try
            myTextStream = New StreamReader(pFileName)
            myXMLReader = New XmlTextReader(myTextStream)

            myXMLFileName = pFileName
            returnResult = True
            isInitialized = True

        Catch exIO As FileNotFoundException


            ' Handle File not Found issue here
            ' find Message in message file based on locale
            ' 

            lastErrorMessage = exIO.ToString
            hasErrors = True

        Catch exXML As XmlException
            '
            ' Handle XML issue here

            lastErrorMessage = exXML.ToString
            hasErrors = True

        Catch ex As Exception
            '            '
            '

            lastErrorMessage = ex.ToString
            hasErrors = True

        End Try

        Return returnResult

    End Function

    Public Function getAllElementValues(ByVal pElementName As String, ByVal pIndex As Integer, ByVal pDepth As Integer) As String

        Dim tmpStr As String
        Dim hasFoundElement As Boolean
        Dim tmpIndex As Integer
        Dim resultValue As String
        If Not Me.isInitialized Then
            hasErrors = True
            Me.lastErrorMessage = "Please call setXMLFile to intialize the XML file name."

            resultValue = Nothing
            Exit Function

        End If

        hasFoundElement = False

        Try

            tmpIndex = 0
            myXMLReader.Read()
            While Not myXMLReader.EOF And Not hasFoundElement

                If myXMLReader.NodeType = XmlNodeType.Element Then

                    If myXMLReader.Name = pElementName And _
                         myXMLReader.Depth = pDepth _
                    Then

                        tmpIndex = tmpIndex + 1
                        If tmpIndex = pIndex Then
                            resultValue = myXMLReader.ReadElementString
                            hasFoundElement = True
                        End If

                    End If

                End If

                If myXMLReader.EOF Then
                    '
                    ' do any cleanup
                Else

                    ' read next
                    myXMLReader.Read()

                End If

            End While

        Catch ex As Exception


            lastErrorMessage = ex.ToString
            hasErrors = True

        End Try

        resetXMLFile()

        Return resultValue

    End Function
    Public Function getAllElementValues(ByVal pElementName As String, ByVal pDepth As Integer) As Collection
        Dim resultValue As Collection

        If Not Me.isInitialized Then
            hasErrors = True
            Me.lastErrorMessage = "Please call setXMLFile to intialize the XML file name."

            resultValue = Nothing
            Exit Function

        End If
        Try

            myXMLReader.Read()
            While Not myXMLReader.EOF
                If myXMLReader.NodeType = XmlNodeType.Element Then

                    If myXMLReader.Name = pElementName And _
                         myXMLReader.Depth = pDepth _
                    Then

                        If resultValue Is Nothing Then
                            resultValue = New Collection
                        End If
                        resultValue.Add(myXMLReader.ReadElementString)

                    End If

                End If


                If myXMLReader.EOF Then
                    '
                    ' do any cleanup
                Else

                    ' read next
                    myXMLReader.Read()

                End If

            End While

        Catch ex As Exception

            lastErrorMessage = ex.ToString
            hasErrors = True

        End Try
        resetXMLFile()
        Return resultValue

    End Function

    Public Function getElementValue(ByVal pElementName As String, ByVal pIndex As Integer, ByVal pDepth As Integer) As String

        Dim tmpStr As String
        Dim hasFoundElement As Boolean
        Dim tmpIndex As Integer
        Dim resultValue As String

        If Not Me.isInitialized Then
            hasErrors = True
            Me.lastErrorMessage = "Please call setXMLFile to intialize the XML file name."

            resultValue = Nothing
            Exit Function

        End If
        hasFoundElement = False

        Try

            tmpIndex = 0
            myXMLReader.Read()
            While Not myXMLReader.EOF And Not hasFoundElement

                If myXMLReader.NodeType = XmlNodeType.Element Then

                    If myXMLReader.Name = pElementName And _
                         myXMLReader.Depth = pDepth _
                    Then

                        tmpIndex = tmpIndex + 1
                        If tmpIndex = pIndex Then
                            resultValue = myXMLReader.ReadElementString
                            hasFoundElement = True
                        End If

                    End If

                End If


                If myXMLReader.EOF Then
                    '
                    ' do any cleanup future

                Else

                    ' read next XML Line 
                    myXMLReader.Read()

                End If

            End While

        Catch ex As Exception


            lastErrorMessage = ex.ToString
            hasErrors = True

        End Try

        resetXMLFile()

        Return resultValue

    End Function

    '
    ' Returns a collection of hashtables matching the requested Node Element Names and Depth
    '
    Public Function getAllElementValues(ByVal pElementNames As Collection) As Collection
        Dim resultValue As Collection
        Dim tmpElementList As Collection

        Dim tmpResultHashTable As Hashtable

        Try

            myXMLReader.Read()
            While Not myXMLReader.EOF


                If myXMLReader.NodeType = XmlNodeType.Element Then

                    For Each tmpStr As String In pElementNames

                        If myXMLReader.Name = tmpStr _
                        Then
                            If resultValue Is Nothing Then
                                resultValue = New Collection
                            End If

                            Dim tmpStr2 As String
                            tmpStr2 = myXMLReader.ReadElementString

                            If tmpResultHashTable Is Nothing Then

                                tmpElementList = New Collection
                                tmpResultHashTable = New Hashtable

                                tmpElementList.Add(tmpStr2)

                                tmpResultHashTable.Add(tmpStr, tmpElementList)
                                resultValue.Add(tmpResultHashTable)

                            Else

                                tmpElementList = tmpResultHashTable(tmpStr)
                                If tmpElementList Is Nothing Then
                                    tmpElementList = New Collection
                                    tmpElementList.Add(tmpStr2)

                                Else

                                    tmpElementList.Add(tmpStr2)

                                End If

                                '
                                ' Must be a better way to update a Hashtable
                                '
                                If tmpResultHashTable(tmpStr) Is Nothing Then

                                    tmpResultHashTable.Add(tmpStr, tmpElementList)

                                Else

                                    tmpResultHashTable.Remove(tmpStr)
                                    tmpResultHashTable.Add(tmpStr, tmpElementList)

                                End If

                                resultValue.Add(tmpResultHashTable.Clone)

                                tmpResultHashTable.Clear()
                                tmpElementList = Nothing

                            End If

                        End If
                    Next

                End If



                If myXMLReader.EOF Then
                    '
                    ' do any cleanup
                Else

                    ' read next
                    myXMLReader.Read()

                End If

            End While

        Catch ex As Exception

            lastErrorMessage = ex.ToString
            hasErrors = True

        End Try
        resetXMLFile()
        Return resultValue

    End Function

End Class
