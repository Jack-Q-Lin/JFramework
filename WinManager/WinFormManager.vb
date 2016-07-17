'Imports System.IO
'Imports System.Windows.Forms
'Imports System.Xml

'Public Class WinFormManager

'    'Public Const gnVerticalScrollBarWidth As Integer = 20       'Keep this value, no change.

'    'Public Function SetWindowsFolder(ByVal SearchDescription As String, Optional ByVal sInitialFolder As String = "", Optional ByVal NewFolderButton As Boolean = False) As String
'    '    'Typically, after creating a new FolderBrowserDialog, you set the RootFolder to the location from which to start browsing. 
'    '    'Optionally, you can set the SelectedPath to an absolute path of a subfolder of RootFolder that will initially be selected.
'    '    'Are you possibly setting the SelectedPath to a location that doesn't equate to a subfolder of RootFolder (i.e. My Computer)? 
'    '    'That would probably cause it to dive back to the RootFolder as the presented location.

'    '    Dim SearchPathDialog As New Windows.Forms.FolderBrowserDialog
'    '    If sInitialFolder = "" Then sInitialFolder = Environment.SpecialFolder.MyComputer
'    '    SearchPathDialog.ShowNewFolderButton = NewFolderButton
'    '    SearchPathDialog.Description = SearchDescription
'    '    SearchPathDialog.RootFolder = Environment.SpecialFolder.MyComputer
'    '    SearchPathDialog.SelectedPath = sInitialFolder
'    '    SearchPathDialog.ShowDialog()
'    '    Return SearchPathDialog.SelectedPath
'    'End Function

'    'Public Function Title_String(ByVal sTitle As String, Optional ByVal sStyle As String = "Normal") As String
'    '    Dim sReturn, sChar As String
'    '    Dim bStartNewWord As Boolean
'    '    Dim nP As Integer

'    '    sReturn = ""
'    '    sTitle = Replace(Replace(sTitle, "_", " "), "  ", " ")

'    '    If sStyle.ToLower = "ucase" Then
'    '        sReturn = sTitle.ToUpper
'    '    ElseIf sStyle.ToLower = "lcase" Then
'    '        sReturn = sTitle.ToLower
'    '    ElseIf sStyle.ToLower = "normal" Then
'    '        bStartNewWord = True
'    '        For nP = 0 To sTitle.Length - 1
'    '            sChar = sTitle.Substring(nP, 1)
'    '            If sChar >= "a" And sChar <= "z" Or sChar >= "A" And sChar <= "Z" Or sChar >= "0" And sChar <= "9" Then
'    '                If bStartNewWord Then
'    '                    sReturn = sReturn + sChar.ToUpper
'    '                    bStartNewWord = False
'    '                Else
'    '                    sReturn = sReturn + sChar.ToLower
'    '                End If
'    '            Else
'    '                sReturn = sReturn + sChar.ToLower
'    '                bStartNewWord = True
'    '            End If
'    '        Next
'    '    End If

'    '    Title_String = sReturn
'    'End Function

'    'Public Sub Generic_Form_Load(ByRef frmCurrent As System.Windows.Forms.Form)

'    '    ''Generic routine for all form loads
'    '    'Dim cntCurrent As System.Windows.Forms.Control
'    '    'Dim iCounter As Short
'    '    'Dim sControlIndicator As String

'    '    'Dim sMsg As String
'    '    'Dim sTag As String
'    '    'Dim sTranslateValue As String = String.Empty
'    '    'Dim sFormID As String

'    '    'Dim iFontSize As Short
'    '    'Dim bBold As Boolean

'    '    'Dim oScrollForm As Object

'    '    '' If there are some controls those will be on all of the forms, we can do some generic work here for them.

'    '    ''In order to not trigger this form's Activate event during executing following codes, set EnableFormActivate = False
'    '    ''Dim lEFA As Boolean = CType(frmCurrent, MSC.BaseForm).EnableFormActivate
'    '    ''CType(frmCurrent, MSC.BaseForm).EnableFormActivate = False
'    '    'frmCurren.EnableFormActivate = False

'    '    'On Error Resume Next

'    '    '' If you want to set the size of all forms to be same, enable following codes
'    '    ''frmCurrent.Width = giDeviceScreenWidth
'    '    ''frmCurrent.Height = giDeviceScreenHeight

'    '    ''Check flag if user requires multi-language support
'    '    'sFormID = Left(CStr(frmCurrent.Name), 30)
'    '    'If (UCase(gsUseMultiLanguage) = "Y") Then
'    '    '    Call ControlTransValueSend(frmCurrent)

'    '    '    frmCurrent.Text = TranslateValueGet(Left(frmCurrent.Name, 30), frmCurrent.Text)

'    '    '    If (gsFont <> "") Then frmCurrent.Font = New Font(gsFont, frmCurrent.Font.Size, frmCurrent.Font.Style)

'    '    '    Call TransactionProcess(sCurrent)

'    '    'End If

'    '    '' Customized for Pocket PC

'    '    'oScrollForm = Nothing

'    '    '' Setting form level properties
'    '    'frmCurrent.KeyPreview = True

'    '    'Err.Clear()

'    '    ''loop through control and set their fonts
'    '    'For iCounter = 0 To frmCurrent.Controls.Count() - 1

'    '    '    Err.Clear()
'    '    '    cntCurrent = CType(frmCurrent.Controls(iCounter), Object)

'    '    '    If (Err.Number <> 0) Then
'    '    '        'no control found
'    '    '        Err.Clear()
'    '    '        Exit For
'    '    '    End If

'    '    '    sControlIndicator = UCase(Mid(cntCurrent.Name, 1, 3))

'    '    '    Err.Clear()

'    '    '    'if font size is available, set it and the font bold
'    '    '    iFontSize = cntCurrent.Font.Size
'    '    '    If (Err.Number = 0) Then

'    '    '        If (cntCurrent.Font.Style = FontStyle.Bold) Then
'    '    '            bBold = True
'    '    '        End If

'    '    '        If (UCase(gsUseMultiLanguage) = "Y") Then
'    '    '            sTranslateValue = TranslateValueGet(sFormID, cntCurrent.Text)
'    '    '        End If

'    '    '        'if multilanguage is on and there is a font passed from the server, change
'    '    '        'the Font.Name of the controla
'    '    '        If (UCase(gsUseMultiLanguage) = "Y" And gsFont <> "") Then
'    '    '            cntCurrent.Font = New Font(gsFont, cntCurrent.Font.Size, cntCurrent.Font.Style)
'    '    '        End If

'    '    '        Select Case sControlIndicator
'    '    '            Case gsTypeTextBox
'    '    '                iFontSize = giFontSize

'    '    '                If (UCase(gsUseMultiLanguage) = "Y" And sTranslateValue <> "") Then
'    '    '                    cntCurrent.Text = sTranslateValue
'    '    '                End If

'    '    '                bBold = True

'    '    '            Case gsTypeLabel
'    '    '                iFontSize = giFontSizeLabel
'    '    '                cntCurrent.BackColor = frmCurrent.BackColor

'    '    '                If (UCase(gsUseMultiLanguage) = "Y" And sTranslateValue <> "") Then
'    '    '                    cntCurrent.Text = sTranslateValue
'    '    '                End If

'    '    '                bBold = True

'    '    '                If cntCurrent.Name.ToUpper() = "LBLKEYHELP" Then
'    '    '                    If gsShowKeyHelp = "Y" Then
'    '    '                        cntCurrent.Visible = True
'    '    '                    Else
'    '    '                        cntCurrent.Visible = False
'    '    '                    End If
'    '    '                End If
'    '    '            Case gsTypeButton
'    '    '                iFontSize = giFontSizeButton
'    '    '                bBold = False

'    '    '                If (cntCurrent.Width > 0) Then
'    '    '                    cntCurrent.Width = 75
'    '    '                    cntCurrent.TabStop = False
'    '    '                End If

'    '    '                If (UCase(gsUseMultiLanguage) = "Y" And sTranslateValue <> "") Then
'    '    '                    cntCurrent.Text = sTranslateValue
'    '    '                End If

'    '    '                'We never want to use this default property
'    '    '                ' This prevents us from detecting a carriage
'    '    '                ' return during KeyPress event.
'    '    '                ' We want to customize how we deal with carriage returns
'    '    '                ' Basicly if We are on last editable field, we will select
'    '    '                ' the submit button.  If we are not on the last editable
'    '    '                ' field, we will send a tab to the key board buffer.
'    '    '                'cntCurrent.Default = False

'    '    '            Case gsTypeListBox
'    '    '                iFontSize = giFontSizeListBox
'    '    '                bBold = True

'    '    '            Case "FRM"
'    '    '                If (UCase(gsUseMultiLanguage) = "Y" And sTranslateValue <> "") Then
'    '    '                    cntCurrent.Text = sTranslateValue
'    '    '                End If

'    '    '            Case Else
'    '    '                iFontSize = giFontSize
'    '    '                bBold = True
'    '    '        End Select
'    '    '        cntCurrent.Font = New Font(cntCurrent.Font.Name, cntCurrent.Font.Size, FontStyle.Bold)

'    '    '    End If

'    '    '    If (sControlIndicator = "TXT" Or cntCurrent.GetType().Name = "TextBox") Then
'    '    '        'working with a text box.   Check the
'    '    '        ' ini configurations file and set the textboxes edit
'    '    '        ' length and edit mask
'    '    '        Dim txtCntCurrent As TextBox = CType(cntCurrent, TextBox)
'    '    '        If (ConfigurationSubSetLoad(cntCurrent.Name) > 0) Then
'    '    '            txtCntCurrent.MaxLength = CInt(ConfigurationSubSetGetData2(cntCurrent.Name, "MaxLength"))

'    '    '            sTag = ""

'    '    '            'add the mask and dataidentifier to the tag these are used to filter data from the scanner
'    '    '            ' and verify user input
'    '    '            ' DataID is a data identifier, a prefix that must be on a barcode for the scanner to accept the data scanned
'    '    '            sTag = gsMask & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, "Mask")
'    '    '            sTag = sTag & gsTagDelimiter & gsDataId & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, "DATAID")

'    '    '            'add the mode to the tag
'    '    '            sTag = sTag & gsTagDelimiter & gsMode & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, "MODE")

'    '    '            sTag = sTag & gsTagDelimiter & gsDisplay & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, "DISPLAY")

'    '    '            'Add additional fields to control multiple scans into a single field
'    '    '            sTag = sTag & gsTagDelimiter & gsFieldConfig_NumberOfScans & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, gsFieldConfig_NumberOfScans)
'    '    '            sTag = sTag & gsTagDelimiter & gsFieldConfig_ScanDelimiter & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, gsFieldConfig_ScanDelimiter)

'    '    '            'add Strip Leading alpha field to control scans.
'    '    '            sTag = sTag & gsTagDelimiter & gsFieldConfig_StripLeadingAlpha & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, gsFieldConfig_StripLeadingAlpha)

'    '    '            'Added to format the string based on the INI format
'    '    '            sTag = sTag & gsTagDelimiter & gsFieldConfig_Format & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, gsFieldConfig_Format)

'    '    '            cntCurrent.Tag = sTag
'    '    '        End If
'    '    '        If (txtCntCurrent.MaxLength = 0) Then
'    '    '            'length not set, initialize to 20
'    '    '            txtCntCurrent.MaxLength = 20
'    '    '        End If
'    '    '        'cleanup, failure to do so results in memory leak
'    '    '        Call ConfigurationSubSetClear()
'    '    '    End If

'    '    '    'loading INI configuration on Combo Boxes
'    '    '    If (sControlIndicator = gsTypeCombo) Then
'    '    '        If (ConfigurationSubSetLoad(cntCurrent.Name) > 0) Then
'    '    '            sTag = ""
'    '    '            sTag = sTag & gsTagDelimiter & gsDataId & "=" & ConfigurationSubSetGetData2(cntCurrent.Name, "DATAID")

'    '    '            cntCurrent.Tag = sTag
'    '    '        End If

'    '    '        'cleanup, failure to do so results in memory leak
'    '    '        Call ConfigurationSubSetClear()

'    '    '    End If


'    '    '    Err.Clear()
'    '    'Next iCounter

'    '    '' After calling FormDimensionHideStart, the form will lose the focus
'    '    '' therefore, we call this function to set the focus back to the form
'    '    ''Call PostMessage(frmCurrent.hwnd, WM_SETFOCUS, 0, 0)

'    '    ''memory cleanup
'    '    'cntCurrent = Nothing
'    '    'Call FormSetStandardButton(frmCurrent)
'    '    'Err.Clear()

'    '    ''If there are some controls those are on all of the forms, and they are keep same size and location, rearrange their size and location here.
'    '    'Call ReArrangeControls(frmCurrent)

'    '    'CType(frmCurrent, MSC.BaseForm).EnableFormActivate = lEFA   'True

'    'End Sub

'#Region "Common used functions for DataGridView control>"

'    ''' <summary>
'    ''' 
'    ''' </summary>
'    ''' <param name="dgv"></param>
'    ''' <param name="aDgvColumns"> Here is a sample:
'    ''' The FieldNames must be match with the Datasource bound to this DGV.
'    ''' The Width can be percentage value or integer value, percentage Width is the percentage of the dgv's clientsize.width, the integer width is the absolute width of the column.
'    ''' The SUM of the percentage width must be 100. Percentage Width is used for the columns at left.
'    '''Dim aDgvColumns(,) As String = { _
'    '''    {"FieldName     ", "DataType", "HeaderText      ", "Width  ", "Alignment", "BackColor", "FontColor"}, _
'    '''    {"build_seq     ", "Int32   ", "                ", "8%     ", "         ", "         ", "         "}, _
'    '''    {"solution_name ", "String  ", "                ", "22%    ", "         ", "         ", "         "}, _
'    '''    {"solution_file ", "String  ", "                ", "24%    ", "         ", "         ", "         "}, _
'    '''    {"solution_path ", "String  ", "solution_Path   ", "46%    ", "         ", "         ", "         "}, _
'    '''    {"Note          ", "String  ", "                ", "80     ", "         ", "         ", "         "} _
'    '''}
'    ''' </param>
'    ''' <remarks></remarks>
'    Public Sub Generic_DGV_Create(ByRef dgv As DataGridView, ByVal aDgvColumns(,) As String)
'        Dim nCol, nColWidth, nSumPercentageWidth As Integer
'        Dim colNew As DataGridViewTextBoxColumn
'        Dim sErrMsg As String

'        sErrMsg = ""

'        ' Set RowHeadersWidth, ColumnHeadersHeight and other attributes
'        dgv.RowHeadersWidth = 25            'Must >= 24
'        dgv.ColumnHeadersHeight = 25
'        dgv.ScrollBars = ScrollBars.Both
'        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
'        dgv.MultiSelect = False
'        dgv.ReadOnly = True

'        ' Clear then add columns
'        dgv.Columns.Clear()
'        nSumPercentageWidth = 0
'        For nCol = 1 To UBound(aDgvColumns)
'            nColWidth = InStr(aDgvColumns(nCol, 3), "%")
'            If nColWidth > 0 Then
'                nColWidth = CInt(aDgvColumns(nCol, 3).ToString.Substring(0, nColWidth - 1))
'                nSumPercentageWidth = nSumPercentageWidth + nColWidth
'                If nSumPercentageWidth <= 100 Then
'                    nColWidth = nColWidth * (dgv.Width - dgv.RowHeadersWidth - gnVerticalScrollBarWidth) / 100
'                Else
'                    sErrMsg = "The sum of Width of aDgvColumns should be 100. Please change the configuration of aDgvColumns."
'                    GoTo lExit
'                End If
'            Else
'                nColWidth = CInt(Trim(aDgvColumns(nCol, 3)))
'            End If

'            colNew = New DataGridViewTextBoxColumn
'            colNew.Name = Trim(aDgvColumns(nCol, 0))
'            colNew.SortMode = DataGridViewColumnSortMode.NotSortable
'            If Trim(aDgvColumns(nCol, 2)) <> "" Then
'                colNew.HeaderText = Title_String(Trim(aDgvColumns(nCol, 2)))
'            Else
'                colNew.HeaderText = Title_String(Trim(aDgvColumns(nCol, 0)))
'            End If
'            colNew.DataPropertyName = Trim(aDgvColumns(nCol, 0))
'            colNew.Width = nColWidth

'            dgv.Columns.Add(colNew)
'        Next

'        If nSumPercentageWidth < 100 Then
'            sErrMsg = "The sum of Width of aDgvColumns should be 100. Please change the configuration of aDgvColumns."
'            GoTo lExit
'        End If
'lExit:
'        If sErrMsg <> "" Then
'            MsgBox(sErrMsg)
'            Application.Exit()
'        End If
'    End Sub

'    ''' <summary>
'    ''' 
'    ''' </summary>
'    ''' <param name="dgv"></param>
'    ''' <param name="dt"></param>
'    ''' <param name="sSortBy">Fields must be in the Datatable.</param>
'    ''' <remarks></remarks>
'    Public Sub Generic_DGV_Populate(ByRef dgv As DataGridView, ByVal dt As DataTable, Optional ByVal sSortBy As String = "")
'        Dim bs As BindingSource

'        bs = New BindingSource()
'        bs.DataSource = dt
'        bs.Sort = sSortBy
'        dgv.DataSource = bs

'        ' No need to call Generic_DGV_HighlightRow because it will be triggered by CurrentCellChanged event of each DGV

'    End Sub
'    ''' <summary>
'    ''' Highlight a row by indicating nToRowIndex, or searchig a row with ColumnsValues.
'    ''' </summary>
'    ''' <param name="dgv"></param>
'    ''' <param name="nCurrentRowIndex"></param>
'    ''' <param name="nToRowIndex"></param>
'    ''' <param name="sSearchRowHavingColumnsValues">Format: Col(n)="xxxx", Col(n)="yyyy", ...</param>
'    ''' <remarks></remarks>
'    Public Sub Generic_DGV_HighlightRow(ByRef dgv As DataGridView, ByRef nCurrentRowIndex As Integer, _
'                    ByVal nToRowIndex As Integer, Optional ByVal sSearchRowHavingColumnsValues As String = "", _
'                    Optional ByRef btnMoveUp As System.Windows.Forms.Button = Nothing, _
'                    Optional ByRef btnMoveDown As System.Windows.Forms.Button = Nothing, _
'                    Optional ByRef btnDelete As System.Windows.Forms.Button = Nothing, _
'                    Optional ByRef btnAddNew As System.Windows.Forms.Button = Nothing, _
'                    Optional ByRef btnUpdate As System.Windows.Forms.Button = Nothing, _
'                    Optional ByRef btnScanRows As System.Windows.Forms.Button = Nothing _
')
'        Dim bFound As Boolean
'        Dim aCol_Val(20, 2) As String
'        Dim nP1, nP2, nCols, nCol As Integer

'        If dgv.RowCount <= 0 Then Exit Sub

'        nCols = 0
'        sSearchRowHavingColumnsValues = sSearchRowHavingColumnsValues.Trim().ToUpper()
'        If sSearchRowHavingColumnsValues.Trim <> "" Then
'            sSearchRowHavingColumnsValues = sSearchRowHavingColumnsValues.Replace(" COL(", "COL(").Replace(" ,COL(", ",COL(")
'            sSearchRowHavingColumnsValues = sSearchRowHavingColumnsValues.Replace(" =", "=").Replace("= ", "=") + ",COL("
'            If sSearchRowHavingColumnsValues.Substring(0, 4) <> "COL(" Then
'                MsgBox("Generic_DGV_HighlightRow: Parameter sSearchRowHavingColumnsValues Error. ")
'                Exit Sub
'            End If
'            While sSearchRowHavingColumnsValues.Length > 5
'                nP1 = InStr(sSearchRowHavingColumnsValues, ")=")
'                nP2 = InStr(sSearchRowHavingColumnsValues, """,COL(")
'                If nP1 <= 0 Or nP2 <= 0 Then
'                    MsgBox("Generic_DGV_HighlightRow: Parameter sSearchRowHavingColumnsValues Error. ")
'                    Exit Sub
'                End If
'                aCol_Val(nCols, 0) = sSearchRowHavingColumnsValues.Substring(4, nP1 - 5)
'                aCol_Val(nCols, 1) = sSearchRowHavingColumnsValues.Substring(nP1 + 2, nP2 - nP1 - 3)
'                If Not IsNumeric(aCol_Val(nCols, 0)) Then
'                    MsgBox("Generic_DGV_HighlightRow: Parameter sSearchRowHavingColumnsValues Error. ")
'                    Exit Sub
'                End If
'                sSearchRowHavingColumnsValues = sSearchRowHavingColumnsValues.Substring(nP2 + 1, sSearchRowHavingColumnsValues.Length - nP2 - 1)
'                nCols = nCols + 1
'            End While

'            nToRowIndex = -1
'            For Each row As DataGridViewRow In dgv.Rows
'                bFound = True
'                For nCol = 0 To nCols - 1
'                    If row.Cells(CInt(aCol_Val(nCol, 0))).Value.ToString.ToUpper <> aCol_Val(nCol, 1) Then
'                        bFound = False
'                        Exit For
'                    End If
'                Next

'                If bFound Then
'                    nToRowIndex = row.Index
'                    Exit For
'                End If
'            Next

'        End If

'        nCurrentRowIndex = nToRowIndex

'        If btnMoveUp IsNot Nothing Then
'            If dgv.RowCount = 0 Then
'                btnMoveUp.Enabled = False
'            ElseIf dgv.RowCount = 1 Then
'                btnMoveUp.Enabled = False
'            ElseIf nCurrentRowIndex = dgv.RowCount - 1 Then
'                btnMoveUp.Enabled = True
'            ElseIf nCurrentRowIndex = 0 Then
'                btnMoveUp.Enabled = False
'            ElseIf nCurrentRowIndex < 0 Then
'                btnMoveUp.Enabled = False
'            Else
'                btnMoveUp.Enabled = True
'            End If
'        End If

'        If btnMoveDown IsNot Nothing Then
'            If dgv.RowCount = 0 Then
'                btnMoveDown.Enabled = False
'            ElseIf dgv.RowCount = 1 Then
'                btnMoveDown.Enabled = False
'            ElseIf nCurrentRowIndex = dgv.RowCount - 1 Then
'                btnMoveDown.Enabled = False
'            ElseIf nCurrentRowIndex = 0 Then
'                btnMoveDown.Enabled = True
'            ElseIf nCurrentRowIndex < 0 Then
'                btnMoveDown.Enabled = False
'            Else
'                btnMoveDown.Enabled = True
'            End If
'        End If

'        If btnDelete IsNot Nothing Then
'            btnDelete.Enabled = True
'            If dgv.RowCount = 0 Or nCurrentRowIndex < 0 Then
'                btnDelete.Enabled = False
'            End If
'        End If

'        If btnAddNew IsNot Nothing Then
'            btnAddNew.Enabled = True
'            btnAddNew.Text = "Add New"      ' For cancelling "Adding". by clicking dgv.
'            If dgv.RowCount = 0 Or nCurrentRowIndex < 0 Then
'                btnAddNew.Text = "Add"
'            End If
'        End If

'        If btnUpdate IsNot Nothing Then
'            btnUpdate.Enabled = False       'Will be enabled when the textBoxes of the current row are changed.
'        End If

'        If btnScanRows IsNot Nothing Then
'            btnScanRows.Enabled = True
'            If dgv.RowCount = 0 Or nCurrentRowIndex < 0 Then
'                btnScanRows.Enabled = False
'            End If
'        End If

'        If nToRowIndex < 0 Then Exit Sub

'        nCurrentRowIndex = nToRowIndex

'        dgv.Rows(nToRowIndex).Selected = True
'        If nCols > 0 Then         ' The target row may be out of the current page.
'            dgv.FirstDisplayedScrollingRowIndex = nToRowIndex
'            dgv.PerformLayout()
'        End If

'        dgv.Focus()

'    End Sub

'    Public Sub Generic_DGV_Row_Delete(ByRef dgv As DataGridView)
'        Dim nRowToBeHighlight As Integer

'        nRowToBeHighlight = dgv.CurrentRow.Index
'        If nRowToBeHighlight = dgv.Rows.Count - 1 Then nRowToBeHighlight = dgv.Rows.Count - 2
'        dgv.Rows.Remove(dgv.CurrentRow)

'        ' No need to call Generic_DGV_HighlightRow because it will be triggered by CurrentCellChanged event of each DGV

'    End Sub
'    ''' <summary>
'    ''' .Cells(0).Selected = True will trigger _CurrentCellChanged(dgvSolution_CurrentCellChanged() then Call _HighlightRow()
'    ''' </summary>
'    ''' <param name="dgv"></param>
'    ''' <remarks></remarks>
'    Public Sub Generic_DGV_Row_MoveDown(ByRef dgv As DataGridView)
'        Dim sTemp As String
'        Dim nCol, nCurrentRowIndex As Integer

'        nCurrentRowIndex = dgv.CurrentRow.Index
'        If nCurrentRowIndex = dgv.Rows.Count - 1 Then Exit Sub

'        For nCol = 1 To dgv.Rows(nCurrentRowIndex).Cells.Count - 1
'            sTemp = dgv.Rows(nCurrentRowIndex).Cells(nCol).Value
'            dgv.Rows(nCurrentRowIndex).Cells(nCol).Value = dgv.Rows(nCurrentRowIndex + 1).Cells(nCol).Value
'            dgv.Rows(nCurrentRowIndex + 1).Cells(nCol).Value = sTemp
'        Next
'        dgv.Rows(nCurrentRowIndex + 1).Cells(0).Selected = True

'        ' No need to call Generic_DGV_HighlightRow because it will be triggered by CurrentCellChanged event of each DGV

'    End Sub

'    ''' <summary>
'    ''' .Cells(0).Selected = True will trigger _CurrentCellChanged(dgvSolution_CurrentCellChanged() then Call _HighlightRow()
'    ''' </summary>
'    ''' <param name="dgv"></param>
'    ''' <remarks></remarks>
'    Public Sub Generic_DGV_Row_MoveUp(ByRef dgv As DataGridView)
'        Dim sTemp As String
'        Dim nCol, nCurrentRowIndex As Integer

'        nCurrentRowIndex = dgv.CurrentRow.Index
'        If nCurrentRowIndex = 0 Then Exit Sub

'        For nCol = 1 To dgv.Rows(nCurrentRowIndex).Cells.Count - 1
'            sTemp = dgv.Rows(nCurrentRowIndex).Cells(nCol).Value
'            dgv.Rows(nCurrentRowIndex).Cells(nCol).Value = dgv.Rows(nCurrentRowIndex - 1).Cells(nCol).Value
'            dgv.Rows(nCurrentRowIndex - 1).Cells(nCol).Value = sTemp
'        Next
'        dgv.Rows(nCurrentRowIndex - 1).Cells(0).Selected = True

'        ' No need to call Generic_DGV_HighlightRow because it will be triggered by CurrentCellChanged event of each DGV

'    End Sub

'#End Region

'#Region "Common used functions for TreeView control>"

'    Public Sub TreeView_Load_From_Xml(ByVal trMenu As TreeView, ByVal sSettingFile As String, ByVal sTagName As String, ByVal sAttributesSaved As String, Optional sdelimiter As String = "/")
'        Dim parentNode As TreeNode = Nothing
'        Dim i As Integer

'        Dim reader As XmlTextReader = Nothing

'        sAttributesSaved = sdelimiter + sAttributesSaved + sdelimiter

'        Try
'            If Not File.Exists(sSettingFile) Then Exit Sub

'            ' disabling re-drawing of trMenu till all nodes are added
'            trMenu.BeginUpdate()

'            ' Load the reader with the data file and ignore all white space nodes.         
'            reader = New XmlTextReader(sSettingFile)
'            reader.WhitespaceHandling = WhitespaceHandling.None

'            ' Parse the file and display each of the nodes. 
'            While reader.Read()
'                Select Case reader.NodeType
'                    Case XmlNodeType.Element
'                        If (reader.Name = sTagName) Then
'                            Dim newNode As TreeNode = New TreeNode()
'                            Dim isEmptyElement As Boolean = reader.IsEmptyElement       ' This line is mandatory

'                            ' loading node attributes
'                            For i = 0 To reader.AttributeCount - 1
'                                reader.MoveToAttribute(i)
'                                'SetAttributeValue(newNode, reader.Name, reader.Value)
'                                If sAttributesSaved.IndexOf(sdelimiter + reader.Name + sdelimiter) >= 0 Then
'                                    Select Case reader.Name.ToLower
'                                        Case "text"
'                                            newNode.Text = reader.Value
'                                        Case "imageindex"
'                                            newNode.ImageIndex = reader.Value
'                                        Case "tag"
'                                            newNode.Tag = reader.Value
'                                    End Select
'                                End If
'                            Next i

'                            ' add new node to Parent Node or trMenu
'                            If (parentNode Is Nothing) Then
'                                trMenu.Nodes.Add(newNode)
'                            Else
'                                parentNode.Nodes.Add(newNode)
'                            End If

'                            ' Making current node 'ParentNode' if its not empty
'                            If (Not isEmptyElement) Then parentNode = newNode
'                        End If

'                    Case XmlNodeType.Text
'                        parentNode.Nodes.Add(reader.Value)
'                    Case XmlNodeType.CDATA
'                    Case XmlNodeType.ProcessingInstruction
'                    Case XmlNodeType.Comment
'                    Case XmlNodeType.XmlDeclaration
'                    Case XmlNodeType.Document
'                    Case XmlNodeType.DocumentType
'                    Case XmlNodeType.EntityReference
'                    Case XmlNodeType.EndElement
'                        If (reader.Name = sTagName) Then parentNode = parentNode.Parent
'                End Select
'            End While

'        Finally
'            ' enabling redrawing of trMenu after all nodes are added
'            trMenu.EndUpdate()

'            If Not (reader Is Nothing) Then
'                reader.Close()
'            End If
'        End Try

'    End Sub

'    ''Used by Deserialize method for setting properties of TreeNode from xml node attributes
'    Public Sub TreeView_Save_Into_Xml(ByVal trMenu As TreeView, ByVal sSettingFile As String, ByVal sTagName As String, ByVal sAttributesSaved As String, Optional sdelimiter As String = "/", Optional ByRef bTreeHasBeenUpdated As Boolean = True)
'        If bTreeHasBeenUpdated Then
'            Dim textWriter As XmlTextWriter = New XmlTextWriter(sSettingFile, System.Text.Encoding.ASCII)
'            ' writing the xml declaration tag
'            textWriter.WriteStartDocument()
'            'textWriter.WriteRaw("\r\n");
'            ' writing the main tag that encloses all node tags
'            textWriter.WriteStartElement("TreeView")
'            ' save the nodes, recursive method
'            TreeView_Save_Node_Into_Xml(trMenu.Nodes, textWriter, sTagName, sAttributesSaved)

'            textWriter.WriteEndElement()
'            textWriter.Close()

'            bTreeHasBeenUpdated = False
'        End If
'    End Sub

'    Public Sub TreeView_Save_Node_Into_Xml(ByVal nodesCollection As TreeNodeCollection, ByVal textWriter As XmlTextWriter, ByVal sTagName As String, ByVal sAttributesToBeSaved As String, Optional sdelimiter As String = "/")
'        Dim node As TreeNode
'        Dim i, j As Integer
'        Dim aAttributesToBeSaved() As String = sAttributesToBeSaved.Split(sdelimiter)

'        For i = 0 To nodesCollection.Count - 1
'            node = nodesCollection(i)

'            textWriter.WriteStartElement(sTagName)

'            For j = aAttributesToBeSaved.Length - 1 To 0 Step -1
'                Select Case aAttributesToBeSaved(j).ToLower
'                    Case "text"
'                        textWriter.WriteAttributeString("text", node.Text)
'                    Case "imageindex"
'                        textWriter.WriteAttributeString("imageindex", node.ImageIndex.ToString())
'                    Case "tag"
'                        If (node.Tag IsNot Nothing) Then textWriter.WriteAttributeString("tag", node.Tag.ToString())
'                End Select
'            Next

'            ' Add other node properties to serialize here  
'            If (node.Nodes.Count > 0) Then TreeView_Save_Node_Into_Xml(node.Nodes, textWriter, sTagName, sAttributesToBeSaved)

'            textWriter.WriteEndElement()
'        Next i
'    End Sub

'    Public Function TreeView_Move_Node_Up(ByVal node As TreeNode) As Boolean
'        Dim tree As TreeView
'        Dim parent As TreeNode
'        Dim index As Integer
'        Dim bMoved As Boolean = False
'        Try
'            If node Is Nothing Then
'                MsgBox("Failed moving up because no position was indicated.")
'                Exit Try
'            End If

'            tree = node.TreeView
'            parent = node.Parent

'            If (parent IsNot Nothing) Then
'                index = parent.Nodes.IndexOf(node)
'                If (index > 0) Then
'                    parent.Nodes.RemoveAt(index)
'                    parent.Nodes.Insert(index - 1, node)

'                    ' bw : add this line to restore the originally selected node as selected
'                    node.TreeView.SelectedNode = node
'                    bMoved = True
'                End If
'                node.TreeView.Focus()
'            Else
'                index = tree.Nodes.IndexOf(node)
'                If (index > 0) Then
'                    tree.Nodes.RemoveAt(index)
'                    tree.Nodes.Insert(index - 1, node)

'                    ' bw : add this line to restore the originally selected node as selected
'                    tree.SelectedNode = node
'                    bMoved = True
'                End If
'                tree.Focus()
'            End If
'        Catch ex As Exception
'            MsgBox("Exception: " + ex.Message)

'        Finally

'        End Try

'        Return bMoved

'    End Function

'    Public Function TreeView_Move_Node_Down(ByVal node As TreeNode) As Boolean
'        Dim tree As TreeView
'        Dim parent As TreeNode
'        Dim index As Integer
'        Dim bMoved As Boolean = False
'        Try
'            If node Is Nothing Then
'                MsgBox("Failed moving up because no position was indicated.")
'                Exit Try
'            End If

'            tree = node.TreeView
'            parent = node.Parent

'            If (parent IsNot Nothing) Then
'                index = parent.Nodes.IndexOf(node)
'                If (index < parent.Nodes.Count - 1) Then
'                    parent.Nodes.RemoveAt(index)
'                    parent.Nodes.Insert(index + 1, node)

'                    ' bw : add this line to restore the originally selected node as selected
'                    node.TreeView.SelectedNode = node
'                    bMoved = True
'                End If
'                node.TreeView.Focus()
'            Else
'                index = tree.Nodes.IndexOf(node)
'                If (index < tree.Nodes.Count - 1) Then
'                    tree.Nodes.RemoveAt(index)
'                    tree.Nodes.Insert(index + 1, node)

'                    ' bw : add this line to restore the originally selected node as selected
'                    tree.SelectedNode = node
'                    bMoved = True
'                End If
'                tree.Focus()
'            End If
'        Catch ex As Exception
'            MsgBox("Exception: " + ex.Message)

'        Finally

'        End Try

'        Return bMoved

'    End Function

'    Public Function TreeView_Add_Sibling(ByVal node As TreeNode, Optional ByVal treeView As TreeView = Nothing) As Boolean
'        Dim bAdded As Boolean = False
'        Dim newNode As TreeNode
'        Try
'            node.TreeView.BeginUpdate()

'            If node IsNot Nothing Then
'                newNode = New TreeNode("<New ...>")
'                If node.Parent IsNot Nothing Then
'                    node.Parent.Nodes.Add(newNode)
'                    node.Parent.ExpandAll()
'                    node.TreeView.SelectedNode = newNode
'                    node.TreeView.Focus()
'                Else
'                    node.TreeView.Nodes.Add(newNode)
'                    node.TreeView.ExpandAll()
'                    node.TreeView.SelectedNode = newNode
'                    node.TreeView.Focus()
'                End If
'                bAdded = True
'            ElseIf treeView IsNot Nothing Then
'                If treeView.Nodes.Count = 0 Then
'                    newNode = New TreeNode("<New ...>")
'                    treeView.Nodes.Add(newNode)
'                    treeView.ExpandAll()
'                    treeView.SelectedNode = newNode
'                    treeView.Focus()
'                    bAdded = True
'                Else
'                    MsgBox("Could not add sibing because no position was indicated.")
'                    Exit Try
'                End If
'            Else
'                MsgBox("Parameters are missing.")
'                Exit Try
'            End If

'        Catch ex As Exception
'            MsgBox("Exception: " + ex.Message)

'        Finally
'            node.TreeView.EndUpdate()
'        End Try

'        Return bAdded

'    End Function

'    Public Function TreeView_Add_Child(ByVal node As TreeNode, Optional ByVal treeView As TreeView = Nothing) As Boolean
'        Dim newNode As TreeNode
'        Dim bAdded As Boolean = False
'        Try
'            If node IsNot Nothing Then
'                newNode = New TreeNode("<New ...>")
'                node.Nodes.Add(newNode)
'                node.ExpandAll()
'                node.TreeView.SelectedNode = newNode
'                node.TreeView.Focus()
'                bAdded = True
'            ElseIf treeView IsNot Nothing Then
'                If treeView.Nodes.Count = 0 Then
'                    newNode = New TreeNode("<New ...>")
'                    treeView.Nodes.Add(newNode)
'                    treeView.ExpandAll()
'                    treeView.SelectedNode = newNode
'                    treeView.Focus()
'                    bAdded = True
'                Else
'                    MsgBox("Could not add child because no position was indicated.")
'                    Exit Try
'                End If
'            Else
'                MsgBox("Parameters are missing.")
'                Exit Try
'            End If
'        Catch ex As Exception
'            MsgBox("Exception: " + ex.Message)

'        Finally

'        End Try

'        Return bAdded

'    End Function

'    Public Function TreeView_Delete_Node(ByVal node As TreeNode, Optional ByVal treeView As TreeView = Nothing) As Boolean
'        Dim bDeleted As Boolean = False
'        Dim tree As TreeView
'        Dim index As Integer
'        Try
'            If node IsNot Nothing Then
'                If node.Parent IsNot Nothing Then
'                    tree = node.TreeView
'                    index = node.Parent.Nodes.IndexOf(node)
'                    If index >= 0 Then
'                        node.Parent.Nodes.RemoveAt(index)
'                        tree.Focus()
'                        bDeleted = True
'                    Else
'                        MsgBox("Could not find the node to be deleted.")
'                    End If
'                Else
'                    tree = node.TreeView
'                    index = tree.Nodes.IndexOf(node)
'                    If index >= 0 Then
'                        tree.Nodes.RemoveAt(index)
'                        tree.Focus()
'                        bDeleted = True
'                    Else
'                        MsgBox("Could not find the node to be deleted.")
'                    End If
'                End If
'            ElseIf treeView IsNot Nothing Then
'                If treeView.Nodes.Count = 0 Then
'                    MsgBox("This TreeView is empty.")
'                Else
'                    MsgBox("Failed deleting because no node was indicated.")
'                End If
'            Else
'                MsgBox("Parameters are missing.")
'                Exit Try
'            End If

'        Catch ex As Exception
'            MsgBox("Exception: " + ex.Message)

'        Finally

'        End Try

'        Return bDeleted

'    End Function

'#End Region




'    ' This Sub is not necessary Because each DGV must have its own CurrentCellChanged event handler
'    'Public Sub Generic_DGV_CurrentCellChanged(ByRef dgv As DataGridView)
'    '    Dim nRowIndex As Integer

'    '    If dgv.RowCount <= 0 Then Exit Sub

'    '    nRowIndex = dgv.CurrentCellAddress.Y
'    '    If nRowIndex < -1 Then Exit Sub

'    '    Call Generic_DGV_HighlightRow(dgv, -1, nRowIndex)

'    'End Sub

'    'Sub ClearImmediateWindow()
'    '    Dim currentActiveWindow As Window = DTE.ActiveWindow
'    '    DTE.Windows.Item("{ECB7191A-597B-41F5-9843-03A4CF275DDE}").Activate() 'Immediate Window
'    '    DTE.ExecuteCommand("Edit.SelectAll")
'    '    DTE.ExecuteCommand("Edit.ClearAll")
'    '    currentActiveWindow.Activate()
'    'End Sub

'    'Sub Reset_Output_Window(ByVal dte As DTE2)

'    '    Dim dte As EnvDTE.DTE
'    '    dte = System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE")

'    '    ' Get an instance of the currently running Visual Studio IDE.
'    '    Dim DTE2 As EnvDTE80.DTE2
'    '    DTE2 = System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.12.0")

'    '    'below uses the 7.1 DTE for 2003
'    '    dte = System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.7.1")
'    '    Dim mwin As Window = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput)
'    '    Dim ow As OutputWindow = mwin.Object
'    '    ow.ActivePane.Clear()

'    '    ' Retrieve the Output window.
'    '    Dim outputWin As OutputWindow = dte.ToolWindows.OutputWindow

'    '    ' Find the "Test Pane" Output window pane; if it does not exist, 
'    '    ' create it.
'    '    Dim pane As OutputWindowPane
'    '    Try
'    '        pane = outputWin.OutputWindowPanes.Item("Test Pane")
'    '    Catch
'    '        pane = outputWin.OutputWindowPanes.Add("Test Pane")
'    '    End Try

'    '    ' Show the Output window and activate the new pane.
'    '    outputWin.Parent.AutoHides = False
'    '    outputWin.Parent.Activate()
'    '    pane.Activate()

'    '    ' Add a line of text to the new pane.
'    '    pane.OutputString("Some text." & vbCrLf)

'    '    If MsgBox("Clear the Output window pane?", MsgBoxStyle.YesNo) = _
'    '        MsgBoxResult.Yes Then
'    '        pane.Clear()
'    '    End If

'    'End Sub

'End Class
