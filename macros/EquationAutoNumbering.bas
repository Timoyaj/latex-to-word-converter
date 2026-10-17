Attribute VB_Name = "EquationAutoNumbering"
' =============================================================================
' LaTeX to Word Equation Converter & Auto-Numbering Suite for Microsoft Word
' =============================================================================
' Entry points:
'   1. ConvertAllLatexInDocument : Scans the document for $$...$$ and $...$
'      (including multi-line display math) and converts them to native,
'      interactive Word equations. Optional auto-numbering via SEQ fields.
'   2. ConvertSelectionLatex     : Converts the selected LaTeX formula.
'   3. RenumberAllEquations      : Renumbers all tagged/legacy equation tables.
'   4. InsertEquationNumber      : Inserts a pre-numbered 1x3 equation table.
'   5. InsertEquationReference   : Inserts a LIVE cross-reference (REF field)
'                                  to an equation — stays correct after
'                                  renumbering.
'   6. CountEquations            : Returns the count of numbered equation tables.
'
' Design notes:
'   * Display equations are located by pairing "$$" markers with plain Find
'     (Word wildcard Find cannot cross paragraph marks, so the common
'      multi-line $$...$$ form would otherwise be invisible).
'   * Inline $...$ candidates are validated in code: escaped \$ and currency
'     amounts ($5.00 ... $10.00) are never converted.
'   * Equation tables are tagged with EqTbl<n> bookmarks and their number
'     cells with EqNum<n> bookmarks so cross-references survive renumbering.
'   * Number cells use a SEQ Eq field, so Word itself maintains numbering;
'     legacy literal labels such as (3) or [Eq. 3] created by the Python app
'     are also understood (renumbering preserves the surrounding format).
'   * The whole run is wrapped in one custom undo record.
' =============================================================================

Option Explicit

' --- configuration -----------------------------------------------------------
Private Const EQ_SEQ_FIELD As String = "Eq \* ARABIC"   ' SEQ field instruction
Private Const EQ_BOOKMARK_PREFIX As String = "EqTbl"    ' tags equation tables
Private Const EQ_NUM_PREFIX As String = "EqNum"         ' bookmarks number cells

' =============================================================================
' ConvertAllLatexInDocument
' =============================================================================
Public Sub ConvertAllLatexInDocument()
    Dim optChoice As VbMsgBoxResult
    optChoice = MsgBox("Convert all LaTeX equations in this document to native Word formulas?" & vbCrLf & vbCrLf & _
                       "• Click [Yes] to Convert and AUTO-NUMBER display equations (1), (2)..." & vbCrLf & _
                       "• Click [No] to Convert display equations WITHOUT numbers (centered)" & vbCrLf & _
                       "• Click [Cancel] to Abort", _
                       vbYesNoCancel + vbQuestion, "LaTeX to Word Converter")

    If optChoice = vbCancel Then Exit Sub

    Dim autoNumber As Boolean
    autoNumber = (optChoice = vbYes)

    Dim displayCount As Long, inlineCount As Long, failCount As Long
    displayCount = 0
    inlineCount = 0
    failCount = 0

    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert LaTeX Equations"

    On Error GoTo ErrorHandler

    ' -------------------------------------------------------------
    ' Step 1: display equations ($$...$$), multi-line safe
    ' -------------------------------------------------------------
    Dim searchStart As Long
    searchStart = 0

    Do
        Dim openRng As Range
        Set openRng = FindLiteral("$$", searchStart, ActiveDocument.Content.End)
        If openRng Is Nothing Then Exit Do

        Dim closeRng As Range
        Set closeRng = FindLiteral("$$", openRng.End, ActiveDocument.Content.End)
        If closeRng Is Nothing Then Exit Do   ' unpaired opener: stop

        Dim wholeRng As Range
        Set wholeRng = ActiveDocument.Range(openRng.Start, closeRng.End)

        Dim rawText As String
        rawText = wholeRng.Text

        Dim cleanFormula As String
        cleanFormula = StripDisplayDelimiters(rawText)
        displayCount = displayCount + 1

        If autoNumber Then
            Set wholeRng = InsertNumberedEquation(wholeRng, cleanFormula, displayCount)
            searchStart = wholeRng.End + 1
        Else
            Set wholeRng = InsertPlainEquation(wholeRng, cleanFormula)
            searchStart = wholeRng.End + 1
        End If

        If searchStart >= ActiveDocument.Content.End Then Exit Do
    Loop

    ' -------------------------------------------------------------
    ' Step 2: inline equations ($...$) with currency/escape guards
    ' -------------------------------------------------------------
    Dim rngInline As Range
    Set rngInline = ActiveDocument.Content

    With rngInline.Find
        .ClearFormatting
        .MatchWildcards = True
        .Text = "\$[!\$]@\$"
        .Wrap = wdFindStop

        Do While .Execute()
            If Not LooksLikeCurrency(rngInline) Then
                Dim rawInline As String
                rawInline = Trim(rngInline.Text)

                Dim cleanInline As String
                cleanInline = StripInlineDelimiters(rawInline)
                inlineCount = inlineCount + 1

                If BuildUpRange(rngInline, cleanInline) Then
                    rngInline.SetRange rngInline.End, ActiveDocument.Content.End
                Else
                    failCount = failCount + 1
                    rngInline.SetRange rngInline.End + 1, ActiveDocument.Content.End
                End If
            Else
                ' currency / escaped dollar: skip past it
                rngInline.SetRange rngInline.End + 1, ActiveDocument.Content.End
            End If
        Loop
    End With

    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True

    Dim summaryMsg As String
    summaryMsg = "LaTeX Conversion Complete!" & vbCrLf & vbCrLf & _
                 "• Display Equations: " & displayCount & vbCrLf & _
                 "• Inline Formulas: " & inlineCount & vbCrLf & _
                 "• Auto-Numbering: " & IIf(autoNumber, "Yes (1 to " & displayCount & ")", "None") & vbCrLf
    If failCount > 0 Then
        summaryMsg = summaryMsg & "• Failed (kept as text): " & failCount & vbCrLf
    End If
    summaryMsg = summaryMsg & vbCrLf & "All formulas are now native editable Word equations."

    MsgBox summaryMsg, vbInformation, "LaTeX to Word Converter"
    Exit Sub

ErrorHandler:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox "An error occurred during conversion: " & Err.Description, vbCritical, "Conversion Error"
End Sub

' =============================================================================
' ConvertSelectionLatex
' =============================================================================
Public Sub ConvertSelectionLatex()
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please highlight a LaTeX formula before running this command.", vbExclamation, "No Selection"
        Exit Sub
    End If

    Dim raw As String
    raw = Trim(Selection.Text)

    ' strip trailing paragraph/row marks from the selection
    Do While Len(raw) > 0
        Dim lastCh As String
        lastCh = Right$(raw, 1)
        If lastCh = vbCr Or lastCh = vbVerticalTab Or lastCh = Chr$(7) Then
            raw = Left$(raw, Len(raw) - 1)
        Else
            Exit Do
        End If
    Loop

    Dim isDisplay As Boolean
    isDisplay = (Left$(raw, 2) = "$$" And Right$(raw, 2) = "$$")

    Dim clean As String
    If isDisplay Then
        clean = StripDisplayDelimiters(raw)
    Else
        clean = StripInlineDelimiters(raw)
    End If

    If Len(clean) = 0 Then Exit Sub

    Application.UndoRecord.StartCustomRecord "Convert LaTeX Selection"
    If isDisplay Then
        BuildUpRange Selection.Range, clean
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Else
        BuildUpRange Selection.Range, clean
    End If
    Application.UndoRecord.EndCustomRecord
End Sub

' =============================================================================
' RenumberAllEquations
'   Updates SEQ fields and rewrites legacy literal labels, preserving each
'   label's format: "(3)" -> "(4)", "[Eq. 3]" -> "[Eq. 4]", "Eq. 3" -> "Eq. 4".
' =============================================================================
Public Sub RenumberAllEquations()
    Dim tbl As Table
    Dim eqNum As Long
    eqNum = 0

    Application.UndoRecord.StartCustomRecord "Renumber Equations"

    For Each tbl In ActiveDocument.Tables
        If IsEquationTable(tbl) Then
            eqNum = eqNum + 1

            Dim rightCellRng As Range
            Set rightCellRng = CellContentRange(tbl.Cell(1, 3))

            If rightCellRng.Fields.Count > 0 Then
                ' SEQ-managed number: Word recomputes it
                rightCellRng.Fields.Update
            Else
                ' legacy literal label: rewrite the number, keep the format
                Dim prefix As String, suffix As String, ignoreNum As String
                SplitLabel PlainText(rightCellRng.Text), prefix, ignoreNum, suffix
                If Len(prefix) = 0 And Len(suffix) = 0 Then
                    prefix = "("
                    suffix = ")"
                End If
                rightCellRng.Text = prefix & CStr(eqNum) & suffix
            End If

            With tbl.Cell(1, 3).Range
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Name = "Cambria Math"
                .Font.Size = 12
            End With

            ' make sure the number cell is bookmarked for cross-references
            EnsureNumberBookmark tbl, eqNum
        End If
    Next tbl

    Application.UndoRecord.EndCustomRecord

    MsgBox "Renumbered " & eqNum & " equations sequentially (1) to (" & eqNum & ").", vbInformation, "Equation Numbering"
End Sub

' =============================================================================
' InsertEquationNumber
' =============================================================================
Public Sub InsertEquationNumber()
    Dim nextNum As Long
    nextNum = CountEquations() + 1

    Application.UndoRecord.StartCustomRecord "Insert Equation Number"

    Dim tbl As Table
    Set tbl = InsertNumberedEquation(Selection.Range, "", nextNum)
    tbl.Cell(1, 2).Range.Select

    Application.UndoRecord.EndCustomRecord
End Sub

' =============================================================================
' InsertEquationReference
'   Inserts a LIVE cross-reference (REF field) to an equation number, so the
'   reference updates when equations are renumbered or moved.
' =============================================================================
Public Sub InsertEquationReference()
    Dim eqRef As String
    eqRef = InputBox("Enter equation number to reference:", "Equation Reference", "1")
    If eqRef = "" Then Exit Sub
    If Not IsNumeric(eqRef) Then
        MsgBox "Please enter a valid numeric equation number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    Dim totalEqs As Long
    totalEqs = CountEquations()

    If CLng(eqRef) > totalEqs Or CLng(eqRef) < 1 Then
        MsgBox "Equation (" & eqRef & ") not found. Document has " & totalEqs & " numbered equations.", vbExclamation, "Equation Not Found"
        Exit Sub
    End If

    Application.UndoRecord.StartCustomRecord "Insert Equation Reference"

    Dim tbl As Table
    Set tbl = EquationTableByNumber(CLng(eqRef))
    If tbl Is Nothing Then
        MsgBox "Could not locate equation (" & eqRef & ") in the document.", vbExclamation, "Equation Not Found"
        Application.UndoRecord.EndCustomRecord
        Exit Sub
    End If

    ' make sure the number cell carries a bookmark, then REF it
    Dim bmName As String
    bmName = EnsureNumberBookmark(tbl, CLng(eqRef))

    Dim label As String
    label = PlainText(CellContentRange(tbl.Cell(1, 3)).Text)

    Dim prefix As String, suffix As String, numPart As String
    SplitLabel label, prefix, numPart, suffix
    If Len(prefix) = 0 And Len(suffix) = 0 Then
        prefix = "("
        suffix = ")"
    End If

    ' "Eq. (n)" style for plain parens; the label's own format otherwise
    If prefix = "(" Then
        Selection.TypeText Text:="Eq. ("
    Else
        Selection.TypeText Text:=prefix
    End If

    ' live REF field that displays just the equation number
    Dim insRng As Range
    Set insRng = Selection.Range
    insRng.Collapse wdCollapseStart
    Dim fld As Field
    Set fld = ActiveDocument.Fields.Add(Range:=insRng, Type:=wdFieldRef, _
                                        Text:=bmName & " \h", PreserveFormatting:=False)
    fld.Result.Select
    Selection.Collapse wdCollapseEnd
    Selection.TypeText Text:=suffix

    Application.UndoRecord.EndCustomRecord
End Sub

' =============================================================================
' CountEquations
' =============================================================================
Public Function CountEquations() As Long
    Dim tbl As Table
    Dim count As Long
    count = 0

    For Each tbl In ActiveDocument.Tables
        If IsEquationTable(tbl) Then count = count + 1
    Next tbl

    CountEquations = count
End Function

' =============================================================================
' Internals
' =============================================================================

' --- FindLiteral: plain-text Find that can cross paragraph marks -------------
Private Function FindLiteral(ByVal needle As String, ByVal startPos As Long, _
                             ByVal endPos As Long) As Range
    Dim rng As Range
    Set rng = ActiveDocument.Range(Start:=startPos, End:=endPos)

    With rng.Find
        .ClearFormatting
        .MatchWildcards = False
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindStop
        .Text = needle
        If .Execute() Then
            Set FindLiteral = rng
        Else
            Set FindLiteral = Nothing
        End If
    End With
End Function

' --- StripDisplayDelimiters: "$$...$$" (any layout) -> formula ---------------
Private Function StripDisplayDelimiters(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If Left$(s, 2) = "$$" Then s = Mid$(s, 3)
    If Right$(s, 2) = "$$" Then s = Left$(s, Len(s) - 2)

    ' join multi-line formulas with spaces
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbVerticalTab, " ")
    s = Replace(s, vbLf, " ")

    StripDisplayDelimiters = Trim$(s)
End Function

' --- StripInlineDelimiters: "$...$" -> formula --------------------------------
Private Function StripInlineDelimiters(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If Left$(s, 2) = "$$" Then
        s = Mid$(s, 3, Len(s) - 4)
    ElseIf Left$(s, 1) = "$" And Right$(s, 1) = "$" Then
        s = Mid$(s, 2, Len(s) - 2)
    End If
    StripInlineDelimiters = Trim$(s)
End Function

' --- LooksLikeCurrency: guard for $5.00 ... $10.00 and escaped \$ ------------
Private Function LooksLikeCurrency(ByVal rng As Range) As Boolean
    Dim doc As Document
    Set doc = rng.Document

    ' escaped dollar: "\$5.00 and \$10.00" — preceded by a backslash
    If rng.Start > 0 Then
        Dim prevRng As Range
        Set prevRng = doc.Range(rng.Start - 1, rng.Start)
        If prevRng.Text = "\" Then
            LooksLikeCurrency = True
            Exit Function
        End If
    End If

    ' closing "$" followed by a digit => currency-style ("... is $10.00")
    If rng.End < doc.Content.End Then
        Dim nextRng As Range
        Set nextRng = doc.Range(rng.End, rng.End + 1)
        If nextRng.Text Like "[0-9]" Then
            LooksLikeCurrency = True
            Exit Function
        End If
    End If

    ' opening "$" followed by whitespace => not math
    Dim innerRng As Range
    Set innerRng = doc.Range(rng.Start + 1, rng.Start + 2)
    If innerRng.Text = " " Or innerRng.Text = vbTab Then
        LooksLikeCurrency = True
        Exit Function
    End If

    LooksLikeCurrency = False
End Function

' --- BuildUpRange: replace text with a Word equation -------------------------
Private Function BuildUpRange(ByVal rng As Range, ByVal formula As String) As Boolean
    On Error Resume Next
    Err.Clear

    rng.Text = formula
    ActiveDocument.OMaths.Add rng
    If Err.Number = 0 Then
        ActiveDocument.OMaths.Item(ActiveDocument.OMaths.Count).BuildUp
        If Err.Number = 0 Then
            BuildUpRange = True
            On Error GoTo 0
            Exit Function
        End If
    End If

    ' fallback: leave the plain formula text in place
    Err.Clear
    On Error GoTo 0
    BuildUpRange = False
End Function

' --- InsertNumberedEquation: 1x3 borderless table with SEQ number ------------
Private Function InsertNumberedEquation(ByVal anchor As Range, ByVal formula As String, _
                                        ByVal eqIndex As Long) As Table
    ' clear the anchor content (formula text / selection)
    anchor.Text = ""

    Dim tbl As Table
    Set tbl = ActiveDocument.Tables.Add(anchor, 1, 3)
    tbl.Borders.Enable = False

    ' proportional widths derived from the usable page width
    Dim usable As Single
    usable = ActiveDocument.PageSetup.PageWidth _
             - ActiveDocument.PageSetup.LeftMargin _
             - ActiveDocument.PageSetup.RightMargin
    tbl.Columns(1).PreferredWidth = usable * 0.1
    tbl.Columns(2).PreferredWidth = usable * 0.77
    tbl.Columns(3).PreferredWidth = usable * 0.13

    ' centre cell: the equation (may be empty, e.g. InsertEquationNumber)
    Dim cEq As Range
    Set cEq = tbl.Cell(1, 2).Range
    cEq.ParagraphFormat.Alignment = wdAlignParagraphCenter
    If Len(formula) > 0 Then
        BuildUpRange cEq, formula
    End If

    ' right cell: ( SEQ Eq )
    Dim cNum As Range
    Set cNum = tbl.Cell(1, 3).Range
    cNum.ParagraphFormat.Alignment = wdAlignParagraphRight

    Dim writeRng As Range
    Set writeRng = CellContentRange(tbl.Cell(1, 3))
    writeRng.Text = "("

    Set writeRng = CellContentRange(tbl.Cell(1, 3))
    writeRng.Collapse wdCollapseEnd
    ActiveDocument.Fields.Add Range:=writeRng, Type:=wdFieldSequence, _
                              Text:=EQ_SEQ_FIELD, PreserveFormatting:=False

    Set writeRng = CellContentRange(tbl.Cell(1, 3))
    writeRng.Collapse wdCollapseEnd
    writeRng.InsertAfter ")"

    With tbl.Cell(1, 3).Range
        .Font.Name = "Cambria Math"
        .Font.Size = 12
    End With

    ' tag the table and bookmark its number for cross-references
    TagEquationTable tbl, eqIndex
    EnsureNumberBookmark tbl, eqIndex

    Set InsertNumberedEquation = tbl
End Function

' --- InsertPlainEquation: centered, unnumbered -------------------------------
Private Function InsertPlainEquation(ByVal anchor As Range, ByVal formula As String) As Range
    anchor.ParagraphFormat.Alignment = wdAlignParagraphCenter
    BuildUpRange anchor, formula
    Set InsertPlainEquation = anchor
End Function

' --- CellContentRange: cell range without the end-of-cell marks --------------
Private Function CellContentRange(ByVal cel As Cell) As Range
    Dim rng As Range
    Set rng = cel.Range
    rng.End = rng.End - 1
    Set CellContentRange = rng
End Function

' --- PlainText: strip cell end marks from raw cell text ----------------------
Private Function PlainText(ByVal raw As String) As String
    Dim s As String
    s = raw
    If Len(s) >= 1 Then
        If Right$(s, 1) = Chr$(7) Then s = Left$(s, Len(s) - 1)
    End If
    If Len(s) >= 1 Then
        If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
    End If
    PlainText = Trim$(s)
End Function

' --- SplitLabel: "(3)" -> "(", "3", ")"   |   "[Eq. 3]" -> "[Eq. ", "3", "]" -
' Rewrites the LAST digit run so section-style labels "(2.3)" keep their
' prefix intact ("(2." & n & ")").
Private Sub SplitLabel(ByVal label As String, ByRef prefix As String, _
                       ByRef numPart As String, ByRef suffix As String)
    prefix = ""
    numPart = ""
    suffix = ""

    Dim i As Long
    i = Len(label)
    Do While i >= 1
        If Mid$(label, i, 1) Like "[0-9]" Then
            Dim j As Long
            j = i
            Do While j >= 1
                If Mid$(label, j, 1) Like "[0-9]" Then
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            prefix = Left$(label, j)
            numPart = Mid$(label, j + 1, i - j)
            suffix = Mid$(label, i + 1)
            Exit Sub
        End If
        i = i - 1
    Loop
End Sub

' --- TagEquationTable: bookmark the table so we can find it later ------------
Private Sub TagEquationTable(ByVal tbl As Table, ByVal eqIndex As Long)
    Dim bmName As String
    bmName = EQ_BOOKMARK_PREFIX & CStr(eqIndex)
    On Error Resume Next
    ActiveDocument.Bookmarks.Add Name:=bmName, Range:=tbl.Range
    On Error GoTo 0
End Sub

' --- EnsureNumberBookmark: bookmark the number cell for REF fields -----------
' Bookmarks only the digits of the label so a REF field displays "3", not "(3)".
Private Function EnsureNumberBookmark(ByVal tbl As Table, ByVal eqIndex As Long) As String
    Dim bmName As String
    bmName = EQ_NUM_PREFIX & CStr(eqIndex)

    Dim cellRng As Range
    Set cellRng = CellContentRange(tbl.Cell(1, 3))

    Dim target As Range
    If cellRng.Fields.Count > 0 Then
        ' SEQ field: bookmark its result (REF then shows just the number)
        Set target = cellRng.Fields(1).Result
    Else
        ' legacy literal label: bookmark only the digit run
        Dim prefix As String, numPart As String, suffix As String
        SplitLabel PlainText(cellRng.Text), prefix, numPart, suffix
        If Len(numPart) > 0 Then
            Set target = ActiveDocument.Range(cellRng.Start + Len(prefix), _
                                              cellRng.Start + Len(prefix) + Len(numPart))
        Else
            Set target = cellRng
        End If
    End If

    On Error Resume Next
    ActiveDocument.Bookmarks.Add Name:=bmName, Range:=target
    On Error GoTo 0

    EnsureNumberBookmark = bmName
End Function

' --- IsEquationTable: tagged table, or legacy 1x3 with a number label --------
Private Function IsEquationTable(ByVal tbl As Table) As Boolean
    If tbl.Columns.Count <> 3 Then
        IsEquationTable = False
        Exit Function
    End If

    ' fast path: our bookmark tag
    Dim bm As Bookmark
    For Each bm In ActiveDocument.Bookmarks
        If Left$(bm.Name, Len(EQ_BOOKMARK_PREFIX)) = EQ_BOOKMARK_PREFIX Then
            If bm.Range.Start <= tbl.Range.Start And bm.Range.End >= tbl.Range.Start Then
                IsEquationTable = True
                Exit Function
            End If
        End If
    Next bm

    ' fallback: legacy heuristic (right cell holds a label with a number)
    On Error Resume Next
    Dim label As String
    label = PlainText(tbl.Cell(1, 3).Range.Text)
    On Error GoTo 0

    Dim prefix As String, numPart As String, suffix As String
    SplitLabel label, prefix, numPart, suffix
    IsEquationTable = (Len(numPart) > 0)
End Function

' --- EquationTableByNumber: find the table currently displaying the number ---
Private Function EquationTableByNumber(ByVal n As Long) As Table
    Dim tbl As Table

    For Each tbl In ActiveDocument.Tables
        If IsEquationTable(tbl) Then
            Dim label As String
            label = PlainText(CellContentRange(tbl.Cell(1, 3)).Text)

            Dim prefix As String, numPart As String, suffix As String
            SplitLabel label, prefix, numPart, suffix

            If numPart = CStr(n) Then
                Set EquationTableByNumber = tbl
                Exit Function
            End If
        End If
    Next tbl

    Set EquationTableByNumber = Nothing
End Function
