Attribute VB_Name = "EquationAutoNumbering"
' ==============================================================================
' LaTeX to Word Equation Converter & Auto-Numbering Suite for Microsoft Word
' ==============================================================================
' Features:
'   1. ConvertAllLatexInDocument: Scans document for $$...$$ and $...$ and converts
'      all formulas into native interactive Word OMML equations with auto-numbering.
'   2. ConvertSelectionLatex: Converts currently selected LaTeX formula.
'   3. RenumberAllEquations: Sequentially renumbers all equation tables (1) to (N).
'   4. InsertEquationNumber: Inserts a 3-column pre-numbered equation table at cursor.
'   5. InsertEquationReference: Prompts for equation number and inserts "Eq. (n)".
' ==============================================================================

Option Explicit

' --- ConvertAllLatexInDocument ---
' Primary entry point: Scans the entire active Word document, finds all LaTeX
' display equations ($$...$$) and inline expressions ($...$), and converts them
' into native, interactive Word equations. Prompts user for auto-numbering.
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
    
    Dim displayCount As Long, inlineCount As Long
    displayCount = 0
    inlineCount = 0
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' -------------------------------------------------------------
    ' Step 1: Find and convert display equations ($$...$$)
    ' -------------------------------------------------------------
    Dim rngDisplay As Range
    Set rngDisplay = ActiveDocument.Content
    
    With rngDisplay.Find
        .ClearFormatting
        .MatchWildcards = True
        .Text = "\$\$*\$\$"
        
        Do While .Execute()
            displayCount = displayCount + 1
            Dim rawText As String
            rawText = Trim(rngDisplay.Text)
            
            ' Strip leading and trailing $$
            Dim cleanFormula As String
            cleanFormula = rawText
            If Left(cleanFormula, 2) = "$$" Then cleanFormula = Mid(cleanFormula, 3)
            If Right(cleanFormula, 2) = "$$" Then cleanFormula = Left(cleanFormula, Len(cleanFormula) - 2)
            cleanFormula = Trim(cleanFormula)
            
            If autoNumber Then
                ' Insert 1x3 borderless layout table
                rngDisplay.Text = ""
                Dim tbl As Table
                Set tbl = ActiveDocument.Tables.Add(rngDisplay, 1, 3)
                tbl.Borders.Enable = False
                
                ' Set proportional column widths (0.6 in, 5.0 in, 1.0 in)
                tbl.Columns(1).PreferredWidth = InchesToPoints(0.6)
                tbl.Columns(2).PreferredWidth = InchesToPoints(5.0)
                tbl.Columns(3).PreferredWidth = InchesToPoints(1.0)
                
                ' Center cell: OMML equation
                Dim cEq As Range
                Set cEq = tbl.Cell(1, 2).Range
                cEq.ParagraphFormat.Alignment = wdAlignParagraphCenter
                cEq.Text = cleanFormula
                ActiveDocument.OMaths.Add cEq
                ActiveDocument.OMaths.Item(ActiveDocument.OMaths.Count).BuildUp
                
                ' Right cell: Equation number tag (n)
                Dim cNum As Range
                Set cNum = tbl.Cell(1, 3).Range
                cNum.ParagraphFormat.Alignment = wdAlignParagraphRight
                cNum.Text = "(" & displayCount & ")"
                cNum.Font.Name = "Cambria Math"
                cNum.Font.Size = 12
                
                ' Advance search range past the created table
                rngDisplay.SetRange tbl.Range.End + 1, ActiveDocument.Content.End
            Else
                ' Simple centered display equation
                rngDisplay.Text = cleanFormula
                rngDisplay.ParagraphFormat.Alignment = wdAlignParagraphCenter
                ActiveDocument.OMaths.Add rngDisplay
                ActiveDocument.OMaths.Item(ActiveDocument.OMaths.Count).BuildUp
                rngDisplay.SetRange rngDisplay.End + 1, ActiveDocument.Content.End
            End If
        Loop
    End With
    
    ' -------------------------------------------------------------
    ' Step 2: Find and convert inline equations ($...$)
    ' -------------------------------------------------------------
    Dim rngInline As Range
    Set rngInline = ActiveDocument.Content
    
    With rngInline.Find
        .ClearFormatting
        .MatchWildcards = True
        .Text = "\$[!\$]@\$"
        
        Do While .Execute()
            inlineCount = inlineCount + 1
            Dim rawInline As String
            rawInline = Trim(rngInline.Text)
            
            Dim cleanInline As String
            cleanInline = rawInline
            If Left(cleanInline, 1) = "$" Then cleanInline = Mid(cleanInline, 2)
            If Right(cleanInline, 1) = "$" Then cleanInline = Left(cleanInline, Len(cleanInline) - 1)
            cleanInline = Trim(cleanInline)
            
            rngInline.Text = cleanInline
            ActiveDocument.OMaths.Add rngInline
            ActiveDocument.OMaths.Item(ActiveDocument.OMaths.Count).BuildUp
            
            rngInline.SetRange rngInline.End, ActiveDocument.Content.End
        Loop
    End With
    
    Application.ScreenUpdating = True
    
    Dim summaryMsg As String
    summaryMsg = "LaTeX Conversion Complete!" & vbCrLf & vbCrLf & _
                 "• Display Equations: " & displayCount & vbCrLf & _
                 "• Inline Formulas: " & inlineCount & vbCrLf & _
                 "• Auto-Numbering: " & IIf(autoNumber, "Yes (1 to " & displayCount & ")", "None") & vbCrLf & vbCrLf & _
                 "All formulas are now native editable Word equations."
                 
    MsgBox summaryMsg, vbInformation, "LaTeX to Word Converter"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred during conversion: " & Err.Description, vbCritical, "Conversion Error"
End Sub

' --- ConvertSelectionLatex ---
' Converts only the currently selected LaTeX text into a Word equation.
Public Sub ConvertSelectionLatex()
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please highlight a LaTeX formula before running this command.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Dim raw As String
    raw = Trim(Selection.Text)
    
    Dim isDisplay As Boolean
    isDisplay = (Left(raw, 2) = "$$" And Right(raw, 2) = "$$")
    
    Dim clean As String
    clean = raw
    If isDisplay Then
        clean = Mid(clean, 3, Len(clean) - 4)
    ElseIf Left(clean, 1) = "$" And Right(clean, 1) = "$" Then
        clean = Mid(clean, 2, Len(clean) - 2)
    End If
    clean = Trim(clean)
    
    Selection.Text = clean
    If isDisplay Then
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
    
    ActiveDocument.OMaths.Add Selection.Range
    ActiveDocument.OMaths.Item(ActiveDocument.OMaths.Count).BuildUp
End Sub

' --- RenumberAllEquations ---
' Walks through the document and renumbers all equations
' that have equation number markers in 3-column layout tables.
Public Sub RenumberAllEquations()
    Dim tbl As Table
    Dim eqNum As Long
    eqNum = 0
    
    For Each tbl In ActiveDocument.Tables
        If tbl.Columns.Count = 3 Then
            Dim rightCell As String
            rightCell = Trim(tbl.Cell(1, 3).Range.Text)
            ' Remove cell end markers (Chr(13) & Chr(7))
            If Len(rightCell) >= 2 Then
                rightCell = Left(rightCell, Len(rightCell) - 2)
            End If
            rightCell = Trim(rightCell)
            
            ' Check if it matches equation number pattern like (1), (2) etc
            If Left(rightCell, 1) = "(" And Right(rightCell, 1) = ")" Then
                eqNum = eqNum + 1
                tbl.Cell(1, 3).Range.Text = "(" & eqNum & ")"
                tbl.Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                tbl.Cell(1, 3).Range.Font.Name = "Cambria Math"
                tbl.Cell(1, 3).Range.Font.Size = 12
            End If
        End If
    Next tbl
    
    MsgBox "Renumbered " & eqNum & " equations sequentially (1) to (" & eqNum & ").", vbInformation, "Equation Numbering"
End Sub

' --- InsertEquationNumber ---
' Inserts an equation number at the current cursor position
' by creating a 3-column invisible layout table.
Public Sub InsertEquationNumber()
    Dim nextNum As Long
    nextNum = CountEquations() + 1
    
    Dim tbl As Table
    Set tbl = ActiveDocument.Tables.Add( _
        Range:=Selection.Range, _
        NumRows:=1, _
        NumColumns:=3)
    
    tbl.Borders.Enable = False
    
    tbl.Columns(1).PreferredWidth = InchesToPoints(0.6)
    tbl.Columns(2).PreferredWidth = InchesToPoints(5.0)
    tbl.Columns(3).PreferredWidth = InchesToPoints(1.0)
    
    tbl.Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    tbl.Cell(1, 3).Range.Text = "(" & nextNum & ")"
    tbl.Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    tbl.Cell(1, 3).Range.Font.Name = "Cambria Math"
    tbl.Cell(1, 3).Range.Font.Size = 12
    
    tbl.Cell(1, 2).Range.Select
End Sub

' --- CountEquations ---
' Returns the count of numbered equations in the document
Public Function CountEquations() As Long
    Dim tbl As Table
    Dim count As Long
    count = 0
    
    For Each tbl In ActiveDocument.Tables
        If tbl.Columns.Count = 3 Then
            Dim rightCell As String
            rightCell = Trim(tbl.Cell(1, 3).Range.Text)
            If Len(rightCell) >= 2 Then
                rightCell = Left(rightCell, Len(rightCell) - 2)
            End If
            rightCell = Trim(rightCell)
            If Left(rightCell, 1) = "(" And Right(rightCell, 1) = ")" Then
                count = count + 1
            End If
        End If
    Next tbl
    
    CountEquations = count
End Function

' --- InsertEquationReference ---  
' Inserts a cross-reference to a specific equation number.
' Prompts user for the equation number and inserts "Eq. (n)" text.
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
    
    Selection.TypeText Text:="Eq. (" & eqRef & ")"
End Sub
