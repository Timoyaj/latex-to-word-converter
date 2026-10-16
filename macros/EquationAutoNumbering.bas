Attribute VB_Name = "EquationAutoNumbering"
' ============================================================
' Equation Auto-Numbering Macros for Microsoft Word
' Provides automatic equation numbering and cross-referencing
' for documents with native OMML equations.
' ============================================================

' --- RenumberAllEquations ---
' Walks through the document and renumbers all equations
' that have equation number markers (right-aligned numbers in
' equation tables created by the converter).
Public Sub RenumberAllEquations()
    ' Implementation:
    ' 1. Search for all 1x3 tables that look like equation tables
    '    (invisible borders, 3 columns, with a number pattern in column 3)
    ' 2. Renumber them sequentially starting from 1
    ' 3. Format: (1), (2), (3), etc.
    ' 4. Update any cross-reference bookmarks
    Dim tbl As Table
    Dim eqNum As Long
    eqNum = 0
    
    For Each tbl In ActiveDocument.Tables
        If tbl.Columns.Count = 3 Then
            ' Check if this looks like an equation table
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
                ' Right-align the number
                tbl.Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                tbl.Cell(1, 3).Range.Font.Name = "Cambria Math"
                tbl.Cell(1, 3).Range.Font.Size = 12
            End If
        End If
    Next tbl
    
    MsgBox "Renumbered " & eqNum & " equations.", vbInformation, "Equation Numbering"
End Sub

' --- InsertEquationNumber ---
' Inserts an equation number at the current cursor position
' by converting the current paragraph into a 3-column invisible table.
Public Sub InsertEquationNumber()
    ' Find the next equation number
    Dim nextNum As Long
    nextNum = CountEquations() + 1
    
    ' Create a 3-column table at cursor
    Dim tbl As Table
    Set tbl = ActiveDocument.Tables.Add( _
        Range:=Selection.Range, _
        NumRows:=1, _
        NumColumns:=3)
    
    ' Remove all borders
    tbl.Borders.Enable = False
    
    ' Set column widths (proportional)
    tbl.Columns(1).PreferredWidth = InchesToPoints(0.5)
    tbl.Columns(2).PreferredWidth = InchesToPoints(4.5)
    tbl.Columns(3).PreferredWidth = InchesToPoints(1.5)
    
    ' Center the equation column
    tbl.Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    ' Add equation number
    tbl.Cell(1, 3).Range.Text = "(" & nextNum & ")"
    tbl.Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    tbl.Cell(1, 3).Range.Font.Name = "Cambria Math"
    tbl.Cell(1, 3).Range.Font.Size = 12
    
    ' Place cursor in center cell for equation entry
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
        MsgBox "Equation (" & eqRef & ") not found. Document has " & totalEqs & " equations.", vbExclamation, "Equation Not Found"
        Exit Sub
    End If
    
    Selection.TypeText Text:="Eq. (" & eqRef & ")"
End Sub
