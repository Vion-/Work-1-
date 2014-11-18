Attribute VB_Name = "Functions"
Sub SelectHosts2()

Sheets("Hosts").Select
Range("A2").Select

End Sub

Sub SelectHosts()

Sheets("Hosts").Select
Range("A1").Select

End Sub

Sub RemoveDups()

'removes duplicates from single middleware cells in hosts sheet

Dim dic As Object, cell As Range, temp As Variant
Dim i As Long
Set dic = CreateObject("scripting.dictionary")
With dic
    For Each cell In Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row)
        .RemoveAll
        If Len(cell.Value) > 0 Then
            temp = Split(Replace(cell.Value, "", ""), ",")
            For i = 0 To UBound(temp)
                If Not .Exists(temp(i)) Then .Add temp(i), temp(i)
            Next i
            cell.Value = Join(.Keys, ",")
        End If
    Next cell
End With

End Sub

Sub About()

UserForm2.Show

End Sub

Sub RoundedRectangle5_Click()

    UserForm1.Show
    
End Sub

Sub Clean()

'Cleans all sheets and adds column headers

Application.ScreenUpdating = False
    Sheets("BRIO - ABOVE").Cells.Clear
    Sheets("Hosts").Cells.Clear
    Sheets("Hosts").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Scp Hostname"
    ActiveCell.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Sw Type2"
   ActiveCell.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1:B1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Sheets("Hosts").Select
    Range("A1").Select
    Selection.Copy
    Sheets("BRIO - ABOVE").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("C1").Select
    ActiveSheet.Paste
    Sheets("Hosts").Select
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("BRIO - ABOVE").Select
    Range("Z1").Select
    ActiveSheet.Paste
    Columns("C:C").Select
    Application.CutCopyMode = False
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Columns("Z:Z").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Application.ScreenUpdating = True
    Sheets("Home").Select
    MsgBox ("All sheets cleaned, Please insert raw Data"), vbInformation
    
End Sub
