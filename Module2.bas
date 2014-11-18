Attribute VB_Name = "Module2"
Sub Process()

Dim LastRow As Integer
LastRow = 0

Application.ScreenUpdating = False
'Finds last cell with data on row for both sheets
LastRow = Sheets("BRIO - ABOVE").Range("C1").End(xlDown).Row - 1
Call SelectHosts
'---------------------------------------------------------------------
'---------------------------------------------------------------------
'---------------------------------------------------------------------
'---------------------------------------------------------------------
'repeats all the below structure on ever cell containing hostname
Dim a As Integer

For a = 2 To LastRow
'------------------------------------
'------------------------------------
'------------------------------------
'scans entire column for same hostname and saves middlewware on middleware variable
Dim i As Integer
Dim HostName As String
Dim Middleware As String
Dim CellEmpty2 As Byte

Sheets("BRIO - ABOVE").Select
Range("C" & a).Select

HostName = ActiveCell.Value
ActiveCell.Offset(0, 23).Select
Middleware = ActiveCell.Value
ActiveCell.Offset(0, -23).Select

For i = 2 To LastRow
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value = HostName Then
        ActiveCell.Offset(0, 23).Select
        Middleware = Middleware & "," & ActiveCell.Value
        ActiveCell.Offset(0, -23).Select
        'MsgBox (Middleware)
    End If
'----------------------------------------------------------
'----------------------------------------------------------
'Passes data on to new sheet
    If i = LastRow Then
    Call SelectHosts2
    CellEmpty2 = 0
        Do While CellEmpty2 = 0
            If ActiveCell.Value <> "" Then
'--------------------------------------------------------------
'Checks if hostname has already been included into results with all Middlewares
                If ActiveCell.Value = HostName Then
                    CellEmpty2 = 1
                End If
'--------------------------------------------------------------
                ActiveCell.Offset(1, 0).Select
            Else
                ActiveCell.Value = HostName
                ActiveCell.Offset(0, 1).Select
                ActiveCell.Value = Middleware
                ActiveCell.Offset(0, -1).Select
                CellEmpty2 = 1
            End If
        Loop
'----------------------------------------------------------
'----------------------------------------------------------
    End If
Next i
'------------------------------------
'------------------------------------
'------------------------------------
Next a
'---------------------------------------------------------------------
'---------------------------------------------------------------------
'---------------------------------------------------------------------
'---------------------------------------------------------------------
Call SelectHosts
Call RemoveDups
Application.ScreenUpdating = True

End Sub
