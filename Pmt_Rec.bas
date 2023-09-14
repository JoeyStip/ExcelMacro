Attribute VB_Name = "Pmt_Rec"
Sub InvoiceLoadFile()

Dim WB As Workbook
Set WB = ActiveWorkbook
Dim recSH As Worksheet
Set recSH = WB.ActiveSheet

MsgBox ("Choose folder to save invoice load file in")

Dim fDialog As FileDialog
Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

'Show the dialog. -1 means success!
If fDialog.Show = -1 Then
   Debug.Print fDialog.SelectedItems(1) 'The full path to the file selected by the user
End If

Dim templatepath As String
templatepath = "\\emcshr01.spwynet.com\General Accounting\Lease Payables\Basic User\Joey\Macros\pmt rec\Invoice Load File\SPW July 2023 Invoice Load File - EMPTY.xlsx"

Dim CloseMo As String
CloseMo = InputBox("Which month/year is this close file for?", "Select Month", Format(Date, "Mmm yyyy"))


Dim savePathInv As String
savePathInv = fDialog.SelectedItems(1) & "\SPW " & CloseMo & " Invoice Load File.xlsx"

FileCopy templatepath, savePathInv




Dim RecData As range
Dim recDataLastRow As Integer

recDataLastRow = recSH.range("a7").End(xlDown).Row
Set RecData = recSH.range("a7", recSH.Cells(recDataLastRow, 15).Address)

Dim recArr As Variant
recArr = RecData

Dim i As Integer
Dim j As Integer
CloseMo = Format(CloseMo, "yyyymm")
Debug.Print (CloseMo)

Dim ILarr(1 To 5000, 1 To 15) As Variant

For i = 1 To UBound(recArr)
    ILarr(i, 2) = Left(recArr(i, 2), 4)
    ILarr(i, 3) = recArr(i, 3)
    ILarr(i, 4) = CloseMo
    ILarr(i, 5) = recArr(i, 4)
    ILarr(i, 6) = "Principal"
    If recArr(i, 12) = "" Then
        ILarr(i, 12) = 0
    Else
        ILarr(i, 7) = recArr(i, 12)
    End If
    ILarr(i, 9) = recArr(i, 6)
    ILarr(i, 1) = Left(recArr(i, 2), 4) & recArr(i, 3) & Left(recArr(i, 4), InStr(1, recArr(i, 4), "-") - 1)
Next

Dim invWB As Workbook
Set invWB = Workbooks.Open(savePathInv)
Dim invWS As Worksheet
Set invWS = invWB.Sheets(1)

invWS.range("a7", "k9999") = ILarr

End Sub
Sub SelectFileMacro2()

Dim fileSelect As Variant

Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogFilePicker)

Dim vrtSelectedItem As Variant

With fd

    If .Show = -1 Then
    
        For Each vrtSelectedItem In .SelectedItems
        
            MsgBox "The path is: " & vrtSelectedItem
            
            fileSelect = vrtSelectedItem
            
            
        Next vrtSelectedItem
        Else
    End If
End With

End Sub

Sub Format_New_rec()

Dim PmtRec As Workbook
Dim RecSheet As Worksheet

Set PmtRec = ActiveWorkbook()
Set RecSheet = PmtRec.ActiveSheet()

Debug.Print (RecSheet.Name)

'change date in upper right hand corner of rec sheet
RecSheet.range("a5").Value = Format(Date, "Mmm-yy")

'change date in name of pmt recon sheet
RecSheet.Name = Format(Date, "MMM yyyy") & " PMT Recon"


End Sub


Sub GL_Pull()


Dim databook As Workbook
Dim allComp As Worksheet
Dim headers As range
Dim dataPull As range

Set databook = ActiveWorkbook
Dim resp As Integer
resp = MsgBox("Process " & databook.Name & "?", vbYesNo)
If resp = 7 Then
    End
End If
    
Dim Pullnum As Integer
Pullnum = InputBox("which pull is this?", "Pull #")

Dim month As String
month = InputBox("which month?", "Month", Format(Date, "Mmmm"))

Set headers = ActiveSheet.range("a1", range("a1").End(xlToRight))

Set allComp = databook.ActiveSheet
Set dataPull = allComp.UsedRange

allComp.Name = "All Companies " & month & " pull " & Pullnum

Dim arrRow As Integer
Dim arrCol As Integer

arrRow = dataPull.Rows.Count
arrCol = dataPull.Columns.Count

Dim arr As Variant
arr = dataPull

Dim i As Integer
Dim j As Integer
Dim Jrow As Integer
Dim Krow As Integer
Dim Arow As Integer
Jrow = 1
Krow = 1
Arow = 1

Dim AngieArr(1 To 999, 1 To 36) As Variant
Dim KarenArr(1 To 999, 1 To 36) As Variant
Dim JoeyArr(1 To 999, 1 To 36) As Variant

For i = 1 To UBound(arr)
    If arr(i, 18) = "5200" Or arr(i, 18) = "5235" Or arr(i, 18) = "5257" Then
        For j = 1 To 36
            JoeyArr(Jrow, j) = arr(i, j)
        Next
        Jrow = Jrow + 1
        
    ElseIf arr(i, 18) = "5243" Or arr(i, 18) = "5245" Or arr(i, 18) = "5247" Or arr(i, 18) = "5242" Then
        For j = 1 To 36
            AngieArr(Arow, j) = arr(i, j)
        Next
        Arow = Arow + 1
        
    ElseIf arr(i, 18) = "5241" Or arr(i, 18) = "5244" Or arr(i, 18) = "5246" Or arr(i, 18) = "5248" Then
        For j = 1 To 36
            KarenArr(Krow, j) = arr(i, j)
        Next
        Krow = Krow + 1
    End If
Next

Dim KarenSHname As String
Dim AngieSHname As String
Dim JoeySHname As String

Dim sheetsToMake As New Collection

sheetsToMake.Add "Karen " & month & " pull " & Pullnum
sheetsToMake.Add "Joey " & month & " pull " & Pullnum
sheetsToMake.Add "Angie " & month & " pull " & Pullnum

For x = 1 To sheetsToMake.Count
    Sheets.Add.Name = sheetsToMake(x)
Next

Dim arrays As New Collection

arrays.Add KarenArr
arrays.Add JoeyArr
arrays.Add AngieArr

For x = 1 To sheetsToMake.Count
    Worksheets(sheetsToMake(x)).range("a2", "aj999").Value = arrays(x)
Next

For x = 1 To sheetsToMake.Count
    headers.Copy Worksheets(sheetsToMake(x)).range("a1", "aj1")
    'Worksheets(sheetsToMake(x)).UsedRange.Columns.AutoFit
Next

Sheets.Add.Name = "Stats"

Dim stats As Worksheet
Set stats = Worksheets("Stats")

Dim lastrow As Integer
Dim sh As Worksheet
Dim sumCell As range

For x = 1 To sheetsToMake.Count
    Set sh = Worksheets(sheetsToMake(x))
    lastrow = sh.range("j2").End(xlDown).Row
    
    Set sumCell = sh.Cells(lastrow + 2, 10)
    
    With sumCell
        .Formula = "=sum(j2:j" & lastrow & ")"
        .Style = "Comma"
    End With
    
    stats.Cells(x + 1, 1).Value = sheetsToMake(x)
    stats.Cells(x + 1, 2).Formula = "='" & sheetsToMake(x) & "'!" & sumCell.Address
    
    sh.Columns(2).Insert
    sh.Columns(4).Insert
    With sh.range("b1")
        .Value = "Join Code"
        .Interior.ColorIndex = 34
    End With
    With sh.range("d1")
        .Value = "MLA #"
        .Interior.ColorIndex = 34
    End With
    
    For y = 2 To lastrow
        sh.Cells(y, 2).FormulaR1C1 = "=rc[1]&rc[31]"
        sh.Cells(y, 4).FormulaR1C1 = "=left(rc[-1],9)"
    Next
    
    sh.UsedRange.Columns.AutoFit
    sh.Activate
    ActiveWindow.ScrollRow = lastrow - 40
    
Next

'inserting total for all companies pull
lastrow = allComp.range("j2").End(xlDown).Row
Set sumCell = allComp.Cells(lastrow + 2, 10)
sumCell.Formula = "=sum(j2:j" & lastrow & ")"
allComp.UsedRange.Columns.AutoFit

With stats
    .range("b5").Formula = "=sum(b2:b4)"
    .range("a5").Value = "Total"

    .range("a7").Value = "All Companies Pull"
    .range("b7").Formula = "='" & allComp.Name & "'!" & sumCell.Address

    .range("a8").Value = "Diff"
    .range("b8").Formula = "=b5-b7"

    .UsedRange.Columns.AutoFit

End With

End Sub

Sub UnmatchedPmts()


Dim databook As Workbook
Dim allComp As Worksheet
Dim headers As range
Dim dataPull As range

Set databook = ActiveWorkbook
Dim resp As Integer
resp = MsgBox("Process " & databook.Name & "?", vbYesNo)
If resp = 7 Then
    End
End If
    
Dim Pullnum As Integer
Pullnum = InputBox("which pull is this?", "Pull #")

Dim month As String
month = InputBox("which month?", "Month", Format(Date, "Mmm"))

Set headers = ActiveSheet.range("a1", range("a1").End(xlToRight))

Set allComp = databook.ActiveSheet
Set dataPull = allComp.UsedRange

allComp.Name = "All Comp " & month & " Unmchd " & Pullnum

Dim arrRow As Integer
Dim arrCol As Integer

arrRow = dataPull.Rows.Count
arrCol = dataPull.Columns.Count

Dim arr As Variant
arr = dataPull

Dim i As Integer
Dim j As Integer
Dim Jrow As Integer
Dim Krow As Integer
Dim Arow As Integer
Jrow = 1
Krow = 1
Arow = 1

Dim AngieArr(1 To 999, 1 To 20) As Variant
Dim KarenArr(1 To 999, 1 To 20) As Variant
Dim JoeyArr(1 To 999, 1 To 20) As Variant

For i = 1 To UBound(arr)
    If arr(i, 2) = "5200-Speedway" Or arr(i, 2) = "5235-TRMC Retail" Or arr(i, 2) = "5257-SWTO LLC" Then
        For j = 1 To 20
            JoeyArr(Jrow, j) = arr(i, j)
        Next
        Jrow = Jrow + 1
        
    ElseIf arr(i, 2) = "5243-Tesoro Northstore Company" Or arr(i, 2) = "5245-Western Refining Retail" Or arr(i, 2) = "5247-Giant Four Corners LLC" Or arr(i, 2) = "5242-Tesoro West Coast Company" Then
        For j = 1 To 20
            AngieArr(Arow, j) = arr(i, j)
        Next
        Arow = Arow + 1
        
    ElseIf arr(i, 2) = "5241-Northern Tier Bakery LLC" Or arr(i, 2) = "5244-Northern Tier Retail LLC" Or arr(i, 2) = "5246-Giant Stop-N-Go of New Mexico" Or arr(i, 2) = "5248-Tesoro Sierra Properties" Then
        For j = 1 To 20
            KarenArr(Krow, j) = arr(i, j)
        Next
        Krow = Krow + 1
    End If
Next

Dim KarenSHname As String
Dim AngieSHname As String
Dim JoeySHname As String

Dim sheetsToMake As New Collection

sheetsToMake.Add "Karen " & month & " Unmatched " & Pullnum
sheetsToMake.Add "Joey " & month & " Unmatched " & Pullnum
sheetsToMake.Add "Angie " & month & " Unmatched " & Pullnum

For x = 1 To sheetsToMake.Count
    Sheets.Add.Name = sheetsToMake(x)
Next

Dim arrays As New Collection

arrays.Add KarenArr
arrays.Add JoeyArr
arrays.Add AngieArr

For x = 1 To sheetsToMake.Count
    Worksheets(sheetsToMake(x)).range("a2", "t999").Value = arrays(x)
Next

For x = 1 To sheetsToMake.Count
    headers.Copy Worksheets(sheetsToMake(x)).range("a1", "aj1")
    'Worksheets(sheetsToMake(x)).UsedRange.Columns.AutoFit
Next

Sheets.Add.Name = "Stats"

Dim stats As Worksheet
Set stats = Worksheets("Stats")

Dim lastrow As Integer
Dim sh As Worksheet
Dim sumCell As range

For x = 1 To sheetsToMake.Count
    Set sh = Worksheets(sheetsToMake(x))
    lastrow = sh.range("j2").End(xlDown).Row
    
    Set sumCell = sh.Cells(lastrow + 2, 5)
    
    With sumCell
        .Formula = "=sum(e2:e" & lastrow & ")"
        .Style = "Comma"
    End With
    
    stats.Cells(x + 1, 1).Value = sheetsToMake(x)
    stats.Cells(x + 1, 2).Formula = "='" & sheetsToMake(x) & "'!" & sumCell.Address

    
    sh.UsedRange.Columns.AutoFit
    sh.Activate
    ActiveWindow.ScrollRow = lastrow - 40
    
Next

'inserting total for all companies pull
lastrow = allComp.range("e2").End(xlDown).Row
Set sumCell = allComp.Cells(lastrow + 2, 5)
sumCell.Formula = "=sum(e2:e" & lastrow & ")"
allComp.UsedRange.Columns.AutoFit

With stats
    .range("b5").Formula = "=sum(b2:b4)"
    .range("a5").Value = "Total"

    .range("a7").Value = "All Companies Unmatched"
    .range("b7").Formula = "='" & allComp.Name & "'!" & sumCell.Address

    .range("a8").Value = "Diff"
    .range("b8").Formula = "=b5-b7"

    .UsedRange.Columns.AutoFit

End With

End Sub

Sub updateUnmatched2()

Dim WB As Workbook
Set WB = ActiveWorkbook
Dim recon As Worksheet
Dim Unmatched As Worksheet

Dim wksh As Worksheet

For Each wksh In WB.Sheets
    If InStr(1, wksh.Name, "Recon") > 0 Then
    Dim resp As Integer
    resp = MsgBox("Update unmatched on this sheet: " & wksh.Name & "?", vbYesNo)
        If resp = 6 Then
            Set recon = wksh
        End If
    End If
Next

If recon Is Nothing Then
    MsgBox ("Could not locate Recon sheet")
    End
End If

For Each wksh In WB.Sheets
    If InStr(1, wksh.Name, "Unmatched") > 0 Then
    resp = MsgBox("Use " & wksh.Name & " date to update unmatched payments on the Recon sheet?", vbYesNo)
        If resp = 6 Then
            Set Unmatched = wksh
            Exit For
        End If
    End If
Next

If Unmatched Is Nothing Then
    MsgBox ("Could not locate Unmatched Payments")
    End
End If
' 6 = yes, 7 = no

Debug.Print (recon.Name)

Debug.Print (Unmatched.Name)

Dim RecData As range
Dim UnmData As range
Dim RecdataRows As Integer
Dim UNMdataRows As Integer
Dim recArr As Variant
Dim UNMarr As Variant
Dim expPMT(1 To 5000, 1 To 1) As Long


dataRows = recon.range("a7").End(xlDown).Row
Set RecData = recon.range("a7", recon.Cells(dataRows, 11))
recArr = RecData

UNMdataRows = Unmatched.range("a1").End(xlDown).Row
Set unmrecdata = Unmatched.range("a1", Unmatched.Cells(UNMdataRows, 14))
UNMarr = unmrecdata

Dim i As Integer
Dim j As Integer
Dim checkCellRec As String
Dim checkCellUNM As String


For i = 1 To UBound(recArr)
        checkCellRec = recArr(i, 2) & recArr(i, 4) & recArr(i, 6)
    For j = 1 To UBound(UNMarr)
        checkCellUNM = UNMarr(i + 1, 2) & UNMarr(i + 1, 4) & UNMarr(i + 1, 14)
        If checkCellUNM = checkCellRec Then
            expPMT(i, 1) = UNMarr(j, 5)
            Exit For
        End If
    Next
Next
            

End Sub

Sub updateUnmatched()

Dim Recbook As Workbook
Dim PmtRec As Worksheet
Dim databook As Workbook
Dim Unmatched As Worksheet

Set Recbook = ActiveWorkbook
Set PmtRec = Recbook.ActiveSheet

Set databook = Workbooks("unmatched payments 2")
Set Unmatched = databook.ActiveSheet


Dim recArr(1 To 999, 1 To 11) As Variant
Dim UnmatchedArr(1 To 999, 1 To 20) As Variant

Dim x As Integer

For x = 1 To UBound(recArr)
    For y = 1 To UBound(recArr)
        If recArr(x, 3) = UnmatchedArr(y, 3) Then
        Else
        
        
        

End Sub

Sub PushPull()

Dim PullBook As Workbook
Dim pullSh As Worksheet
Dim x As Integer

Set PullBook = ActiveWorkbook


'Debug.Print (PullBook.Sheets.Count)

Dim sh As Worksheet

For Each sh In Worksheets
    If Left(sh.Name, InStr(1, sh.Name, " ") - 1) = "Angie" Then
        Dim AngiePull As Worksheet
        Set AngiePull = sh
        Debug.Print ("Angie ws name set")
    ElseIf Left(sh.Name, InStr(1, sh.Name, " ") - 1) = "Joey" Then
        Dim JoeyPull As Worksheet
        Set JoeyPull = sh
    ElseIf Left(sh.Name, InStr(1, sh.Name, " ") - 1) = "Karen" Then
        Dim KarenPull As Worksheet
        Set KarenPull = sh
    End If
Next





End Sub


Sub resetSub()

Workbooks("July all companies PULL 2 unedited").Close savechanges:=wdDoNotSaveChanges
Workbooks.Open ("Z:\Lease Payables\PowerPlan Monthly\Payment Reconciliation\PP_07-2023\Joey\July all companies PULL 2 unedited.XLSX")

End Sub
Sub InvoiceLoad()

Dim recWB As Workbook
Dim recSH As Worksheet

Dim ILtemplate As Workbook

Dim RecSaveName As String
Dim resp As Integer

Set recWB = ActiveWorkbook
Set recSH = recWB.ActiveSheet

'Debug.Print (recWB.Name)
'Debug.Print (RecSh.Name)


Set ILtemplate = Workbooks.Open("\\emcshr01.spwynet.com\General Accounting\Lease Payables\Basic User\Joey\Macros\Payment Rec Macro reference files\SPW July 2023 Invoice Load File.xlsx")

RecSaveName = Left(ILtemplate.Name, 4) & Format(Date - 1, "Mmmm yyyy") & Right(ILtemplate.Name, 23)

resp = MsgBox("Select save folder", vbokaycancel)
Debug.Print (resp)
If resp = 6 Then
    End
End If

'Debug.Print (RecSaveName)

'ILtemplate.SaveAs Filename:=RecSaveName

End Sub

Function GetFolder() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .title = "Select a Folder"
    .AllowMultiSelect = False
    '.InitialFileName = strPath
    If .Show <> -1 Then GoTo nextcode
    sItem = .SelectedItems(1)
End With
nextcode:
GetFolder = sItem
Set fldr = Nothing
End Function
Function GetFile() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFilePicker)
With fldr
    .title = "Select a File"
    .AllowMultiSelect = False
    '.InitialFileName = strPath
    If .Show <> -1 Then GoTo nextcode
    sItem = .SelectedItems(1)
End With
nextcode:
GetFile = sItem
Set fldr = Nothing
End Function

Sub SortPull()
'
' SortPull Macro
' Sorts pull data between Karen, Angie and Joey
' Made using Macro recorder

'
    Sheets("Sheet1").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy Before:=Sheets(1)
    Sheets("Sheet1 (2)").Select
    Sheets("Sheet1 (2)").Name = "May Pull 6 Karen"
    Sheets("May Pull 6 Karen").Select
    Sheets("May Pull 6 Karen").Copy Before:=Sheets(1)
    Sheets("May Pull 6 Karen (2)").Select
    Sheets("May Pull 6 Karen (2)").Name = "May Pull 6 Angie"
    Sheets("May Pull 6 Angie").Select
    Sheets("May Pull 6 Angie").Copy Before:=Sheets(1)
    Sheets("May Pull 6 Angie (2)").Select
    Sheets("May Pull 6 Angie (2)").Name = "May Pull 6 Joey"
    Sheets("May Pull 6 Joey").Select
    range("B23").Select
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "Stats"
    Sheets("Stats").Select
    Sheets("Stats").Move After:=Sheets(4)
    range("A1").Select
    ActiveCell.FormulaR1C1 = "Stats"
    range("A1").Select
    Selection.Font.Bold = True
    range("A3").Select
    ActiveCell.FormulaR1C1 = "Joey"
    range("A4").Select
    ActiveCell.FormulaR1C1 = "Karen"
    range("A5").Select
    ActiveCell.FormulaR1C1 = "Angie"
    range("A6").Select
    ActiveCell.FormulaR1C1 = "total"
    range("A8").Select
    ActiveCell.FormulaR1C1 = "Raw data"
    range("A9").Select
    ActiveCell.FormulaR1C1 = "diff"
    range("B9").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C-R[-3]C"
    range("A1:B9").Select
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
    range("A6").Select
    Selection.Font.Underline = xlUnderlineStyleSingle
    range("A9").Select
    Selection.Font.Underline = xlUnderlineStyleSingle
    range("D11").Select
End Sub
