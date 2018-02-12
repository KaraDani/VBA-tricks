```vba
Sub selectpath()

Dim path As Variant
With Application.FileDialog(msoFileDialogFolderPicker)
    .Show
    If .SelectedItems.Count Then
    [C3] = .SelectedItems(1)
    End If
End With

End Sub

Sub selectpath2()

Dim path As Variant
With Application.FileDialog(msoFileDialogFolderPicker)
    .Show
    If .SelectedItems.Count Then
    [C5] = .SelectedItems(1)
    End If
End With

End Sub

Sub opencopyall()

Application.ScreenUpdating = False

On Error Resume Next
If Workbooks("PERSONAL.XLS") Is Nothing Then
    If Workbooks("PERSONAL.XLSB") Is Nothing Then
        If Application.Workbooks.Count <> 1 Then
        MsgBox ("Close all the other workbooks")
        End
        End If
    ElseIf Application.Workbooks.Count <> 2 Then
    MsgBox ("Close all the other workbooks")
    End
    End If
End If


Dim macro As Workbook
Set macro = Application.ActiveWorkbook
    Dim excelfile As Variant


    path = [C3].Value
    ChDir path
    excelfile = Dir("*.*")

    
    Do While excelfile <> ""
        Workbooks.Open Filename:=path & "\" & excelfile
        excelfile = Dir
    Loop

 
 Workbooks.Add
ActiveWorkbook.SaveAs Filename:=macro.Sheets("Controll").[C5].Value & "\ConcatResults.xlsx"
 CopyTargetBookmark = 1
 For Each Workbook In Application.Workbooks
 If Workbook.Name <> "ConcatResults.xlsx" And Workbook.Name <> "PERSONAL.XLS" And Workbook.Name <> macro.Name Then
 Workbook.Activate
 Workbook.Worksheets(1).UsedRange.Copy
 Workbooks("ConcatResults.xlsx").Activate
 Range("A" & CopyTargetBookmark).Select
 ActiveSheet.Paste
 CopyTargetBookmark = CopyTargetBookmark + Workbook.Worksheets(1).UsedRange.Rows.Count
 End If
 Next Workbook


Workbooks("ConcatResults.xlsx").Activate
Workbooks("ConcatResults.xlsx").Save

Dim wkbWorkbook As Workbook
 For Each wkbWorkbook In Application.Workbooks
 If wkbWorkbook.Name <> ActiveWorkbook.Name And wkbWorkbook.Name <> macro.Name Then wkbWorkbook.Close
 Next wkbWorkbook
 
 Application.ScreenUpdating = True
 
 End Sub
 ```
