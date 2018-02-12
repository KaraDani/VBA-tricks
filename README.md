# VBA-tricks

 Dim sFil As String
 Dim sPath As String

 With Application.FileDialog(msoFileDialogFolderPicker)
    .Show
        If .SelectedItems.Count Then
        sPath = .SelectedItems(1)
        Else: Exit Sub
        End If
 End With

'X021 BWP01R01
 ChDir sPath
 sFil = Dir("*")
 Do While sFil <> ""
 If InStr(sFil, "21") <> 0 And InStr(sFil, "R01") <> 0 And InStr(sFil, "BW") <> 0 Then
 [C10] = (sPath & "\" & sFil)
 [D10] = sFil
 End If
 sFil = Dir
 Loop
