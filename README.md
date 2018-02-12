# VBA-tricks

## Searche files in a folder with some part of its name

```
 Dim sFil As String
 Dim sPath As String

 With Application.FileDialog(msoFileDialogFolderPicker)
    .Show
        If .SelectedItems.Count Then
        sPath = .SelectedItems(1)
        Else: Exit Sub
        End If
 End With

 ChDir sPath
 sFil = Dir("*")
 Do While sFil <> ""
 If InStr(sFil, "21") <> 0 And InStr(sFil, "R01") <> 0 And InStr(sFil, "BW") <> 0 Then 'Searching file with "21" and "R01" and "BW" in its name
 [C10] = (sPath & "\" & sFil) 'Filepath as output
 [D10] = sFil 'Filename as output
 End If
 sFil = Dir
 Loop
 ```
