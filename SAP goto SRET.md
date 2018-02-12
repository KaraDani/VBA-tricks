```vba
' Linenumbers Help:
'3  NET Operative reports&metrics       9  PP     Production Planning
'4  BM     Batch Management             10 PS     Project System
'5  CO Controlling                      11 QM     Quality Management
'6  CS     Customer Service             12 SD Sales And Distribution
'7  FI     Financial Accounting         13 WM/IM  Warehouse/Inventory
'8  MM     Materials Management         14 Data Archiving
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub goSret(Childrennumber As String, Sessionnum As String, Linenumber As String, Reportname As String, ActivateOrNot As Boolean, NewSession As Boolean)

'create new session if needed
If NewSession = True Then
    If Childrennumber = 0 Then
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
    session.createsession
    Sleep (2000)
    End If
    
    If Childrennumber = 1 Then
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(1).Children(0)
    session.createsession
    Sleep (2000)
    End If
    
    If Childrennumber = 2 Then
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(2).Children(0)
    session.createsession
    Sleep (2000)
    End If
End If


'connecting to the selected SAP session
If Childrennumber = 0 Then
If Sessionnum = 1 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
If Sessionnum = 2 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(1)
If Sessionnum = 3 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(2)
If Sessionnum = 4 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(3)
If Sessionnum = 5 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(4)
If Sessionnum = 6 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(5)
End If

If Childrennumber = 1 Then
If Sessionnum = 1 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(1).Children(0)
If Sessionnum = 2 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(1).Children(1)
If Sessionnum = 3 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(1).Children(2)
If Sessionnum = 4 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(1).Children(3)
If Sessionnum = 5 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(1).Children(4)
If Sessionnum = 6 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(1).Children(5)
End If

If Childrennumber = 2 Then
If Sessionnum = 1 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(2).Children(0)
If Sessionnum = 2 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(2).Children(1)
If Sessionnum = 3 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(2).Children(2)
If Sessionnum = 4 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(2).Children(3)
If Sessionnum = 5 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(2).Children(4)
If Sessionnum = 6 Then Set session = GetObject("SAPGUI").GetScriptingEngine.Children(2).Children(5)
End If



session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n sret" 'type in SRET to the tcode field
session.findById("wnd[0]").sendVKey 0 'enter into SRET

'activate window if it was requested
If ActivateOrNot = True Then AppActivate ("P20(" & Sessionnum & ")/005 Application Tree Report Selection General")

session.findById("wnd[0]/usr/lbl[5," & Linenumber & "]").SetFocus 'set focus to the selected main treeline
session.findById("wnd[0]").sendVKey 33 'expand the tree of the selected main line
session.findById("wnd[0]").sendVKey 71 'open search
session.findById("wnd[1]/usr/txtRSYSF-STRING").Text = Reportname 'enter report name
session.findById("wnd[1]").sendVKey 0 'execute search
On Error GoTo Errhandler
session.findById("wnd[2]/usr/lbl[0,2]").SetFocus 'select the report which has been found
On Error GoTo 0
session.findById("wnd[2]").sendVKey 2 'press ok to the search box
session.findById("wnd[0]").sendVKey 2 'enter to report

Exit Sub


'if the user has not the same main line numbering in the tree _
code will use this part which will expand the whole tree for the search, _
it will work, but it's slower
Errhandler:

session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[12]").press

'go back to the main menu
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "sret" 'type in SRET to the tcode field
session.findById("wnd[0]").sendVKey 0 'enter into SRET
session.findById("wnd[0]").sendVKey 33 'expand the whole tree
session.findById("wnd[0]").sendVKey 71 'open search
session.findById("wnd[1]/usr/txtRSYSF-STRING").Text = Reportname 'enter report name
session.findById("wnd[1]").sendVKey 0 'execute search
Resume 'go back and resume where the error occured


End Sub
```
