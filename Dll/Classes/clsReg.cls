VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Private hCurKey As Long
Public Sub PublicReg(Index As Integer, AppPath As String, ExeName As String)

    Select Case Index
     Case 1: Call AssociateFile(".mp3", AppPath, ExeName)
     Case 2: Call AssociateFile(".wav", AppPath, ExeName)
     Case 3: Call AssociateFile(".wma", AppPath, ExeName)
     Case 4: Call AssociateFile(".mid", AppPath, ExeName)
     Case 5: Call AssociateFile(".midi", AppPath, ExeName)
     Case 6: Call AssociateList(".m3u", AppPath, ExeName)
     Case 7: Call AssociateList(".pls", AppPath, ExeName)
    End Select

End Sub
Public Sub SaveString(hKey As Long, Key As String, Value As String, KeyValue As String)

    On Error Resume Next
    Call RegCreateKey(hKey, Key, hCurKey)
    Call RegSetValueEx(hCurKey, Value, 0, 1, ByVal KeyValue, Len(KeyValue))
    Call RegCloseKey(hCurKey)

End Sub
Private Sub AssociateFile(Ext As String, AppPath As String, ExeName As String)

    On Error GoTo AError
    Call SaveString(HKEY_CLASSES_ROOT, Ext, "", "CoolPlayer.file")
    Call SaveString(HKEY_CLASSES_ROOT, "CoolPlayer.file", "", "CoolPlayer media file")
    Call SaveString(HKEY_CLASSES_ROOT, "CoolPlayer.file" & "\DefaultIcon", "", AppPath & "\" & ExeName & ".exe, 0")
    Call SaveString(HKEY_CLASSES_ROOT, "CoolPlayer.file" & "\Shell\Open\command", "", AppPath & "\" & ExeName & ".exe %1")
    Call SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0)

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub AssociateList(Ext As String, AppPath As String, ExeName As String)

    On Error GoTo AError
    Call SaveString(HKEY_CLASSES_ROOT, Ext, "", "CoolPlayer.list")
    Call SaveString(HKEY_CLASSES_ROOT, "CoolPlayer.list", "", "CoolPlayer playlist")
    Call SaveString(HKEY_CLASSES_ROOT, "CoolPlayer.list" & "\DefaultIcon", "", AppPath & "\" & ExeName & ".exe, 0")
    Call SaveString(HKEY_CLASSES_ROOT, "CoolPlayer.list" & "\Shell\Open\command", "", AppPath & "\" & ExeName & ".exe %1")
    Call SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0)

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub RunStartup(AppTitle As String, AppPath As String, Run As Boolean)

    If Run Then
     Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", AppTitle, AppPath & "\" & AppTitle & ".exe")
    Else
     Call DeleteString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", AppTitle, AppPath & "\" & AppTitle & ".exe")
    End If

End Sub
Public Sub DeleteString(hKey As Long, Key As String, Value As String, KeyValue As String)

    On Error Resume Next
    Call RegOpenKey(hKey, Key, hCurKey)
    Call RegDeleteValue(hCurKey, Value)
    Call RegCloseKey(hCurKey)

End Sub
