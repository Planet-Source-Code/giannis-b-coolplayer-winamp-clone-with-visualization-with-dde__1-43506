VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDex"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private G As New clsGraph

Public Enum SFlags
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
    CDLCANCEL = 32755
End Enum

Private Type OPENFILENAME
    StructureSize As Long
    hOwner As Long
    Instance As Long
    Filter As String
    CustomFilter As String
    MaxCustFilter As Long
    FilterIndex As Long
    Filename As String
    MaxFile As Long
    FileTitle As String
    MaxFileTitle As Long
    InitialDir As String
    DialogTitle As String
    Flags As SFlags
    FileOffset As Integer
    FileExtension As Integer
    DefaultExt As String
    CustData As Long
    Hook As Long
    TemplateName As String
End Type

Private OFN As OPENFILENAME
Private R As Long

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Property Get hOwner() As Long
    hOwner = OFN.hOwner
End Property
Public Property Let hOwner(ByVal Own As Long)
    OFN.hOwner = Own
End Property
Public Property Get Flags() As SFlags
    Flags = OFN.Flags
End Property
Public Property Let Flags(ByVal NFlag As SFlags)
    OFN.Flags = NFlag
End Property
Public Property Get DialogTitle() As String
    DialogTitle = OFN.DialogTitle
End Property
Public Property Let DialogTitle(Tit As String)
    OFN.DialogTitle = Tit
End Property
Public Property Get DefaultExt() As String
    DefaultExt = OFN.DefaultExt
End Property
Public Property Let DefaultExt(Def As String)
    OFN.DefaultExt = Def
End Property
Public Property Get Filename() As String
    Filename = OFN.Filename
End Property
Public Property Let Filename(Name As String)
    OFN.Filename = Name
End Property
Public Property Get FileTitle() As String
    FileTitle = OFN.FileTitle
End Property
Public Property Let FileTitle(Tit As String)
    OFN.FileTitle = Tit
End Property
Public Property Get Filter() As String
    Filter = OFN.Filter
End Property
Public Property Let Filter(Fil As String)

    Dim i As Integer
    With OFN
     .Filter = Fil
     For i = 1 To Len(.Filter)
      If Mid(.Filter, i, 1) = "|" Then
       Mid(.Filter, i, 1) = vbNullChar
      End If
     Next i
     .Filter = .Filter & String(2, 0)
    End With

End Property
Public Property Get FilterIndex() As Long
    FilterIndex = OFN.FilterIndex
End Property
Public Property Let FilterIndex(Ind As Long)
    OFN.FilterIndex = Ind
End Property
Public Property Get InitDir() As String
    InitDir = OFN.InitialDir
End Property
Public Property Let InitDir(sDir As String)
    OFN.InitialDir = sDir
End Property
Public Sub ShowOpen(Top As Boolean)
  
    On Error GoTo SError
    Call InitializeOFN(Top)
    R = GetOpenFileName(OFN)

    With OFN
     If R > 0 Then
      R = InStr(.FileTitle, vbNullChar)
      If R Then .FileTitle = Left(.FileTitle, R - 1)
      R = InStr(.Filename, vbNullChar)
      If R Then .Filename = Left(.Filename, R - 1)
     End If
    End With

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub ShowSave(Top As Boolean)
  
    On Error GoTo SError
    Call InitializeOFN(Top)
    R = GetSaveFileName(OFN)

    With OFN
     If R > 0 Then
      R = InStr(.FileTitle, vbNullChar)
      If R Then .FileTitle = Left(.FileTitle, R - 1)
      R = InStr(.Filename, vbNullChar)
      If R Then .Filename = Left(.Filename, R - 1)
     End If
    End With

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub InitializeOFN(Top As Boolean)
    
    On Error GoTo IError
    With OFN
     Call G.OnTop(.hOwner, Top)
     .StructureSize = Len(OFN)
     .Filename = String(255, 0)
     .FileTitle = String(255, 0)
     .MaxFile = 32768
     .MaxFileTitle = 255
     .Instance = App.hInstance
     .Flags = .Flags Or OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT
    End With

IError:
    If Err.Number <> 0 Then Exit Sub

End Sub
