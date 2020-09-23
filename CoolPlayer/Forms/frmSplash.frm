VERSION 5.00
Begin VB.Form frmSp 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   3285
   ClientLeft      =   4905
   ClientTop       =   3840
   ClientWidth     =   4605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSt 
      Interval        =   2000
      Left            =   0
      Top             =   3120
   End
End
Attribute VB_Name = "frmSp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Click()
    frmSp.Hide
End Sub
Private Sub Form_Initialize()
    Call CreateKey
End Sub
Private Sub Form_Load()

    On Error GoTo LError
    Call CheckSplash: ReDim PL.Tr(0 To 0)
    Call Graph.LoadElliptic(frmSp)
    Call Graph.Ontop(frmSp.hwnd, True)
    Call AddCommands(Command): Load frmDDE

LError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub tmrSt_Timer()

    Unload frmSp: DoEvents
    Call LoadMain

End Sub
