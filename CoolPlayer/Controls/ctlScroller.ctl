VERSION 5.00
Begin VB.UserControl ctlScroller 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   LockControls    =   -1  'True
   ScaleHeight     =   406
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   20
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   75
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   0
      Top             =   0
      Width           =   120
   End
End
Attribute VB_Name = "ctlScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Scroll()

Private iMax As Long, Val As Long
Private iY As Single, Down As Boolean
Public Function MSlide(l As Long)

    Dim X As Long
    l = IIf(l <= 0, 1, l)
    l = IIf(l >= 388, 388, l)
    S.Top = l: DoEvents
    X = (l * iMax) / 388
    Val = IIf(X = 0, 1, X)

End Function
Public Sub Update()

    On Error Resume Next
    Dim i As Integer
    For i = 0 To 14
     Call UserControl.PaintPicture(frmMn.Pledit, 0, i * 28, 20, 29, 31, 42, 20, 29)
    Next i
    If Down = False Then Call USlider

End Sub
Public Sub USlider()

    On Error Resume Next
    Call S.PaintPicture(frmMn.Pledit, 0, 0, 8, 18, 52, 53, 8, 18)
    
End Sub
Public Sub DSlider()

    On Error Resume Next
    Call S.PaintPicture(frmMn.Pledit, 0, 0, 8, 18, 61, 53, 8, 18)

End Sub
Private Sub S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     Down = True: iY = Y
     Call DSlider
    End If

End Sub
Private Sub s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo MError
    If Down Then
     If iMax <> 0 Then RaiseEvent Scroll
     Call DSlider
     Call MSlide(S.Top + Y - iY)
    End If

MError:
    If Err.Number <> 0 Then Exit Sub
    
End Sub
Private Sub s_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 And Down Then
     Down = False: Call USlider
    End If

End Sub
Private Sub UserControl_Resize()
    
    With UserControl
     .Width = 300: .Height = 6090
    End With
    Call Update

End Sub
Public Property Get Value() As Long
    Value = Val
End Property
Public Property Let Value(l As Long)

    On Error Resume Next
    Val = l
    S.Top = (Val / iMax) * 388
    Call USlider

End Property
Public Property Let Max(l As Long)

    iMax = IIf(l < 0, 0, l)
    If iMax = 0 Then S.Top = 0
    Call Update

End Property
