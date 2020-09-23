Attribute VB_Name = "basFuncs"
Option Explicit

Dim X As Integer, Y As Integer
Dim H(3) As Integer
Public Device As Long
Public sOut(512) As Integer
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Sub DrawBack(s As PictureBox, X As Long, Y As Long)

    Call SetPixel(s.hdc, X, Y, 0)
    Call SetPixel(s.hdc, X, Y, RGB(0, 50, 0))

End Sub
Private Sub Scope(s As PictureBox)

    On Error Resume Next
    s.Cls
    Call FillBox(s, 0, RGB(0, 75, 0))

    For X = 0 To 74
     H(0) = sOut(X) / 775
     For H(1) = 0 To Abs(H(0))
      H(2) = IIf(H(1) ^ 3 > 255, 255, H(1) ^ 3)
      H(3) = Abs(H(0)) - H(1)
      Call SetPixel(s.hdc, X, 8 - 1 - H(3), RGB(H(2), 0, 0))
      Call SetPixel(s.hdc, X, 8, RGB(0, 75, 0))
      Call SetPixel(s.hdc, X, 8 + 1 + H(3), RGB(0, H(2), 0))
     Next H(1)
    Next X

End Sub
Private Sub FillBox(s As PictureBox, cl As Long, cl1 As Long)

    On Error Resume Next
    s.Cls: s.BackColor = cl
    For X = 0 To 37
     For Y = 0 To 8
      Call SetPixel(s.hdc, 2 * X, 2 * Y + 1, cl1)
     Next Y
    Next X

End Sub
Public Sub Visualize(s As PictureBox, DType As Integer)

    Select Case DType
     Case 0: Call Scope(s)
     Case 1: Call Spectrum(s)
     Case 2: Call FillBox(s, 0, RGB(0, 75, 0))
    End Select

End Sub
Private Sub Spectrum(s As PictureBox)

    On Error Resume Next
    s.Cls: s.BackColor = RGB(0, 0, 0)

    For X = 0 To 74
     H(0) = Abs(sOut(X)) / 1600
     H(0) = IIf(H(0) > 13, 13, H(0))
     For H(1) = 0 To H(0)
      H(2) = H(1) ^ 2.1
      'H(3) = IIf(H(2) > 255, 255, H(2))
      Call SetPixel(s.hdc, X, 16 - H(1), RGB(94, 175, 94)) 'RGB(255 - H(3), H(3), 0))
     Next H(1)
     For H(2) = 0 To (16 - H(1))
      Call SetPixel(s.hdc, 2 * X, 2 * H(2) + 1, RGB(0, 75, 0))
     Next H(2)
    Next X

End Sub
