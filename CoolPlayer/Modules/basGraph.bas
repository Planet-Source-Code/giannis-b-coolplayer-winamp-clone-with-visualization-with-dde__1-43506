Attribute VB_Name = "basGr"
Option Explicit

Private Type Globe
    lPos As Long
    xSli As Long


    X As Single
    Y As Single
    sOff As Single

    vDrag As Boolean
    bDrag As Boolean
    sDrag As Boolean
    lDown As Boolean
    bClick As Boolean
    lMin As Boolean
    mMin As Boolean


    sTrack As String

End Type

Private Type Plobe
    Tr() As Long
    TrS As Long
    Ct As Long

    Vis As Integer
    St As String
End Type

Public GL As Globe
Public PL As Plobe
Dim ix(512) As Integer

Public Sub KeyChoice(k As Integer)

    Select Case k
     Case 13
      Call GetPlay(True)
     Case 27
      Call StopPlay
     Case 37
      Call GotoTime(-5)
     Case 39
      Call GotoTime(5)
     Case 46
      Call RemoveItem
    End Select

End Sub
Public Function SelectTracks()

    On Error Resume Next
    Dim i As Long, j As Long

    With frmPro.lstP
     For i = 0 To .ListCount - 1
      If .Selected(i) = True Then
       j = j + 1: PL.Tr(j) = (i + 1)
       PL.Ct = j
      End If
     Next i: PL.TrS = UBound(PL.Tr)
     frmPro.lblSl.Caption = j & " from " & .ListCount & " items, " & _
                            CLng((j * 100) / .ListCount) & "% of list."
    End With

End Function

Public Sub ShowTab(STab As Integer)

    On Error Resume Next
    Dim i As Integer

    For i = 0 To 4
     If i = STab - 1 Then
      frmSet.fraGen(STab - 1).Visible = True
     Else
      frmSet.fraGen(i).Visible = False
     End If
    Next i

End Sub
Public Sub TEvent(X As Long)

    On Error GoTo TError
    Select Case X
     Case &H202
      If frmMn.Enabled = False And frmPl.Enabled = False Then Exit Sub
      Call HideForms(True)
     Case &H205
      If frmMn.Enabled And frmPl.Enabled Then
       Call frmMnu.PopupMenu(frmMnu.mnuMTray, True)
      ElseIf frmID3.Visible Then
       Call frmID3.PopupMenu(frmID3.mnuT, True)
      ElseIf frmAb.Visible Then
       Call frmAb.PopupMenu(frmAb.mnuT, True)
      ElseIf frmTim.Visible Then
       Call frmTim.PopupMenu(frmTim.mnuT, True)
      ElseIf frmSkn.Visible Then
       Call frmSkn.PopupMenu(frmSkn.mnuT, True)
      End If
    End Select

TError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub SetLay(l As Byte)

    If l = 0 Then
     Call Graph.SetLayered(frmMn.hwnd, 255)
     Call Graph.SetLayered(frmPl.hwnd, 255)
     CI.yLay = 255
    Else
     Call Graph.SetLayered(frmMn.hwnd, l)
     Call Graph.SetLayered(frmPl.hwnd, l)
    End If

End Sub
Public Sub CheckPics()

    Call PlUp
    Call ShuffUp
    Call LoopUp
    Call TopUp
    Call TExitUp

End Sub
Public Sub PaintPlaylist()

    Dim i As Integer, O As Object
    Set O = frmMn.Pledit

    With frmPl
     For i = 0 To 14
      Call .PaintPicture(O, 0, (i * 435) + 0, 180, 435, 0, 630, 180, 435)
     Next i
     Call .PaintPicture(O, 0, 0, 375, 300, 0, 0, 375, 300)
     For i = 1 To 10
      Call .PaintPicture(O, i * 375, 0, 375, 300, 1905, 0, 375, 300)
     Next i
     Call .PaintPicture(O, 1305, 0, 1500, 300, 390, 315, 1500, 300)
     Call .PaintPicture(O, 3750, 0, 375, 300, 2295, 0, 375, 300)
     Call .PaintPicture(O, 0, 6390, 1875, 570, 0, 1080, 1875, 570)
     Call .PaintPicture(O, 1875, 6390, 2250, 570, 1890, 1080, 2250, 570)

     'Call .AddBar.PaintPicture(O, 0, 0, 45, 810, 720, 1665, 45, 810)
     'Call .RemBar.PaintPicture(O, 0, 0, 45, 1080, 1500, 1665, 45, 1080)
     'Call .SelBar.PaintPicture(O, 0, 0, 45, 810, 2250, 1665, 45, 810)
     'Call .MisBar.PaintPicture(O, 0, 0, 45, 810, 3000, 1665, 45, 810)
     'Call .ListBar.PaintPicture(O, 0, 0, 45, 810, 3750, 1665, 45, 810)

     'Call .picAddDir.PaintPicture(O, 0, 0, 330, 270, 0, 1955, 330, 270)
     'Call .picAddUrl.PaintPicture(O, 0, 0, 330, 270, 0, 1665, 330, 270)

     'Call .picCrop.PaintPicture(O, 0, 0, 330, 270, 815, 1955, 330, 270)
     'Call .picRemAll.PaintPicture(O, 0, 0, 330, 270, 815, 1665, 330, 270)
     'Call .picRemMisc.PaintPicture(O, 0, 0, 330, 270, 815, 2515, 330, 270)

     'Call .picInv.PaintPicture(O, 0, 0, 330, 270, 1560, 1665, 330, 270)
     'Call .picSelZero.PaintPicture(O, 0, 0, 330, 270, 1560, 1955, 330, 270)

     'Call .picInfo.PaintPicture(O, 0, 0, 330, 270, 2310, 1955, 330, 270)
     'Call .picSort.PaintPicture(O, 0, 0, 330, 270, 2310, 1665, 330, 270)

     'Call .picLoad.PaintPicture(O, 0, 0, 330, 270, 3065, 1955, 330, 270)
     'Call .picNew.PaintPicture(O, 0, 0, 330, 270, 3065, 1665, 330, 270)
    End With

End Sub
Public Sub Shuffle(Start As Boolean, Optional Check As Boolean)

    If Start Then GL.lPos = Lst.random(1, frmPl.l.ListItems.Count, GL.lPos)
    Call SetScroller(GL.lPos): Call GetPlay(True)

End Sub
Public Sub DialogBottom()

    Call Graph.Ontop(frmMn.hwnd, False)
    Call Graph.Ontop(frmPl.hwnd, False)
    Call Graph.Ontop(frmSet.hwnd, False)

End Sub
Public Sub DialogTop()

    Call Graph.Ontop(frmMn.hwnd, True)
    Call Graph.Ontop(frmPl.hwnd, True)
    Call Graph.Ontop(frmSet.hwnd, True)

End Sub
Public Sub LoadPictures()

    On Error Resume Next
    Call PaintPlaylist
    Call DrawTitleBar
    Call AboutUp
    Call MinUp
    Call TopUp
    Call ExitUp

    Call SliderUp
    Call EqUp
    Call VolUp
    Call BalUp
    Call BackBalance
    Call BackVolume
    Call BackPic
    Call PlUp
    Call ShuffUp
    Call LoopUp
    Call PrevUp
    Call PlayUp
    Call StopUp
    Call PauseUp
    Call NextUp
    Call OpenUp

    Call PExitUp
    Call TExitUp
    Call PlFileUp
    Call PlOptUp
    Call PlListUp
    Call PlTrackUp
    Call PlRemUp

End Sub
