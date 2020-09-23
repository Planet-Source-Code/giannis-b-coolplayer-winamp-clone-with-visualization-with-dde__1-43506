Attribute VB_Name = "basPl"
Option Explicit

Public Const Def As String = "\Default.m3u"
Public Sub LoadM3U(Name As String)

    On Error GoTo LoadError
    Dim strData As String
    Dim strD As String
    Dim i As Long

    With frmPl.l
     .ListItems.Clear
     Close #1
     Open Name For Input As #1
      Line Input #1, strData
      If strData = "#EXTM3U" Then
       Do Until EOF(1)
        Line Input #1, strData
        If Left(strData, 8) = "#EXTINF:" Then
         strD = Right(strData, Len(strData) - 8)
         Call .ListItems.Add(.ListItems.Count + 1, , Right(strD, Len(strD) - InStr(Right(strData, Len(strData) - 8), ",")))
         .ListItems.Item(.ListItems.Count).SubItems(1) = File.gettime(Left(strD, InStr(strD, ",") - 1))
        ElseIf Left(strData, 8) <> "#EXTINF:" Then
         .ListItems(.ListItems.Count).Tag = Ini.getshortpath(strData)
        End If
       Loop
      End If
     Close #1
     Call SetScroller(1)
    End With

LoadError:
    If Err.Number <> 0 Then Close #1: Exit Sub

End Sub
Public Sub LoadPLS(Name As String)

    On Error GoTo LoadError
    Dim strData As String, Temp As String

    With frmPl.l.ListItems
     .Clear
     Close #1
     Open Name For Input As #1
      Line Input #1, strData
      If strData = "[playlist]" Then
       Do Until EOF(1)
        Line Input #1, strData
        If Left(strData, 4) = "File" Then
         Temp = Right(strData, Len(strData) - Len(Left(strData, InStrRev(strData, "="))))
        ElseIf Left(strData, 5) = "Title" Then
         Call .Add(.Count + 1, , Right(strData, Len(strData) - InStr(strData, "=")))
         .Item(.Count).Tag = Ini.getshortpath(Temp)
        ElseIf Left(strData, 6) = "Length" Then
         Temp = Right(strData, Len(strData) - Len(Left(strData, InStrRev(strData, "="))))
         .Item(.Count).SubItems(1) = File.gettime(Temp)
        End If
       Loop
      End If
     Close #1
     Call SetScroller(1)
    End With

LoadError:
    If Err.Number <> 0 Then Close #1: Exit Sub

End Sub
Public Sub DisableForms(S As Boolean)

    frmMn.Enabled = S
    frmPl.Enabled = S

End Sub
Public Sub GetPlay(G As Boolean)

    If frmPl.l.ListItems.Count = 0 Then
     Call OpenForFile(frmMn): Call PlayUp
    Else
     Call FullName(G): Call Play
    End If

End Sub
Public Sub RemDoubs(l As MSComctlLib.ListView, S As Boolean)

    On Error GoTo RError
    Dim i As Single

    With l.ListItems
     If S Then
      For i = 1 To .Count
      If i <= .Count - 1 Then
       If LCase(.Item(i)) Like LCase(.Item(i + 1)) Then
         Call .Remove(i + 1)
         i = i - 1
       End If
        End If
      Next i
     End If
     Call SetScroller(1)
    End With

RError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub AddFile(Path As String, Name As String, Start As Boolean, Optional DoPlay As Boolean)

    On Error GoTo OError
    With frmPl.l.ListItems
     If Path <> "" Then
      Call .Add(.Count + 1, , Left(Name, InStrRev(Name, ".") - 1))
      .Item(.Count).SubItems(1) = MP3.gettime(Path, True)
      .Item(.Count).Tag = Ini.getshortpath(Path)
      If Start Then Call SetScroller(.Count)
      If DoPlay Then Call GetPlay(True)
     End If
    End With

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub FullName(G As Boolean)

    On Error GoTo GError
    With frmPl.l
     GL.sTrack = .ListItems.Item(.SelectedItem.Index).Tag
     If G Then PL.St = .SelectedItem.Text
    End With

GError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub LoadList(Name As String, Sort As Boolean)

    On Error GoTo LError
    If Name = "" Then Exit Sub
    With frmPl
     If Lst.getext(Name) = "m3u" Then
      Call LoadM3U(Name)
     ElseIf Lst.getext(Name) = "pls" Then
      Call LoadPLS(Name)
     End If
    End With
    If Sort Then Call CheckSort(CI.bSort)

LError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub MP3PlayInfo(Name As String)

    On Error GoTo MP3Error
    With MP3
     .FileName = Name
     .GetHeader
     If Not .ValidHeader Then Exit Sub
    End With

    With frmMn
     Call SkinString(.Bit, MP3.Bitrate)
     Call SkinString(.Hrz, Left(MP3.Frequency, 2))
     Call Ster
    End With

MP3Error:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub NextP()

    On Error GoTo PError
    Dim l As Long

    With frmPl.l
     If .SelectedItem.Index >= .ListItems.Count Then Exit Sub
     l = l + 1
     Call SetScroller(GL.lPos + 1)
     Call GetPlay(True)
    End With

PError:
     If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub Pause()

    On Error GoTo PauseError
    With MP
     If .playstate = 1 Then
      .Pause
      Call ButtonChoice("paused")
     ElseIf .playstate <> 1 And GL.sTrack <> "" Then
      .Play
      Call ButtonChoice("resumed")
     End If
    End With
    
PauseError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub PrevP()

    On Error GoTo PError
    Dim l As Long

    With frmPl.l
     If .SelectedItem.Index <= 1 Then Exit Sub
     l = l - 1
     Call SetScroller(GL.lPos - 1)
     Call GetPlay(True)
    End With

PError:
     If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub TextChange(S As String)

    On Error GoTo TError
    Dim i As Long

    If S = "" Then Exit Sub
    With frmPl.l.ListItems
     For i = 1 To .Count
      If S Like Mid(.Item(i), 1, Len(S)) Then
       Call SetScroller(i): Exit For
      End If
     Next
    End With

TError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub ScrollText(l As PictureBox)

    On Error GoTo SError
    If GL.vDrag = False And GL.bDrag = False And GL.sDrag = False Then
     Call NameLabels(PL.St, MP.duration)
    End If

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
