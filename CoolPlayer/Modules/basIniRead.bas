Attribute VB_Name = "basIR"
Option Explicit

Public Type Ini
    iIcon As Integer
    iSkin As Integer
    yLay As Byte

    BList As Boolean
    bRand As Boolean
    bSnap As Boolean
    bGraph As Boolean
    bClick As Boolean
    bSplash As Boolean
    bSort As Boolean
    bInst As Boolean
    bAss As Boolean
    bMute As Boolean
    bTop As Boolean
    bLoop As Boolean
    bDoub As Boolean
    bTray As Boolean
    bScroll As Boolean
    bStartUp As Boolean

    sPath As String
End Type

Public CI As Ini
Public Sub Card()

    On Error Resume Next
    Call Shell("rundll32.exe shell32,Control_RunDLL mmsys.cpl @1", 5)

End Sub
Public Sub CheckSplash()

    On Error Resume Next
    With frmMn
     .Ini = Ini.LoadIni("Sets", "Splash")
     If .Ini = "true" Then
      CI.bSplash = True
     Else
      CI.bSplash = False
     End If

     .Ini = Ini.LoadIni("Sets", "Inst")
     If .Ini = "true" Then
      CI.bInst = True
     Else
      CI.bInst = False
     End If

     If App.PrevInstance And CI.bInst = False Then
      Call Execute(Command)
      End
     End If
 
     .Ini = Ini.LoadIni("List", "Dubs")
     If .Ini = "true" Then
      CI.bDoub = True
     Else
      CI.bDoub = False
     End If
 
     .Ini = Ini.LoadIni("List", "Sort")
     If .Ini = "true" Then
      CI.bSort = True
     Else
      CI.bSort = False
     End If
    End With
    frmSp.Visible = IIf(CI.bSplash, True, False)
    Call NoMode

End Sub
Private Sub Execute(S As String)

    On Error Resume Next
    With frmMn.Ini
     .LinkTopic = "CoolPlayer|frmDDE"
     .LinkMode = 2
     .LinkExecute Command
    End With

End Sub


Public Sub LoadTransparency(Apply As Boolean)

    On Error Resume Next
    Dim O As Object
    Set O = frmSet

    If Ini.LoadIni("Sets", "Trans") = 0 Or Ini.LoadIni("Sets", "Trans") > 255 Then
     If Apply Then Call SetLay(255)
     CI.yLay = 255
    End If

    CI.yLay = CByte(Ini.LoadIni("Sets", "Trans"))
    If Apply Then Call SetLay(CI.yLay)
    frmSet.sliT.Value = CInt(CI.yLay)

    CI.iIcon = CInt(Ini.LoadIni("Sets", "Icon"))
    frmSet.sliI.Value = CInt(CI.iIcon)
    Call ChangeIcon

    With frmMn
     .Ini = Ini.LoadIni("Sets", "Tray")
     If .Ini = "true" Then
      O.chkMin.Value = 1
      CI.bTray = True
      If Apply Then Call HideForms(False)
     Else
      O.chkMin.Value = 0
      CI.bTray = False
     End If

     .Ini = Ini.LoadIni("Sets", "Inst")
     If .Ini = "true" Then
      O.chkInst.Value = 1
      CI.bInst = True
     Else
      O.chkInst.Value = 0
      CI.bInst = False
     End If

     .Ini = Ini.LoadIni("List", "Sort")
     If .Ini = "true" Then
      O.chkSort.Value = 1
      CI.bSort = True
     Else
      O.chkSort.Value = 0
      CI.bSort = False
     End If

     .Ini = Ini.LoadIni("List", "Single")
     If .Ini = "true" Then
      O.chkSingl.Value = 1
      CI.bClick = True
     Else
      O.chkSingl.Value = 0
      CI.bClick = False
     End If

     .Ini = Ini.LoadIni("Sets", "Splash")
     If .Ini = "true" Then
      O.chkSplash.Value = 1
      CI.bSplash = True
     Else
      O.chkSplash.Value = 0
      CI.bSplash = False
     End If

     .Ini = Ini.LoadIni("Sets", "Graph")
     If .Ini = "true" Then
      O.chkGraph.Value = 1
      CI.bGraph = True
     Else
      O.chkGraph.Value = 0
      CI.bGraph = False
     End If

     .Ini = Ini.LoadIni("Sets", "Snap")
     If .Ini = "true" Then
      O.chkSnap.Value = 1
      CI.bSnap = True
     Else
      O.chkSnap.Value = 0
      CI.bSnap = False
     End If

     .Ini = Ini.LoadIni("Sets", "Assoc")
     If .Ini = "true" Then
      O.chkAss.Value = 1
      CI.bAss = True
     Else
      O.chkAss.Value = 0
      CI.bAss = False
     End If

     .Ini = Ini.LoadIni("Sets", "Scroll")
     If .Ini = "true" Then
      O.chkScroll.Value = 1
      CI.bScroll = True
     Else
      O.chkScroll.Value = 0
      CI.bScroll = False
     End If
    End With

End Sub
Public Sub SaveIniSettings(Sav As Boolean)
    
    On Error GoTo SError
    With CI
     If .bTop Then
      Call Ini.saveini("Sets", "Top", "true")
     Else
      Call Ini.saveini("Sets", "Top", "false")
     End If

     If .bStartUp Then
      Call Ini.saveini("Sets", "Startup", "true")
     Else
      Call Ini.saveini("Sets", "Startup", "false")
     End If

     If .bInst Then
      Call Ini.saveini("Sets", "Inst", "true")
     Else
      Call Ini.saveini("Sets", "Inst", "false")
     End If

     If .bSnap Then
      Call Ini.saveini("Sets", "Snap", "true")
     Else
      Call Ini.saveini("Sets", "Snap", "false")
     End If

     If .bGraph Then
      Call Ini.saveini("Sets", "Graph", "true")
     Else
      Call Ini.saveini("Sets", "Graph", "false")
     End If

     If .bSort Then
      Call Ini.saveini("List", "Sort", "true")
     Else
      Call Ini.saveini("List", "Sort", "false")
     End If

     If .bClick Then
      Call Ini.saveini("List", "Single", "true")
     Else
      Call Ini.saveini("List", "Single", "false")
     End If

     If .bSplash Then
      Call Ini.saveini("Sets", "Splash", "true")
     Else
      Call Ini.saveini("Sets", "Splash", "false")
     End If

     If .bAss Then
      Call Ini.saveini("Sets", "Assoc", "true")
     Else
      Call Ini.saveini("Sets", "Assoc", "false")
     End If

     If .BList Then
      Call Ini.saveini("List", "Show", "true")
     Else
      Call Ini.saveini("List", "Show", "false")
     End If

     If .bLoop Then
      Call Ini.saveini("List", "Loop", "true")
     Else
      Call Ini.saveini("List", "Loop", "false")
     End If

     If .bRand Then
      Call Ini.saveini("List", "Rand", "true")
     Else
      Call Ini.saveini("List", "Rand", "false")
     End If

     If .bMute Then
      Call Ini.saveini("List", "Mute", "true")
     Else
      Call Ini.saveini("List", "Mute", "false")
     End If

     If .bDoub Then
      Call Ini.saveini("List", "Dubs", "true")
     Else
      Call Ini.saveini("List", "Dubs", "false")
     End If

     If .bTray Then
      Call Ini.saveini("Sets", "Tray", "true")
     Else
      Call Ini.saveini("Sets", "Tray", "false")
     End If

     If .bScroll Then
      Call Ini.saveini("Sets", "Scroll", "true")
     Else
      Call Ini.saveini("Sets", "Scroll", "false")
     End If

     Call Ini.saveini("Sets", "Vol", frmMn.picMVol.Left)
     Call Ini.saveini("Sets", "X", frmMn.Left)
     Call Ini.saveini("Sets", "Y", frmMn.Top)
     Call Ini.saveini("Sets", "Trans", CStr(.yLay))
     Call Ini.saveini("Sets", "Icon", CStr(.iIcon))

     .iSkin = frmSkn.lstSkins.ListIndex
     If Sav Then Call Ini.saveini("Sets", "Skin", CStr(.iSkin))
     Call Ini.saveini("Sets", "SPath", .sPath)
    End With

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub LoadIniSettings(Apply As Boolean, Frm As Long)

    On Error Resume Next
    Dim O As Object
    Set O = frmSet

    With frmMn
     .Ini = Ini.LoadIni("Sets", "Top")
     If .Ini = "true" Then
      O.chkTop.Value = 1
      CI.bTop = True
      Call DialogTop
     Else
      O.chkTop.Value = 0
      CI.bTop = False
      Call DialogBottom
     End If

     If Apply Then
      .Left = Ini.LoadIni("Sets", "X")
      .Top = Ini.LoadIni("Sets", "Y")
      Call ListLeft
     End If

     .Ini = Ini.LoadIni("List", "Show")
     If Apply Then Call ListLeft
     If .Ini = "true" Then
      CI.BList = True
      frmPl.Visible = True
     Else
      CI.BList = False
      frmPl.Visible = False
     End If

     .Ini = Ini.LoadIni("List", "Loop")
     If .Ini = "true" Then
      O.chkLoop.Value = 1
      CI.bLoop = True
     Else
      O.chkLoop.Value = 0
      CI.bLoop = False
     End If

     .Ini = Ini.LoadIni("List", "Rand")
     If .Ini = "true" Then
      O.chkRand.Value = 1
      CI.bRand = True
     Else
      O.chkRand.Value = 0
      CI.bRand = False
     End If

     .Ini = Ini.LoadIni("List", "Mute")
     If .Ini = "true" Then
      O.chkMute.Value = 1
      CI.bMute = True
      MP.Mute = True
      frmMnu.mnuMuteP.Checked = True
     Else
      O.chkMute.Value = 0
      CI.bMute = False
      MP.Mute = False
      frmMnu.mnuMuteP.Checked = False
     End If

     .Ini = Ini.LoadIni("List", "Dubs")
     If .Ini = "true" Then
      O.chkDubs.Value = 1
      CI.bDoub = True
     Else
      O.chkDubs.Value = 0
      CI.bDoub = False
     End If

     .Ini = Ini.LoadIni("Sets", "Startup")
     If .Ini = "true" Then
      O.chkStart.Value = 1
      CI.bStartUp = True
     Else
      O.chkStart.Value = 0
      CI.bStartUp = False
     End If
     Call MoveVol(CSng(Ini.LoadIni("Sets", "Vol")))
     Call LoadTransparency(Apply)
    End With

End Sub
Public Sub GetAllColors()
    
    On Error GoTo AllError
    Dim O As Object
    Set O = frmPl.l

    With frmMn.Ini
     .Text = Ini.loadcolor("Text", "Normal", frmSkn.Files.Path)
     If Len(.Text) = 6 Then .Text = "#" & .Text
     If Len(.Text) = 7 Then
      O.ForeColor = Graph.GetRGB(.Text)
     ElseIf .Text = "#Error" Or .Text = "Error" Then
      O.ForeColor = &HFF00&
     End If

     .Text = Ini.loadcolor("Text", "mbBG", frmSkn.Files.Path)
     If Len(.Text) = 6 Then .Text = "#" & .Text
     If Len(.Text) = 7 Then
      O.BackColor = Graph.GetRGB(.Text)
     ElseIf .Text = "#Error" Or .Text = "Error" Then
      O.BackColor = &H0&
     End If
    End With

AllError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub SaveOptions()

    On Error GoTo SError
    With frmSet
     CI.bLoop = IIf(.chkLoop.Value = 1, True, False)
     CI.bRand = IIf(.chkRand.Value = 1, True, False)
     CI.bGraph = IIf(.chkGraph.Value = 1, True, False)
     CI.bClick = IIf(.chkSingl.Value = 1, True, False)
     CI.bSort = IIf(.chkSort.Value = 1, True, False)
     CI.bSplash = IIf(.chkSplash.Value = 1, True, False)
     CI.bSnap = IIf(.chkSnap.Value = 1, True, False)
     CI.bInst = IIf(.chkInst.Value = 1, True, False)
     CI.bAss = IIf(.chkAss.Value = 1, True, False)
     CI.bDoub = IIf(.chkDubs.Value = 1, True, False)
     CI.bTray = IIf(.chkMin.Value = 1, True, False)
     CI.bStartUp = IIf(.chkStart.Value = 1, True, False)

     With frmMn
      If frmSet.chkMute.Value = 1 Then
       frmMnu.mnuMuteP.Checked = True
       CI.bMute = True
       MP.Mute = True
      ElseIf frmSet.chkMute.Value = 0 Then
       frmMnu.mnuMuteP.Checked = False
       CI.bMute = False
       MP.Mute = False
      End If
     End With

     If .chkStart.Value = 0 Then
      Call Reg.runstartup(App.Title, App.Path, False)
     ElseIf .chkStart.Value = 1 Then
      Call Reg.runstartup(App.Title, App.Path, True)
     End If

     If .chkTop.Value = 0 Then
      CI.bTop = False
      Call DialogBottom
     ElseIf .chkTop.Value = 1 Then
      CI.bTop = True
      Call DialogTop
     End If

     If .chkScroll.Value = 1 Then
      CI.bScroll = True
     ElseIf .chkScroll.Value = 0 Then
      CI.bScroll = False
     End If

     CI.yLay = CByte(.sliT.Value + 5)
     Call SaveIniSettings(False)
     Call SetLay(CI.yLay)
    End With

SError:
    If Err.Number <> 0 Then Call SaveIniSettings(False): Exit Sub

End Sub
Public Sub AddCommands(Comm As String)

    With frmPl
     Select Case Comm
      Case ""
       Call LoadM3U(App.Path & Def)
       Call CheckSort(CI.bSort)

      Case Else
       Comm = Ini.GetLongPath(Comm)
       If Lst.getext(Comm) = "m3u" Or Lst.getext(Comm) = "pls" Then
        Call LoadList(Comm, True)
       Else
        Call AddFile(Comm, Right(Comm, Len(Comm) - InStrRev(Comm, "\")), True, True)
       End If
     End Select
    End With

End Sub
Public Sub CreateKey()

    On Error GoTo IErr
    If Len(GetSetting("CoolPlayer", "Plugins", "Registered")) = 0 Then
     Call CreateKeys
    End If
    Set Ini = CreateObject("Misc_v1.clsIni")
    Set Vis = CreateObject("visOut.clsMain")
    Set Lst = CreateObject("Misc_v1.clsList")
    Set Tray = CreateObject("Misc_v1.clsTray")
    Set MP = CreateObject("Misc_v1.clsMCI")
    Set File = CreateObject("Misc_v1.clsFiles")
    Set CD = CreateObject("Misc_v1.clsCDex")
    Set IDx = CreateObject("Misc_v1.clsID3")
    Set Graph = CreateObject("Misc_v1.clsGraph")
    Set MP3 = CreateObject("Misc_v1.clsMP3info")
    Set Reg = CreateObject("Misc_v1.clsReg")
    Call LoadPictures: DoEvents: Call LoadSkinNumber

IErr:
    If Err.Number <> 0 Then
     Call DeleteKeys
     MsgBox ("Failed to load:  " & vbCrLf & App.Path & "\Plugins\Misc_v1.dll") _
            , vbCritical, "Object not found!": End
    End If

End Sub
