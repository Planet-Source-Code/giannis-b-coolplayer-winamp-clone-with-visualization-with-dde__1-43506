VERSION 5.00
Begin VB.Form frmMnu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menus"
   ClientHeight    =   15
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuO 
      Caption         =   "Main options"
      Begin VB.Menu mnuOOpt 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuOS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOID3 
         Caption         =   "&ID3 Editor..."
      End
      Begin VB.Menu mnuOAb 
         Caption         =   "&CoolPlayer..."
      End
      Begin VB.Menu mnuOS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOEx 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuPM 
      Caption         =   "Main popup"
      Begin VB.Menu mnuMPFile 
         Caption         =   "F&ile"
         Begin VB.Menu mnuPMDir 
            Caption         =   "Add &dir..."
         End
         Begin VB.Menu mnuPMFile 
            Caption         =   "Add &file..."
         End
      End
      Begin VB.Menu mnuMPList 
         Caption         =   "&Playlist"
         Begin VB.Menu mnuPMLo 
            Caption         =   "&Load playlist"
         End
         Begin VB.Menu mnuPMSa 
            Caption         =   "&Save playlist"
         End
      End
      Begin VB.Menu mnuPMS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMPPre 
         Caption         =   "&Options..."
         Begin VB.Menu mnuPMOp 
            Caption         =   "&Preferences..."
         End
         Begin VB.Menu mnuPMSk 
            Caption         =   "S&kin browser..."
         End
      End
      Begin VB.Menu mnuPMAb 
         Caption         =   "&CoolPlayer..."
      End
      Begin VB.Menu mnuPMS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMPCars 
         Caption         =   "Sound card..."
      End
      Begin VB.Menu mnuPMS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPMEx 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuMTray 
      Caption         =   "Main tray"
      Begin VB.Menu mnuMTrayRes 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuMTrayS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTrayOpt 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuMTrayAb 
         Caption         =   "CoolPlayer..."
      End
      Begin VB.Menu mnuMTrayS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTrayEx 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMute 
      Caption         =   "Mute"
      Begin VB.Menu mnuMuteP 
         Caption         =   "Mute"
      End
   End
   Begin VB.Menu mnuP 
      Caption         =   "Playlist options"
      Begin VB.Menu mnuPOpt 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuPS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPEd 
         Caption         =   "&ID3 Editor..."
      End
      Begin VB.Menu mnuPSe 
         Caption         =   "&Text search..."
      End
      Begin VB.Menu mnuPS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPEx 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuPP 
      Caption         =   "Playlist popup"
      Begin VB.Menu mnuPPFilep 
         Caption         =   "&File"
         Begin VB.Menu mnuPPDir 
            Caption         =   "Add &dir..."
         End
         Begin VB.Menu mnuPPFile 
            Caption         =   "&Add file..."
         End
      End
      Begin VB.Menu mnuPPList 
         Caption         =   "&Playlist"
         Begin VB.Menu mnuPPLo 
            Caption         =   "&Load playlist"
         End
         Begin VB.Menu mnuPPSa 
            Caption         =   "&Save playlist"
         End
      End
      Begin VB.Menu mnuPPS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPPPre 
         Caption         =   "Options..."
         Begin VB.Menu mnuPPPr 
            Caption         =   "P&references..."
         End
         Begin VB.Menu mnuPPSk 
            Caption         =   "S&kin browser..."
         End
         Begin VB.Menu mnuPPS3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPPPro 
            Caption         =   "&Program..."
         End
         Begin VB.Menu mnuPPSe 
            Caption         =   "&Text search..."
         End
      End
      Begin VB.Menu mnuPPIn 
         Caption         =   "File &info..."
      End
      Begin VB.Menu mnuPPS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPPExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuF 
      Caption         =   "File"
      Begin VB.Menu mnuFDir 
         Caption         =   "Add &dir..."
      End
      Begin VB.Menu mnuFS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFFile 
         Caption         =   "Add &file..."
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "Track"
      Begin VB.Menu mnuTEr 
         Caption         =   "&Remove    Del"
      End
      Begin VB.Menu mnuTS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTPa 
         Caption         =   "Pa&use"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTSt 
         Caption         =   "&Stop          Esc"
      End
      Begin VB.Menu mnuTS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTPr 
         Caption         =   "P&revius"
      End
      Begin VB.Menu mnuTNe 
         Caption         =   "&Next"
      End
      Begin VB.Menu mnuTS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTPl 
         Caption         =   "&Play           Enter"
      End
   End
   Begin VB.Menu mnuL 
      Caption         =   "List"
      Begin VB.Menu mnuLSo 
         Caption         =   "&Sorting"
         Begin VB.Menu mnuLSoA 
            Caption         =   "Sort by &filetitle"
         End
         Begin VB.Menu mnuLSoT 
            Caption         =   "Sort by &time"
         End
         Begin VB.Menu mnuLSoZ 
            Caption         =   "&Reverse"
         End
         Begin VB.Menu mnuLSoS 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLSoR 
            Caption         =   "R&andomize"
         End
      End
      Begin VB.Menu mnuLNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuLS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLLo 
         Caption         =   "Loa&d"
      End
      Begin VB.Menu mnuLSa 
         Caption         =   "Sa&ve"
      End
   End
   Begin VB.Menu mnuMisc 
      Caption         =   "Misc"
      Begin VB.Menu mnuMiscAb 
         Caption         =   "&CoolPlayer..."
      End
      Begin VB.Menu mnuMiscS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMiscC 
         Caption         =   "&Sound card..."
      End
   End
End
Attribute VB_Name = "frmMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuLSoR_Click()

    With frmPl
     Call Lst.randomizelist(.l)
     .Scroller.Value = .l.SelectedItem.Index
    End With

End Sub

Private Sub mnuLSoT_Click()
    Call SortList(frmPl.l, 1)
End Sub

Private Sub mnuMPCars_Click()
    Call Card
End Sub

Private Sub mnuMuteP_Click()

    Dim O As Object
    Set O = mnuMuteP
    If O.Checked Then
     CI.bMute = False
     MP.Mute = False
     O.Checked = False
    Else
     CI.bMute = True
     MP.Mute = True
     O.Checked = True
    End If
    Call SaveIniSettings(False)

End Sub
Private Sub mnuOOpt_Click()

    On Error GoTo OError
    Call LoadfrmSet

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuOID3_Click()

    On Error GoTo EError
    Call LoadfrmID3(False)

EError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuOAb_Click()

    On Error GoTo AError
    Call LoadfrmAb

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuOEx_Click()
    Call ProgramExit
End Sub

Private Sub mnuPpDir_Click()
    Call OpenForFolder(frmPl)
End Sub

Private Sub mnuPPExit_Click()
    Call ProgramExit
End Sub

Private Sub mnuPpfile_Click()

    On Error GoTo MError
    Call OpenForFile(frmPl)

MError:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub mnuPPPro_Click()
    Call LoadfrmPro
End Sub

Private Sub mnuPpSk_Click()

    On Error GoTo OError
    Call LoadfrmSkn

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnumTrayAb_Click()

    On Error GoTo AError
    Call LoadfrmAb

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnumTrayEx_Click()
    Call ProgramExit
End Sub
Private Sub mnumTrayOpt_Click()

    On Error GoTo OError
    Call LoadfrmSet

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnumTrayRes_Click()
      
    On Error GoTo RError
    Call HideForms(True)

RError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnumiscAb_Click()

    On Error GoTo AError
    Call LoadfrmAb

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuFDir_Click()
    Call OpenForFolder(frmPl)
End Sub
Private Sub mnuFFile_Click()

    On Error GoTo MError
    Call OpenForFile(frmPl)

MError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnumiscC_Click()
    Call Card
End Sub
Private Sub mnupEd_Click()

    On Error GoTo EError
    Call LoadfrmID3(False)

EError:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub mnupEx_Click()
    Call ProgramExit
End Sub
Private Sub mnupOpt_Click()

    On Error GoTo OError
    Call LoadfrmSet

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub mnupSe_Click()
    Call LoadfrmTxt
End Sub
Private Sub mnulLo_Click()

    On Error GoTo OError
    Call OpenForLoad(frmPl)

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnulNew_Click()

    frmPl.l.ListItems.Clear
    Call SetMax
    Call Lst.saveM3U(App.Path & Def, frmPl.l)

End Sub
Private Sub mnulSa_Click()

    On Error GoTo SError
    Call OpenForSave(frmPl)

SError:
   If Err <> 0 Then Exit Sub

End Sub
Private Sub mnulSoA_Click()
    Call SortList(frmPl.l, 0)
End Sub
Private Sub mnulSoZ_Click()
    Call ReverseList(frmPl.l)
End Sub
Private Sub mnuPpIn_Click()

    On Error GoTo ID3Error
    Call LoadfrmID3(True)

ID3Error:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuPpLo_Click()

    On Error GoTo OError
    Call OpenForLoad(frmPl)

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuPpPr_Click()

    On Error GoTo OError
    Call LoadfrmSet

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuPpSa_Click()

    On Error GoTo SError
    Call OpenForSave(frmPl)

SError:
   If Err <> 0 Then Exit Sub

End Sub
Private Sub mnuPpSe_Click()
    Call LoadfrmTxt
End Sub
Private Sub mnuTEr_Click()
    Call RemoveItem
End Sub

Private Sub mnuTNe_Click()
    Call NextP
End Sub
Private Sub mnuTPa_Click()
    Call Pause
End Sub
Private Sub mnuTPl_Click()
    Call GetPlay(True)
End Sub
Private Sub mnuTPr_Click()
    Call PrevP
End Sub
Private Sub mnuTSt_Click()
    Call StopPlay
End Sub
Private Sub mnuPmAb_Click()

    On Error GoTo AError
    Call LoadfrmAb

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuPmdir_Click()
    Call OpenForFolder(frmMn)
End Sub
Private Sub mnuPmfile_Click()

    On Error GoTo MError
    Call OpenForFile(frmMn)

MError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuPmex_Click()
    Call ProgramExit
End Sub
Private Sub mnuPmlo_Click()

    On Error GoTo OError
    Call OpenForLoad(frmMn)

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuPmOp_Click()

    On Error GoTo OError
    Call LoadfrmSet

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnuPmsa_Click()

    On Error GoTo SError
    Call OpenForSave(frmMn)

SError:
   If Err <> 0 Then Exit Sub

End Sub
Private Sub mnuPmSk_Click()

    On Error GoTo OError
    Call LoadfrmSkn

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
