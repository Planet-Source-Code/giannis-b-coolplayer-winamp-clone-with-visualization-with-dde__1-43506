VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVol"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000

Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (&H1000& + 8)
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_CONTROLDETAILSF_VALUE = &H0&

Private Type Target
    dwType As Long
    dwDeviceID As Long
    wMid As Integer
    wPid As Integer
    vVersion As Long
    szPname As String * 32
End Type

Private Type MIXERLINE
    cbStruct As Long
    dwDestination As Long
    dwSource As Long
    dwLineID As Long
    fdwLine As Long
    dwUser As Long
    dwComponentType As Long
    cChannels As Long
    cConnections As Long
    cControls As Long
    szShortName As String * 16
    szName As String * 64
    tTarget As Target
End Type

Private Type MIXERCONTROL
    cbStruct        As Long
    dwControlID     As Long
    dwControlType   As Long
    fdwControl      As Long
    cMultipleItems  As Long
    szShortName(1 To 16) As Byte
    szName(1 To 64) As Byte
    'Bounds(1 To 6)  As Long
    'Metrics(1 To 6) As Long
    lMin As Long
    lMax As Long
    Reserved(10) As Long
End Type

Private Type MIXERLINECONTROLS
    cbStruct As Long
    dwLineID As Long
    dwControl As Long
    cControls As Long
    cbmxctrl As Long
   'pamxctrl As MIXERCONTROL
    pamxctrl As Long
End Type

Private Type MIXERCONTROLDETAILS
    cbStruct As Long
    dwControlID As Long
    cChannels As Long
    item As Long
    cbDet As Long
    paDet As Long
End Type

Private Type MIXERCONTROLDETAILS_UNSIGNED
    Value As Long
End Type

Private Declare Function mixerGetNumDevs Lib "winmm" () As Long
Private Declare Function mixerOpen Lib "winmm" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function MixerClose Lib "winmm" Alias "mixerClose" (ByVal hmx As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal HMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal HMem As Long) As Long

Private VOL As MIXERCONTROLDETAILS_UNSIGNED
Private mVOL(1) As MIXERCONTROLDETAILS_UNSIGNED
Private MCon(20) As MIXERCONTROL
Private MXCD As MIXERCONTROLDETAILS
Private MXLC As MIXERLINECONTROLS
Private MXL As MIXERLINE

Private MHandle As Long, HMem As Long
Private Ret As Long, Stereo As Boolean

Public Enum VOL_CONTROL
    SPEAKER = 0
    LINEIN = 1
    MICROPHONE = 2
    SYNTHESIZER = 3
    COMPACTDISC = 4
    WAVEOUT = 5
    AUXILIARY = 6
End Enum

Public Enum MUTE_CONTROL
    mSPEAKER = 7
    mLINEIN = 8
    mMICROPHONE = 9
    mSYNTHESIZER = 10
    mCOMPACTDISC = 11
    mWAVEOUT = 12
    mAUXILIARY = 13
End Enum
Private Function MixerStart(ByVal Number As Long) As Boolean

    On Error GoTo OError
    If Number < 0 Or Number > mixerGetNumDevs - 1 Then Exit Function
    Ret = mixerOpen(MHandle, Number, 0, 0, 0)
    If Ret <> 0 Then
     MixerStart = False: Exit Function
    End If
    Ret = MixerGet(MHandle, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_VOLUME, MCon(WAVEOUT))
    Ret = MixerGet(MHandle, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_MUTE, MCon(mWAVEOUT))
    Ret = GetVolume(WAVEOUT)
    Ret = GetMute(mWAVEOUT)
    MixerStart = True

OError:
    If Err.Number <> 0 Then MixerStart = False: Exit Function

End Function
Private Function MixerExit() As Long

    MixerExit = MixerClose(MHandle)
    MHandle = 0

End Function
Public Function GetMute(Cont As MUTE_CONTROL) As Boolean

    On Error GoTo GError
    With MXCD
     .item = 0
     .cbStruct = Len(MXCD)
     .dwControlID = MCon(Cont).dwControlID
     .cChannels = 1
     .cbDet = Len(VOL)

     HMem = GlobalAlloc(&H40, Len(VOL))
     .paDet = GlobalLock(HMem)
    End With

    Ret = mixerGetControlDetails(MHandle, MXCD, MIXER_CONTROLDETAILSF_VALUE)
    Call CopyStructFromPtr(VOL, MXCD.paDet, Len(VOL))
    Call GlobalFree(HMem)
    GetMute = CBool(VOL.Value)

GError:
    If Err.Number <> 0 Then GetMute = False: Exit Function

End Function
Public Function SetVolume(ByVal Balance As Long, ByVal mValue As Long, Cont As VOL_CONTROL) As Boolean

    On Error GoTo SError
    If Stereo Then
     SetVolume = SetBalance(Balance, mValue, Cont)
    Else
     With MXCD
      .item = MCon(Cont).cMultipleItems
      .dwControlID = MCon(Cont).dwControlID
      .cbStruct = Len(MXCD)
      .cbDet = Len(VOL)
      .cChannels = 1

      HMem = GlobalAlloc(&H40, Len(VOL))
      .paDet = GlobalLock(HMem)
     End With

     With VOL
      If MCon(Cont).lMax > 100 Then
       .Value = mValue * (MCon(Cont).lMax \ 100)
      Else
       .Value = mValue
      End If
      If .Value > MCon(Cont).lMax Then .Value = MCon(Cont).lMax
      If .Value < MCon(Cont).lMin Then .Value = MCon(Cont).lMin
     End With

     Call CopyPtrFromStruct(MXCD.paDet, VOL, Len(VOL) * MXCD.cChannels)
     Ret = mixerSetControlDetails(MHandle, MXCD, MIXER_CONTROLDETAILSF_VALUE)
     Call GlobalFree(HMem)
     SetVolume = IIf(Ret = 0, True, False)
    End If

SError:
    If Err.Number <> 0 Then SetVolume = False: Exit Function

End Function
Public Function SetBalance(ByVal mBalVal As Long, ByVal mVolVal As Long, Cont As VOL_CONTROL) As Boolean

    On Error GoTo SError
    Dim volL As Long, volR As Long, tmpVol As Long

    tmpVol = MCon(Cont).lMax - ((mVolVal * MCon(Cont).lMax) / 100)
    tmpVol = IIf(tmpVol = 0, 1, tmpVol)
    volR = tmpVol * (IIf(mBalVal >= 0, 1, (100 + mBalVal) / 100))
    volL = tmpVol * (IIf(mBalVal <= 0, 1, (100 - mBalVal) / 100))

    With MXCD
     .item = MCon(Cont).cMultipleItems
     .dwControlID = MCon(Cont).dwControlID
     .cbStruct = Len(MXCD)
     .cbDet = Len(mVOL(0))
     .cChannels = 2

     HMem = GlobalAlloc(&H40, Len(mVOL(0)))
     .paDet = GlobalLock(HMem)
    End With

    mVOL(1).Value = volR: mVOL(0).Value = volL
    Call CopyPtrFromStruct(MXCD.paDet, mVOL(1).Value, Len(mVOL(0)) * MXCD.cChannels)
    Call CopyPtrFromStruct(MXCD.paDet, mVOL(0).Value, Len(mVOL(1)) * MXCD.cChannels)
    Ret = mixerSetControlDetails(MHandle, MXCD, MIXER_CONTROLDETAILSF_VALUE)
    Call GlobalFree(HMem)
    SetBalance = IIf(Ret = 0, True, False)

SError:
    If Err.Number <> 0 Then SetBalance = False: Exit Function

End Function
Public Function SetMute(ByVal MuteState As Boolean, Cont As MUTE_CONTROL) As Boolean

    On Error GoTo SError
    With MXCD
     .item = 0
     .cbStruct = Len(MXCD)
     .dwControlID = MCon(Cont).dwControlID
     .cChannels = 1
     .cbDet = Len(VOL)

     HMem = GlobalAlloc(&H40, Len(VOL))
     .paDet = GlobalLock(HMem)
    End With

    VOL.Value = MuteState
    Call CopyPtrFromStruct(MXCD.paDet, VOL, Len(VOL))
    Ret = mixerSetControlDetails(MHandle, MXCD, MIXER_CONTROLDETAILSF_VALUE)
    Call GlobalFree(HMem)
    SetMute = IIf(Ret = 0, MuteState, Not MuteState)

SError:
    If Error <> 0 Then SetMute = False: Exit Function

End Function
Public Function GetVolume(Cont As VOL_CONTROL) As Long

    On Error GoTo GError
    With MXCD
     .item = 0
     .dwControlID = MCon(Cont).dwControlID
     .cbStruct = Len(MXCD)
     .cbDet = Len(VOL)
     .cChannels = 1

     HMem = GlobalAlloc(&H40, Len(VOL))
     .paDet = GlobalLock(HMem)
    End With

    Ret = mixerGetControlDetails(MHandle, MXCD, MIXER_CONTROLDETAILSF_VALUE)
    Call CopyStructFromPtr(VOL, MXCD.paDet, Len(VOL))
    Call GlobalFree(HMem)
    If Ret <> 0 Then GetVolume = 0: Exit Function

    If MCon(Cont).lMax > 100 Then
     GetVolume = (VOL.Value * 100) / (MCon(Cont).lMax - MCon(Cont).lMin)
    Else
     GetVolume = VOL.Value
    End If

GError:
    If Err.Number <> 0 Then GetVolume = 0: Exit Function

End Function
Private Function MixerGet(ByVal HMixer As Long, ByVal cType As Long, ByVal ctrlType As Long, MXC As MIXERCONTROL) As Boolean

    On Error GoTo MError
    With MXL
     .cbStruct = Len(MXL)
     .dwComponentType = cType
    End With

    Ret = mixerGetLineInfo(HMixer, MXL, MIXER_GETLINEINFOF_COMPONENTTYPE)
    If Ret = 0 Then
     With MXLC
      .cbStruct = Len(MXLC)
      .dwLineID = MXL.dwLineID
      .dwControl = ctrlType
      .cControls = 1
      .cbmxctrl = Len(MXC)
      HMem = GlobalAlloc(&H40, Len(MXC))
      .pamxctrl = GlobalLock(HMem)
      MXC.cbStruct = Len(MXC)
      Stereo = MXL.cChannels > 1
     End With

     Ret = mixerGetLineControls(HMixer, MXLC, MIXER_GETLINECONTROLSF_ONEBYTYPE)
     If Ret = 0 Then
      MixerGet = True
      Call CopyStructFromPtr(MXC, MXLC.pamxctrl, Len(MXC))
     Else
      MixerGet = False
     End If
     Call GlobalFree(HMem)
    Else
     MixerGet = False
    End If

MError:
    If Err.Number <> 0 Then MixerGet = False: Exit Function

End Function
Public Function GetBalance(Cont As VOL_CONTROL) As Long
   
    On Error GoTo GError
    With MXCD
     .item = MCon(Cont).cMultipleItems
     .dwControlID = MCon(Cont).dwControlID
     .cbStruct = Len(MXCD)
     .cbDet = Len(mVOL(0))
     .cChannels = 2

     HMem = GlobalAlloc(&H40, Len(mVOL(0)))
     .paDet = GlobalLock(HMem)
    End With

    Call CopyStructFromPtr(mVOL(0).Value, MXCD.paDet, Len(mVOL(1)) * MXCD.cChannels)
    Ret = mixerGetControlDetails(MHandle, MXCD, MIXER_CONTROLDETAILSF_VALUE)
    Call GlobalFree(HMem)
    If Ret <> 0 Then GetBalance = 0: Exit Function

    If MCon(Cont).lMax > 100 Then
     GetBalance = ((mVOL(1).Value * 100) - (mVOL(0).Value * 100)) / (MCon(Cont).lMax - MCon(Cont).lMin)
    Else
     GetBalance = mVOL(1).Value - mVOL(0).Value
    End If

GError:
    If Err.Number <> 0 Then GetBalance = 0: Exit Function
   
End Function
Private Sub Class_Initialize()
    Call MixerStart(0)
End Sub
Private Sub Class_Terminate()
    Call MixerExit
End Sub
