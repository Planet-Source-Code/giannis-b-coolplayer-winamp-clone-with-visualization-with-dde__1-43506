VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type WaveFormat
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type Wave
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Private PCM As Wave
Private Const WAVE_FORMAT_1M08 = &H1&
Private Const WAVE_FORMAT_PCM = 1
Private Const Flags = &H1 Or &H2

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Sub AboutPlugin(Top As Boolean)

    On Error GoTo AError
    With frmAb
     Load frmAb
     Call OnTop(.hWnd, Top)
     .Caption = "About visOut v" & App.Major & _
     "." & App.Minor & "." & App.Revision
     .lblDate.Caption = "Date: " & Date
     .Show vbModal
    End With

AError:
    If Err.Number <> 0 Then Unload frmAb

End Sub
Public Property Get Version() As String
    Version = "visOut v" & App.Major & "." & App.Minor & "." & App.Revision
End Property
Private Sub OnTop(Frm As Long, Top As Boolean)
    
    On Error GoTo BError
    Dim i As Integer
    i = IIf(Top, -1, -2)
    Call SetWindowPos(Frm, i, 0, 0, 0, 0, Flags)

BError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub GetData(PictureToDisplay As Object, VisType As Integer)

    On Error GoTo GError
    Call waveInPrepareHeader(Device, VarPtr(PCM), Len(PCM))
    Call waveInAddBuffer(Device, VarPtr(PCM), Len(PCM))
    Call waveInUnprepareHeader(Device, VarPtr(PCM), Len(PCM))
    Call Visualize(PictureToDisplay, VisType)

GError:
    If Err.Number <> 0 Then Call StopDevice

End Sub
Public Sub StartDevice()

    On Error GoTo SError
    Dim WAV As WaveFormat

    Call StopDevice
    With WAV
     .FormatTag = WAVE_FORMAT_1M08
     .Channels = 2
     .SamplesPerSec = 44100
     .BitsPerSample = 16
     .BlockAlign = (.Channels * .BitsPerSample) \ 8
     .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
     .ExtraDataSize = 0
    End With
    Call waveInOpen(Device, 0, VarPtr(WAV), 0, 0, 0)
    Call waveInStart(Device)
    With PCM
     .lpData = VarPtr(sOut(0))
     .dwBufferLength = 1024
     .dwFlags = 0
    End With

SError:
    If Err.Number <> 0 Then Call StopDevice

End Sub
Public Sub StopDevice()

    On Error Resume Next
    Call waveInReset(Device)
    Call waveInStop(Device)
    Call waveInClose(Device): Device = 0

End Sub
Private Sub Class_Initialize()
    Call StartDevice
End Sub
Private Sub Class_Terminate()
    Call StopDevice
End Sub
