VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3Info"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type MPEG
    ModeExt As Integer
    Bitrate As Integer
    Frames As Integer

    Frequency As Long
    FileLength As Long
    Len As Long

    ValidHeader As Boolean
    ProtectionBitSet As Boolean
    Padded As Boolean
    PrivateBitSet As Boolean
    Copyrighted As Boolean
    Original As Boolean

    Emphasis As String
    Mode As String
    Filename As String
    ID As String
    Layer As String
    ProtectionChecksum As String
End Type

Private MP3 As MPEG, Fi As New clsFiles
Dim TMP(1) As Byte, FileNum As Integer, B(4) As Byte
Dim i As Long, SByte As Long, ARG(2) As Integer
Public Property Let Filename(sInput As String)

    Call ClearData
    MP3.Filename = sInput

End Property
Public Property Get Filename() As String
    Filename = MP3.Filename
End Property
Public Function GetTime(Name As String, C As Boolean) As String

    On Error GoTo GError
    With MP3
     FileNum = FreeFile
     .FileLength = FileLen(Name)

     Close #FileNum
     Open Name For Binary As #FileNum
      For i = 1 To LOF(FileNum) - 1
       Get #FileNum, i, TMP(0)
       If TMP(0) = &HFF Then
        Get #FileNum, i + 1, TMP(1)
        If TMP(1) And &HE0 = &HE0 Then
         .ValidHeader = True
         SByte = i + 1: Exit For
        End If
       End If
      Next

      If Not .ValidHeader Then
       GetTime = "0"
       Close #FileNum: Exit Function
      End If
    
      Get #FileNum, SByte, B(0)
      Get #FileNum, SByte + 1, B(1)
     Close #FileNum
     ARG(0) = CInt(B(0) And &H8) / 8
     ARG(1) = CInt(B(0) And &H6) / 2
     ARG(2) = CInt(B(1) And &HF0) / 16

     i = Int((Int(.FileLength * 0.8)) / (Int(GetBitrate(ARG(0), ARG(1), ARG(2)) * 100)))
     GetTime = IIf(C = True, Fi.GetTime(CStr(i)), CStr(i))
    End With

GError:
    If Err.Number <> 0 Then
     GetTime = "0": Close #FileNum
     MP3.ValidHeader = False
    End If

End Function

Public Property Get ValidHeader() As Boolean
    ValidHeader = MP3.ValidHeader
End Property
Public Property Get ID() As String
    ID = MP3.ID & " " & MP3.Layer
End Property
Public Property Get ProtectionBitSet() As Boolean
    ProtectionBitSet = MP3.ProtectionBitSet
End Property
Public Property Get ProtectionChecksum() As String
    ProtectionChecksum = "Protection: " & MP3.ProtectionChecksum
End Property
Public Property Get Bitrate() As String
    Bitrate = MP3.Bitrate
End Property
Public Property Get Length() As String
    Length = "Length: " & Int((Int(MP3.FileLength * 0.8)) / (Int(MP3.Bitrate * 100))) & " Seconds "
End Property
Public Property Get Frames() As String
    Frames = "Frames: " & MP3.Frames
End Property
Public Property Get Frequency() As String
    Frequency = MP3.Frequency
End Property
Public Property Get Padded() As String
    Padded = "Padded: " & MP3.Padded
End Property
Public Property Get PrivateBitSet() As Boolean
    PrivateBitSet = MP3.PrivateBitSet
End Property
Public Property Get PrivateBit() As String
    PrivateBit = "Private: " & MP3.PrivateBitSet
End Property
Public Property Get Mode() As String
    Mode = MP3.Mode
End Property
Public Property Get ModeExt() As String
    ModeExt = "Mode ext: " & MP3.ModeExt
End Property
Public Property Get Copyrighted() As String
    Copyrighted = "Copyrighted: " & MP3.Copyrighted
End Property
Public Property Get Original() As String
    Original = "Original: " & MP3.Original
End Property
Public Property Get Emphasis() As String
    Emphasis = MP3.Emphasis
End Property
Public Property Get Filesize() As String
    Filesize = "File size: " & MP3.FileLength & " bytes"
End Property
Public Sub GetHeader()

    On Error GoTo GError
    With MP3
     FileNum = FreeFile
     .FileLength = FileLen(.Filename)

     Close #FileNum
     Open .Filename For Binary As #FileNum
      For i = 1 To LOF(FileNum) - 1
       Get #FileNum, i, TMP(0)
       If TMP(0) = &HFF Then
        Get #FileNum, i + 1, TMP(1)
        If TMP(1) And &HE0 = &HE0 Then
         .ValidHeader = True
         SByte = i + 1: Exit For
        End If
       End If
      Next
      If Not .ValidHeader Then
       .ValidHeader = False: Close #FileNum: Exit Sub
      End If

      Get #FileNum, SByte, B(0)
      Get #FileNum, SByte + 1, B(1)
      Get #FileNum, SByte + 2, B(2)
      Get #FileNum, SByte + 3, B(3)
      Get #FileNum, SByte + 4, B(4)
     Close #FileNum

     Select Case CInt(B(0) And &H18) / 8
      Case 0: .ID = "MPEG 2.5"
      Case 1: .ID = "MPEG Not Defined"
      Case 2: .ID = "MPEG 2"
      Case 3: .ID = "MPEG 1"
     End Select

     Select Case (B(0) And &H6)
      Case &H0: .Layer = "Layer Not Defined"
      Case &H2: .Layer = "Layer III"
      Case &H4: .Layer = "Layer II"
      Case &H6: .Layer = "Layer I"
     End Select

     If (B(0) And &H1) = &H1 Then
      .ProtectionBitSet = True
     Else
      .ProtectionBitSet = False
      .ProtectionChecksum = Hex(B(3)) & " " & Hex(B(4))
     End If

     ARG(0) = CInt(B(0) And &H8) / 8
     ARG(1) = CInt(B(0) And &H6) / 2
     ARG(2) = CInt(B(1) And &HF0) / 16
     .Bitrate = GetBitrate(ARG(0), ARG(1), ARG(2))

     Select Case CInt(B(0) And &H18) / 8
      Case 0: i = 1
      Case 1: i = 0
      Case 2: i = 2
      Case 3: i = 4
     End Select
    
     Select Case CInt(B(1) And &HC) / 4
      Case 0: .Frequency = i * 11025
      Case 1: .Frequency = i * 12000
      Case 2: .Frequency = i * 8000
     End Select

     .Frames = CInt(.FileLength / ((.Bitrate * 144000) / (.Frequency)))
     If (B(1) And &H2) = &H2 Then .Padded = True
     If (B(1) And &H1) = &H1 Then .PrivateBitSet = True

     Select Case CInt(B(2) And &HC0) / CInt(&H40)
      Case 0: .Mode = "Stereo"
      Case 1: .Mode = "Joint Stereo"
      Case 2: .Mode = "Dual Channel"
      Case 3: .Mode = "Mono"
     End Select

     .ModeExt = CInt(B(2) And &H30) / CInt(&H10)
     If (B(2) And &H8) = &H8 Then .Copyrighted = True
     If (B(2) And &H4) = &H4 Then .Original = True

     Select Case B(2) And &H3
      Case 0: .Emphasis = "Empasis: None"
      Case 2: .Emphasis = "Emphasis: Not defined"
      Case 1: .Emphasis = "Emphasis: 50/15 ms"
      Case 3: .Emphasis = "Emphasis: CCITT J. 17"
     End Select
    End With

GError:
    If Err.Number <> 0 Then
     MP3.ValidHeader = False: Close #FileNum: Exit Sub
    End If

End Sub
Private Function GetBitrate(ARG1 As Integer, ARG2 As Integer, ARG3 As Integer) As Integer

    On Error GoTo BitError
    Dim a(1, 3, 15) As Integer
    
    For i = 1 To 14
     a(1, 3, i) = i * 32
     If i < 5 Then
      a(1, 2, i) = 8 * (i + 4)
      a(1, 1, i) = 8 * (i + 3)
     Else
      a(1, 2, i) = a(1, 2, i - 4) * 2
      a(1, 1, i) = a(1, 1, i - 4) * 2
     End If
     If i < 9 Then
      a(0, 1, i) = i * 8
     Else
      a(0, 1, i) = (i - 4) * 16
     End If
     a(0, 2, i) = a(0, 1, i)
    Next
    
    a(1, 2, 1) = 32: a(0, 3, 1) = 32: a(0, 3, 2) = 48
    a(0, 3, 3) = 56: a(0, 3, 4) = 64: a(0, 3, 5) = 80
    a(0, 3, 6) = 96: a(0, 3, 7) = 112: a(0, 3, 8) = 128
    a(0, 3, 9) = 144: a(0, 3, 10) = 160: a(0, 3, 11) = 176
    a(0, 3, 12) = 192: a(0, 3, 13) = 224: a(0, 3, 14) = 256

    GetBitrate = a(ARG1, ARG2, ARG3)
    If ARG3 = 15 Then GetBitrate = 1
    If ARG3 = 0 Then GetBitrate = 0

BitError:
    If Err.Number <> 0 Then GetBitrate = 0: Exit Function

End Function
Private Sub ClearData()

    With MP3
     .Bitrate = 0: .Copyrighted = False
     .Emphasis = "": .FileLength = 0
     .Filename = "": .Frames = 0
     .Frequency = 0: .ID = ""
     .Layer = "": .Len = 0
     .Mode = "": .ModeExt = 0
     .Original = False: .Padded = False
     .PrivateBitSet = False: .ProtectionBitSet = False
     .ProtectionChecksum = "": .ValidHeader = False
    End With

End Sub
