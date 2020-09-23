VERSION 5.00
Begin VB.Form frmAb 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About visOut.dll"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Caption         =   "visOut.dll by John"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         ToolTipText     =   "Close the dialog"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblAb 
         Caption         =   "Visualization plugin for CoolPlayer. It streams data from the soundcard, using the waveInStart API. Enjoy it..."
         Height          =   495
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "About"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblDate 
         Caption         =   "Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "The date..."
         Top             =   840
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload frmAb
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload frmAb
End Sub
