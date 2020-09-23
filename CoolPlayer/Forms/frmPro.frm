VERSION 5.00
Begin VB.Form frmPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Caption         =   "Program list"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton cmdSet 
         Caption         =   "&Set tracks"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         ToolTipText     =   "Just save the changes."
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "Select &none"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         ToolTipText     =   "Select none."
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Select &all"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Select all list."
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         ToolTipText     =   "Close the dialog and save the changes."
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear tracks"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         ToolTipText     =   "Clear all data."
         Top             =   720
         Width           =   1215
      End
      Begin VB.ListBox lstP 
         Height          =   3375
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblSl 
         Height          =   855
         Left            =   3000
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblS 
         Caption         =   "Total selected:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Done As Boolean
Private Sub cmdAll_Click()
    Call SelList(True, lstP): Done = True
End Sub
Private Sub cmdClear_Click()

    Call SelList(False, lstP)
    ReDim PL.Tr(0 To 0): PL.Ct = 0
    Call SelectTracks

End Sub
Private Sub cmdClose_Click()

    If Done = False Then Call SelectTracks
    Call DisableForms(True)
    Unload frmPro

End Sub
Private Sub cmdNone_Click()
    Call SelList(False, lstP): Done = True
End Sub
Private Sub cmdSet_Click()
    Call SelectTracks: Done = False
End Sub
Private Sub Form_Unload(Cancel As Integer)

    If Done = False Then Call SelectTracks
    Call DisableForms(True)

End Sub
