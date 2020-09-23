VERSION 5.00
Begin VB.Form frmDDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DDE form"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   LinkMode        =   1  'Source
   LinkTopic       =   "frmDDE"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   2295
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "frmDDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    Call AddCommands(CmdStr)
End Sub
