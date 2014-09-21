VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Remove %20"
   ClientHeight    =   1644
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7044
   LinkTopic       =   "Form1"
   ScaleHeight     =   1644
   ScaleWidth      =   7044
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   432
      Left            =   3036
      TabIndex        =   3
      Top             =   1080
      Width           =   972
   End
   Begin VB.Frame frameSource 
      Caption         =   "Source Directory"
      Height          =   672
      Left            =   186
      TabIndex        =   0
      Top             =   180
      Width           =   6672
      Begin VB.TextBox txtSource 
         Height          =   288
         Left            =   60
         TabIndex        =   2
         Top             =   240
         Width           =   5472
      End
      Begin VB.CommandButton cmdBrowseSource 
         Caption         =   "&Browse"
         Height          =   312
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   912
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBrowseSource_Click()
    cmdSourceBrowseHit = True
    frmChooseFolder.txtPath.Text = txtSource.Text
    frmChooseFolder.Show vbModal
    If frmChooseFolder.txtPath.Text <> "" Then
        txtSource.Text = frmChooseFolder.txtPath.Text
    End If
    cmdSourceBrowseHit = False
End Sub

