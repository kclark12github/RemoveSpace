VERSION 5.00
Begin VB.Form frmChooseFolder 
   Caption         =   "Choose Folder"
   ClientHeight    =   4092
   ClientLeft      =   72
   ClientTop       =   360
   ClientWidth     =   5616
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4092
   ScaleWidth      =   5616
   StartUpPosition =   1  'CenterOwner
   Begin VB.DriveListBox drvFolders 
      Height          =   288
      Left            =   120
      TabIndex        =   7
      Top             =   3660
      Width           =   4032
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   4260
      TabIndex        =   6
      Top             =   540
      Width           =   1272
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   4260
      TabIndex        =   5
      Top             =   60
      Width           =   1272
   End
   Begin VB.DirListBox dirFolders 
      Height          =   1800
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   4032
   End
   Begin VB.TextBox txtPath 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   3972
   End
   Begin VB.Label lblDrives 
      AutoSize        =   -1  'True
      Caption         =   "&Drives:"
      Height          =   192
      Left            =   120
      TabIndex        =   8
      Top             =   3420
      Width           =   504
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Folders:"
      Height          =   192
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   588
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Path:"
      Height          =   192
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   360
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Please select a folder:"
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1596
   End
End
Attribute VB_Name = "frmChooseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Me.Hide
   txtPath.Text = ""
End Sub

Private Sub cmdOK_Click()
   Me.Hide
End Sub

Private Sub dirFolders_Change()
   txtPath.Text = dirFolders.Path
End Sub

Private Sub drvFolders_Change()
   dirFolders.Path = UCase(drvFolders.Drive)
End Sub

Private Sub Form_Load()
   dirFolders_Change
End Sub

Private Sub txtPath_Change()
   Dim Drive As String
   Dim Directory As String
   
   If txtPath.Text = "" Then Exit Sub
   
   Drive = Left(txtPath.Text, InStr(txtPath.Text, ":"))
   drvFolders.Drive = Drive
   dirFolders.Path = txtPath.Text
End Sub
