VERSION 5.00
Begin VB.Form frmConfirmFileReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm File Replace"
   ClientHeight    =   4044
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   7344
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmConfirmFileReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   7344
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   5880
      TabIndex        =   14
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton cmdNoToAll 
      Caption         =   "N&o to All"
      Height          =   372
      Left            =   4440
      TabIndex        =   13
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton cmdYesToAll 
      Caption         =   "Y&es to All"
      Height          =   372
      Left            =   1560
      TabIndex        =   12
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   372
      Left            =   3000
      TabIndex        =   11
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   1332
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   1320
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   2280
      Width           =   540
   End
   Begin VB.PictureBox picTarget 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   1320
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   1320
      Width           =   540
   End
   Begin VB.PictureBox picReplaceIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   240
      Picture         =   "frmConfirmFileReplace.frx":000C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   240
      Width           =   384
   End
   Begin VB.Label lblSourceDate 
      AutoSize        =   -1  'True
      Caption         =   "modified on <long-date>"
      Height          =   192
      Left            =   2040
      TabIndex        =   9
      Top             =   2640
      Width           =   1728
   End
   Begin VB.Label lblSourceSize 
      AutoSize        =   -1  'True
      Caption         =   "<file-size>KB"
      Height          =   192
      Left            =   2040
      TabIndex        =   8
      Top             =   2400
      Width           =   912
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "with this one?"
      Height          =   192
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   936
   End
   Begin VB.Label lblTargetDate 
      AutoSize        =   -1  'True
      Caption         =   "modified on <long-date>"
      Height          =   192
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   1728
   End
   Begin VB.Label lblTargetSize 
      AutoSize        =   -1  'True
      Caption         =   "<file-size>KB"
      Height          =   192
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   912
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Would you like to replace the existing file"
      Height          =   192
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   2880
   End
   Begin VB.Label lblFileName 
      Caption         =   "This folder already contains a file called '<filename>'."
      Height          =   672
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   6156
   End
End
Attribute VB_Name = "frmConfirmFileReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gfCancel As Boolean
Public gfNo As Boolean
Public gfNoToAll As Boolean
Public gfYes As Boolean
Public gfYesToAll As Boolean
Public Sub InitFlags()
   gfCancel = False
   gfNo = False
   gfNoToAll = False
   gfYes = False
   gfYesToAll = False
End Sub
Private Sub cmdCancel_Click()
   gfCancel = True
   Me.Hide
End Sub

Private Sub cmdNo_Click()
   gfNo = True
   Me.Hide
End Sub

Private Sub cmdNoToAll_Click()
   gfNoToAll = True
   Me.Hide
End Sub

Private Sub cmdYes_Click()
   gfYes = True
   Me.Hide
End Sub

Private Sub cmdYesToAll_Click()
   gfYesToAll = True
   Me.Hide
End Sub

Private Sub Form_GotFocus()
   InitFlags
End Sub

Private Sub Form_Load()
   InitFlags
End Sub
