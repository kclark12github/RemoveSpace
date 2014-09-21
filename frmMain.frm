VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove %20"
   ClientHeight    =   1836
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   7044
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1836
   ScaleWidth      =   7044
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   432
      Left            =   3036
      TabIndex        =   2
      Top             =   1080
      Width           =   972
   End
   Begin VB.Frame frameSource 
      Caption         =   "Source Directory"
      Height          =   672
      Left            =   186
      TabIndex        =   3
      Top             =   180
      Width           =   6672
      Begin VB.TextBox txtSource 
         Height          =   288
         Left            =   60
         TabIndex        =   0
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
   Begin ComctlLib.StatusBar sbBottom 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   4
      Top             =   1524
      Width           =   7044
      _ExtentX        =   12425
      _ExtentY        =   550
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9821
            TextSave        =   ""
            Key             =   "Status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "Size"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const KB& = 1024
Const MB& = KB * 1024
Const GB& = MB * 1024
Private Sub cmdBrowseSource_Click()
    frmChooseFolder.txtPath.Text = txtSource.Text
    frmChooseFolder.Show vbModal
    If frmChooseFolder.txtPath.Text <> "" Then
        txtSource.Text = frmChooseFolder.txtPath.Text
    End If
End Sub
Public Function DoFiles(strSourcePath As String, ByRef lTotalBytes As Long) As Long
    Dim strDir As String
    Dim strNew As String
    Dim Icon As Integer
    Dim strIcon As String
    Dim lTotalFiles As Integer
   
    If Right(strSourcePath, 1) <> "\" Then strSourcePath = strSourcePath & "\"
   
    strDir = Dir(strSourcePath, vbDirectory)
    Do While strDir <> ""   ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        Select Case Left(strDir, 1)
            Case "."
            Case Else
                'Debug.Print "Processing: " & strSourcePath & strDir & "..."
                If InStr(strDir, "%20") Then
                    strNew = Replace(strDir, "%20", " ")
                    Name strSourcePath & strDir As strSourcePath & strNew
                    lTotalBytes = lTotalBytes + FileLen(strSourcePath & strNew)
                    lTotalFiles = lTotalFiles + 1
                    strDir = strNew
                End If
                
                If (GetAttr(strSourcePath & strDir) And vbDirectory) = vbDirectory Then
                    sbBottom.Panels("Status").Text = "Processing " & strSourcePath & strDir & "..."
                    lTotalFiles = lTotalFiles + DoFiles(strSourcePath & strDir, lTotalBytes)
                    RePosition strSourcePath, strDir
                    DoEvents
                End If
        End Select
        strDir = Dir   ' Get next entry.
    Loop
    DoFiles = lTotalFiles
End Function
Private Sub RePosition(strPath As String, strDirectory As String)
    Dim strDir As String
   
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    strDir = Dir(strPath, vbDirectory)
    Do While strDir <> ""
        If strDir = strDirectory Then Exit Sub
        strDir = Dir
    Loop
End Sub
Private Sub cmdOK_Click()
    Dim lTotalFiles As Long
    Dim TotalBytes As Long
   
    If txtSource.Text = "" Then
        MsgBox "Source Directory must be specified.", vbExclamation + vbOKOnly
        txtSource.SetFocus
        Exit Sub
    End If
    
    On Error GoTo 0
    cmdOK.Enabled = False
    txtSource.Enabled = False
    lTotalFiles = DoFiles(txtSource.Text, TotalBytes)
    sbBottom.Panels("Status").Text = Format(lTotalFiles, "###,##0") & " file(s) found."
        
    If TotalBytes > GB Then
        sbBottom.Panels("Size").Text = Format(TotalBytes / GB, "#,##0.00 GB")
    Else
        If TotalBytes > MB Then
            sbBottom.Panels("Size").Text = Format(TotalBytes / MB, "#,##0.00 MB")
        Else
            If TotalBytes > KB Then
                sbBottom.Panels("Size").Text = Format(TotalBytes / KB, "#,##0.00 KB")
            Else
                If TotalBytes > 0 Then sbBottom.Panels("Size").Text = Format(TotalBytes, "#,##0 Bytes")
            End If
        End If
    End If
    cmdOK.Enabled = True
    txtSource.Enabled = True
End Sub
Private Sub txtSource_Validate(Cancel As Boolean)
    If txtSource.Text = "" Then
        cmdBrowseSource_Click
        Exit Sub
    End If
   
    If Right(txtSource.Text, 1) = "\" Then txtSource.Text = Left(txtSource.Text, Len(txtSource.Text) - 1)
    On Error Resume Next
    If Dir(txtSource.Text, vbDirectory) <> "" Then
        If (GetAttr(txtSource.Text) And vbDirectory) <> vbDirectory Then
            MsgBox "Path specified does not represent a directory.", vbExclamation + vbOKOnly
            Cancel = True
        End If
    Else
        MsgBox "Path specified does not exist.", vbExclamation + vbOKOnly
        Cancel = True
    End If
    On Error GoTo 0
End Sub
