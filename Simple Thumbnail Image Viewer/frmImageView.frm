VERSION 5.00
Begin VB.Form frmImageView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image View"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   Icon            =   "frmImageView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   213
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox filImages 
      Height          =   480
      Left            =   0
      Pattern         =   "*.gif"
      TabIndex        =   7
      Top             =   2430
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   1020
      TabIndex        =   6
      Top             =   3270
      Width           =   825
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   3270
      Width           =   825
   End
   Begin VB.ComboBox cmbImagesType 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2910
      Width           =   1935
   End
   Begin VB.DirListBox dirImages 
      Height          =   2565
      Left            =   0
      TabIndex        =   3
      Top             =   330
      Width           =   1935
   End
   Begin VB.DriveListBox drvImages 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.VScrollBar vscImages 
      Height          =   3195
      LargeChange     =   10
      Left            =   2910
      SmallChange     =   2
      TabIndex        =   1
      Top             =   420
      Width           =   255
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   1950
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   0
      Top             =   390
      Width           =   1245
      Begin VB.PictureBox picImages 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3105
         Left            =   90
         ScaleHeight     =   205
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   9
         Top             =   45
         Width           =   750
      End
   End
   Begin VB.Label lblImagesNumber 
      Alignment       =   2  'Center
      Caption         =   "0 Images"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2010
      TabIndex        =   8
      Top             =   90
      Width           =   1125
   End
End
Attribute VB_Name = "frmImageView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_RESTORE As Long = &H9&
Private strLastToolTip As String

Private Sub ShowImages()
  Dim lngcounter As Integer
  Dim strImageName As String
  Dim strTemp As String
  Dim picTemp As IPictureDisp
  
  lblImagesNumber.Caption = CStr(filImages.ListCount) + " Images"
  picImages.Cls
  picImages.Height = 0
  picImages.Top = 5
  vscImages.Enabled = False
  If (filImages.ListCount = 0) Then Exit Sub
  strTemp = ""
  Screen.MousePointer = 11
  cmdAbout.Enabled = False
  cmdExit.Enabled = False
  For lngcounter = 0 To filImages.ListCount - 1
    strImageName = dirImages.Path + "\" + filImages.List(lngcounter)
    Set picTemp = LoadPicture(strImageName)
    picImages.Height = picImages.Height + 58
    picImages.Line (0, lngcounter * 58 - 1)-(49, lngcounter * 58 + 48), vbBlack, B
    Call picImages.PaintPicture(picTemp, 1, lngcounter * 58, 48, 48)
    Caption = "Image View " + CStr(lngcounter + 1) + " / " + CStr(filImages.ListCount)
    DoEvents
  Next lngcounter
  If (picView.Height < picImages.Height) Then
    vscImages.Max = 1000
    vscImages.Enabled = True
  End If
  cmdExit.Enabled = True
  cmdAbout.Enabled = True
  Screen.MousePointer = 0
  Caption = "Image View"
  vscImages.Value = vscImages.Min
End Sub

Private Sub cmbImagesType_Click()
  filImages.Pattern = cmbImagesType.Text
  filImages.Refresh
End Sub

Private Sub cmdAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub dirImages_Change()
  filImages.Path = dirImages.Path
End Sub

Private Sub drvImages_Change()
  dirImages.Path = drvImages.Drive
End Sub

Private Sub filImages_PathChange()
  Call ShowImages
End Sub

Private Sub filImages_PatternChange()
  Call ShowImages
End Sub

Private Sub Form_Load()
  drvImages.Drive = App.Path
  cmbImagesType.AddItem "*.bmp"
  cmbImagesType.AddItem "*.gif"
  cmbImagesType.AddItem "*.jpg"
  cmbImagesType.ListIndex = 1
  picImages.BorderStyle = 0
  Call ShowImages
End Sub

Private Sub picImages_Click()
  Dim lngcounter As Integer
  Dim strFileName As String
  
  strFileName = ""
  If (picImages.ToolTipText <> "") Then strFileName = dirImages.Path + "\" + picImages.ToolTipText
  If (strFileName <> "") Then Call ShellExecute(Me.hWnd, "Open", strFileName, &H0&, &H0&, SW_RESTORE)
End Sub

Private Sub picImages_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lngcounter As Integer
  Dim lngFileIndex As Integer
  Dim strToolTipText As String
  
  lngFileIndex = -1
  If (filImages.ListCount = 0) Then Exit Sub
  For lngcounter = 0 To filImages.ListCount - 1
    If ((Y > lngcounter * 58) And (Y < lngcounter * 58 + 48)) Then
      lngFileIndex = lngcounter
    End If
  Next lngcounter
  If (lngFileIndex <> -1) Then
    strToolTipText = filImages.List(lngFileIndex)
  Else
    strToolTipText = ""
  End If
  If (strLastToolTip <> strToolTipText) Then
    picImages.ToolTipText = strToolTipText
    strLastToolTip = strToolTipText
  End If
End Sub

Private Sub vscImages_Change()
  Dim lngTop As Long
  
  lngTop = (((vscImages.Value) / vscImages.Max) * (picView.Height - picImages.Height)) + 5
  picImages.Top = lngTop
End Sub
