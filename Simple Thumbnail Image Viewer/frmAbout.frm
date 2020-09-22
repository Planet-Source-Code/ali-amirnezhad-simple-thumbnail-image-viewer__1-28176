VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About ""Simple Thumbnail Image Viewer"""
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   4
      Top             =   660
      Width           =   945
   End
   Begin VB.Label lblWriter 
      Alignment       =   2  'Center
      Caption         =   "Written by: Ali Amirnezhad (amirnezhad@yahoo.com)"
      Height          =   465
      Left            =   2850
      TabIndex        =   3
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version: 1.02.0001"
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
      Left            =   450
      TabIndex        =   2
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label lblNameUp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Simple Thumbnail Image Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   705
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblNameDown 
      Alignment       =   2  'Center
      Caption         =   "Simple Thumbnail Image Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  lblVersion.Caption = "Version: " + CStr(App.Major) + "." + Format(App.Minor, "0#") + "." + Format(App.Revision, "0###")
  Me.Icon = frmImageView.Icon
End Sub

