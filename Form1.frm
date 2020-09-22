VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "YouTube Video Downloader 1.1"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "FLV Player"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      TabIndex        =   12
      Top             =   3960
      Width           =   7485
      Begin VB.CommandButton Command3 
         Caption         =   "&Download"
         Height          =   330
         Left            =   6120
         TabIndex        =   5
         ToolTipText     =   "Click to download the FLV Player"
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Videos downloaded are in FLV format thus a FLV Player is needed to watch videos."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   315
         Width           =   5880
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Download Video"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   45
      TabIndex        =   7
      Top             =   2160
      Width           =   7485
      Begin VB.CommandButton Command2 
         Caption         =   "&Get Video"
         Default         =   -1  'True
         Height          =   330
         Left            =   6120
         TabIndex        =   2
         ToolTipText     =   "Click to download the video from youtube.com"
         Top             =   585
         Width           =   1230
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   585
         Width           =   5910
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note : Files need to include .flv extension."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   1395
         Width           =   2955
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample : http://www.youtube.com/watch?v=urtWHqN3O78"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   1035
         Width           =   4290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the URL of the target video in youtube.com"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Width           =   3480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "YouTube Video Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   45
      TabIndex        =   6
      Top             =   810
      Width           =   7485
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   330
         Left            =   6075
         TabIndex        =   4
         ToolTipText     =   "Click to make a search in youtube.com"
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   585
         Width           =   5820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the keywords for a search in youtube.com"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   315
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample : Clappers Ramady"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   990
         Width           =   1905
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      Picture         =   "Form1.frx":628A
      ScaleHeight     =   720
      ScaleWidth      =   7545
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Visit http://youtube.com"
      Top             =   0
      Width           =   7575
      Begin VB.Label Label0 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coded by Ramci"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5895
         TabIndex        =   14
         Top             =   450
         Width           =   1545
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    If Len(Text1.Text) = 0 Then Exit Sub
    Call OpenURL("http://www.youtube.com/results?search_type=search_videos&search_query=" + Replace(Text1.Text, " ", "+"), True)

End Sub

Private Sub Command2_Click()

    Dim CurrentURL$

    If InStr(1, LCase(Text2.Text), "http://www.youtube.com/watch?v=") = 0 Then Exit Sub
    CurrentURL = VideoDownloadURL(Text2.Text)
    Call OpenURL(CurrentURL, False)

End Sub

Private Sub Command3_Click()

    Call OpenURL("http://wcarchive.cdrom.com/pub/simtelnet/win95/mmedmisc/flvplayer_setup.exe", True)

End Sub

Private Sub Label2_Click()

    Text1.Text = Mid(Label2.Caption, Len("Sample : ") + 1)

End Sub

Private Sub Label4_Click()

    Text2.Text = Mid(Label4.Caption, Len("Sample : ") + 1)

End Sub

Private Sub Picture1_Click()

    Call OpenURL("http://www.youtube.com", True)

End Sub

Private Sub Text1_GotFocus()

    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text2_GotFocus()

    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)

End Sub
