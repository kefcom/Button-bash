VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "game_engine"
   ClientHeight    =   11490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_changegame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   11535
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   15255
      Begin VB.Timer waiter_changegame 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   240
         Top             =   720
      End
      Begin VB.CommandButton btn_capture_changegame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Height          =   135
         Left            =   15120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Width           =   135
      End
      Begin VB.FileListBox File1 
         Height          =   675
         Left            =   840
         Pattern         =   "*.ini"
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer resetadminkeypresses 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   240
         Top             =   240
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "to select a game and wait 5 seconds to load it."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   50
         Top             =   1920
         Width           =   9855
      End
      Begin VB.Label lbl_changeletter 
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4920
         TabIndex        =   49
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image img_changebutton 
         Height          =   615
         Left            =   4680
         Picture         =   "Form1.frx":0442
         Stretch         =   -1  'True
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   11
         Left            =   0
         TabIndex        =   46
         Top             =   10680
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   10
         Left            =   0
         TabIndex        =   45
         Top             =   9960
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   9
         Left            =   0
         TabIndex        =   44
         Top             =   9240
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   8
         Left            =   0
         TabIndex        =   43
         Top             =   8520
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   7
         Left            =   0
         TabIndex        =   42
         Top             =   7800
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   6
         Left            =   0
         TabIndex        =   41
         Top             =   7080
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   5
         Left            =   0
         TabIndex        =   40
         Top             =   6360
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   4
         Left            =   0
         TabIndex        =   39
         Top             =   5640
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   3
         Left            =   0
         TabIndex        =   38
         Top             =   4920
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   2
         Left            =   0
         TabIndex        =   37
         Top             =   4200
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label lbl_filename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Config file name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   36
         Top             =   3480
         Visible         =   0   'False
         Width           =   15255
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Press "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   34
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Change game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   0
         TabIndex        =   33
         Top             =   120
         Width           =   15255
      End
   End
   Begin VB.Frame frm_startgame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "start_newgame"
      Height          =   3375
      Left            =   3960
      TabIndex        =   28
      Top             =   5880
      Width           =   4095
      Begin VB.Image img_adminstart 
         Height          =   1695
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbl_admin_start 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1680
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TO START"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.ListBox average 
      Height          =   1035
      Index           =   3
      Left            =   14520
      TabIndex        =   26
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox average 
      Height          =   1035
      Index           =   2
      Left            =   14520
      TabIndex        =   25
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox average 
      Height          =   1035
      Index           =   1
      Left            =   14520
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox average 
      Height          =   1035
      Index           =   0
      Left            =   14520
      TabIndex        =   23
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer click_sec 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12840
      Top             =   4200
   End
   Begin VB.Frame frm_gameover 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   2760
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Timer switch_gameover_newgame 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   120
         Top             =   480
      End
      Begin VB.Label overall_average 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total average:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   6495
      End
      Begin VB.Label lbl_playerwin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   6495
      End
      Begin VB.Image img_winner 
         Height          =   1215
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Winner:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Game over!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   6495
      End
   End
   Begin VB.Timer game_clock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   10920
   End
   Begin VB.Frame frm_countdown 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   2760
      TabIndex        =   16
      Top             =   6120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Timer countdown_clock 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   120
      End
      Begin VB.Label lbl_countdown 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   6495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Get ready..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   6495
      End
   End
   Begin VB.Frame frm_invisible 
      Caption         =   "invisible objects"
      Height          =   4575
      Left            =   15480
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
      Begin VB.Image img_nothing 
         Height          =   495
         Left            =   120
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Image img_bg 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Image img_pl4_img2 
         Height          =   615
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image img_pl4_img1 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image img_pl3_img2 
         Height          =   615
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image img_pl3_img1 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image img_pl2_img2 
         Height          =   615
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image img_pl2_img1 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image img_pl1_img2 
         Height          =   615
         Left            =   840
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image img_pl1_img1 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1920
      Top             =   2640
      Width           =   375
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp_pl4 
      Height          =   495
      Left            =   1200
      TabIndex        =   56
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   873
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp_pl3 
      Height          =   495
      Left            =   840
      TabIndex        =   55
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   873
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp_pl2 
      Height          =   495
      Left            =   480
      TabIndex        =   54
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   873
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp_pl1 
      Height          =   495
      Left            =   1200
      TabIndex        =   53
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   873
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   495
      Left            =   840
      TabIndex        =   52
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   873
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp_winner 
      Height          =   495
      Left            =   480
      TabIndex        =   51
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   873
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4080
      TabIndex        =   15
      Top             =   10560
      Width           =   11175
   End
   Begin VB.Label gametime 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2880
      TabIndex        =   14
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Game time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   10560
      Width           =   2535
   End
   Begin VB.Label clicks_ps 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 / second"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   3
      Left            =   12360
      TabIndex        =   12
      Top             =   9720
      Width           =   2775
   End
   Begin VB.Label clicks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   3
      Left            =   10800
      TabIndex        =   11
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label player_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   3
      Left            =   10800
      TabIndex        =   10
      Top             =   9120
      Width           =   4335
   End
   Begin VB.Shape infobox 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   3
      Left            =   10680
      Top             =   9000
      Width           =   4575
   End
   Begin VB.Label clicks_ps 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 / second"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   12360
      TabIndex        =   9
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label clicks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   10800
      TabIndex        =   8
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label player_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   10800
      TabIndex        =   7
      Top             =   7680
      Width           =   4335
   End
   Begin VB.Shape infobox 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   2
      Left            =   10680
      Top             =   7560
      Width           =   4575
   End
   Begin VB.Label clicks_ps 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 / second"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   12360
      TabIndex        =   6
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label clicks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   10800
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label player_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   10800
      TabIndex        =   4
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Shape infobox 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   1
      Left            =   10680
      Top             =   6120
      Width           =   4575
   End
   Begin VB.Label clicks_ps 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 / second"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   12360
      TabIndex        =   3
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label clicks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   10800
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label player_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   10800
      TabIndex        =   1
      Top             =   4800
      Width           =   4335
   End
   Begin VB.Shape infobox 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1455
      Index           =   0
      Left            =   10680
      Top             =   4680
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   5775
      Left            =   10560
      Top             =   4680
      Width           =   135
   End
   Begin VB.Shape track 
      BorderColor     =   &H00000000&
      Height          =   1455
      Index           =   3
      Left            =   360
      Top             =   9000
      Width           =   10215
   End
   Begin VB.Image img_pl1 
      Height          =   1215
      Index           =   3
      Left            =   480
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Shape track 
      BorderColor     =   &H00000000&
      Height          =   1455
      Index           =   2
      Left            =   360
      Top             =   7560
      Width           =   10215
   End
   Begin VB.Image img_pl1 
      Height          =   1215
      Index           =   2
      Left            =   480
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Shape track 
      BorderColor     =   &H00000000&
      Height          =   1455
      Index           =   1
      Left            =   360
      Top             =   6120
      Width           =   10215
   End
   Begin VB.Image img_pl1 
      Height          =   1215
      Index           =   1
      Left            =   480
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Image img_pl1 
      Height          =   1215
      Index           =   0
      Left            =   480
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Shape track 
      BorderColor     =   &H00000000&
      Height          =   1455
      Index           =   0
      Left            =   360
      Top             =   4680
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pl1img1 As String
Dim pl1img2 As String
Dim pl2img1 As String
Dim pl2img2 As String
Dim pl3img1 As String
Dim pl3img2 As String
Dim pl4img1 As String
Dim pl4img2 As String
Dim bgimg As String
Dim maxcountdown As Integer
Dim countdown As Integer
Dim game_in_progress As Boolean
Dim key_admin_start As String
Dim key_1_1 As String
Dim key_1_2 As String
Dim key_2_1 As String
Dim key_2_2 As String
Dim key_3_1 As String
Dim key_3_2 As String
Dim key_4_1 As String
Dim key_4_2 As String
Dim lastkey1 As Integer
Dim lastkey2 As Integer
Dim lastkey3 As Integer
Dim lastkey4 As Integer
Dim lastclicks1 As Integer
Dim lastclicks2 As Integer
Dim lastclicks3 As Integer
Dim lastclicks4 As Integer
Dim player1name As String
Dim player2name As String
Dim player3name As String
Dim player4name As String
Dim buttonmode As String
Dim img_startbutton As String
Dim summarytime As Integer
Dim game_timeout As Integer
Dim move_camel As Boolean
Dim adminkeypresses As Integer
Dim selectednewgame As Integer
Dim newconfig As String
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpChar As Integer, ByVal uFlags As Long) As Long

Private Function TranslateKeyCode(ByVal KeyCode As Long, ByVal ShiftDown As Boolean) As String
    Dim pKBdata(255) As Byte
    Dim retChar As Integer
    
    If ShiftDown Then pKBdata(vbKeyShift) = &H80 'Set shift key to down
    ToAscii KeyCode, 0, pKBdata(0), retChar, 0
    
    On Error Resume Next
    TranslateKeyCode = Chr$(retChar)
End Function



Private Sub btn_capture_changegame_KeyUp(KeyCode As Integer, Shift As Integer)
If TranslateKeyCode(KeyCode, False) = key_admin_start Then
    selectednewgame = selectednewgame + 1
    teller = 0
    Do
        If teller <> selectednewgame Then
            lbl_filename(teller).ForeColor = &H80FF&
        Else
            If lbl_filename(teller).Visible = False Then
                selectednewgame = 0
                lbl_filename(0).ForeColor = &HFF00&
            Else
                lbl_filename(teller).ForeColor = &HFF00&
            End If
        End If
        teller = teller + 1
    Loop Until teller = lbl_filename.Count
    waiter_changegame.Enabled = False
    waiter_changegame.Enabled = True
End If
End Sub

Private Sub click_sec_Timer()
clicks_ps(0).Caption = (clicks(0).Caption - lastclicks1) & " / second"
clicks_ps(1).Caption = (clicks(1).Caption - lastclicks2) & " / second"
clicks_ps(2).Caption = (clicks(2).Caption - lastclicks3) & " / second"
clicks_ps(3).Caption = (clicks(3).Caption - lastclicks4) & " / second"

average(0).AddItem (clicks(0).Caption - lastclicks1)
average(1).AddItem (clicks(1).Caption - lastclicks2)
average(2).AddItem (clicks(2).Caption - lastclicks3)
average(3).AddItem (clicks(3).Caption - lastclicks4)

lastclicks1 = clicks(0).Caption
lastclicks2 = clicks(1).Caption
lastclicks3 = clicks(2).Caption
lastclicks4 = clicks(3).Caption



End Sub

Private Sub countdown_clock_Timer()
countdown = countdown - 1
lbl_countdown.Caption = countdown

If countdown = 0 Then
    frm_countdown.Visible = False
    startgame
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'admin codes (only work when NO game is in progress)
If game_in_progress = False Then
    If TranslateKeyCode(KeyCode, False) = key_admin_start Then
        If adminkeypresses = 10 Then
            changegame
        Else
            newgame
            adminkeypresses = adminkeypresses + 1
            resetadminkeypresses.Enabled = False
            resetadminkeypresses.Enabled = True
        End If
    End If
End If

Randomize Timer
playSound = Int(Rnd * 5) + 1

'user codes (only work when in game)
If move_camel = True Then
    check_finish
    If TranslateKeyCode(KeyCode, False) = key_1_1 Then
        If lastkey1 = 1 Then
            img_pl1(0).Picture = img_pl1_img1
            img_pl1(0).Left = img_pl1(0).Left + 20
            If wmp_pl1.URL <> "" And playSound = 1 Then
                wmp_pl1.Controls.play
            End If
            lastkey1 = 0
            clicks(0).Caption = clicks(0).Caption + 1
            Exit Sub
        End If
    End If
    If TranslateKeyCode(KeyCode, False) = key_1_2 Then
        If lastkey1 = 0 Then
            img_pl1(0).Picture = img_pl1_img2
            img_pl1(0).Left = img_pl1(0).Left + 20
            lastkey1 = 1
            clicks(0).Caption = clicks(0).Caption + 1
            Exit Sub
        End If
    End If
    
    If TranslateKeyCode(KeyCode, False) = key_2_1 Then
        If lastkey2 = 1 Then
            img_pl1(1).Picture = img_pl2_img1
            img_pl1(1).Left = img_pl1(1).Left + 20
            If wmp_pl2.URL <> "" And playSound = 1 Then
                wmp_pl2.Controls.play
            End If
            lastkey2 = 0
            clicks(1).Caption = clicks(1).Caption + 1
            Exit Sub
        End If
    End If
    If TranslateKeyCode(KeyCode, False) = key_2_2 Then
        If lastkey2 = 0 Then
            img_pl1(1).Picture = img_pl2_img2
            img_pl1(1).Left = img_pl1(1).Left + 20
            lastkey2 = 1
            clicks(1).Caption = clicks(1).Caption + 1
            Exit Sub
        End If
    End If
    
    If TranslateKeyCode(KeyCode, False) = key_3_1 Then
        If lastkey3 = 1 Then
            img_pl1(2).Picture = img_pl3_img1
            img_pl1(2).Left = img_pl1(2).Left + 20
            If wmp_pl3.URL <> "" And playSound = 1 Then
                wmp_pl3.Controls.play
            End If
            lastkey3 = 0
            clicks(2).Caption = clicks(2).Caption + 1
            Exit Sub
        End If
    End If
    If TranslateKeyCode(KeyCode, False) = key_3_2 Then
        If lastkey3 = 0 Then
            img_pl1(2).Picture = img_pl3_img2
            img_pl1(2).Left = img_pl1(2).Left + 20
            lastkey3 = 1
            clicks(2).Caption = clicks(2).Caption + 1
            Exit Sub
        End If
    End If
    
    If TranslateKeyCode(KeyCode, False) = key_4_1 Then
        If lastkey4 = 1 Then
            img_pl1(3).Picture = img_pl4_img1
            img_pl1(3).Left = img_pl1(3).Left + 20
            If wmp_pl4.URL <> "" And playSound = 1 Then
                wmp_pl4.Controls.play
            End If
            lastkey4 = 0
            clicks(3).Caption = clicks(3).Caption + 1
            Exit Sub
        End If
    End If
    If TranslateKeyCode(KeyCode, False) = key_4_2 Then
        If lastkey4 = 0 Then
            img_pl1(3).Picture = img_pl4_img2
            img_pl1(3).Left = img_pl1(3).Left + 20
            lastkey4 = 1
            clicks(3).Caption = clicks(3).Caption + 1
            Exit Sub
        End If
    End If
End If
        
End Sub

Private Sub Form_Load()
init_game
End Sub
Function FileExists(FileName As String) As Boolean
On Error GoTo ErrorHandler
' get the attributes and ensure that it isn't a directory
FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
' if an error occurs, this function returns False
End Function

Sub newgame()
'reset images/stats
img_pl1(0).Left = 480
img_pl1(1).Left = 480
img_pl1(2).Left = 480
img_pl1(3).Left = 480

img_pl1(0).Picture = img_pl1_img1.Picture
img_pl1(1).Picture = img_pl2_img1.Picture
img_pl1(2).Picture = img_pl3_img1.Picture
img_pl1(3).Picture = img_pl4_img1.Picture

lastclicks1 = 0
lastclicks2 = 0
lastclicks3 = 0
lastclicks4 = 0

clicks(0).Caption = 0
clicks(1).Caption = 0
clicks(2).Caption = 0
clicks(3).Caption = 0

clicks_ps(0).Caption = "0 / second"
clicks_ps(1).Caption = "0 / second"
clicks_ps(2).Caption = "0 / second"
clicks_ps(3).Caption = "0 / second"

gametime.Caption = 0

lastkey1 = 1
lastkey2 = 1
lastkey3 = 1
lastkey4 = 1

average(0).Clear
average(1).Clear
average(2).Clear
average(3).Clear


countdown = maxcountdown
lbl_countdown.Caption = maxcountdown

frm_startgame.Visible = False
frm_gameover.Visible = False
countdown_clock.Enabled = True
frm_countdown.Visible = True

End Sub

Sub startgame()
game_in_progress = True
move_camel = True
game_clock.Enabled = True
click_sec.Enabled = True
End Sub

Private Sub game_clock_Timer()
gametime.Caption = gametime.Caption + 1
If gametime.Caption >= game_timeout Then
    gametimeout
End If

    
End Sub

Sub check_finish()
    If (img_pl1(0).Left + img_pl1(0).Width) > (track(0).Left + track(0).Width) Then
        endgame (0)
    End If
    If (img_pl1(1).Left + img_pl1(1).Width) > (track(1).Left + track(1).Width) Then
        endgame (1)
    End If
    If (img_pl1(2).Left + img_pl1(2).Width) > (track(2).Left + track(2).Width) Then
        endgame (2)
    End If
    If (img_pl1(3).Left + img_pl1(3).Width) > (track(3).Left + track(3).Width) Then
        endgame (3)
    End If
End Sub

Sub endgame(player As Integer)
Dim sum_average As Integer
Dim teller As Integer
wmp.Controls.stop
wmp_winner.Controls.play

move_camel = False
game_clock.Enabled = False
click_sec.Enabled = False

teller = 0
sum_average = 0
Do
    sum_average = sum_average + average(player).List(teller)
    teller = teller + 1
Loop Until teller = average(player).ListCount


overall_average.Caption = "Game average: " & Int((sum_average / average(player).ListCount)) & " hits/sec."
img_winner.Picture = img_pl1(player).Picture
lbl_playerwin.Caption = player_name(player).Caption
frm_gameover.Visible = True
switch_gameover_newgame.Enabled = True
End Sub

Sub gametimeout()
frm_gameover.Visible = True
game_clock.Enabled = False
click_sec.Enabled = False
img_winner.Picture = img_nothing.Picture
lbl_playerwin.Caption = "Timeout exceeded"
switch_gameover_newgame.Enabled = True
End Sub


Private Sub resetadminkeypresses_Timer()
adminkeypresses = 0
resetadminkeypresses.Enabled = False
End Sub

Private Sub switch_gameover_newgame_Timer()
frm_gameover.Visible = False
move_camel = False
game_in_progress = False
frm_startgame.Visible = True
switch_gameover_newgame.Enabled = False
wmp.Controls.play
wmp_winner.Controls.stop
End Sub

Sub changegame()
selectednewgame = 0
If buttonmode = "keyboard" Then
    lbl_changeletter.Visible = True
    img_changebutton.Visible = False
Else
    img_changebutton.Visible = True
    lbl_changeletter.Visible = False
End If
frm_changegame.Visible = True
resetadminkeypresses = False
adminkeypresses = 0
File1.Refresh
teller = 0
Do
    lbl_filename(teller).Visible = False
    teller = teller + 1
Loop Until teller = lbl_filename.Count
DoEvents
teller = 0
    Do
        lbl_filename(teller).Caption = File1.List(teller)
        lbl_filename(teller).Visible = True
        teller = teller + 1
    Loop Until teller = File1.ListCount Or teller = lbl_filename.Count
End Sub

Sub init_game()
'stop all sounds and clear player URL's
wmp.Controls.stop
wmp_pl1.Controls.stop
wmp_pl2.Controls.stop
wmp_pl3.Controls.stop
wmp_pl4.Controls.stop
wmp_winner.Controls.stop

wmp.URL = ""
wmp_pl1.URL = ""
wmp_pl2.URL = ""
wmp_pl3.URL = ""
wmp_pl4.URL = ""
wmp_winner.URL = ""

'zoek config
If FileExists(App.Path & "\config.ini") = False And Command = "" Then
    MsgBox "config file not found", vbCritical, "Error"
    End
Else
    If newconfig <> "" Then
        Open App.Path & "\" & newconfig For Input As #1
    Else
        If Command = "" Then
            Open App.Path & "\config.ini" For Input As #1
        Else
            Open Command For Input As #1
        End If
    End If
    Do
        Input #1, lijn
            If Left(lijn, 1) = "#" Then
                'lijn is comment
                'doe niks...
                GoTo nextitem
            End If
            If InStr(1, lijn, "player_1_img1:", vbTextCompare) > 0 Then
                'player 1 image 1
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                pl1img1 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player_1_img2:", vbTextCompare) > 0 Then
                'player 1 image 2
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                pl1img2 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player_2_img1:", vbTextCompare) > 0 Then
                'player 2 image 1
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                pl2img1 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player_2_img2:", vbTextCompare) > 0 Then
                'player 2 image 2
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                pl2img2 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player_3_img1:", vbTextCompare) > 0 Then
                'player 3 image 1
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                pl3img1 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player_3_img2:", vbTextCompare) > 0 Then
                'player 3 image 2
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                pl3img2 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player_4_img1:", vbTextCompare) > 0 Then
                'player 4 image 1
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                pl4img1 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player_4_img2:", vbTextCompare) > 0 Then
                'player 4 image 2
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                pl4img2 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "background_img:", vbTextCompare) > 0 Then
                'background image
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                bgimg = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "countdown:", vbTextCompare) > 0 Then
                'countdown seconds
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                maxcountdown = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            
            If InStr(1, lijn, "key_user1_button1:", vbTextCompare) > 0 Then
                'key user 1_1
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_1_1 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "key_user1_button2:", vbTextCompare) > 0 Then
                'key user 1_2
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_1_2 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "key_user2_button1:", vbTextCompare) > 0 Then
                'key user 2_1
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_2_1 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "key_user2_button2:", vbTextCompare) > 0 Then
                'key user 2_2
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_2_2 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "key_user3_button1:", vbTextCompare) > 0 Then
                'key user 3_1
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_3_1 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "key_user3_button2:", vbTextCompare) > 0 Then
                'key user 3_2
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_3_2 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "key_user4_button1:", vbTextCompare) > 0 Then
                'key user 4_1
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_4_1 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "key_user4_button2:", vbTextCompare) > 0 Then
                'key user 4_2
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_4_2 = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "player1name:", vbTextCompare) > 0 Then
                'player 1 name
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                player1name = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player2name:", vbTextCompare) > 0 Then
                'player 2 name
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                player2name = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player3name:", vbTextCompare) > 0 Then
                'player 3 name
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                player3name = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            If InStr(1, lijn, "player4name:", vbTextCompare) > 0 Then
                'player 4 name
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                player4name = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "key_admin_start:", vbTextCompare) > 0 Then
                'admin key start
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                key_admin_start = Right(lijn, (Len(lijn) - pos1))
                lbl_changeletter.Caption = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "button_mode:", vbTextCompare) > 0 Then
                'button mode
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                buttonmode = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "img_startbutton:", vbTextCompare) > 0 Then
                'start button image
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                img_startbutton = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "summarytime:", vbTextCompare) > 0 Then
                'summary time
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                summarytime = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "gametimeout:", vbTextCompare) > 0 Then
                'game timeout
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                game_timeout = Right(lijn, (Len(lijn) - pos1))
                GoTo nextitem
            End If
            
            If InStr(1, lijn, "background_sound:", vbTextCompare) > 0 Then
                'background music
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                wmp.Controls.stop
                wmp.URL = Right(lijn, (Len(lijn) - pos1))
                wmp.Controls.play
                wmp.settings.setMode "Loop", True
                GoTo nextitem
            End If
            If InStr(1, lijn, "player1_sound:", vbTextCompare) > 0 Then
                'background music
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                wmp_pl1.Controls.stop
                wmp_pl1.URL = Right(lijn, (Len(lijn) - pos1))
                wmp_pl1.settings.setMode "Loop", False
                wmp_pl1.Controls.stop
                GoTo nextitem
            End If
            If InStr(1, lijn, "player2_sound:", vbTextCompare) > 0 Then
                'background music
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                wmp_pl2.Controls.stop
                wmp_pl2.URL = Right(lijn, (Len(lijn) - pos1))
                wmp_pl2.settings.setMode "Loop", False
                wmp_pl2.Controls.stop
                GoTo nextitem
            End If
            If InStr(1, lijn, "player3_sound:", vbTextCompare) > 0 Then
                'background music
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                wmp_pl3.Controls.stop
                wmp_pl3.URL = Right(lijn, (Len(lijn) - pos1))
                wmp_pl3.settings.setMode "Loop", False
                wmp_pl3.Controls.stop
                GoTo nextitem
            End If
            If InStr(1, lijn, "player4_sound:", vbTextCompare) > 0 Then
                'background music
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                wmp_pl4.Controls.stop
                wmp_pl4.URL = Right(lijn, (Len(lijn) - pos1))
                wmp_pl4.settings.setMode "Loop", False
                wmp_pl4.Controls.stop
                GoTo nextitem
            End If
            If InStr(1, lijn, "winner_sound:", vbTextCompare) > 0 Then
                'background music
                pos1 = InStr(1, lijn, ":", vbTextCompare)
                wmp_winner.Controls.stop
                wmp_winner.URL = Right(lijn, (Len(lijn) - pos1))
                wmp_winner.settings.setMode "Loop", False
                wmp_winner.Controls.stop
                GoTo nextitem
            End If

nextitem:
        Loop Until EOF(1)
    Close #1
End If

'setup temp images
img_pl1_img1.Picture = LoadPicture(App.Path & "\" & pl1img1)
img_pl1_img2.Picture = LoadPicture(App.Path & "\" & pl1img2)
img_pl2_img1.Picture = LoadPicture(App.Path & "\" & pl2img1)
img_pl2_img2.Picture = LoadPicture(App.Path & "\" & pl2img2)
img_pl3_img1.Picture = LoadPicture(App.Path & "\" & pl3img1)
img_pl3_img2.Picture = LoadPicture(App.Path & "\" & pl3img2)
img_pl4_img1.Picture = LoadPicture(App.Path & "\" & pl4img1)
img_pl4_img2.Picture = LoadPicture(App.Path & "\" & pl4img2)
img_bg.Picture = LoadPicture(App.Path & "\" & bgimg)



'setup real images
Me.Picture = img_bg.Picture
img_pl1(0).Picture = img_pl1_img1.Picture
img_pl1(1).Picture = img_pl2_img1.Picture
img_pl1(2).Picture = img_pl3_img1.Picture
img_pl1(3).Picture = img_pl4_img1.Picture


'set vars
game_in_progress = False
player_name(0).Caption = player1name
player_name(1).Caption = player2name
player_name(2).Caption = player3name
player_name(3).Caption = player4name
switch_gameover_newgame.Interval = (summarytime * 1000)
adminkeypresses = 0

If buttonmode = "keyboard" Then
    lbl_admin_start.Caption = key_admin_start
    lbl_admin_start.Visible = True
Else
    img_adminstart.Picture = LoadPicture(App.Path & "\" & img_startbutton)
    img_adminstart.Visible = True
End If
game_in_progress = False
frm_startgame.Visible = True
game_clock.Enabled = False

End Sub

Private Sub waiter_changegame_Timer()
teller = 0
Do
    If lbl_filename(teller).ForeColor = &HFF00& Then
        newconfig = lbl_filename(teller).Caption
        frm_changegame.Visible = False
        waiter_changegame.Enabled = False
        game_in_progress = False
        game_clock.Enabled = False
        move_camel = False
        click_sec.Enabled = False
        init_game
        Exit Sub
    End If
    teller = teller + 1
Loop Until teller = lbl_filename.Count
End Sub
