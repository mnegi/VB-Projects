VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E1E1E1&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4800
      Left            =   3765
      TabIndex        =   13
      Top             =   2955
      Visible         =   0   'False
      Width           =   7425
      Begin VB.OptionButton themeNames 
         Caption         =   "Gray"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   405
         Index           =   0
         Left            =   390
         TabIndex        =   20
         Tag             =   "1"
         Top             =   1695
         Width           =   1425
      End
      Begin VB.OptionButton themeNames 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   405
         Index           =   1
         Left            =   420
         TabIndex        =   19
         Tag             =   "1"
         Top             =   2265
         Width           =   1425
      End
      Begin VB.OptionButton themeNames 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   405
         Index           =   2
         Left            =   390
         TabIndex        =   18
         Tag             =   "1"
         Top             =   2865
         Width           =   1425
      End
      Begin VB.OptionButton themeNames 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   405
         Index           =   3
         Left            =   390
         TabIndex        =   17
         Tag             =   "1"
         Top             =   3465
         Width           =   1425
      End
      Begin ManoharButton.MyButton Command3 
         Height          =   420
         Left            =   5325
         TabIndex        =   14
         Tag             =   "1"
         Top             =   3240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   741
         BTYPE           =   3
         TX              =   "&Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ManoharButton.MyButton Command2 
         Height          =   420
         Left            =   5310
         TabIndex        =   15
         Tag             =   "1"
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   741
         BTYPE           =   3
         TX              =   "&Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ManoharButton.MyButton Command1 
         Height          =   420
         Left            =   5310
         TabIndex        =   16
         Tag             =   "1"
         Top             =   2025
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   741
         BTYPE           =   3
         TX              =   "&Preview"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   4830
         Left            =   -15
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   7455
      End
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Tag             =   "1"
      Top             =   885
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "I&nsurance Database"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4470
      Top             =   60
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1531
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "O&rder Processing"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Tag             =   "1"
      Top             =   2177
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "S&tudent Enrollment"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Tag             =   "1"
      Top             =   2823
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "B&ook Dealer"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":00A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Tag             =   "1"
      Top             =   3469
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "B&anking Enterprise"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":00C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Tag             =   "1"
      Top             =   4115
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "S&QL Engine 1.0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":00E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   6
      Left            =   240
      TabIndex        =   8
      Tag             =   "1"
      Top             =   4761
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Data Export 1.0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":00FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Tag             =   "1"
      Top             =   6699
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Settings"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0118
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Tag             =   "1"
      Top             =   7345
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "A&bout Application"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0134
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   9
      Left            =   240
      TabIndex        =   11
      Tag             =   "1"
      Top             =   7995
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Q&uit To Desktop"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0150
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   735
      Left            =   3720
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1296
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   10
      Left            =   240
      TabIndex        =   21
      Tag             =   "1"
      Top             =   5407
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&BackUp && Recovery"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":016C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   480
      Index           =   11
      Left            =   240
      TabIndex        =   22
      Tag             =   "1"
      Top             =   6053
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "T&heme Explorer"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0188
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   1620
      Top             =   1800
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Image imgRestore 
      Height          =   270
      Left            =   11265
      Stretch         =   -1  'True
      ToolTipText     =   "Restore Position"
      Top             =   75
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Left            =   120
      Stretch         =   -1  'True
      Top             =   60
      Width           =   330
   End
   Begin VB.Image imgMin 
      Height          =   270
      Left            =   10950
      Stretch         =   -1  'True
      Top             =   75
      Width           =   285
   End
   Begin VB.Image imgClose 
      Height          =   270
      Left            =   11595
      Stretch         =   -1  'True
      Top             =   75
      Width           =   285
   End
   Begin VB.Label lbltime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   11790
      TabIndex        =   1
      Top             =   405
      Width           =   45
   End
   Begin VB.Image imgWallpaper 
      Height          =   7860
      Left            =   90
      Stretch         =   -1  'True
      Top             =   720
      Width           =   11850
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Application Launcher"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   915
      TabIndex        =   0
      Tag             =   "1"
      Top             =   75
      Width           =   2670
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Option Explicit
Dim iiindex As Integer
Private Sub Btns_Click(Index As Integer)
On Error GoTo q
Select Case Index
Case 0
Unload Me
frmInsuranceMain.Show
Case 1
Unload Me
frmOrderMain.Show
Case 2
Unload Me
frmStudentMain.Show
Case 3
Unload Me
frmBookMain.Show
Case 4
Unload Me
frmBankMain.Show
Case 5
Me.Hide
frmSQLEngine.Show
Case 6
Me.Hide
frmExportData.Show
Case 7
Me.Hide
frmAgentViewer.Show
Case 8
frmSplashVDL.Show vbModal
Case 9
ManiExtras1.DesktopIconsShow
ManiExtras1.TaskBarShow
End

Case 10
Me.Hide
frmBackUpRecovery.Show
Case 11
Frame1.Visible = True
Image1.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\wallpapers\frmMainwallpaper.jpg")

End Select
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub Command1_Click()
Select Case iiindex
Case 0
Call Change_Theme("Gray")
Image1.Picture = LoadPicture(App.Path & "\Themes\Gray\wallpapers\frmMainWallpaper.jpg")
Case 1
Call Change_Theme("Green")
Image1.Picture = LoadPicture(App.Path & "\Themes\Green\wallpapers\frmMainWallpaper.jpg")
Case 2
Call Change_Theme("Red")
Image1.Picture = LoadPicture(App.Path & "\Themes\Red\wallpapers\frmMainWallpaper.jpg")
Case 3
Call Change_Theme("Blue")
Image1.Picture = LoadPicture(App.Path & "\Themes\Blue\wallpapers\frmMainWallpaper.jpg")
End Select
Call Apply_Theme(Me, 3)
End Sub
Private Sub Command2_Click()
SaveSetting App.title, "Theme", "Name", themeNames(iiindex).Caption
Call Save_Theme
Call Form_Load
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
Image1.Picture = Nothing
Call Get_Theme
Call Apply_Theme(Me, 3)
End Sub

Private Sub Command4_Click()
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from tab"
Set rs = Cmd.Execute
If rs.RecordCount <> 0 Then
While Not rs.EOF
    Call Recover(rs.Fields(0))
    rs.MoveNext
Wend
    
End If
End Sub
Function Recover(TableName As String) As Integer
On Error GoTo errorHandle


'****************************************
'Open the selected table
'****************************************
Dim Recovered As Integer
Dim RsR As New ADODB.Recordset
RsR.Open "select * from " & TableName, Cn, adOpenStatic, adLockOptimistic

'****************************************
'count the total no of fileds in table
'****************************************
Dim RCount As Integer
RCount = RsR.Fields.Count - 1

'****************************************
'check the field name and its type
'****************************************
'For jp = 0 To RsR.Fields.Count - 1
'MsgBox RsR.Fields(jp).NAME & " : " & RsR.Fields(jp).Type
'Next

Dim Counter As Integer

Dim CmdTxt As String
Dim newline As String

'****************************************
'Open the backup file (saved with .bkp)
'read that file and do the recovery
'****************************************

Open App.Path & "\BackUp\" & TableName & ".bkp" For Input As #1
Do Until EOF(1)
    Line Input #1, newline
    Counter = 0
    CmdTxt = "insert into " & TableName & " values("
M:
    For k = 1 To Len(newline)
        If Mid(newline, k, 1) = "^" Then
'        MsgBox Counter + 1 & " : " & Mid(newline, 1, k - 1)
        '*************************************
        'check if it is string type
        '200 represents string type
        '*************************************
        If RsR.Fields(Counter).Type = 200 Then
            'check if it is last field
            If Counter = RCount Then
                CmdTxt = CmdTxt & "'" & Mid(newline, 1, k - 1) & "'"
            Else
                CmdTxt = CmdTxt & "'" & Mid(newline, 1, k - 1) & "',"
            End If
        End If
        
        '*************************************
        'check if it is date type
        '135 represents string type
        '*************************************
        If RsR.Fields(Counter).Type = 135 Then
            'check if it is last field
            If Counter = RCount Then
                CmdTxt = CmdTxt & "'" & Format(Mid(newline, 1, k - 1), "dd-mmm-yy") & "'"
            Else
                CmdTxt = CmdTxt & "'" & Format(Mid(newline, 1, k - 1), "dd-mmm-yy") & "',"
            End If
        End If
    '*************************************
    'check if it is number
    '131 represents number type
    '*************************************
    If RsR.Fields(Counter).Type = 131 Then
        'check if it is last field
        If Counter = RCount Then
            CmdTxt = CmdTxt & Mid(newline, 1, k - 1)
        Else
            CmdTxt = CmdTxt & Mid(newline, 1, k - 1) & ","
        End If
    End If
    
    
    Counter = Counter + 1

    newline = Mid(newline, k + 1, Len(newline))
        GoTo M
End If
Next
CmdTxt = CmdTxt & ")"
Cmd.ActiveConnection = Cn
Cmd.CommandText = CmdTxt
Cmd.Execute
Recovered = Recovered + 1
Loop
Close #1

Recover = Recovered
errorHandle:
If Err.Number <> 0 Then
'MsgBox Err.Number & " : " & Err.Description
If Recovered > -1 Then
Recovered = Recovered - 1
End If
Resume Next
End If
End Function

End Function

Private Sub Form_Activate()
Call Get_Theme
Call Apply_Theme(Me, 3)
ManiExtras1.MinimizeAll
'Dim rc As Long
'rc = BringWindowToTop(hWnd)
'q:
'If Err.Number <> 0 Then
'Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
'frmMsgbox.Show vbModal
'End If


End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Move.cur")

'MousePointer = 15
Call ReleaseCapture
Call SendMessage(hWnd, &HA1, 2, 0&)
'MousePointer = 1
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")

  '*********************************
  ' hold down left mouse button and
  ' then move mouse for moving form
  '*********************************
  
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


  
End Sub

Private Sub Form_Load()
On Error GoTo q
Timer1.Enabled = True
Me.Caption = "Main Application Launcher"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmMain.ico")
imgWallpaper.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Wallpapers\frmMainWallpaper.jpg")
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "000")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
End Sub
Private Sub imgRestore_Click()
ManiExtras1.DesktopIconsHide
ManiExtras1.TaskBarHide
Me.Left = 0
Me.Top = 0
End Sub
Private Sub imgRestore_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "010")
q:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub imgClose_Click()
ManiExtras1.DesktopIconsShow
ManiExtras1.TaskBarShow
End
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "001")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub imgMin_Click()
On Error GoTo q
Me.WindowState = 1
ManiExtras1.DesktopIconsShow
ManiExtras1.TaskBarShow
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
End Sub
Private Sub imgMin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
Call ShowButtons(Me, "100")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub
Private Sub imgWallpaper_DblClick()
'
'Call SetWallpaper(Me.Picture)
End Sub
Private Sub themeNames_Click(Index As Integer)
iiindex = Index
End Sub

Private Sub Timer1_Timer()
lbltime.Caption = Format(Now, "hh:mm:ss", vbUseSystemDayOfWeek)
End Sub
