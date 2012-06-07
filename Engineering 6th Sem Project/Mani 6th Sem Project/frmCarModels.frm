VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmCarModels 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmCarModels.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbMakes 
      Height          =   315
      Left            =   1980
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Tag             =   "1"
      ToolTipText     =   "Make Name"
      Top             =   1440
      Width           =   4440
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6810
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbManufacturer 
      Height          =   315
      Left            =   1965
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Manufacture Name"
      Top             =   975
      Width           =   4440
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Navigation"
      Height          =   2670
      Left            =   9690
      TabIndex        =   20
      Tag             =   "1"
      Top             =   810
      Width           =   2175
      Begin ManoharButton.MyButton Btns 
         Height          =   375
         Index           =   7
         Left            =   135
         TabIndex        =   10
         Tag             =   "1"
         Top             =   420
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&First"
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
         MICON           =   "frmCarModels.frx":000C
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
         Height          =   375
         Index           =   8
         Left            =   150
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1005
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Previous"
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
         MICON           =   "frmCarModels.frx":0028
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
         Height          =   375
         Index           =   9
         Left            =   135
         TabIndex        =   12
         Tag             =   "1"
         Top             =   1580
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Next"
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
         MICON           =   "frmCarModels.frx":0044
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
         Height          =   375
         Index           =   10
         Left            =   150
         TabIndex        =   13
         Tag             =   "1"
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Last"
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
         MICON           =   "frmCarModels.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Menu"
      Height          =   4890
      Left            =   9690
      TabIndex        =   19
      Tag             =   "1"
      Top             =   3645
      Width           =   2175
      Begin ManoharButton.MyButton Btns 
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Tag             =   "1"
         Top             =   330
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Add New"
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
         MICON           =   "frmCarModels.frx":007C
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
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   3
         Tag             =   "1"
         Top             =   997
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCarModels.frx":0098
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
         Height          =   375
         Index           =   3
         Left            =   135
         TabIndex        =   7
         Tag             =   "1"
         Top             =   2331
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Edit"
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
         MICON           =   "frmCarModels.frx":00B4
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
         Height          =   375
         Index           =   4
         Left            =   135
         TabIndex        =   8
         Tag             =   "1"
         Top             =   2998
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Update"
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
         MICON           =   "frmCarModels.frx":00D0
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
         Height          =   375
         Index           =   5
         Left            =   135
         TabIndex        =   4
         Tag             =   "1"
         Top             =   3665
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Cancel"
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
         MICON           =   "frmCarModels.frx":00EC
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
         Height          =   375
         Index           =   2
         Left            =   135
         TabIndex        =   6
         Tag             =   "1"
         Top             =   1664
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Delete"
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
         MICON           =   "frmCarModels.frx":0108
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
         Height          =   375
         Index           =   6
         Left            =   135
         TabIndex        =   9
         Tag             =   "1"
         Top             =   4335
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Main"
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
         MICON           =   "frmCarModels.frx":0124
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame fraManufacturer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6600
      Left            =   285
      TabIndex        =   17
      Tag             =   "1"
      Top             =   1965
      Width           =   9270
      Begin VB.ComboBox cmbDisp 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Tag             =   "2"
         ToolTipText     =   "Select an option to diaplay."
         Top             =   345
         Width           =   2895
      End
      Begin VB.ComboBox cmbTypes 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Tag             =   "1"
         ToolTipText     =   "Model Type"
         Top             =   1057
         Width           =   3405
      End
      Begin VB.ComboBox cmbModels 
         Height          =   315
         Left            =   1725
         TabIndex        =   26
         Tag             =   "1"
         Text            =   "cmbModels"
         ToolTipText     =   "Make Name"
         Top             =   360
         Width           =   3405
      End
      Begin XPTEXTBOX.text txtPrice 
         Height          =   420
         Left            =   1710
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "Price in Rupees"
         Top             =   1770
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   741
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   25
         FontBold        =   -1  'True
         LineColor       =   11643476
         Text            =   ""
         BackColor       =   16777215
         ForeColor       =   4787463
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   5520
         Index           =   1
         Left            =   5310
         TabIndex        =   47
         Top             =   945
         Width           =   3900
         Begin XPTEXTBOX.text txtEngine 
            Height          =   420
            Index           =   0
            Left            =   1890
            TabIndex        =   48
            Tag             =   "1"
            ToolTipText     =   "Engine"
            Top             =   1140
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtEngine 
            Height          =   420
            Index           =   1
            Left            =   1875
            TabIndex        =   49
            Tag             =   "1"
            ToolTipText     =   "Engine"
            Top             =   1794
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtEngine 
            Height          =   420
            Index           =   2
            Left            =   1890
            TabIndex        =   50
            Tag             =   "1"
            ToolTipText     =   "Engine"
            Top             =   2460
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtEngine 
            Height          =   420
            Index           =   3
            Left            =   1875
            TabIndex        =   51
            Tag             =   "1"
            ToolTipText     =   "Engine"
            Top             =   3120
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtEngine 
            Height          =   420
            Index           =   4
            Left            =   1890
            TabIndex        =   52
            Tag             =   "1"
            ToolTipText     =   "Engine"
            Top             =   3870
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtEngine 
            Height          =   420
            Index           =   5
            Left            =   1875
            TabIndex        =   53
            Tag             =   "1"
            ToolTipText     =   "Engine"
            Top             =   4620
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line6 
            BorderWidth     =   2
            X1              =   105
            X2              =   870
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Installation"
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
            Left            =   195
            TabIndex        =   60
            Tag             =   "1"
            Top             =   4620
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Power To Weight Ratio"
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
            Height          =   675
            Left            =   195
            TabIndex        =   59
            Tag             =   "1"
            Top             =   3870
            Width           =   1410
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Specific OutPut"
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
            Left            =   195
            TabIndex        =   58
            Tag             =   "1"
            Top             =   3120
            Width           =   1515
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max Torque"
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
            Left            =   195
            TabIndex        =   57
            Tag             =   "1"
            Top             =   2460
            Width           =   1230
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max Power"
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
            Left            =   180
            TabIndex        =   56
            Tag             =   "1"
            Top             =   1794
            Width           =   1125
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Layout"
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
            Left            =   180
            TabIndex        =   55
            Tag             =   "1"
            Top             =   1140
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Engine"
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
            Left            =   120
            TabIndex        =   54
            Tag             =   "1"
            Top             =   375
            Width           =   690
         End
      End
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   5520
         Index           =   0
         Left            =   5295
         TabIndex        =   29
         Top             =   945
         Width           =   3900
         Begin XPTEXTBOX.text txtAcc 
            Height          =   420
            Index           =   0
            Left            =   1920
            TabIndex        =   40
            Tag             =   "1"
            ToolTipText     =   "Acceleration"
            Top             =   1140
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtAcc 
            Height          =   420
            Index           =   1
            Left            =   1935
            TabIndex        =   41
            Tag             =   "1"
            ToolTipText     =   "Acceleration"
            Top             =   1650
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtAcc 
            Height          =   420
            Index           =   2
            Left            =   1920
            TabIndex        =   42
            Tag             =   "1"
            ToolTipText     =   "Acceleration"
            Top             =   2160
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtAcc 
            Height          =   420
            Index           =   3
            Left            =   1920
            TabIndex        =   43
            Tag             =   "1"
            ToolTipText     =   "Acceleration"
            Top             =   2685
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtAcc 
            Height          =   420
            Index           =   4
            Left            =   1920
            TabIndex        =   44
            Tag             =   "1"
            ToolTipText     =   "Acceleration"
            Top             =   3195
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtSpeed 
            Height          =   420
            Index           =   0
            Left            =   1920
            TabIndex        =   45
            Tag             =   "1"
            ToolTipText     =   "Speed"
            Top             =   4335
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtSpeed 
            Height          =   420
            Index           =   1
            Left            =   1920
            TabIndex        =   46
            Tag             =   "1"
            ToolTipText     =   "Speed"
            Top             =   4950
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line9 
            BorderWidth     =   2
            X1              =   180
            X2              =   810
            Y1              =   4335
            Y2              =   4335
         End
         Begin VB.Line Line8 
            BorderWidth     =   2
            X1              =   165
            X2              =   1425
            Y1              =   915
            Y2              =   915
         End
         Begin VB.Line Line7 
            X1              =   90
            X2              =   1320
            Y1              =   465
            Y2              =   465
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Average Speed"
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
            Left            =   180
            TabIndex        =   39
            Tag             =   "1"
            Top             =   5040
            Width           =   1500
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Top Speed"
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
            Left            =   180
            TabIndex        =   38
            Tag             =   "1"
            Top             =   4470
            Width           =   1050
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
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
            Left            =   180
            TabIndex        =   37
            Tag             =   "1"
            Top             =   3990
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 500 KPH"
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
            Left            =   180
            TabIndex        =   36
            Tag             =   "1"
            Top             =   3150
            Width           =   1230
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 400 KPH"
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
            Left            =   180
            TabIndex        =   35
            Tag             =   "1"
            Top             =   2640
            Width           =   1230
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 300 KPH"
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
            Left            =   180
            TabIndex        =   34
            Tag             =   "1"
            Top             =   2130
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 200 KPH"
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
            Left            =   180
            TabIndex        =   33
            Tag             =   "1"
            Top             =   1635
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 100 KPH"
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
            Left            =   180
            TabIndex        =   32
            Tag             =   "1"
            Top             =   1125
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acceleration"
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
            Left            =   180
            TabIndex        =   31
            Tag             =   "1"
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "How Fast ? "
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
            Left            =   90
            TabIndex        =   30
            Tag             =   "1"
            Top             =   150
            Width           =   1170
         End
      End
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   5520
         Index           =   2
         Left            =   5295
         TabIndex        =   61
         Top             =   930
         Width           =   3900
         Begin XPTEXTBOX.text txtFuel 
            Height          =   420
            Index           =   0
            Left            =   1830
            TabIndex        =   62
            Tag             =   "1"
            ToolTipText     =   "Fuel Economy"
            Top             =   720
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtFuel 
            Height          =   420
            Index           =   1
            Left            =   1845
            TabIndex        =   63
            Tag             =   "1"
            ToolTipText     =   "Fuel Economy"
            Top             =   1230
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtFuel 
            Height          =   420
            Index           =   2
            Left            =   1845
            TabIndex        =   64
            Tag             =   "1"
            ToolTipText     =   "Fuel Economy"
            Top             =   1740
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtSuspension 
            Height          =   420
            Index           =   0
            Left            =   1845
            TabIndex        =   65
            Tag             =   "1"
            ToolTipText     =   "Suspension"
            Top             =   3765
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtSuspension 
            Height          =   420
            Index           =   1
            Left            =   1845
            TabIndex        =   66
            Tag             =   "1"
            ToolTipText     =   "Suspension"
            Top             =   4290
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtFuel 
            Height          =   420
            Index           =   3
            Left            =   1860
            TabIndex        =   120
            Tag             =   "1"
            ToolTipText     =   "Fuel Economy"
            Top             =   2280
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuel Capacity"
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
            Left            =   195
            TabIndex        =   121
            Tag             =   "1"
            Top             =   2295
            Width           =   1335
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            X1              =   105
            X2              =   1185
            Y1              =   3645
            Y2              =   3660
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   120
            X2              =   1515
            Y1              =   510
            Y2              =   525
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rear"
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
            Left            =   180
            TabIndex        =   73
            Tag             =   "1"
            Top             =   4380
            Width           =   480
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Front"
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
            Left            =   180
            TabIndex        =   72
            Tag             =   "1"
            Top             =   3900
            Width           =   540
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suspension"
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
            Left            =   120
            TabIndex        =   71
            Tag             =   "1"
            Top             =   3315
            Width           =   1110
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Highway"
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
            Left            =   195
            TabIndex        =   70
            Tag             =   "1"
            Top             =   1710
            Width           =   840
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City"
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
            Left            =   180
            TabIndex        =   69
            Tag             =   "1"
            Top             =   1215
            Width           =   405
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Average"
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
            Left            =   180
            TabIndex        =   68
            Tag             =   "1"
            Top             =   705
            Width           =   840
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuel Economy"
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
            Left            =   90
            TabIndex        =   67
            Tag             =   "1"
            Top             =   150
            Width           =   1395
         End
      End
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   5520
         Index           =   3
         Left            =   5310
         TabIndex        =   74
         Top             =   945
         Width           =   3900
         Begin XPTEXTBOX.text txtSteering 
            Height          =   420
            Index           =   0
            Left            =   1950
            TabIndex        =   75
            Tag             =   "1"
            ToolTipText     =   "Steering"
            Top             =   810
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtSteering 
            Height          =   420
            Index           =   1
            Left            =   1950
            TabIndex        =   76
            Tag             =   "1"
            ToolTipText     =   "Steering"
            Top             =   1320
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtSteering 
            Height          =   420
            Index           =   2
            Left            =   1950
            TabIndex        =   77
            Tag             =   "1"
            ToolTipText     =   "Steering"
            Top             =   1830
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtWheels 
            Height          =   420
            Index           =   0
            Left            =   1950
            TabIndex        =   78
            Tag             =   "1"
            ToolTipText     =   "Wheels"
            Top             =   3630
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtWheels 
            Height          =   420
            Index           =   1
            Left            =   1950
            TabIndex        =   79
            Tag             =   "1"
            ToolTipText     =   "Wheels"
            Top             =   4140
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtWheels 
            Height          =   420
            Index           =   2
            Left            =   1950
            TabIndex        =   87
            Tag             =   "1"
            ToolTipText     =   "Wheels"
            Top             =   4650
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   60
            X2              =   945
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   60
            X2              =   945
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Made Of"
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
            Left            =   195
            TabIndex        =   88
            Tag             =   "1"
            Top             =   4605
            Width           =   870
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wheels"
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
            Left            =   105
            TabIndex        =   86
            Tag             =   "1"
            Top             =   2865
            Width           =   735
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rear Wheel Size"
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
            Left            =   195
            TabIndex        =   85
            Tag             =   "1"
            Top             =   4095
            Width           =   1620
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Front Wheel Size"
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
            Left            =   180
            TabIndex        =   84
            Tag             =   "1"
            Top             =   3600
            Width           =   1680
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Turns"
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
            Left            =   180
            TabIndex        =   83
            Tag             =   "1"
            Top             =   1800
            Width           =   585
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Power"
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
            Left            =   180
            TabIndex        =   82
            Tag             =   "1"
            Top             =   1305
            Width           =   615
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
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
            Left            =   180
            TabIndex        =   81
            Tag             =   "1"
            Top             =   795
            Width           =   495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Steering"
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
            Left            =   75
            TabIndex        =   80
            Tag             =   "1"
            Top             =   195
            Width           =   825
         End
      End
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   5520
         Index           =   5
         Left            =   5295
         TabIndex        =   104
         Top             =   930
         Width           =   3900
         Begin XPTEXTBOX.text txtBig 
            Height          =   420
            Index           =   0
            Left            =   1860
            TabIndex        =   105
            Tag             =   "1"
            ToolTipText     =   "How Big"
            Top             =   1140
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtBig 
            Height          =   420
            Index           =   1
            Left            =   1860
            TabIndex        =   106
            Tag             =   "1"
            ToolTipText     =   "How Big"
            Top             =   1650
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtBig 
            Height          =   420
            Index           =   2
            Left            =   1860
            TabIndex        =   107
            Tag             =   "1"
            ToolTipText     =   "How Big"
            Top             =   2160
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtBig 
            Height          =   420
            Index           =   3
            Left            =   1860
            TabIndex        =   108
            Tag             =   "1"
            ToolTipText     =   "How Big"
            Top             =   2685
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtBig 
            Height          =   420
            Index           =   4
            Left            =   1860
            TabIndex        =   109
            Tag             =   "1"
            ToolTipText     =   "How Big"
            Top             =   3195
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtBig 
            Height          =   420
            Index           =   5
            Left            =   1860
            TabIndex        =   116
            Tag             =   "1"
            ToolTipText     =   "How Big"
            Top             =   3705
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtGearBox 
            Height          =   420
            Left            =   1860
            TabIndex        =   122
            Tag             =   "1"
            ToolTipText     =   "GearBox Type"
            Top             =   4860
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
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
            Left            =   180
            TabIndex        =   124
            Tag             =   "1"
            Top             =   4815
            Width           =   495
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GearBox"
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
            Left            =   90
            TabIndex        =   123
            Tag             =   "1"
            Top             =   4275
            Width           =   885
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   150
            X2              =   960
            Y1              =   4635
            Y2              =   4650
         End
         Begin VB.Line Line10 
            BorderWidth     =   2
            X1              =   120
            X2              =   1110
            Y1              =   675
            Y2              =   660
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rear Track"
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
            Left            =   180
            TabIndex        =   117
            Tag             =   "1"
            Top             =   3660
            Width           =   1125
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "How Big ? "
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
            Left            =   90
            TabIndex        =   115
            Tag             =   "1"
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
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
            Left            =   180
            TabIndex        =   114
            Tag             =   "1"
            Top             =   1125
            Width           =   705
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
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
            Left            =   180
            TabIndex        =   113
            Tag             =   "1"
            Top             =   1635
            Width           =   585
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
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
            Left            =   180
            TabIndex        =   112
            Tag             =   "1"
            Top             =   2130
            Width           =   675
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WheelBase"
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
            Left            =   180
            TabIndex        =   111
            Tag             =   "1"
            Top             =   2640
            Width           =   1125
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Front Track"
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
            Left            =   180
            TabIndex        =   110
            Tag             =   "1"
            Top             =   3150
            Width           =   1185
         End
      End
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   5520
         Index           =   4
         Left            =   5295
         TabIndex        =   89
         Top             =   930
         Width           =   3900
         Begin XPTEXTBOX.text txtTyres 
            Height          =   420
            Index           =   0
            Left            =   1965
            TabIndex        =   90
            Tag             =   "1"
            ToolTipText     =   "Tyres"
            Top             =   810
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtTyres 
            Height          =   420
            Index           =   1
            Left            =   1950
            TabIndex        =   91
            Tag             =   "1"
            ToolTipText     =   "Tyres"
            Top             =   1320
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtTyres 
            Height          =   420
            Index           =   2
            Left            =   1950
            TabIndex        =   92
            Tag             =   "1"
            ToolTipText     =   "Tyres"
            Top             =   1830
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtBreakes 
            Height          =   420
            Index           =   0
            Left            =   1965
            TabIndex        =   93
            Tag             =   "1"
            ToolTipText     =   "Breakes"
            Top             =   3630
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtBreakes 
            Height          =   420
            Index           =   1
            Left            =   1950
            TabIndex        =   94
            Tag             =   "1"
            ToolTipText     =   "Breakes"
            Top             =   4140
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPTEXTBOX.text txtBreakes 
            Height          =   420
            Index           =   2
            Left            =   1950
            TabIndex        =   95
            Tag             =   "1"
            ToolTipText     =   "Breakes"
            Top             =   4650
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   25
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   16777215
            ForeColor       =   4787463
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line12 
            BorderWidth     =   2
            X1              =   180
            X2              =   990
            Y1              =   3255
            Y2              =   3240
         End
         Begin VB.Line Line11 
            BorderWidth     =   2
            X1              =   120
            X2              =   720
            Y1              =   585
            Y2              =   585
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tyres"
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
            Left            =   150
            TabIndex        =   103
            Tag             =   "1"
            Top             =   225
            Width           =   570
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Models"
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
            Left            =   180
            TabIndex        =   102
            Tag             =   "1"
            Top             =   795
            Width           =   750
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Front"
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
            Left            =   180
            TabIndex        =   101
            Tag             =   "1"
            Top             =   1305
            Width           =   540
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rear"
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
            Left            =   180
            TabIndex        =   100
            Tag             =   "1"
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Front"
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
            Left            =   180
            TabIndex        =   99
            Tag             =   "1"
            Top             =   3600
            Width           =   540
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rear"
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
            Left            =   195
            TabIndex        =   98
            Tag             =   "1"
            Top             =   4095
            Width           =   480
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Breakes"
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
            Left            =   180
            TabIndex        =   97
            Tag             =   "1"
            Top             =   2865
            Width           =   840
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discs"
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
            Left            =   195
            TabIndex        =   96
            Tag             =   "1"
            Top             =   4605
            Width           =   540
         End
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5280
         TabIndex        =   119
         Tag             =   "1"
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Car Model"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Tag             =   "1"
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lblImgLogo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click herto add a picture"
         Height          =   195
         Left            =   1710
         TabIndex        =   23
         ToolTipText     =   "Click Here to Add Logo"
         Top             =   4995
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblLogo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PHOTO"
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
         Left            =   90
         TabIndex        =   22
         Tag             =   "1"
         Top             =   3045
         Width           =   825
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price ( Rupees )"
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
         Left            =   90
         TabIndex        =   21
         Tag             =   "1"
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   3030
         Left            =   165
         Stretch         =   -1  'True
         ToolTipText     =   "Click here to add a picture"
         Top             =   3465
         Width           =   5025
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Tag             =   "1"
         Top             =   1065
         Width           =   1200
      End
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   165
      Left            =   780
      TabIndex        =   15
      ToolTipText     =   "Manufacture Name"
      Top             =   465
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   291
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   75
      Top             =   15
   End
   Begin VB.Label lblFounder 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   285
      TabIndex        =   25
      Tag             =   "1"
      Top             =   1425
      Width           =   1215
   End
   Begin VB.Image imgCarManu 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   6930
      Stretch         =   -1  'True
      Top             =   930
      Width           =   2595
   End
   Begin VB.Label lblRegNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   285
      TabIndex        =   16
      Tag             =   "1"
      Top             =   915
      Width           =   1335
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   135
      Left            =   11775
      TabIndex        =   14
      Top             =   405
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cars Models"
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
      Left            =   900
      TabIndex        =   0
      Tag             =   "1"
      Top             =   75
      Width           =   1275
   End
End
Attribute VB_Name = "frmCarModels"
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

Dim Saved As Boolean
Dim loaded As Boolean
Dim logo As String
Dim photoDisPath As String
Dim rs1 As New ADODB.Recordset
Dim disp As Boolean
Private Sub Btns_Click(Index As Integer)
''On Error GoTo q
Select Case Index

Case 0
'add new record
Saved = False
Call CLEAR

cmbManufacturer.CLEAR
Dim RSM As New ADODB.Recordset
RSM.Open "SELECT * FROM MANUFACTURER", Cn, adOpenStatic
If RSM.RecordCount <> 0 Then
RSM.MoveFirst
While Not RSM.EOF
cmbManufacturer.AddItem RSM.Fields(0)
RSM.MoveNext
Wend
If cmbManufacturer.ListCount <> 0 Then
cmbManufacturer.Text = cmbManufacturer.List(0)
End If
End If

Call LockTheControls(False)
Call btnEnable("01000100000")
imgLogo.Picture = Nothing
lblImgLogo.Visible = True
Call LoadCarTypes
Case 1
'save record
'On Error GoTo HandleErr

Cmd.ActiveConnection = Cn
Cmd.CommandText = "insert into carmodels values('" & _
cmbModels.Text & "','" & cmbMakes.Text & "','" & cmbTypes.Text & "','" & txtPrice.Text & "','" & _
txtAcc(0).Text & "','" & txtAcc(1).Text & "','" & txtAcc(2).Text & "','" & txtAcc(3).Text & "','" & _
txtAcc(4).Text & "','" & txtSpeed(0).Text & "','" & txtSpeed(1).Text & "','" & txtEngine(0).Text & "','" & _
txtEngine(1).Text & "','" & txtEngine(2).Text & "','" & txtEngine(3).Text & "','" & txtEngine(4).Text & "','" & _
txtEngine(5).Text & "','" & txtFuel(0).Text & "','" & txtFuel(1).Text & "','" & txtFuel(2).Text & "','" & _
txtFuel(3).Text & "','" & txtGearBox.Text & "','" & _
txtSuspension(0).Text & "','" & txtSuspension(1).Text & "','" & txtSteering(0).Text & "','" & txtSteering(1).Text & "','" & _
txtSteering(2).Text & "','" & txtWheels(0).Text & "','" & txtWheels(1).Text & "','" & txtWheels(2).Text & "','" & txtTyres(0).Text & "','" & txtTyres(1).Text & "','" & _
txtTyres(2).Text & "','" & txtBreakes(0).Text & "','" & txtBreakes(1).Text & "','" & _
txtBreakes(2).Text & "','" & txtBig(0).Text & "','" & txtBig(1).Text & "','" & _
txtBig(2).Text & "','" & txtBig(3).Text & "','" & txtBig(4).Text & "','" & txtBig(5).Text & "')"
On Error GoTo M
MsgBox Cmd.CommandText
Cmd.Execute
M:
If Err.Number <> 0 Then
MsgBox Err.Number & " : " & Err.Description
Exit Sub
End If

If ManiExtras1.Path_Exist(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models") Then
'do nothing
Else
MyFile.CreateFolder (App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models")
End If

If MyFile.FolderExists(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text) Then
'do nothing
Else
MyFile.CreateFolder (App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text)
Call ManiExtras1.Copy_File(CommonDialog1.Filename, App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text & "\")
If CommonDialog1.FileTitle <> "Logo.jpg" Then
Call ManiExtras1.RenameFile(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text & "\" & CommonDialog1.FileTitle, App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text & "\" & cmbModels.Text & ".jpg")
End If
End If

Saved = True
Call btnEnable("10110111111")
Call Load_Records
lblImgLogo.Visible = False
rs1.Requery

HandleErr:
If Err.Number <> 0 Then
Call Handle_Error(Err.Number, "Error : " & Err.Number, Err.Description, "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
Exit Sub
End If

Case 2
'delete
Call Handle_Error("Confirm Delete", "Deleting Current record", "This will delete the currently displayed record. Are you sure to delete ?", "Information1.jpg", "information.ico", 2, 0)
frmMsgbox.cmdCancel.Caption = "&No"
frmMsgbox.cmdok.Caption = "&Yes"
frmMsgbox.Show vbModal
If Not MsgBOx_R_Value Then
Exit Sub
Else
Cmd.ActiveConnection = Cn
Dim Cstring As String
Cstring = "delete from carmodels where name='" & cmbModels.Text & "' and make = '" & cmbMakes.Text & "'"
Cmd.CommandText = Cstring
MsgBox Cmd.CommandText
Cmd.Execute

If MyFile.FolderExists(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text) Then
MyFile.DeleteFolder (App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text)
End If
rs1.Requery
End If
Call CLEAR
Call Load_Records
Call LockTheControls(True)


Case 3
'edit
LockTheControls (False)
Saved = False
btnEnable ("0000110000")

Case 4
'update
Cmd.ActiveConnection = Cn

'Dim Cstring As String
Cstring = "update carmodels set M_TYPE ='" & cmbTypes.Text & "',PRICE='" & txtPrice.Text & "',ACC0TO100='" & txtAcc(0).Text & "',ACC0TO200='" & txtAcc(1).Text & "',ACC0TO300='" & txtAcc(2).Text & "',ACC0TO400='" & txtAcc(3).Text & "',ACC0TO500='" & txtAcc(4).Text & "',TOPSPEED='" & txtSpeed(0).Text & "',AVGSPEED='" & txtSpeed(1).Text & "',E_LAYOUT='" & txtEngine(0).Text & "',E_MAXPOWER='" & txtEngine(1).Text & "',E_MAXTORQUE='" & txtEngine(2).Text & "',E_SOUTPUT='" & txtEngine(3).Text & "',E_PTWR='" & txtEngine(4).Text & "',E_INSTALL='" & txtEngine(5).Text & "',F_AVG='" & txtFuel(0).Text & "',F_CITY='" & txtFuel(1).Text & "',F_HIGHWAY='" & txtFuel(2).Text & "',F_CAPACITY='" & txtFuel(3).Text & "',GEARBOX='" & txtGearBox.Text & "',S_FRONT='" & txtSuspension(0).Text & "',S_REAR='" & txtSuspension(1).Text & "',STR_TYPE='" & txtSteering(0).Text & "',STR_POWER='" & txtSteering(1).Text & "',STR_TURNS='" & txtSteering(2).Text & "',W_SSIZE='" & txtWheels(0).Text & "'," & _
"W_RSIZE='" & txtWheels(1).Text & "',W_MADEOF='" & txtWheels(2).Text & "',T_MODELS='" & txtTyres(0).Text & "',T_FRONT='" & txtTyres(1).Text & "',T_REAR='" & txtTyres(2).Text & "',B_FRONT='" & txtBreakes(0).Text & "',B_REAR='" & txtBreakes(1).Text & "',B_DISCS='" & txtBreakes(2).Text & "',B_LENGTH='" & txtBig(0).Text & "',B_WIDTH='" & txtBig(1).Text & "',B_HEIGHT='" & txtBig(2).Text & "',B_WHEELBASE='" & txtBig(3).Text & "',B_FRONTTRACK='" & txtBig(4).Text & "',B_REARTRACK='" & txtBig(5).Text & "' where NAME='" & cmbModels.Text & "' AND MAKE='" & cmbMakes.Text & "'"
'
Cmd.CommandText = Cstring
MsgBox Cmd.CommandText
Cmd.Execute
'



Saved = True
rs1.Requery
Call Load_Records
lblImgLogo.Visible = False
'rs1.Requery

Case 5
''cancel
Saved = True
LockTheControls (True)
Call btnEnable("10110111111")
Call Load_Records
lblImgLogo.Visible = False

Case 6
Call ExitPrg
Case 7

If Not rs1.BOF Then
rs1.MoveFirst
Call DISPLAY(rs1)
If rs1.BOF = True Then
Btns(7).Enabled = False
End If
End If

Case 8
If Not rs1.BOF Then
rs1.MovePrevious
If rs1.BOF Then rs1.MoveFirst
End If
Call DISPLAY(rs1)

Case 9

If Not rs1.EOF Then
rs1.MoveNext
If rs1.EOF Then rs1.MoveLast

End If
Call DISPLAY(rs1)

Case 10
If Not rs1.EOF Then
rs1.MoveLast
End If
Call DISPLAY(rs1)

End Select



q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub


Private Sub cmbDisp_Click()
Select Case cmbDisp.ListIndex
Case 0
Call DispFrames(0)
Case 1
Call DispFrames(1)
Case 2
Call DispFrames(2)
Case 3
Call DispFrames(3)
Case 4
Call DispFrames(4)
Case 5
Call DispFrames(5)
End Select

End Sub

Private Sub cmbMakes_Click()
If Saved = False Then
Call Carmodels(cmbModels, cmbMakes.Text)
If cmbModels.ListCount <> 0 Then
cmbModels.Text = cmbModels.List(0)
End If
End If
End Sub

Private Sub cmbManufacturer_Click()
If Saved = False Then

Dim rsz As New ADODB.Recordset
rsz.Open "select * from carmakes where vendor= '" & cmbManufacturer.Text & "' order by name", Cn, adOpenStatic
cmbMakes.CLEAR
cmbModels.CLEAR
While Not rsz.EOF
cmbMakes.AddItem rsz.Fields(0)
rsz.MoveNext
Wend
If cmbMakes.ListCount <> 0 Then
cmbMakes.Text = cmbMakes.List(0)
Else
cmbModels.Text = ""
End If
rsz.Close
End If

If ManiExtras1.FileExists(App.Path & "\common\images\Carmanufacturers\" & cmbManufacturer.Text & "\logo.jpg") Then
imgCarManu.Picture = LoadPicture(App.Path & "\common\images\Carmanufacturers\" & cmbManufacturer.Text & "\logo.jpg")
Else
imgCarManu.Picture = Nothing
End If

End Sub


Private Sub cmbModels_KeyPress(KeyAscii As Integer)
Call ALNUMVALID(KeyAscii)
End Sub

Private Sub Form_Activate()
'On Error GoTo q
Call Get_Theme
Call Apply_Theme(Me, 3)
If Saved = True Then
Call LockTheControls(True)
loaded = True
Call Load_Records
End If

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error GoTo q
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


  'Cmd.CommandText = "CREATE TABLE MANUFACTURER(NAME VARCHAR2(25) NOT NULL,PHOTO VARCHAR2(100),PRIMARY KEY (NAME))"
End Sub

Private Sub Form_Load()
On Error GoTo q
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide
disp = False
rs1.Open "select * from carmodels order by name", Cn, adOpenDynamic
Saved = True
loaded = False
Timer1.Enabled = True
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Cars Models"
Me.Icon = LoadPicture(App.Path & "\Themes\Green\Icons\Tree\car.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\Themes\Green\Icons\Tree\car.ico")

cmbDisp.AddItem "Acceleration & Speed"
cmbDisp.AddItem "Engine"
cmbDisp.AddItem "Fuel & Suspension"
cmbDisp.AddItem "Steering & Wheels"
cmbDisp.AddItem "Tyres & Breakes"
cmbDisp.AddItem "Size & GearBox"

cmbDisp.Text = cmbDisp.List(0)

Call LoadCarTypes

Call Load_Manufacturers
Call DispFrames(1)
Call CLEAR
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub
Sub LoadCarTypes()
cmbTypes.AddItem "Small Cars"
cmbTypes.AddItem "Mid Segment Cars"
cmbTypes.AddItem "Premium & Large Cars"
cmbTypes.AddItem "Wagons"
cmbTypes.AddItem "Utility Vehicles"
cmbTypes.AddItem "Multi Utility Vehicles"
cmbTypes.AddItem "Sports Utility Vehicles"
cmbTypes.Text = cmbTypes.List(0)
End Sub
Sub Load_Manufacturers()
Dim rsz As New ADODB.Recordset
rsz.Open "select * from manufacturer order by name", Cn, adOpenStatic
While Not rsz.EOF
cmbManufacturer.AddItem rsz.Fields(0)
rsz.MoveNext
Wend
If cmbManufacturer.ListCount <> 0 Then
cmbManufacturer.Text = cmbManufacturer.List(0)
End If
rsz.Close
End Sub

Sub Load_Records()
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from carmodels"
Set rs = Cmd.Execute
If rs.BOF = True And rs.EOF = True Then
Call CLEAR
Call btnEnable("10000010000")
Else
rs.MoveFirst
Call DISPLAY(rs)
Call btnEnable("10110011111")
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error GoTo q
Call ShowButtons(Me, "000")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
End Sub

Private Sub cmbManufacturer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtYear.SetFocus
End If
Call NAMEVALID(KeyAscii)
End Sub

Private Sub imgCmbCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rtfNotes.SetFocus
Else
NAMEVALID (KeyAscii)
End If
End Sub




Private Sub imgLogo_Click()
If Not Saved Then
Call AddPhoto
End If
End Sub

Private Sub imgLogo_DblClick()
ImagePath = photoDisPath
If Not ImagePath = "" Then
frmImagePreview.Show vbModal
End If
End Sub

Private Sub imgRestore_Click()
ManiExtras1.DesktopIconsHide
ManiExtras1.TaskBarHide
Me.Left = 0
Me.Top = 0
End Sub
Private Sub imgRestore_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error GoTo q
Call ShowButtons(Me, "010")
q:
If Err.Number <> 0 Then
Call Handle_Error("Error", CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub

Private Sub imgClose_Click()
Call ExitPrg
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error GoTo q
Call ShowButtons(Me, "001")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub

Private Sub imgMin_Click()
'On Error GoTo q
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
'On Error GoTo q
Call ShowButtons(Me, "100")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub


Private Sub lblImgLogo_Click()
If Not Saved Then
Call AddPhoto
End If
End Sub
Sub AddPhoto()
lblImgLogo.Visible = False
CommonDialog1.DialogTitle = "Select an image to add"
CommonDialog1.DefaultExt = "jpg"
CommonDialog1.Filter = "Image Files (*.jpg)|*.jpg"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
'MsgBox CommonDialog1.Filename
If Not CommonDialog1.FileTitle = "" Then
imgLogo.Picture = LoadPicture(CommonDialog1.Filename)
End If
End Sub

Private Sub Timer1_Timer()
lbltime.Caption = Format(Now, "hh:mm:ss", vbUseSystemDayOfWeek)
End Sub
Sub btnEnable(btn As String)
'MsgBox Len(btn)
For i = 1 To Len(btn)
If Mid$(btn, i, 1) = "1" Then
Btns(i - 1).Enabled = True
Else
Btns(i - 1).Enabled = False
End If
Next
End Sub

Sub LockTheControls(btn As Boolean)

With Me
For iindex = 0 To .Controls.Count - 1
If .Controls(iindex).Tag = "1" Then
If (btn) Then
If (TypeOf .Controls(iindex) Is Text Or TypeOf .Controls(iindex) Is ComboBox Or TypeOf .Controls(iindex) Is RichTextBox) Then

                    .Controls(iindex).Locked = True
                    .Controls(iindex).BackColor = vbWhite

End If
If (TypeOf .Controls(iindex) Is MaskEdBox Or TypeOf .Controls(iindex) Is ImageCombo) Then

                    .Controls(iindex).Enabled = False
                    .Controls(iindex).BackColor = vbWhite

End If

Else
If (TypeOf .Controls(iindex) Is Text Or TypeOf .Controls(iindex) Is ComboBox Or TypeOf .Controls(iindex) Is RichTextBox) Then
                    .Controls(iindex).Locked = False
                    .Controls(iindex).BackColor = TextBackcolor


End If
If (TypeOf .Controls(iindex) Is MaskEdBox Or TypeOf .Controls(iindex) Is ImageCombo) Then

                    .Controls(iindex).Enabled = True
                    .Controls(iindex).BackColor = TextBackcolor

End If
End If

End If
Next
End With
End Sub

Private Sub txtAddress_GotFocus()
SendKeys "{HOME}+{END}"

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtCity.SetFocus
Else
'Call CHARVALID(KeyAscii)
End If

End Sub

Private Sub txtCity_Change()
txtCity.Text = StrConv(txtCity.Text, vbProperCase)
SendKeys "{END}"
End Sub

Private Sub txtCity_GotFocus()
SendKeys "{HOME}+{END}"
End Sub


Private Sub txtCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'imgCmbCountry.SetFocus
Else
Call NAMEVALID(KeyAscii)
End If

End Sub

Private Sub txtYear_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAddress.SetFocus
End If
Call NUMVALID(KeyAscii)
End Sub
Sub CLEAR()

If cmbManufacturer.ListCount <> 0 Then
cmbManufacturer.Text = cmbManufacturer.List(0)
End If

cmbMakes.CLEAR
cmbModels.Text = ""
cmbTypes.Text = cmbTypes.List(0)
cmbDisp.Text = cmbDisp.List(0)

txtPrice.Text = ""

For jk = 0 To 2
txtAcc(jk).Text = ""
txtTyres(jk).Text = ""
txtBig(jk).Text = ""
txtBreakes(jk).Text = ""
txtSteering(jk).Text = ""
txtWheels(jk).Text = ""
txtFuel(jk).Text = ""
txtEngine(jk).Text = ""
Next

For pk = 3 To 4
txtAcc(pk).Text = ""
txtBig(pk).Text = ""
txtEngine(pk).Text = ""
Next

txtSpeed(0).Text = ""
txtSpeed(1).Text = ""
txtGearBox.Text = ""
txtBig(5).Text = ""
txtEngine(5).Text = ""
txtFuel(3).Text = ""


imgLogo.Picture = Nothing
'imgCarManu.Picture = Nothing
End Sub
Sub DISPLAY(DRS As ADODB.Recordset)
disp = True
For i = 0 To 40
'MsgBox i & " : " & DRS.Fields(i)
Next
If Not IsNull(DRS.Fields(0)) Then
cmbModels.Text = DRS.Fields(0)
End If

cmbMakes.CLEAR
If Not IsNull(DRS.Fields(1)) Then
cmbMakes.AddItem CStr(DRS.Fields(1))
If cmbMakes.ListCount <> 0 Then
cmbMakes.Text = cmbMakes.List(0)
End If
End If

cmbManufacturer.CLEAR
Dim RSM As New ADODB.Recordset
RSM.Open "SELECT * FROM CARMAKES WHERE NAME='" & cmbMakes.Text & "'", Cn, adOpenStatic
If RSM.RecordCount <> 0 Then
RSM.MoveFirst
cmbManufacturer.AddItem RSM.Fields(1)
If cmbManufacturer.ListCount <> 0 Then
cmbManufacturer.Text = cmbManufacturer.List(0)
End If
End If
RSM.Close

cmbTypes.CLEAR
If Not IsNull(DRS.Fields(2)) Then
cmbTypes.AddItem DRS.Fields(2)
If cmbTypes.ListCount <> 0 Then
cmbTypes.Text = cmbTypes.List(0)
End If
End If

If Not IsNull(DRS.Fields(3)) Then
txtPrice.Text = CStr(DRS.Fields(3))
Else
txtPrice.Text = ""
End If

If Not IsNull(DRS.Fields(4)) Then
txtAcc(0).Text = CStr(DRS.Fields(4))
Else
txtAcc(0).Text = ""
End If

If Not IsNull(DRS.Fields(5)) Then
txtAcc(1).Text = CStr(DRS.Fields(5))
Else
txtAcc(1).Text = ""
End If

If Not IsNull(DRS.Fields(6)) Then
txtAcc(2).Text = CStr(DRS.Fields(6))
Else
txtAcc(2).Text = ""
End If

If Not IsNull(DRS.Fields(7)) Then
txtAcc(3).Text = CStr(DRS.Fields(7))
Else
txtAcc(3).Text = ""
End If

If Not IsNull(DRS.Fields(8)) Then
txtAcc(4).Text = CStr(DRS.Fields(8))
Else
txtAcc(4).Text = ""
End If

If Not IsNull(DRS.Fields(9)) Then
txtSpeed(0).Text = CStr(DRS.Fields(9))
Else
txtSpeed(0).Text = ""
End If

If Not IsNull(DRS.Fields(10)) Then
txtSpeed(1).Text = CStr(DRS.Fields(10))
Else
txtSpeed(1).Text = ""
End If

If Not IsNull(DRS.Fields(11)) Then
txtEngine(0).Text = CStr(DRS.Fields(11))
Else
txtEngine(0).Text = ""
End If

If Not IsNull(DRS.Fields(12)) Then
txtEngine(1).Text = CStr(DRS.Fields(12))
Else
txtEngine(1).Text = ""
End If

If Not IsNull(DRS.Fields(13)) Then
txtEngine(2).Text = CStr(DRS.Fields(13))
Else
txtEngine(2).Text = ""
End If

If Not IsNull(DRS.Fields(14)) Then
txtEngine(3).Text = CStr(DRS.Fields(14))
Else
txtEngine(3).Text = ""
End If

If Not IsNull(DRS.Fields(15)) Then
txtEngine(4).Text = CStr(DRS.Fields(15))
Else
txtEngine(4).Text = ""
End If

If Not IsNull(DRS.Fields(16)) Then
txtEngine(5).Text = CStr(DRS.Fields(16))
Else
txtEngine(5).Text = ""
End If

If Not IsNull(DRS.Fields(17)) Then
txtFuel(0).Text = CStr(DRS.Fields(17))
Else
txtFuel(0).Text = ""
End If

If Not IsNull(DRS.Fields(18)) Then
txtFuel(1).Text = CStr(DRS.Fields(18))
Else
txtFuel(1).Text = ""
End If

If Not IsNull(DRS.Fields(19)) Then
txtFuel(2).Text = CStr(DRS.Fields(19))
Else
txtFuel(2).Text = ""
End If

If Not IsNull(DRS.Fields(20)) Then
txtFuel(3).Text = CStr(DRS.Fields(20))
Else
txtFuel(3).Text = ""
End If

If Not IsNull(DRS.Fields(21)) Then
txtGearBox.Text = CStr(DRS.Fields(21))
Else
txtGearBox.Text = ""
End If

If Not IsNull(DRS.Fields(22)) Then
txtSuspension(0).Text = CStr(DRS.Fields(22))
Else
txtSuspension(0).Text = ""
End If

If Not IsNull(DRS.Fields(23)) Then
txtSuspension(1).Text = CStr(DRS.Fields(23))
Else
txtSuspension(1).Text = ""
End If

If Not IsNull(DRS.Fields(24)) Then
txtSteering(0).Text = CStr(DRS.Fields(24))
Else
txtSteering(0).Text = ""
End If

If Not IsNull(DRS.Fields(25)) Then
txtSteering(1).Text = CStr(DRS.Fields(25))
Else
txtSteering(1).Text = ""
End If

If Not IsNull(DRS.Fields(26)) Then
txtSteering(2).Text = CStr(DRS.Fields(26))
Else
txtSteering(2).Text = ""
End If

If Not IsNull(DRS.Fields(27)) Then
txtWheels(0).Text = CStr(DRS.Fields(27))
Else
txtWheels(0).Text = ""
End If

If Not IsNull(DRS.Fields(28)) Then
txtWheels(1).Text = CStr(DRS.Fields(28))
Else
txtWheels(1).Text = ""
End If

If Not IsNull(DRS.Fields(29)) Then
txtWheels(2).Text = CStr(DRS.Fields(29))
Else
txtWheels(2).Text = ""
End If


If Not IsNull(DRS.Fields(30)) Then
txtTyres(0).Text = CStr(DRS.Fields(30))
Else
txtTyres(0).Text = ""
End If

If Not IsNull(DRS.Fields(31)) Then
txtTyres(1).Text = CStr(DRS.Fields(31))
Else
txtTyres(1).Text = ""
End If

If Not IsNull(DRS.Fields(32)) Then
txtTyres(2).Text = CStr(DRS.Fields(32))
Else
txtTyres(2).Text = ""
End If

If Not IsNull(DRS.Fields(33)) Then
txtBreakes(0).Text = CStr(DRS.Fields(33))
Else
txtBreakes(0).Text = ""
End If

If Not IsNull(DRS.Fields(34)) Then
txtBreakes(1).Text = CStr(DRS.Fields(34))
Else
txtBreakes(1).Text = ""
End If

If Not IsNull(DRS.Fields(35)) Then
txtBreakes(2).Text = CStr(DRS.Fields(35))
Else
txtBreakes(2).Text = ""
End If

If Not IsNull(DRS.Fields(36)) Then
txtBig(0).Text = CStr(DRS.Fields(36))
Else
txtBig(0).Text = ""
End If

If Not IsNull(DRS.Fields(37)) Then
txtBig(1).Text = CStr(DRS.Fields(37))
Else
txtBig(1).Text = ""
End If

If Not IsNull(DRS.Fields(38)) Then
txtBig(2).Text = CStr(DRS.Fields(38))
Else
txtBig(2).Text = ""
End If

If Not IsNull(DRS.Fields(39)) Then
txtBig(3).Text = CStr(DRS.Fields(39))
Else
txtBig(3).Text = ""
End If

If Not IsNull(DRS.Fields(40)) Then
txtBig(4).Text = CStr(DRS.Fields(40))
Else
txtBig(4).Text = ""
End If


If Not IsNull(DRS.Fields(41)) Then
txtBig(5).Text = CStr(DRS.Fields(41))
Else
txtBig(5).Text = ""
End If

If MyFile.FileExists(App.Path & "\Common\Images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text & "\" & cmbModels.Text & ".jpg") Then
photoDisPath = App.Path & "\Common\Images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text & "\" & cmbModels.Text & ".jpg"
imgLogo.Picture = LoadPicture(App.Path & "\Common\Images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMakes.Text & "\Models\" & cmbModels.Text & "\" & cmbModels.Text & ".jpg")
Else
imgLogo.Picture = Nothing
photoDisPath = ""
End If

LockTheControls (True)

End Sub
Sub ExitPrg()
'On Error Resume Next
If Saved Then
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
frmInsuranceMain.Show
rs1.Close
Else

Call Handle_Error("Confirm", "Record Not Saved", "Current record not yet saved. Do you want to exit ?", "Information1.jpg", "information.ico", 2, 0)
frmMsgbox.cmdCancel.Caption = "&No"
frmMsgbox.cmdok.Caption = "&Yes"
frmMsgbox.Show vbModal
If MsgBOx_R_Value Then
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
frmInsuranceMain.Show
Else
Exit Sub
End If
End If
End Sub

Sub DispFrames(x As Integer)
For i = 0 To 5
If x = i Then
fraDesc(i).Visible = True
Else
fraDesc(i).Visible = False
End If
Next
End Sub
