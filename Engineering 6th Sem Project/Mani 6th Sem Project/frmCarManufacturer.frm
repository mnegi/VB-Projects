VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmCarManufacturer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmCarManufacturer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6810
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbManufacturer 
      Height          =   315
      Left            =   2505
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Manufacture Name"
      Top             =   1215
      Width           =   6930
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3570
      Pattern         =   "*.ico*"
      TabIndex        =   30
      Top             =   150
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5475
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Navigation"
      Height          =   2670
      Left            =   9690
      TabIndex        =   29
      Tag             =   "1"
      Top             =   810
      Width           =   2175
      Begin ManoharButton.MyButton Btns 
         Height          =   375
         Index           =   7
         Left            =   135
         TabIndex        =   17
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
         MICON           =   "frmCarManufacturer.frx":000C
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
         TabIndex        =   18
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
         MICON           =   "frmCarManufacturer.frx":0028
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
         TabIndex        =   19
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
         MICON           =   "frmCarManufacturer.frx":0044
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
         TabIndex        =   20
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
         MICON           =   "frmCarManufacturer.frx":0060
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
      Height          =   4500
      Left            =   9690
      TabIndex        =   28
      Tag             =   "1"
      Top             =   3690
      Width           =   2175
      Begin ManoharButton.MyButton Btns 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
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
         MICON           =   "frmCarManufacturer.frx":007C
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
         Left            =   120
         TabIndex        =   10
         Tag             =   "1"
         Top             =   942
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
         MICON           =   "frmCarManufacturer.frx":0098
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
         Left            =   120
         TabIndex        =   14
         Tag             =   "1"
         Top             =   2166
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
         MICON           =   "frmCarManufacturer.frx":00B4
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
         Left            =   120
         TabIndex        =   15
         Tag             =   "1"
         Top             =   2778
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
         MICON           =   "frmCarManufacturer.frx":00D0
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
         Left            =   120
         TabIndex        =   11
         Tag             =   "1"
         Top             =   3390
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
         MICON           =   "frmCarManufacturer.frx":00EC
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
         Left            =   120
         TabIndex        =   13
         Tag             =   "1"
         Top             =   1554
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
         MICON           =   "frmCarManufacturer.frx":0108
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
         TabIndex        =   16
         Tag             =   "1"
         Top             =   4005
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
         MICON           =   "frmCarManufacturer.frx":0124
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
      Height          =   6255
      Left            =   345
      TabIndex        =   24
      Tag             =   "1"
      Top             =   1935
      Width           =   9120
      Begin VB.ListBox lstMakes 
         Height          =   3570
         Left            =   6720
         TabIndex        =   38
         Tag             =   "1"
         ToolTipText     =   "Makes"
         Top             =   2475
         Width           =   2115
      End
      Begin RichTextLib.RichTextBox rtfNotes 
         Height          =   1560
         Left            =   1740
         TabIndex        =   9
         Tag             =   "1"
         ToolTipText     =   "Description"
         Top             =   4440
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2752
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmCarManufacturer.frx":0140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageCombo imgCmbCountry 
         Height          =   420
         Left            =   1740
         TabIndex        =   8
         Tag             =   "1"
         ToolTipText     =   "Country"
         Top             =   3780
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   741
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPTEXTBOX.text txtYear 
         Height          =   420
         Left            =   1740
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "The yeat company started"
         Top             =   420
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   741
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   4
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
      Begin XPTEXTBOX.text txtFounder 
         Height          =   420
         Left            =   1740
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Founder' Name"
         Top             =   1095
         Width           =   4455
         _ExtentX        =   7858
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
      Begin XPTEXTBOX.text txtChairman 
         Height          =   420
         Left            =   1740
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "Company's Chairman"
         Top             =   1770
         Width           =   4455
         _ExtentX        =   7858
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
      Begin XPTEXTBOX.text txtAddress 
         Height          =   420
         Left            =   1740
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "Address "
         Top             =   2430
         Width           =   4455
         _ExtentX        =   7858
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
      Begin XPTEXTBOX.text txtCity 
         Height          =   420
         Left            =   1740
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "City"
         Top             =   3105
         Width           =   4455
         _ExtentX        =   7858
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
      Begin XPTEXTBOX.text txtWorth 
         Height          =   420
         Left            =   3690
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "Company's worth"
         Top             =   420
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   741
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   12
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
      Begin VB.Label lblImgLogo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click herto add Logo"
         Height          =   195
         Left            =   7065
         TabIndex        =   40
         ToolTipText     =   "Click Here to Add Logo"
         Top             =   1185
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RS/"
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
         Left            =   5865
         TabIndex        =   39
         Tag             =   "1"
         Top             =   420
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Car Makes"
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
         Left            =   6720
         TabIndex        =   37
         Tag             =   "1"
         Top             =   2085
         Width           =   1110
      End
      Begin VB.Label lblLogo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOGO"
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
         Left            =   6720
         TabIndex        =   36
         Tag             =   "1"
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         TabIndex        =   35
         Tag             =   "1"
         Top             =   4485
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Worth"
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
         Left            =   2925
         TabIndex        =   34
         Tag             =   "1"
         Top             =   420
         Width           =   780
      End
      Begin VB.Label lblCity 
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
         Left            =   150
         TabIndex        =   33
         Tag             =   "1"
         Top             =   3135
         Width           =   570
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         TabIndex        =   32
         Tag             =   "1"
         Top             =   2445
         Width           =   990
      End
      Begin VB.Label lblCounrty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
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
         TabIndex        =   31
         Tag             =   "1"
         Top             =   3810
         Width           =   960
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   6720
         Stretch         =   -1  'True
         ToolTipText     =   "Click Here to Add Logo"
         Top             =   750
         Width           =   2115
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Started"
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
         Left            =   150
         TabIndex        =   27
         Tag             =   "1"
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label lblFounder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Founder"
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
         Left            =   150
         TabIndex        =   26
         Tag             =   "1"
         Top             =   1095
         Width           =   990
      End
      Begin VB.Label lblChairman 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chairman"
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
         TabIndex        =   25
         Tag             =   "1"
         Top             =   1770
         Width           =   1110
      End
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   165
      Left            =   780
      TabIndex        =   22
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
   Begin VB.Label lblRegNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Car Manufacturer"
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
      TabIndex        =   23
      Tag             =   "1"
      Top             =   1185
      Width           =   1755
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
      TabIndex        =   21
      Top             =   405
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cars Manufacturers"
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
      Width           =   1965
   End
End
Attribute VB_Name = "frmCarManufacturer"
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


Private Sub Btns_Click(Index As Integer)
On Error GoTo q
Select Case Index

Case 0
'add new record
Saved = False
Call CLEAR
Call LockTheControls(False)
imgCmbCountry.SelectedItem = imgCmbCountry.ComboItems(76)
Call btnEnable("01000100000")
cmbManufacturer.SetFocus
cmbManufacturer.CLEAR
Call CarManufacturers(cmbManufacturer)
cmbManufacturer.Text = cmbManufacturer.List(0)
imgLogo.Picture = Nothing
lblImgLogo.Visible = True

Case 1
'save record

On Error GoTo HandleErr

Cmd.ActiveConnection = Cn
'Cmd.CommandText = "insert into manufacturer values('manohar',1000,100,'manohar','manohar','kaakda','delhi','india')"
Cmd.CommandText = "insert into manufacturer values('" & cmbManufacturer.Text & "'," & txtYear.Text & "," & txtWorth.Text & ",'" & txtFounder.Text & "','" & txtChairman.Text & "','" & txtAddress.Text & "','" & txtCity.Text & "','" & imgCmbCountry.Text & "')"
MsgBox Cmd.CommandText
Cmd.Execute


If ManiExtras1.Path_Exist(App.Path & "\common\images\CarManufacturers") Then
'do nothing
Else
MyFile.CreateFolder (App.Path & "\common\images\CarManufacturers")
End If

If ManiExtras1.Path_Exist(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text) Then
'do nothing
Else
MyFile.CreateFolder (App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text)
rtfNotes.SaveFile (App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\" & cmbManufacturer.Text & ".txt")
If CommonDialog1.Filename <> "" Then
Call ManiExtras1.Copy_File(CommonDialog1.Filename, App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\")
If CommonDialog1.FileTitle <> "Logo.jpg" Then
Call ManiExtras1.RenameFile(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\" & CommonDialog1.FileTitle, App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Logo.jpg")
End If
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
Cstring = "delete from manufacturer where name='" & cmbManufacturer.Text & "'"
Cmd.CommandText = Cstring
Cmd.Execute

MsgBox App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\"
If ManiExtras1.Path_Exist(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text) Then

MyFile.DeleteFolder App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text, True
MsgBox "deleted"
End If
'btnEnable ("1000001")
End If
rs1.Requery
Call CLEAR
Call Load_Records
Call LockTheControls(True)


Case 3
LockTheControls (False)
Saved = False
btnEnable ("0000110000")

Case 4
Cmd.ActiveConnection = Cn
'Dim Cstring As String
Cstring = "update manufacturer set year =" & txtYear.Text & ",worth=" & txtWorth.Text & ",founder='" & txtFounder.Text & "',chairman='" & txtChairman.Text & "',address='" & txtAddress.Text & "',city='" & txtCity.Text & "',country='" & imgCmbCountry.Text & "' where name='" & cmbManufacturer.Text & "'"
'",'" & txtFounder.Text & "','" & txtChairman.Text & "','" & txtAddress.Text & "','" & txtCity.Text & "','" & imgCmbCountry.Text & "')"
Cmd.CommandText = Cstring
MsgBox Cmd.CommandText
Cmd.Execute

If ManiExtras1.Path_Exist(App.Path & "\common\images\CarManufacturers") Then
'do nothing
Else
MyFile.CreateFolder (App.Path & "\common\images\CarManufacturers")
End If

If ManiExtras1.Path_Exist(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text) Then
'do nothing
Else
MyFile.CreateFolder (App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text)
rtfNotes.SaveFile (App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\" & cmbManufacturer.Text & ".txt")
Call ManiExtras1.Copy_File(CommonDialog1.Filename, App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\")
If CommonDialog1.FileTitle <> "Logo.jpg" Then
Call ManiExtras1.RenameFile(App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\" & CommonDialog1.FileTitle, App.Path & "\common\images\CarManufacturers\" & cmbManufacturer.Text & "\Logo.jpg")
End If
End If


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

Private Sub cmbManufacturer_LostFocus()
cmbManufacturer.Text = StrConv(cmbManufacturer.Text, vbProperCase)
End Sub

Private Sub Form_Activate()
On Error GoTo q
Call Get_Theme
Call Apply_Theme(Me, 3)
If loaded = False Then
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


  'Cmd.CommandText = "CREATE TABLE MANUFACTURER(NAME VARCHAR2(25) NOT NULL,PHOTO VARCHAR2(100),PRIMARY KEY (NAME))"
End Sub

Private Sub Form_Load()
On Error GoTo q
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide

rs1.Open "select * from manufacturer order by name", Cn, adOpenDynamic
Saved = True
loaded = False
Timer1.Enabled = True
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Cars Manufacturers"
Me.Icon = LoadPicture(App.Path & "\Themes\Green\Icons\Tree\car.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\Themes\Green\Icons\Tree\car.ico")

'Call btnEnable("1000001")

'ADD THE FILE OF COUNTRY FLAGS TO FILE1 FileListBox
File1.Path = App.Path & "\Common\Icons\Countries\"
'set height and width of imagelist
ImageList1.ImageHeight = 15
ImageList1.ImageWidth = 20
'add the image from FileList to imaglist
For j = 1 To File1.ListCount
ImageList1.ListImages.Add j, CStr(File1.List(j)), LoadPicture(App.Path & "\Common\Icons\Countries\" & File1.List(j - 1))
Next

'set the imagecombo's imagelist to imagelist1
imgCmbCountry.ImageList = ImageList1
Dim dname As String

'Add the text and images to imagecombo, use imagelist1
'Call imgCmbCountry.ComboItems.Add(1, "india", "India", 74)
For k = 1 To File1.ListCount
dname = ""
dname = File1.List(k - 1)
dname = Mid$(dname, 1, Len(dname) - 4)
Call imgCmbCountry.ComboItems.Add(k, CStr(File1.List(j)), StrConv(dname, vbProperCase), k)
Next

'set the selected item to india
'imgCmbCountry.SelectedItem = imgCmbCountry.ComboItems(76)


q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub
Sub Load_Records()
Cmd.ActiveConnection = Cn
Cmd.CommandText = "SELECT * FROM MANUFACTURER ORDER BY NAME"
'rs.CursorType = adOpenDynamic
Set rs = Cmd.Execute
'rs.CursorType = adOpenDynamic
'MsgBox rs.RecordCount
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
On Error GoTo q
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
Call ExitPrg
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

Private Sub txtChairman_Change()
txtChairman.Text = StrConv(txtChairman.Text, vbProperCase)
SendKeys "{END}"
End Sub

Private Sub txtChairman_GotFocus()
SendKeys "{HOME}+{END}"
End Sub


Private Sub txtChairman_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAddress.SetFocus
Else
Call NAMEVALID(KeyAscii)
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

Private Sub txtFounder_Change()
txtFounder.Text = StrConv(txtFounder.Text, vbProperCase)
SendKeys "{END}"
End Sub

Private Sub txtFounder_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtFounder_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtChairman.SetFocus
Else
Call NAMEVALID(KeyAscii)
End If
End Sub

Private Sub txtWorth_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtWorth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFounder.SetFocus
Else
Call NUMVALID(KeyAscii)
End If
End Sub

Private Sub txtYear_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtWorth.SetFocus
End If
Call NUMVALID(KeyAscii)
End Sub
Sub CLEAR()
cmbManufacturer.Text = ""
txtYear.Text = ""
txtWorth.Text = ""
txtFounder.Text = ""
txtChairman.Text = ""
txtAddress.Text = ""
txtCity.Text = ""
imgCmbCountry.Text = ""
imgLogo.Picture = Nothing
rtfNotes.Text = ""
lstMakes.CLEAR
End Sub
Sub DISPLAY(DRS As ADODB.Recordset)
If Not IsNull(DRS.Fields(0)) Then
cmbManufacturer.Text = DRS.Fields(0)
End If
If Not IsNull(DRS.Fields(1)) Then
txtYear.Text = CStr(DRS.Fields(1))
Else
txtYear.Text = ""
End If
If Not IsNull(DRS.Fields(2)) Then
txtWorth.Text = CStr(DRS.Fields(2))
Else
txtWorth.Text = ""
End If
If Not IsNull(DRS.Fields(3)) Then
txtFounder.Text = CStr(DRS.Fields(3))
Else
txtFounder.Text = ""
End If
If Not IsNull(DRS.Fields(4)) Then
txtChairman.Text = CStr(DRS.Fields(4))
Else
txtChairman.Text = ""
End If
If Not IsNull(DRS.Fields(5)) Then
txtAddress.Text = CStr(DRS.Fields(5))
Else
txtAddress.Text = ""
End If
If Not IsNull(DRS.Fields(6)) Then
txtCity.Text = CStr(DRS.Fields(6))
Else
txtCity.Text = ""
End If

For i = 1 To imgCmbCountry.ComboItems.Count
If imgCmbCountry.ComboItems(i).Text = DRS.Fields(7) Then
imgCmbCountry.SelectedItem = imgCmbCountry.ComboItems(i)
End If
Next

If ManiExtras1.Path_Exist(App.Path & "\common\images\carmanufacturers\" & DRS.Fields(0) & "\Logo.jpg") Then
imgLogo.Picture = LoadPicture(App.Path & "\common\images\carmanufacturers\" & DRS.Fields(0) & "\Logo.jpg")
Else
imgLogo.Picture = Nothing
End If
If ManiExtras1.Path_Exist(App.Path & "\common\images\carmanufacturers\" & DRS.Fields(0) & "\" & DRS.Fields(0) & ".TXT") Then

rtfNotes.Filename = App.Path & "\common\images\carmanufacturers\" & DRS.Fields(0) & "\" & DRS.Fields(0) & ".TXT"
End If

lstMakes.CLEAR
Dim rsD As New ADODB.Recordset
Cmd.CommandText = "SELECT * FROM CARMAKES WHERE VENDOR='" & cmbManufacturer.Text & "' ORDER BY NAME"
Set rsD = Cmd.Execute
While Not rsD.EOF
lstMakes.AddItem rsD.Fields(0)
rsD.MoveNext
Wend
rsD.Close

LockTheControls (True)
End Sub
Sub ExitPrg()
On Error Resume Next
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
