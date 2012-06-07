VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmCustomer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Tabs"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMenu 
      BackColor       =   &H00E0E0E0&
      Height          =   1890
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   2850
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Top             =   2010
         Width           =   105
      End
      Begin VB.Image imgFile 
         Height          =   285
         Index           =   3
         Left            =   135
         Stretch         =   -1  'True
         Top             =   1110
         Width           =   315
      End
      Begin VB.Image imgFile 
         Height          =   285
         Index           =   2
         Left            =   135
         Stretch         =   -1  'True
         Top             =   765
         Width           =   315
      End
      Begin VB.Image imgFile 
         Height          =   285
         Index           =   1
         Left            =   135
         Stretch         =   -1  'True
         Top             =   435
         Width           =   315
      End
      Begin VB.Label lblFileItem 
         BackColor       =   &H00C0E0FF&
         Caption         =   "       E&xit                   Ctrl+X"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   21
         Tag             =   "2"
         Top             =   1500
         Width           =   2730
      End
      Begin VB.Label lblFileItem 
         BackColor       =   &H00C0E0FF&
         Caption         =   "       &Cancel                Ctrl+C"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   20
         Tag             =   "2"
         Top             =   1155
         Width           =   2730
      End
      Begin VB.Label lblFileItem 
         BackColor       =   &H00C0E0FF&
         Caption         =   "       &Save                  Ctrl+S"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Tag             =   "2"
         Top             =   480
         Width           =   2730
      End
      Begin VB.Image imgFile 
         Height          =   285
         Index           =   0
         Left            =   150
         Stretch         =   -1  'True
         Top             =   150
         Width           =   315
      End
      Begin VB.Label lblFileItem 
         BackColor       =   &H80000016&
         Caption         =   "       &New                  Ctrl+N"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   13
         Tag             =   "2"
         Top             =   150
         Width           =   2730
      End
      Begin VB.Label lblFileItem 
         BackColor       =   &H00C0E0FF&
         Caption         =   "       Save &As             Ctrl+A"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   19
         Tag             =   "2"
         Top             =   810
         Width           =   2730
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00FDE7B3&
      Height          =   1560
      Index           =   1
      Left            =   1005
      TabIndex        =   27
      Top             =   705
      Visible         =   0   'False
      Width           =   2850
      Begin VB.TextBox txtEdit 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   1515
         Width           =   150
      End
      Begin VB.Label lblEditItem 
         BackColor       =   &H00FEFDD6&
         Caption         =   "       &Delete               Ctrl+D"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   32
         Tag             =   "2"
         Top             =   810
         Width           =   2730
      End
      Begin VB.Image imgEdit 
         Height          =   285
         Index           =   0
         Left            =   150
         Stretch         =   -1  'True
         Top             =   150
         Width           =   315
      End
      Begin VB.Image imgEdit 
         Height          =   285
         Index           =   1
         Left            =   135
         Stretch         =   -1  'True
         Top             =   450
         Width           =   315
      End
      Begin VB.Image imgEdit 
         Height          =   285
         Index           =   2
         Left            =   135
         Stretch         =   -1  'True
         Top             =   780
         Width           =   315
      End
      Begin VB.Image imgEdit 
         Height          =   285
         Index           =   3
         Left            =   135
         Stretch         =   -1  'True
         Top             =   1125
         Width           =   315
      End
      Begin VB.Label lblEditItem 
         BackColor       =   &H00FEFDD6&
         Caption         =   "       &Edit                   Ctrl+E"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   31
         Tag             =   "2"
         Top             =   150
         Width           =   2730
      End
      Begin VB.Label lblEditItem 
         BackColor       =   &H00FEFDD6&
         Caption         =   "       &Update              Ctrl+U"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   30
         Tag             =   "2"
         Top             =   465
         Width           =   2730
      End
      Begin VB.Label lblEditItem 
         BackColor       =   &H00FEFDD6&
         Caption         =   "       &Search               Ctrl+S"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   29
         Tag             =   "2"
         Top             =   1170
         Width           =   2730
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00FDE7B3&
      Height          =   855
      Index           =   3
      Left            =   2730
      TabIndex        =   39
      Top             =   705
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtPh 
         Height          =   285
         Left            =   285
         TabIndex        =   72
         Top             =   1005
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   150
         TabIndex        =   40
         Top             =   2010
         Width           =   105
      End
      Begin VB.Image imgReport 
         Height          =   285
         Index           =   0
         Left            =   135
         Stretch         =   -1  'True
         Top             =   150
         Width           =   315
      End
      Begin VB.Image imgReport 
         Height          =   285
         Index           =   1
         Left            =   135
         Stretch         =   -1  'True
         Top             =   435
         Width           =   315
      End
      Begin VB.Label lblReportItem 
         BackColor       =   &H00FEFDD6&
         Caption         =   "       &Display All                  Ctrl+N"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   42
         Tag             =   "2"
         Top             =   150
         Width           =   2730
      End
      Begin VB.Label lblReportItem 
         BackColor       =   &H00FEFDD6&
         Caption         =   "       &Print All                  Ctrl+S"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   41
         Tag             =   "2"
         Top             =   480
         Width           =   2730
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00FDE7B3&
      Height          =   885
      Index           =   2
      Left            =   1845
      TabIndex        =   34
      Top             =   705
      Visible         =   0   'False
      Width           =   3420
      Begin VB.TextBox txtPhoto 
         Height          =   285
         Left            =   135
         TabIndex        =   35
         Top             =   3705
         Width           =   150
      End
      Begin VB.Image imgPhoto 
         Height          =   285
         Index           =   5
         Left            =   135
         Stretch         =   -1  'True
         Top             =   465
         Width           =   360
      End
      Begin VB.Image imgPhoto 
         Height          =   285
         Index           =   0
         Left            =   150
         Stretch         =   -1  'True
         Tag             =   "2"
         Top             =   150
         Width           =   315
      End
      Begin VB.Label lblPhotoItem 
         BackColor       =   &H00FEFDD6&
         Caption         =   "       Add &Photo             Ctrl+E"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   36
         Tag             =   "2"
         Top             =   150
         Width           =   2985
      End
      Begin VB.Label lblPhotoItem 
         BackColor       =   &H00FEFDD6&
         Caption         =   "       &Remove Photo       Ctrl+R"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   38
         Tag             =   "2"
         Top             =   495
         Width           =   2985
      End
   End
   Begin VB.Frame fraShowAll 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   2865
      TabIndex        =   70
      Tag             =   "22121"
      Top             =   825
      Visible         =   0   'False
      Width           =   8985
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   7260
         Left            =   -45
         TabIndex        =   71
         Top             =   360
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   12806
         _Version        =   393216
         ForeColor       =   4210688
         BackColorSel    =   16777152
         ForeColorSel    =   4194304
         GridColor       =   14737632
         GridColorFixed  =   8421504
      End
      Begin VB.Image imgShowAllClose 
         Height          =   270
         Left            =   8640
         Stretch         =   -1  'True
         Top             =   30
         Width           =   285
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3660
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6045
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListTree 
      Left            =   5340
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   4035
      TabIndex        =   45
      Top             =   3150
      Width           =   7740
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   2085
         Left            =   0
         TabIndex        =   54
         Top             =   3360
         Width           =   8505
         Begin VB.ComboBox cmbStates 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1245
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "1"
            ToolTipText     =   "State"
            Top             =   285
            Width           =   2700
         End
         Begin VB.ComboBox cmbDist 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5010
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "1"
            ToolTipText     =   "District"
            Top             =   300
            Width           =   2700
         End
         Begin XPTEXTBOX.text txtCity 
            Height          =   420
            Left            =   1245
            TabIndex        =   9
            Tag             =   "1"
            ToolTipText     =   "City"
            Top             =   855
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   30
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
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
         Begin XPTEXTBOX.text txtPin 
            Height          =   420
            Left            =   1245
            TabIndex        =   11
            Tag             =   "1"
            ToolTipText     =   "PIN Code"
            Top             =   1395
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   6
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
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
         Begin XPTEXTBOX.text txtPhone 
            Height          =   420
            Left            =   4995
            TabIndex        =   12
            Tag             =   "1"
            ToolTipText     =   "Mobile/Phone"
            Top             =   1425
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   15
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
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
            Left            =   4995
            TabIndex        =   10
            Tag             =   "1"
            ToolTipText     =   "Address"
            Top             =   795
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   741
            FontName        =   "MS Serif"
            FontSize        =   9.75
            Locked          =   -1  'True
            MaxLength       =   30
            FontBold        =   -1  'True
            LineColor       =   11643476
            Text            =   ""
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
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
         Begin VB.Label lblDistt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distt"
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
            Left            =   4065
            TabIndex        =   60
            Tag             =   "1"
            Top             =   345
            Width           =   480
         End
         Begin VB.Label lblState 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State"
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
            Left            =   315
            TabIndex        =   59
            Tag             =   "1"
            Top             =   315
            Width           =   495
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
            Left            =   300
            TabIndex        =   58
            Tag             =   "1"
            Top             =   870
            Width           =   405
         End
         Begin VB.Label lblPin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P I N"
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
            Left            =   315
            TabIndex        =   57
            Tag             =   "1"
            Top             =   1470
            Width           =   510
         End
         Begin VB.Label lblPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            Left            =   4080
            TabIndex        =   56
            Tag             =   "1"
            Top             =   1440
            Width           =   765
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
            Left            =   4080
            TabIndex        =   55
            Tag             =   "1"
            Top             =   855
            Width           =   825
         End
      End
      Begin VB.ComboBox cmbGender 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1665
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "Gender"
         Top             =   2310
         Width           =   3435
      End
      Begin MSMask.MaskEdBox mskDob 
         Height          =   435
         Left            =   1665
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "Date Of Birth"
         Top             =   2850
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   767
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         Mask            =   "##-???-##"
         PromptChar      =   "_"
      End
      Begin XPTEXTBOX.text txtAge 
         Height          =   360
         Left            =   6795
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2895
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   635
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   25
         FontBold        =   -1  'True
         LineColor       =   11643476
         Text            =   ""
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
      Begin XPTEXTBOX.text txtDriverId 
         Height          =   435
         Left            =   1665
         TabIndex        =   1
         Tag             =   "1"
         ToolTipText     =   "Driver Id  [ NOT NULL ] "
         Top             =   165
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   767
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   10
         FontBold        =   -1  'True
         LineColor       =   11643476
         Text            =   ""
         BackColor       =   192
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
      Begin XPTEXTBOX.text txtFname 
         Height          =   435
         Left            =   1680
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "First Name [ NOT NULL ]"
         Top             =   705
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   767
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   20
         FontBold        =   -1  'True
         LineColor       =   11643476
         Text            =   ""
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
      Begin XPTEXTBOX.text txtMname 
         Height          =   435
         Left            =   1665
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "Middle Name"
         Top             =   1245
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   767
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   15
         FontBold        =   -1  'True
         LineColor       =   11643476
         Text            =   ""
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
      Begin XPTEXTBOX.text txtLname 
         Height          =   435
         Left            =   1665
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Last Name"
         Top             =   1770
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   767
         FontName        =   "MS Serif"
         FontSize        =   9.75
         Locked          =   -1  'True
         MaxLength       =   15
         FontBold        =   -1  'True
         LineColor       =   11643476
         Text            =   ""
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
         Caption         =   "Click Here To Add Picture"
         Height          =   195
         Left            =   5655
         TabIndex        =   69
         ToolTipText     =   "Click to add file, Double click to view"
         Top             =   1275
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lblDriverId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   120
         TabIndex        =   53
         Tag             =   "1"
         Top             =   195
         Width           =   1290
      End
      Begin VB.Label lblFName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         TabIndex        =   52
         Tag             =   "1"
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label lblDob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth"
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
         TabIndex        =   51
         Tag             =   "1"
         Top             =   2895
         Width           =   1335
      End
      Begin VB.Label lblLname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
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
         Left            =   135
         TabIndex        =   50
         Tag             =   "1"
         Top             =   1815
         Width           =   1065
      End
      Begin VB.Label lblMidName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mid Name"
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
         TabIndex        =   49
         Tag             =   "1"
         Top             =   1290
         Width           =   1035
      End
      Begin VB.Label lblGender 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         TabIndex        =   48
         Tag             =   "1"
         Top             =   2355
         Width           =   735
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Age"
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
         Left            =   5265
         TabIndex        =   47
         Tag             =   "1"
         Top             =   2895
         Width           =   1245
      End
      Begin VB.Image imgPhotoDisplay 
         BorderStyle     =   1  'Fixed Single
         Height          =   2325
         Left            =   5520
         Stretch         =   -1  'True
         ToolTipText     =   "Click to add file, Double click to view"
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   -615
      TabIndex        =   44
      Top             =   90
      Width           =   105
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   735
      Left            =   4575
      TabIndex        =   24
      Top             =   -15
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1296
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4125
      Top             =   15
   End
   Begin MSComctlLib.TreeView treeCustomer 
      Height          =   6075
      Left            =   150
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   2430
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   10716
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bell MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   150
      TabIndex        =   62
      Top             =   780
      Visible         =   0   'False
      Width           =   2685
      Begin VB.ComboBox cmbSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "2"
         ToolTipText     =   "Gender"
         Top             =   450
         Width           =   2685
      End
      Begin XPTEXTBOX.text txtSearch 
         Height          =   360
         Left            =   30
         TabIndex        =   18
         Tag             =   "2"
         ToolTipText     =   "First Name [ NOT NULL ]"
         Top             =   1215
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   635
         FontName        =   "MS Serif"
         FontSize        =   9.75
         MaxLength       =   20
         FontBold        =   -1  'True
         LineColor       =   11643476
         Text            =   ""
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
      Begin VB.Image imgSearchClose 
         Height          =   225
         Left            =   2385
         Stretch         =   -1  'True
         Top             =   135
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Text"
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
         Left            =   45
         TabIndex        =   64
         Tag             =   "2"
         Top             =   870
         Width           =   1185
      End
      Begin VB.Label lblSearchField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seach Field"
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
         Left            =   45
         TabIndex        =   63
         Tag             =   "2"
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.Frame fraTree 
      BackColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   135
      TabIndex        =   65
      Top             =   780
      Width           =   2700
      Begin VB.ComboBox cmbSortOrder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   15
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Tag             =   "2"
         ToolTipText     =   "Gender"
         Top             =   1200
         Width           =   2685
      End
      Begin VB.ComboBox cmbSort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "2"
         ToolTipText     =   "Gender"
         Top             =   450
         Width           =   2685
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sorting Oreder"
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
         Left            =   45
         TabIndex        =   67
         Tag             =   "2"
         Top             =   870
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By"
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
         Left            =   60
         TabIndex        =   66
         Tag             =   "2"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label lblDisplayName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manohar"
      BeginProperty Font 
         Name            =   "Ribbon131 Bd BT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   885
      Left            =   3615
      TabIndex        =   61
      Tag             =   " 7879"
      Top             =   2310
      Visible         =   0   'False
      Width           =   8145
   End
   Begin VB.Image imgWallpaper 
      Height          =   7695
      Left            =   2895
      Stretch         =   -1  'True
      Top             =   825
      Width           =   8985
   End
   Begin VB.Shape MenuShape 
      BorderColor     =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   2730
      Shape           =   4  'Rounded Rectangle
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label lblMenuName 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " &Photo"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003A3A3A&
      Height          =   270
      Index           =   2
      Left            =   1905
      TabIndex        =   37
      Tag             =   "2"
      Top             =   435
      Width           =   840
   End
   Begin VB.Shape MenuShape 
      BorderColor     =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1860
      Shape           =   4  'Rounded Rectangle
      Top             =   420
      Width           =   885
   End
   Begin VB.Shape MenuShape 
      BorderColor     =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1005
      Shape           =   4  'Rounded Rectangle
      Top             =   420
      Width           =   855
   End
   Begin VB.Label lblMenuName 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " &Edit"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003A3A3A&
      Height          =   270
      Index           =   1
      Left            =   1050
      TabIndex        =   33
      Tag             =   "2"
      Top             =   435
      Width           =   780
   End
   Begin VB.Shape MenuShape 
      BorderColor     =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   420
      Width           =   885
   End
   Begin VB.Label lblMenuName 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " &File"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003A3A3A&
      Height          =   270
      Index           =   0
      Left            =   165
      TabIndex        =   26
      Tag             =   "2"
      Top             =   435
      Width           =   840
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
      Left            =   135
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
      TabIndex        =   23
      Top             =   405
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Database"
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
      TabIndex        =   22
      Tag             =   "1"
      Top             =   75
      Width           =   1935
   End
   Begin VB.Label lblMenuName 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " &Reports"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003A3A3A&
      Height          =   255
      Index           =   3
      Left            =   2775
      TabIndex        =   43
      Tag             =   "2"
      Top             =   435
      Width           =   975
   End
End
Attribute VB_Name = "frmCustomer"
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

Dim formactivated As Boolean
Dim treenode As Node
Dim edit As Boolean

Dim menuSelected As Integer
Dim FileItemSelected As Integer
Dim EditItemSelected As Integer
Dim PhotoItemSelected As Integer
Dim ReportItemselected As Integer
Dim Saved As Boolean
Dim photoDisPath As String

Private Sub cmbSort_Click()
Call UpdateTreeView(cmbSort.Text, cmbSortOrder.Text)
Call MenuVisible("0000")
End Sub
Private Sub cmbSortOrder_Click()
Call UpdateTreeView(cmbSort.Text, cmbSortOrder.Text)
Call MenuVisible("0000")
End Sub

Private Sub Form_Click()
Call MenuVisible("0000")
Call NoMenuSelected
End Sub

Private Sub Form_DblClick()
ManiExtras1.DesktopIconsHide
ManiExtras1.TaskBarHide
Me.Left = 0
Me.Top = 0
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'check for ALT and other combinations
If Shift = 4 Then

    Select Case KeyCode
    
    'check if F is pressed
    Case 70
    menuSelected = 1
   Call MenuVisible("1000")
    FileItemSelected = 0
    txtFile.SetFocus
    
    'check if E is pressed
    Case 69
    menuSelected = 2
    Call MenuVisible("0100")
    EditItemSelected = 0
    txtEdit.SetFocus
    
    'check if P is pressed
    Case 80
    menuSelected = 3
    Call MenuVisible("0010")
    PhotoItemSelected = 0
    txtPhoto.SetFocus
    
    'CHECK IF R IS PRESSED
    Case 82
    menuSelected = 4
    Call MenuVisible("0001")
    ReportItemselected = 0
        
    End Select
    Exit Sub
End If

'check if Escape key is pressed
If KeyCode = 27 And menuSelected <> 0 Then
    
Call MenuVisible("0000")
        menuSelected = 0
        FileItemSelected = 0
        EditItemSelected = 0
        PhotoItemSelected = 0
        ReportItemselected = 0
    
   
End If

'check for right key
If KeyCode = 39 Then

If menuSelected < 4 Then
menuSelected = menuSelected + 1
Text2.SetFocus
Select Case menuSelected
Case 1
Call MenuVisible("1000")
Case 2
Call MenuVisible("0100")
Case 3
Call MenuVisible("0010")
Case 4
Call MenuVisible("0001")
End Select

Else
Call MenuVisible("1000")
menuSelected = 1
End If

End If


'check for left key
If KeyCode = 37 Then

If menuSelected > 1 Then
menuSelected = menuSelected - 1
Text2.SetFocus
Select Case menuSelected
Case 1
Call MenuVisible("1000")
Case 2
Call MenuVisible("0100")
Case 3
Call MenuVisible("0010")
Case 4
Call MenuVisible("0001")
End Select

Else
Call MenuVisible("0001")
menuSelected = 4
End If

End If


'check for down key
If KeyCode = 40 Then
Select Case menuSelected
Case 1
FileItemSelected = FileItemSelected + 1
If FileItemSelected < 6 Then
Select Case FileItemSelected
'check for file menu
Case 1
Call FilesubMenu("10000")
Case 2
Call FilesubMenu("01000")
Case 3
Call FilesubMenu("00100")
Case 4
Call FilesubMenu("00010")
Case 5
Call FilesubMenu("00001")
End Select
Else
FileItemSelected = 1
Call FilesubMenu("10000")
End If

'check for edit menu
Case 2
EditItemSelected = EditItemSelected + 1
If EditItemSelected < 5 Then
Select Case EditItemSelected
Case 1
Call EditsubMenu("1000")
Case 2
Call EditsubMenu("0100")
Case 3
Call EditsubMenu("0010")
Case 4
Call EditsubMenu("0001")
End Select
Else
EditItemSelected = 1
Call EditsubMenu("1000")
End If

'check for photo menu
Case 3
PhotoItemSelected = PhotoItemSelected + 1
If PhotoItemSelected < 3 Then
Select Case PhotoItemSelected
Case 1
Call PhotosubMenu("10")
Case 2
Call PhotosubMenu("01")
End Select
Else
PhotoItemSelected = 1
Call PhotosubMenu("10")
End If

Case 4
ReportItemselected = ReportItemselected + 1
If ReportItemselected < 3 Then
Select Case ReportItemselected
Case 1
Call ReportsubMenu("10")
Case 2
Call ReportsubMenu("01")
End Select
Else
ReportItemselected = 1
Call ReportsubMenu("10")
End If

End Select
End If

'check up key
If KeyCode = 38 Then
Select Case menuSelected
'check for file menu
Case 1
FileItemSelected = FileItemSelected - 1
If FileItemSelected > 0 Then
Select Case FileItemSelected
Case 1
Call FilesubMenu("10000")
Case 2
Call FilesubMenu("01000")
Case 3
Call FilesubMenu("00100")
Case 4
Call FilesubMenu("00010")
Case 5
Call FilesubMenu("00001")
End Select
Else
FileItemSelected = 5
Call FilesubMenu("00001")
End If

'check for edit menu

Case 2
EditItemSelected = EditItemSelected - 1
If EditItemSelected > 0 Then
Select Case EditItemSelected
Case 1
Call EditsubMenu("1000")
Case 2
Call EditsubMenu("0100")
Case 3
Call EditsubMenu("0010")
Case 4
Call EditsubMenu("0001")
End Select
Else
EditItemSelected = 5
Call EditsubMenu("0001")
End If

'check for photo menu

Case 3
PhotoItemSelected = PhotoItemSelected - 1
If PhotoItemSelected > 0 Then
Select Case PhotoItemSelected
Case 1
Call PhotosubMenu("10")
Case 2
Call PhotosubMenu("01")
End Select
Else
PhotoItemSelected = 3
Call PhotosubMenu("01")
End If

Case 4
ReportItemselected = ReportItemselected - 1
If ReportItemselected > 0 Then
Select Case ReportItemselected
Case 1
Call ReportsubMenu("10")
Case 2
Call ReportsubMenu("01")
End Select
Else
ReportItemselected = 3
Call ReportsubMenu("01")
End If

End Select
End If


'check for enter key

If KeyCode = 13 Then
Select Case menuSelected

'check for file menu
Case 1
If lblFileItem(FileItemSelected - 1).Enabled = True Then
Select Case FileItemSelected
Case 1
'MsgBox "selectd new"
Call AddNewRecord
Case 2
'MsgBox "selected save"
Call SaveRecord

Case 3
'MsgBox "selected save as"
Call SaveAs
Case 4
'MsgBox "selected cancel"
Call Cancel
Case 5
Call ExitPrg
End Select
End If
'check for edit menu
Case 2
If lblEditItem(EditItemSelected - 1).Enabled = True Then
Select Case EditItemSelected
Case 1
'MsgBox "selectd Edit"
Call EditRecord
Case 2
'MsgBox "selected update"
Call UpdateRecord
Case 3
'MsgBox "selected delete"
Call DeleteRecord
Case 4
'MsgBox "selected search"
Call SearchRecord
End Select
End If

'check for photo menu
Case 3
If lblPhotoItem(PhotoItemSelected - 1).Enabled = True Then
Select Case PhotoItemSelected
Case 1
'MsgBox "selectd Add photo"
Call AddPhoto
Case 2
'MsgBox "selected remove"
Call RemovePhoto
End Select
End If

Case 4
If lblReportItem(ReportItemselected - 1).Enabled = True Then
Select Case ReportItemselected
Case 1
'MsgBox "selectd DIAPLAY ALL"
Call lblReportItem_Click(0)
Case 2
'MsgBox "selected PRINT ALL"
Call lblReportItem_Click(1)
End Select
End If

End Select
End If

'check for CTRL keys

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
If Button = vbRightButton Then
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Move.cur")

'MousePointer = 15
Call ReleaseCapture
Call SendMessage(hWnd, &HA1, 2, 0&)
'MousePointer = 1
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")

End If
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
Private Sub Form_Activate()
'MsgBox formactivated
If Not formactivated Then
    Call LockTheControls(True)
End If
formactivated = True
'Set RS = db.OpenRecordset("select * from cust", dbOpenDynamic, dbExecDirect, dbOptimistic)
End Sub
Private Sub Form_Load()
On Error GoTo q
Saved = True
formactivated = False
Timer1.Enabled = True
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Insurance DataBase Launcher"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmPerson.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmPerson.ico")
imgWallpaper.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\Wallpapers\FrmPersonWallpaper.jpg")

imgSearchClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\window\c_up.jpg")

Call Get_Theme
Call Apply_Theme(Me, 3)
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide
imgListTree.ImageHeight = 20
imgListTree.ImageWidth = 20
imgListTree.ListImages.Add 1, "Open", LoadPicture(App.Path & "\Themes\" & theme & "\Icons\Tree\OpenFolder.ico")
imgListTree.ListImages.Add 2, "Close", LoadPicture(App.Path & "\Themes\" & theme & "\Icons\Tree\CloseFolder.ico")
imgListTree.ListImages.Add 3, "Male", LoadPicture(App.Path & "\Themes\" & theme & "\Icons\Tree\Male.ico")
imgListTree.ListImages.Add 4, "Female", LoadPicture(App.Path & "\Themes\" & theme & "\Icons\Tree\Female.ico")

Call FileEnable("10001")
Call EditEnable("0001")
Call PhotoEnable("00")

Cmd.ActiveConnection = Cn
'treePerson.ImageList = imgListTree




cmbSearch.AddItem "CUST_ID"
cmbSearch.AddItem "FNAME"
cmbSearch.AddItem "MNAME"
cmbSearch.AddItem "LNAME"
cmbSearch.AddItem "SEX"
cmbSearch.AddItem "DOB"
cmbSearch.AddItem "ADDRESS"
cmbSearch.AddItem "CITY"
cmbSearch.AddItem "DISTT"
cmbSearch.AddItem "STATE"
cmbSearch.AddItem "PIN"
cmbSearch.AddItem "PHONE"
cmbSearch.Text = cmbSearch.List(0)

cmbSort.AddItem "CUST_ID"
cmbSort.AddItem "FNAME"
cmbSort.AddItem "MNAME"
cmbSort.AddItem "LNAME"
cmbSort.AddItem "SEX"
cmbSort.AddItem "DOB"
cmbSort.AddItem "ADDRESS"
cmbSort.AddItem "CITY"
cmbSort.AddItem "DISTT"
cmbSort.AddItem "STATE"
cmbSort.AddItem "PIN"
cmbSort.AddItem "PHONE"
cmbSort.Text = cmbSort.List(0)

cmbSortOrder.AddItem "ASC"
cmbSortOrder.AddItem "DESC"
cmbSortOrder.Text = cmbSortOrder.List(0)


lblDisplayName.ForeColor = TextForecolor
Call UpdateTreeView(cmbSort.List(0), cmbSortOrder.List(0))

'error handle
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
'

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo q
'Call ShowButtons("000")
q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If
End Sub


Private Sub imgPhotoDisplay_Click()
If Not Saved Then
Call AddPhoto
End If
End Sub

Private Sub imgPhotoDisplay_DblClick()
If MyFile.FileExists(photoDisPath) Then
ImagePath = photoDisPath
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
Me.WindowState = 1
ManiExtras1.DesktopIconsShow
ManiExtras1.TaskBarShow

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

Private Sub imgSearchClose_Click()
fraSearch.Visible = False
fraTree.Visible = True
Call UpdateTreeView(cmbSort.Text, cmbSortOrder.Text)
End Sub

Private Sub imgShowAllClose_Click()
imgShowAllClose.Picture = Nothing
fraShowAll.Visible = False
End Sub

Private Sub imgWallpaper_Click()
Call MenuVisible("0000")

End Sub
Sub DeleteRecord()
On Error GoTo q
Cmd.ActiveConnection = Cn
Cmd.CommandText = "delete from cust where cust_id = " & txtDriverId.Text
Cmd.Execute
'MsgBox "deleted"
rs.Close
Call UpdateTreeView(cmbSort.Text, cmbSortOrder.Text)
Call CLEAR
MenuVisible ("0000")


q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If

End Sub
Sub SearchRecord()
fraTree.Visible = False
Call CLEAR
fraSearch.Visible = True
MenuVisible ("0000")
cmbSearch.SetFocus
End Sub

Private Sub lblEditItem_Click(Index As Integer)
lblDisplayName.Visible = False

Select Case Index
Case 0
Call EditRecord
Saved = False
Case 1
Call UpdateRecord
Saved = True
Case 2
Call DeleteRecord
Saved = True
Case 3
Call SearchRecord
End Select
txtSearch.Text = ""
EditItemSelected = 0
menuSelected = 0
EditEnable ("0100")
End Sub

Private Sub lblEditItem_mousemove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
EditItemSelected = Index + 1
Select Case Index
Case 0
Call EditsubMenu("1000")
Case 1
Call EditsubMenu("0100")
Case 2
Call EditsubMenu("0010")
Case 3
Call EditsubMenu("0001")
End Select

End Sub
Sub AddNewRecord()
'MsgBox "add new"
'RS.AddNew
'new
cmbStates.CLEAR
Call Load_States(cmbStates)
cmbStates.Text = cmbStates.List(0)
cmbGender.AddItem "Male", 0
cmbGender.AddItem "Female", 1
cmbGender.Text = cmbGender.List(0)

lblImgLogo.Visible = True
Call NoMenuSelected
FileEnable ("01010")
PhotoEnable ("10")
LockTheControls (False)
MenuVisible ("0000")
Call CLEAR
treeCustomer.Enabled = False
End Sub
Sub SaveRecord()
'on error resume next
treeCustomer.Enabled = True


'If IsUniqueRecord(txtDriverId.Text) = False Then
'Call Handle_Error("DUPLICATE RECORD", "DUPLICATING PRIMARY KEY", "A Record with this Driver Id do exists ,You can't save this with the same driver id. In order to change the existing use Edit method", "Information1.jpg", "information.ico", 1, 0)
'frmMsgbox.Show vbModal
'Exit Sub
'End If

CmdText = "insert into cust values(" & txtDriverId.Text & ",'" & txtFname.Text & "','" & txtMname.Text & "','" & txtLname.Text & "','" & cmbGender.Text & "','" & mskDob.Text & "','" & txtAddress.Text & "','" & txtCity.Text & "','" & cmbDist.Text & "','" & cmbStates.Text & "','" & txtPin.Text & "','" & txtPhone.Text & "','" & CommonDialog1.FileTitle & "')"
Cmd.CommandText = CmdText
'On Error GoTo duplicate
MsgBox Cmd.CommandText
Cmd.Execute
MsgBox "Inserted"


If Not MyFile.FolderExists(App.Path & "\Common\Images\MyPhotos") Then
    MyFile.CreateFolder (App.Path & "\Common\Images\MyPhotos")
End If

If MyFile.FileExists(App.Path & "\Common\Images\MyPhotos\" & CommonDialog1.FileTitle) Then
    'do nothing
Else
    MyFile.CopyFile CommonDialog1.Filename, App.Path & "\Common\Images\MyPhotos\" & CommonDialog1.FileTitle
End If

MsgBox "Inserted"

LockTheControls (True)
NoMenuSelected
Call UpdateTreeView(cmbSort.Text, cmbSortOrder.Text)
FileEnable ("10001")
MenuVisible ("0000")
Call CLEAR
duplicate:

If Err.Number <> 0 Then
Select Case Err.Number
Case 3146
Call Handle_Error("Error : " & CStr(Err.Number), "Duplicate Record", "A record with the Driver Id " & txtDriverId.Text & " already exists", "information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End Select
End If

End Sub
Sub SaveAs()
'txtDriverId.Text = ""
'txtDriverId.Locked = False
LockTheControls (False)
FileEnable ("01010")
PhotoEnable ("11")
treeCustomer.Enabled = True
End Sub
Sub Cancel()
LockTheControls (True)
NoMenuSelected
FileEnable ("10001")
MenuVisible ("0000")
treeCustomer.Enabled = True
'Call CLEAR
End Sub
Sub ExitPrg()
On Error Resume Next
If Saved Then
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
frmOrderMain.Show
rsD.Close
Else

Call Handle_Error("Confirm", "Record Not Saved", "Current record not yet saved. Do you want to exit ?", "Information1.jpg", "information.ico", 2, 0)
frmMsgbox.cmdCancel.Caption = "&No"
frmMsgbox.cmdok.Caption = "&Yes"
frmMsgbox.Show vbModal
If MsgBOx_R_Value Then
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
frmOrderMain.Show
Else
Exit Sub
End If
End If
End Sub
Private Sub lblFileItem_Click(Index As Integer)
lblDisplayName.Visible = False
menuSelected = 0
MenuVisible ("0000")
If lblFileItem(Index).Enabled = True Then
Select Case Index
Case 0
Saved = False
Call AddNewRecord
'save
Case 1
Call SaveRecord
Saved = True
'save as
Case 2
Call SaveAs
Saved = True
'cancel
Case 3
Call Cancel
Saved = True
Case 4
Call ExitPrg
End Select
End If
End Sub

Private Sub lblFileItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
FileItemSelected = Index + 1
Select Case Index
Case 0
Call FilesubMenu("10000")
Case 1
Call FilesubMenu("01000")
Case 2
Call FilesubMenu("00100")
Case 3
Call FilesubMenu("00010")
Case 4
Call FilesubMenu("00001")
End Select
End Sub

Private Sub lblImgLogo_Click()
If Not Saved Then
Call AddPhoto
End If
End Sub

Private Sub lblMenuName_Click(Index As Integer)
Select Case Index
Case 0
menuSelected = 1
Call MenuVisible("1000")
FileItemSelected = 0
txtFile.SetFocus
Case 1

menuSelected = 2
Call MenuVisible("0100")
EditItemSelected = 0
txtEdit.SetFocus

Case 2
menuSelected = 3
Call MenuVisible("0010")
PhotoItemSelected = 0
txtPhoto.SetFocus
Case 3
menuSelected = 4
Call MenuVisible("0001")
ReportItemselected = 0
txtPh.SetFocus

End Select
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
imgPhotoDisplay.Picture = LoadPicture(CommonDialog1.Filename)
photoDisPath = App.Path & "\common\images\MyPhotos\" & CommonDialog1.FileTitle
End If
End Sub
Sub RemovePhoto()
imgPhotoDisplay.Picture = Nothing
If ManiExtras1.FileExists(photoDisPath) Then
MyFile.DeleteFile (photoDisPath)
End If
photoDisPath = "Nothing"

End Sub
Private Sub lblPhotoItem_Click(Index As Integer)
menuSelected = 0
MenuVisible ("0000")
Select Case Index
Case 0
Call AddPhoto
Case 1
Call RemovePhoto
End Select
PhotoEnable ("10")
End Sub

Private Sub lblPhotoItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
PhotoItemSelected = Index + 1
Select Case Index
Case 0
Call PhotosubMenu("10")
Case 1
Call PhotosubMenu("01")
End Select
End Sub


Private Sub MyButton1_Click(Index As Integer)
General.Visible = True
End Sub

Private Sub MyButton2_Click()
home.Visible = True
End Sub

Private Sub MyButton3_Click()
Personal.Visible = True
End Sub



Private Sub lstSearch_Click()
'MsgBox lstSearch.SelCount
Call showRecord(lstSearch.Text)
End Sub
Sub showRecord(srchText As String)
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from cust where " & cmbSearch.Text & "='" & srchText & "'"
Dim rsz As New ADODB.Recordset
Set rsz = Cmd.Execute
Call DISPLAY(rsz)
End Sub

Private Sub lblReportItem_Click(Index As Integer)
menuSelected = 0
Select Case Index
Case 0
imgShowAllClose.Picture = LoadPicture(App.Path & "\Themes\" & theme & "\window\c_up.jpg")
fraShowAll.Visible = True
Dim rsD As New ADODB.Recordset
rsD.Source = "select * from cust order by CUST_ID"
rsD.Open , Cn, adOpenForwardOnly, adLockReadOnly
LoadRecordsetIntoGrid rsD, grd, True, True
Case 1
'print
'drsCustomerAll.Show

End Select
MenuVisible ("0000")
End Sub

Private Sub lblReportItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
ReportItemselected = Index + 1
Select Case Index
Case 0
Call ReportsubMenu("10")
Case 1
Call ReportsubMenu("01")
End Select

End Sub

Private Sub Timer1_Timer()
lbltime.Caption = Format(Now, "hh:mm:ss", vbUseSystemDayOfWeek)
End Sub

Private Sub cmbStates_Change()
Call Load_Districts(cmbDist, cmbStates.Text)
cmbDist.Text = cmbDist.List(0)
End Sub
Private Sub cmbStates_Click()
Call Load_Districts(cmbDist, cmbStates.Text)
If cmbDist.ListCount <> 0 Then
cmbDist.Text = cmbDist.List(0)
End If
End Sub



'*************************************************
'HERE ONWARDS CODE IS FOR USER DEFINED SUBROUTINES
'*************************************************


Sub CLEAR()
On Error Resume Next
With Screen.ActiveForm
For iindex = 0 To .Controls.Count - 1
If .Controls(iindex).Tag = "1" Then
If (TypeOf .Controls(iindex) Is Text) Then
    .Controls(iindex).Text = ""
End If
If (TypeOf .Controls(iindex) Is MaskEdBox) Then
    .Controls(iindex).Format = "DD-MMM-YY"
    .Controls(iindex).Mask = "##-???-##"
End If
End If
Next
End With
imgPhotoDisplay.Picture = Nothing
photoDisPath = "Nothing"
End Sub


Sub LockTheControls(btn As Boolean)
On Error Resume Next
With Screen.ActiveForm
For iindex = 0 To .Controls.Count - 1
If .Controls(iindex).Tag = "1" Then
If (btn) Then
If (TypeOf .Controls(iindex) Is Text Or TypeOf .Controls(iindex) Is ComboBox) Then
                    
                    .Controls(iindex).Locked = True
                    .Controls(iindex).BackColor = vbWhite
                       
End If
If (TypeOf .Controls(iindex) Is MaskEdBox) Then
                    
                    .Controls(iindex).Enabled = False
                    .Controls(iindex).BackColor = vbWhite
                       
End If

Else
If (TypeOf .Controls(iindex) Is Text Or TypeOf .Controls(iindex) Is ComboBox) Then
                    .Controls(iindex).Locked = False
                    .Controls(iindex).BackColor = TextBackcolor
                    
                       
End If
If (TypeOf .Controls(iindex) Is MaskEdBox) Then
                    
                    .Controls(iindex).Enabled = True
                    .Controls(iindex).BackColor = TextBackcolor
                       
End If
End If

End If
Next
End With
End Sub
Sub UpdateTreeView(a As String, b As String)
treeCustomer.ImageList = imgListTree
treeCustomer.Nodes.CLEAR

Set treenode = treeCustomer.Nodes.Add(, , "Customers", "Customer Database", "Close")
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from CUST order by " & a & " " & b
Set rs = Cmd.Execute

'rs.Open "select * from CUSTOMER order by CUST_ID", Cn
If rs.RecordCount <> 0 Then
'RS.MoveFirst

Dim listI As Integer
listI = 1
'Dim listnode As ListView
While Not rs.EOF
    If rs.Fields("SEX") = "Male" Then
        Set treenode = treeCustomer.Nodes.Add("Customers", tvwChild, "key" & rs.Fields("CUST_ID"), rs.Fields("CUST_ID"), "Male")
    Else
        Set treenode = treeCustomer.Nodes.Add("Customers", tvwChild, "key" & rs.Fields("CUST_ID"), rs.Fields("CUST_ID"), "Female")
    End If
rs.MoveNext
'Set listnode = ListView1.ListItems.Add()
'
'
'listI = listI + 1
Wend

Else
'
End If
End Sub

Public Sub DISPLAY(DRS As ADODB.Recordset)
    
    txtDriverId.Text = DRS.Fields("CUST_ID")
    txtFname.Text = DRS.Fields("FNAME")
    
    If IsNull(DRS.Fields("MNAME")) Then
    txtMname.Text = ""
    Else
    txtMname.Text = DRS.Fields("MNAME")
    End If
    
    If IsNull(DRS.Fields("LNAME")) Then
    txtLname.Text = ""
    Else
    txtLname.Text = DRS.Fields("LNAME")
    End If
    
    If IsNull(DRS.Fields("ADDRESS")) Then
    txtAddress.Text = ""
    Else
    txtAddress.Text = DRS.Fields("ADDRESS")
    End If
    
    If IsNull(DRS.Fields("CITY")) Then
    txtCity.Text = ""
    Else
    txtCity.Text = DRS.Fields("CITY")
    End If
    
    If IsNull(DRS.Fields("PIN")) Then
    txtPin.Text = ""
    Else
    txtPin.Text = DRS.Fields("PIN")
    End If
    
    If IsNull(DRS.Fields("PHONE")) Then
    txtPhone.Text = ""
    Else
    txtPhone.Text = DRS.Fields("PHONE")
    End If
    cmbGender.CLEAR
    cmbGender.AddItem DRS.Fields("SEX")
    If cmbGender.ListCount <> 0 Then
    cmbGender.Text = cmbGender.List(0)
    End If
    
    cmbStates.CLEAR
    cmbStates.AddItem DRS.Fields("STATE")
    If cmbStates.ListCount <> 0 Then
    cmbStates.Text = cmbStates.List(0)
    End If
    
    cmbDist.CLEAR
    cmbDist.AddItem DRS.Fields("STATE")
    If cmbDist.ListCount <> 0 Then
    cmbDist.Text = cmbDist.List(0)
    End If
    
    
    mskDob.Text = Format(DRS.Fields("DOB"), "DD-MMM-YY")
    txtAge.Text = CInt(DateDiff("d", Format(DRS.Fields("DOB"), "DD-MMM-YY"), Format(Now, "DD-MMM-YY")) / 365)
    
    If DRS.Fields("photo") = "Nothing" Then
        imgPhotoDisplay.Picture = Nothing
    Else
        'Call PhotoEnable("01")
        photoDisPath = App.Path & "\Common\Images\MyPhotos\" & DRS.Fields("photo")
        If ManiExtras1.FileExists(photoDisPath) Then
        imgPhotoDisplay.Picture = LoadPicture(photoDisPath)
        Else
        imgPhotoDisplay.Picture = Nothing
        'imgPhotoDisplay.Paint drs.Fields("photo")
        End If
    End If
        Call PhotoEnable("00")
   
End Sub




Private Sub treeCustomer_Click()
Call MenuVisible("0000")
fraShowAll.Visible = False
End Sub

Private Sub treeCustomer_Collapse(ByVal Node As MSComctlLib.Node)
Node.Image = "Close"
End Sub
'
'Private Sub treeCustomer_DblClick()
'End Sub

Private Sub treeCustomer_Expand(ByVal Node As MSComctlLib.Node)
Node.Image = "Open"
End Sub

Private Sub treeCustomer_NodeClick(ByVal Node As MSComctlLib.Node)

'***************************************
'TRYING TO UNCHECK ALL OTHER CHECKBOXES AS ONE IS CLICKED
'***************************************
'For i = 0 To treeCustomer.Nodes.Count - 1
'treeCustomer.Nodes(i).Checked = False
'Next
'Node.Checked = True
'
'***************************************
If Node.Text = "Customer Database" Then
Else
If Node.Checked = True Then
'rs.Close
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from cust where cust_id = " & Node.Text

'***************************************
'THIS IS ONE OTHER METHOD TO SPECIFY THE PARAMETERS TO THE COMMAND OBJECT
'Par.Type = adBSTR
'Par.Value = Node.Text
'Cmd.Parameters.Append Par
'***************************************
Set rs = Cmd.Execute
'rs.ActiveCommand = Cmd.


Call DISPLAY(rs)

'***************************************
'WHEN USING PARAMETERS THIS IS TO TRY DELETING THE PREVIOUS PARAMETERS
'Cmd.Parameters.Delete (Par)
'rs.Close
'***************************************

LockTheControls (True)
EditEnable ("1011")
FileEnable ("0011")
Saved = True
lblDisplayName.Visible = True

End If
End If
End Sub

'*****************************************
'CODE HERE ONWARDS IS FOR DATA VALIDATION
'*****************************************



Private Sub txtAddress_Change()
txtAddress.Text = StrConv(txtAddress.Text, vbProperCase)
SendKeys "{END}"
End Sub
Private Sub txtAddress_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
Call ALNUMVALID(KeyAscii)
If KeyAscii = 13 Then
txtPin.SetFocus
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
Call NAMEVALID(KeyAscii)
If KeyAscii = 13 Then
txtAddress.SetFocus
End If
End Sub
Private Sub txtDriverId_Change()
txtDriverId.Text = StrConv(txtDriverId.Text, vbUpperCase)
SendKeys "{end}"
End Sub
Private Sub txtDriverId_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub txtDriverId_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
txtFname.SetFocus
End If
Call NUMVALID(KeyAscii)
End Sub
'
'Private Sub txtDriverId_LostFocus()
'If Saved = False Then
'If IsUniqueRecord(txtDriverId.Text) = False Then
'Call Handle_Error("DUPLICATE RECORD", "DUPLICATING PRIMARY KEY", "A Record with this Driver Id do exists ,You can't save this with the same driver id. In order to change the existing use Edit method", "Information1.jpg", "information.ico", 1, 0)
'frmMsgbox.Show vbModal
'Exit Sub
'End If
'
'
'End If
'
'End Sub

Private Sub txtFname_Change()
txtFname.Text = StrConv(txtFname.Text, vbProperCase)
SendKeys "{END}"
lblDisplayName.Visible = True
lblDisplayName.Caption = ""
lblDisplayName.Caption = txtFname.Text & " " & txtMname.Text & " " & txtLname.Text

End Sub
Private Sub txtFname_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub txtFname_KeyPress(KeyAscii As Integer)
Call NAMEVALID(KeyAscii)
If (KeyAscii = 13) Or (KeyAscii = 32) Then
txtMname.SetFocus
End If
End Sub
Private Sub txtLNAME_Change()
txtLname.Text = StrConv(txtLname.Text, vbProperCase)
SendKeys "{END}"
lblDisplayName.Visible = True
lblDisplayName.Caption = ""
lblDisplayName.Caption = txtFname.Text & " " & txtMname.Text & " " & txtLname.Text
End Sub
Private Sub txtLNAME_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub txtLNAME_KeyPress(KeyAscii As Integer)
Call NAMEVALID(KeyAscii)
If KeyAscii = 13 Then
cmbGender.SetFocus
End If
End Sub
Private Sub txtMNAME_Change()
txtMname.Text = StrConv(txtMname.Text, vbProperCase)
SendKeys "{END}"
lblDisplayName.Visible = True
lblDisplayName.Caption = ""
lblDisplayName.Caption = txtFname.Text & " " & txtMname.Text & " " & txtLname.Text
End Sub
Private Sub txtMNAME_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub txtMNAME_KeyPress(KeyAscii As Integer)
Call NAMEVALID(KeyAscii)
If (KeyAscii = 13) Or (KeyAscii = 32) Then
txtLname.SetFocus
End If
End Sub
Private Sub txtPhone_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub txtPhone_KeyPress(KeyAscii As Integer)
Call PHVALID(KeyAscii)
'If KeyAscii = 13 Then
'btn(6).SetFocus
'End If
End Sub

Private Sub txtPin_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtPin_KeyPress(KeyAscii As Integer)
Call NUMVALID(KeyAscii)
If KeyAscii = 13 Then
txtPhone.SetFocus
End If
End Sub

Sub MenuVisible(menus As String)
'If menuSelected <> 0 Then
For i = 0 To Len(menus) - 1
If Mid$(menus, i + 1, 1) = "1" Then
fraMenu(i).BackColor = MenuFrameColor
fraMenu(i).Visible = True
lblMenuName(i).BackStyle = 1
lblMenuName(i).BackColor = MenuLabelBackColor
Else
fraMenu(i).Visible = False
'lblMenuName(I).BackColor = MenuFrameColor
lblMenuName(i).BackStyle = 0
End If
Next

For i = 0 To lblFileItem.Count - 1
lblFileItem(i).BackColor = MenuFrameColor
Next
For i = 0 To lblEditItem.Count - 1
lblEditItem(i).BackColor = MenuFrameColor
Next
For i = 0 To lblPhotoItem.Count - 1
lblPhotoItem(i).BackColor = MenuFrameColor
Next


'End If
End Sub
Sub FilesubMenu(submenus As String)
If menuSelected = 1 Then
For i = 0 To Len(submenus) - 1
If Mid$(submenus, i + 1, 1) = "1" Then
lblFileItem(i).BackColor = MenuLabelBackColor
Else
lblFileItem(i).BackColor = MenuFrameColor
End If
Next
End If
End Sub
Sub EditsubMenu(submenus As String)
If menuSelected = 2 Then
For i = 0 To Len(submenus) - 1
If Mid$(submenus, i + 1, 1) = "1" Then
lblEditItem(i).BackColor = MenuLabelBackColor
Else
lblEditItem(i).BackColor = MenuFrameColor
End If
Next
End If
End Sub
Sub PhotosubMenu(submenus As String)
If menuSelected = 3 Then
For i = 0 To Len(submenus) - 1
If Mid$(submenus, i + 1, 1) = "1" Then
lblPhotoItem(i).BackColor = MenuLabelBackColor
Else
lblPhotoItem(i).BackColor = MenuFrameColor
End If
Next
End If
End Sub
Sub ReportsubMenu(submenus As String)
If menuSelected = 4 Then
For i = 0 To Len(submenus) - 1
If Mid$(submenus, i + 1, 1) = "1" Then
lblReportItem(i).BackColor = MenuLabelBackColor
Else
lblReportItem(i).BackColor = MenuFrameColor
End If
Next
End If
End Sub
Sub FileEnable(submenus As String)
For i = 0 To Len(submenus) - 1
If Mid$(submenus, i + 1, 1) = "1" Then
lblFileItem(i).Enabled = True
Else
lblFileItem(i).Enabled = False
End If
Next
End Sub
Sub EditEnable(submenus As String)
For i = 0 To Len(submenus) - 1
If Mid$(submenus, i + 1, 1) = "1" Then
lblEditItem(i).Enabled = True
Else
lblEditItem(i).Enabled = False
End If
Next
End Sub
Sub PhotoEnable(submenus As String)
For i = 0 To Len(submenus) - 1
If Mid$(submenus, i + 1, 1) = "1" Then
lblPhotoItem(i).Enabled = True
Else
lblPhotoItem(i).Enabled = False
End If
Next
End Sub
Sub ReportEnable(submenus As String)
For i = 0 To Len(submenus) - 1
If Mid$(submenus, i + 1, 1) = "1" Then
lblReportItem(i).Enabled = True
Else
lblReportItem(i).Enabled = False
End If
Next
End Sub

Sub NoMenuSelected()
menuSelected = 0
FileItemSelected = 0
EditItemSelected = 0
PhotoItemSelected = 0
ReportItemselected = 0
End Sub

Sub EditRecord()
treeCustomer.Enabled = False
'MsgBox rs.Fields("CUST_ID")
'RS.Fields("CUST_ID")
cmbStates.CLEAR
Call Load_States(cmbStates)
cmbStates.Text = cmbStates.List(0)
cmbGender.CLEAR
cmbGender.AddItem "Male"
cmbGender.AddItem "Female"
cmbGender.Text = rs.Fields("SEX")

LockTheControls (False)
txtDriverId.Locked = True
FileEnable ("00010")
EditEnable ("0100")
MenuVisible ("0000")
NoMenuSelected
If rs.Fields("photo") = "Nothing" Then
Call PhotoEnable("10")
Else
Call PhotoEnable("01")
End If
End Sub

Sub UpdateRecord()
treeCustomer.Enabled = True
'********************************************
'USED THE DAO WAY
'Dim rs1 As Recordset
'Set rs1 = DB.OpenRecordset("select * from CUSTOMER where CUST_ID = '" & txtDriverId.Text & "' ", dbOpenDynamic, dbExecDirect, dbOptimistic)
''rs1.Delete
'MsgBox "deleted"
'rs1.AddNew
'MsgBox rs1.RecordCount
'rs1.Fields("CUST_ID") = txtDriverId.Text
'rs1.Fields("FNAME") = txtFname.Text
'rs1.Fields("MNAME") = txtMname.Text
'rs1.Fields("LNAME") = txtLname.Text
'rs1.Fields("SEX") = cmbGender.Text
'rs1.Fields("DOB") = mskDob.Text
'rs1.Fields("DISTT") = cmbDist.Text
'rs1.Fields("STATE") = cmbStates.Text
'rs1.Fields("CITY") = txtCity.Text
'rs1.Fields("ADDRESS") = txtAddress.Text
'rs1.Fields("PIN") = txtPin.Text
'rs1.Fields("PHONE") = txtPhone.Text
'rs1.Fields("PHOTO") = "Nothing"
'rs1.Update
'rs1.Close
'MsgBox "Updated"
'********************************************
If ManiExtras1.FileExists(App.Path & "\common\images\MyPhotos\" & CommonDialog1.FileTitle) Then
'do nothing
Else
Call ManiExtras1.Copy_File(CommonDialog1.Filename, App.Path & "\common\images\MyPhotos\")
End If

Cmd.ActiveConnection = Cn
Cmd.CommandText = "update cust set fname = '" & txtFname.Text & "',lname = '" & txtLname.Text & "',mname = '" & txtMname.Text & "',sex = '" & cmbGender.Text & "',dob = '" & mskDob.Text & "',distt = '" & cmbDist.Text & "',state = '" & cmbStates.Text & "',address = '" & txtAddress.Text & "',city = '" & txtCity.Text & "',pin = '" & txtPin.Text & "',phone = '" & txtPhone.Text & "',photo = '" & CommonDialog1.FileTitle & "' where cust_id = " & txtDriverId.Text
Cmd.Execute
MsgBox "updated"


LockTheControls (True)
FileEnable ("1011")
EditEnable ("1001")
MenuVisible ("0000")
NoMenuSelected
Call UpdateTreeView(cmbSort.Text, cmbSortOrder.Text)
End Sub
Private Sub txtSearch_Change()
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from cust where " & cmbSearch.Text & " like '" & txtSearch.Text & "%'"
Dim rs1 As New ADODB.Recordset
Set rs1 = Cmd.Execute

If rs1.EOF And rs1.EOF Then
Call CLEAR
Else
rs1.MoveFirst
Call DISPLAY(rs1)
End If

treeCustomer.ImageList = imgListTree
treeCustomer.Nodes.CLEAR


While Not rs1.EOF
    If rs1.Fields("SEX") = "Male" Then
        Set treenode = treeCustomer.Nodes.Add(, , rs1.Fields("CUST_ID"), rs1.Fields("CUST_ID"), "Male")
    Else
        Set treenode = treeCustomer.Nodes.Add(, , rs1.Fields("CUST_ID"), rs1.Fields("CUST_ID"), "Female")
    End If
rs1.MoveNext
Wend

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
End Sub

Private Function IsUniqueRecord(DId As String) As Boolean
Dim rsCheck As ADODB.Recordset
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from cust where cust_id=" & DId
Set rsCheck = Cmd.Execute
On Error GoTo F
rsCheck.MoveFirst
IsUniqueRecord = False
F:
If Err.Number <> 0 Then
IsUniqueRecord = True
End If
rsCheck.Close
End Function



