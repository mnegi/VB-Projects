VERSION 5.00
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmCatalog 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Craeting User Account - *"
   ClientHeight    =   5175
   ClientLeft      =   1740
   ClientTop       =   1995
   ClientWidth     =   8790
   Icon            =   "frmCatalog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   Picture         =   "frmCatalog.frx":000C
   ScaleHeight     =   5175
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   435
      Left            =   1395
      TabIndex        =   11
      Top             =   465
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   767
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   345
      Index           =   0
      Left            =   6720
      TabIndex        =   12
      Tag             =   "1"
      Top             =   2302
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":2F52
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
      Height          =   345
      Index           =   1
      Left            =   6720
      TabIndex        =   8
      Tag             =   "1"
      Top             =   2690
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":2F6E
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
      Height          =   345
      Index           =   3
      Left            =   6720
      TabIndex        =   13
      Tag             =   "1"
      Top             =   3078
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":2F8A
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
      Height          =   345
      Index           =   4
      Left            =   6720
      TabIndex        =   10
      Tag             =   "1"
      Top             =   3466
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":2FA6
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
      Height          =   345
      Index           =   5
      Left            =   6720
      TabIndex        =   9
      Tag             =   "1"
      Top             =   3854
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":2FC2
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
      Height          =   345
      Index           =   2
      Left            =   6720
      TabIndex        =   14
      Tag             =   "1"
      Top             =   4245
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":2FDE
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
      Height          =   345
      Index           =   7
      Left            =   6720
      TabIndex        =   15
      Tag             =   "1"
      Top             =   750
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":2FFA
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
      Height          =   345
      Index           =   8
      Left            =   6720
      TabIndex        =   16
      Tag             =   "1"
      Top             =   1138
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":3016
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
      Height          =   345
      Index           =   9
      Left            =   6720
      TabIndex        =   17
      Tag             =   "1"
      Top             =   1526
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":3032
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
      Height          =   345
      Index           =   10
      Left            =   6720
      TabIndex        =   18
      Tag             =   "1"
      Top             =   1914
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":304E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   195
      TabIndex        =   19
      Top             =   720
      Width           =   6450
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   2535
         TabIndex        =   6
         Tag             =   "1"
         ToolTipText     =   "Select Book Id from here"
         Top             =   2880
         Width           =   3435
      End
      Begin VB.ComboBox cmbCatagory 
         Height          =   315
         Left            =   2535
         TabIndex        =   5
         Tag             =   "1"
         Text            =   "cmbCatagory"
         ToolTipText     =   "Select Book Id from here"
         Top             =   2415
         Width           =   3435
      End
      Begin VB.ComboBox cmbPub 
         Height          =   315
         Left            =   2535
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Select the publisher ID"
         Top             =   1920
         Width           =   3435
      End
      Begin VB.ComboBox cmbAuthorId 
         Height          =   315
         Left            =   2535
         TabIndex        =   3
         Tag             =   "1"
         ToolTipText     =   "Select the author"
         Top             =   1455
         Width           =   3435
      End
      Begin XPTEXTBOX.text txtPrice 
         Height          =   420
         Left            =   2535
         TabIndex        =   7
         Tag             =   "1"
         ToolTipText     =   "Enter the price"
         Top             =   3360
         Width           =   3435
         _ExtentX        =   6059
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
      Begin XPTEXTBOX.text txtTitle 
         Height          =   420
         Left            =   2535
         TabIndex        =   2
         Tag             =   "1"
         ToolTipText     =   "Enter the book title"
         Top             =   885
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   741
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
      Begin XPTEXTBOX.text txtBookId 
         Height          =   420
         Left            =   2535
         TabIndex        =   1
         Tag             =   "1"
         ToolTipText     =   "Enter the book id"
         Top             =   300
         Width           =   3435
         _ExtentX        =   6059
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   345
         TabIndex        =   26
         Tag             =   "1"
         ToolTipText     =   "Select the year"
         Top             =   3360
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catagory ID"
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
         Left            =   345
         TabIndex        =   25
         Tag             =   "1"
         Top             =   2415
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher ID"
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
         Left            =   345
         TabIndex        =   24
         Tag             =   "1"
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author ID"
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
         Left            =   345
         TabIndex        =   23
         Tag             =   "1"
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   345
         TabIndex        =   22
         Tag             =   "1"
         ToolTipText     =   "Select the year"
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label lblRno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Title"
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
         Left            =   345
         TabIndex        =   21
         Tag             =   "1"
         Top             =   885
         Width           =   1065
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book ID"
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
         Left            =   345
         TabIndex        =   20
         Tag             =   "1"
         Top             =   300
         Width           =   855
      End
   End
   Begin ManoharButton.MyButton Btns 
      Height          =   345
      Index           =   6
      Left            =   6720
      TabIndex        =   27
      Tag             =   "1"
      Top             =   1140
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   609
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
      MICON           =   "frmCatalog.frx":306A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgRestore 
      Height          =   270
      Left            =   8085
      Stretch         =   -1  'True
      ToolTipText     =   "Restore Position"
      Top             =   90
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Left            =   165
      Stretch         =   -1  'True
      ToolTipText     =   "Application Icon"
      Top             =   75
      Width           =   315
   End
   Begin VB.Image imgMin 
      Height          =   270
      Left            =   7770
      Stretch         =   -1  'True
      ToolTipText     =   "Minimise"
      Top             =   90
      Width           =   285
   End
   Begin VB.Image imgClose 
      Height          =   270
      Left            =   8400
      Stretch         =   -1  'True
      ToolTipText     =   "Close"
      Top             =   90
      Width           =   285
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Catalog Database"
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
      Left            =   975
      TabIndex        =   0
      Tag             =   "1"
      Top             =   75
      Width           =   1710
   End
End
Attribute VB_Name = "frmCatalog"
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
Dim rs1 As New ADODB.Recordset

Private Sub Btns_Click(Index As Integer)
''On Error GoTo q
Select Case Index

Case 0
'add new record
Call CLEAR
Saved = False
LockTheControls (False)
Call LockTheControls(False)
Call btnEnable("01000100000")

Dim rsD As New ADODB.Recordset

cmbAuthorId.CLEAR
rsD.Open "select * from author", Cn, adOpenStatic
If Not (rsD.EOF = True And rsD.BOF = True) Then
rsD.MoveFirst
While Not rsD.EOF
cmbAuthorId.AddItem rsD.Fields(0)
rsD.MoveNext
Wend
rsD.Close
End If
If cmbAuthorId.ListCount <> 0 Then
    cmbAuthorId.Text = cmbAuthorId.List(0)
End If

cmbPub.CLEAR
rsD.Open "select * from publisher", Cn, adOpenStatic
If Not (rsD.EOF = True And rsD.BOF = True) Then
rsD.MoveFirst
While Not rsD.EOF
cmbPub.AddItem rsD.Fields(0)
rsD.MoveNext
Wend
rsD.Close
End If
If cmbPub.ListCount <> 0 Then
    cmbPub.Text = cmbPub.List(0)
End If


cmbCatagory.CLEAR
rsD.Open "select * from category", Cn, adOpenStatic
If Not (rsD.EOF = True And rsD.BOF = True) Then
rsD.MoveFirst
While Not rsD.EOF
cmbCatagory.AddItem rsD.Fields(0)
rsD.MoveNext
Wend
rsD.Close
End If
If cmbCatagory.ListCount <> 0 Then
    cmbCatagory.Text = cmbCatagory.List(0)
End If

cmbYear.CLEAR
For i = 1900 To 3000
    cmbYear.AddItem i
Next
cmbYear.Text = Format(Now, "yyyy")

Case 1
'save record
On Error GoTo HandleErr

Cmd.ActiveConnection = Cn
Cmd.CommandText = "insert into catalog values(" & txtBookId.Text & ",'" & txtTitle.Text & "'," & cmbAuthorId.Text & "," & cmbPub.Text & "," & cmbCatagory.Text & "," & cmbYear.Text & "," & txtPrice.Text & ")"
MsgBox Cmd.CommandText
Cmd.Execute
Saved = True
Call btnEnable("10110111111")
LockTheControls (True)
Call Load_Records
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
Cstring = "delete from catalog where book_id=" & txtBookId.Text
Cmd.CommandText = Cstring
MsgBox Cmd.CommandText
Cmd.Execute
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
txtBookId.Locked = True
txtBookId.BackColor = vbWhite


rsD.Open "select * from author", Cn, adOpenStatic
If Not (rsD.EOF = True And rsD.BOF = True) Then
rsD.MoveFirst
While Not rsD.EOF
cmbAuthorId.AddItem rsD.Fields(0)
rsD.MoveNext
Wend
rsD.Close
End If


rsD.Open "select * from publisher", Cn, adOpenStatic
If Not (rsD.EOF = True And rsD.BOF = True) Then
rsD.MoveFirst
While Not rsD.EOF
cmbPub.AddItem rsD.Fields(0)
rsD.MoveNext
Wend
rsD.Close
End If



rsD.Open "select * from category", Cn, adOpenStatic
If Not (rsD.EOF = True And rsD.BOF = True) Then
rsD.MoveFirst
While Not rsD.EOF
cmbCatagory.AddItem rsD.Fields(0)
rsD.MoveNext
Wend
rsD.Close
End If


For i = 1900 To 3000
    cmbYear.AddItem i
Next



Case 4
'Update
Cmd.ActiveConnection = Cn

Cstring = "update catalog set title ='" & txtTitle.Text & "',author_id=" & cmbAuthorId.Text & ",publish_id=" & cmbPub.Text & ",category_id=" & cmbCatagory.Text & ",year=" & cmbYear.Text & ",price=" & txtPrice.Text & " where book_id=" & txtBookId.Text
Cmd.CommandText = Cstring
MsgBox Cmd.CommandText
Cmd.Execute
Saved = True
rs1.Requery
Call Load_Records
Call LockTheControls(True)

Case 5
''cancel
Saved = True
LockTheControls (True)
Call btnEnable("10110111111")
Call Load_Records

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
Private Sub Form_Activate()
On Error GoTo q
Call Get_Theme
Call Apply_Theme(Me, 4)
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
Private Sub Form_DblClick()
Me.Left = 1740
Me.Top = 1995
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
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide

rs1.Open "select * from catalog", Cn, adOpenDynamic
Saved = True
loaded = False
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")

Me.Caption = "Catalog DataBase"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 588, 345, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmMain.ico")


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
Me.Left = 1740
Me.Top = 1995
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
Sub CLEAR()
txtBookId.Text = ""
txtTitle.Text = ""
cmbAuthorId.CLEAR
cmbPub.CLEAR
cmbCatagory.CLEAR
cmbYear.CLEAR
txtPrice.Text = ""
End Sub
Sub DISPLAY(DRS As ADODB.Recordset)

If Not IsNull(DRS.Fields(0)) Then
txtBookId.Text = DRS.Fields(0)
Else
txtBookId.Text = ""
End If

If Not IsNull(DRS.Fields(1)) Then
txtTitle.Text = DRS.Fields(1)
Else
txtTitle.Text = ""
End If

cmbAuthorId.CLEAR
If Not IsNull(DRS.Fields(2)) Then
cmbAuthorId.AddItem DRS.Fields(2)
cmbAuthorId.Text = cmbAuthorId.List(0)
End If

cmbPub.CLEAR
If Not IsNull(DRS.Fields(3)) Then
cmbPub.AddItem DRS.Fields(3)
cmbPub.Text = cmbPub.List(0)
End If

cmbCatagory.CLEAR
If Not IsNull(DRS.Fields(4)) Then
cmbCatagory.AddItem DRS.Fields(4)
cmbCatagory.Text = cmbCatagory.List(0)
End If

cmbYear.CLEAR
If Not IsNull(DRS.Fields(5)) Then
cmbYear.AddItem DRS.Fields(5)
cmbYear.Text = cmbYear.List(0)
End If

If Not IsNull(DRS.Fields(6)) Then
txtPrice.Text = DRS.Fields(6)
Else
txtPrice.Text = ""
End If



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

Sub Load_Records()
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from catalog"
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

Sub ExitPrg()
'On Error Resume Next
If Saved Then
ManiExtras1.TaskBarShow
ManiExtras1.DesktopIconsShow
Unload Me
frmBookMain.Show
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
frmBookMain.Show
Else
Exit Sub
End If
End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'
Else
Call NUMVALID(KeyAscii)
End If
End Sub

Private Sub txtBookId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTitle.SetFocus
Else
Call NUMVALID(KeyAscii)
End If
End Sub

Private Sub txtTitle_Change()
txtTitle.Text = StrConv(txtTitle.Text, vbProperCase)
SendKeys "{END}"
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbAuthorId.SetFocus
Else
Call ADDVALID(KeyAscii)
End If
End Sub
