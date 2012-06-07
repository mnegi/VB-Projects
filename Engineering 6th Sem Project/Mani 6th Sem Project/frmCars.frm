VERSION 5.00
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Object = "{1059D9DC-C88F-11D5-80C0-0050BA3C6E71}#2.0#0"; "XPtextbox.ocx"
Begin VB.Form frmCars 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmCars.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Height          =   1080
      Left            =   4170
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   1320
      Width           =   5040
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Navigation"
      Height          =   2670
      Left            =   9705
      TabIndex        =   23
      Tag             =   "1"
      Top             =   810
      Width           =   2175
      Begin ManoharButton.MyButton Btns 
         Height          =   375
         Index           =   7
         Left            =   150
         TabIndex        =   24
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
         MICON           =   "frmCars.frx":000C
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
         TabIndex        =   25
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
         MICON           =   "frmCars.frx":0028
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
         TabIndex        =   26
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
         MICON           =   "frmCars.frx":0044
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
         TabIndex        =   27
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
         MICON           =   "frmCars.frx":0060
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
      TabIndex        =   15
      Tag             =   "1"
      Top             =   4005
      Width           =   2175
      Begin ManoharButton.MyButton Btns 
         Height          =   375
         Index           =   0
         Left            =   150
         TabIndex        =   16
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
         MICON           =   "frmCars.frx":007C
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
         TabIndex        =   17
         Tag             =   "1"
         Top             =   945
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
         MICON           =   "frmCars.frx":0098
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
         TabIndex        =   18
         Tag             =   "1"
         Top             =   1550
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
         MICON           =   "frmCars.frx":00B4
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
         TabIndex        =   19
         Tag             =   "1"
         Top             =   2160
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
         MICON           =   "frmCars.frx":00D0
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
         TabIndex        =   20
         Tag             =   "1"
         Top             =   2770
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
         MICON           =   "frmCars.frx":00EC
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
         TabIndex        =   21
         Tag             =   "1"
         Top             =   3380
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
         MICON           =   "frmCars.frx":0108
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
         Left            =   120
         TabIndex        =   22
         Tag             =   "1"
         Top             =   3990
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
         MICON           =   "frmCars.frx":0124
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   5745
      Left            =   210
      TabIndex        =   7
      Tag             =   "1"
      Top             =   2760
      Width           =   9120
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   4740
         TabIndex        =   36
         Top             =   300
         Width           =   4335
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
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
            Left            =   1350
            TabIndex        =   45
            Tag             =   "1"
            Top             =   210
            Width           =   1590
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   4770
         Left            =   4740
         TabIndex        =   14
         Top             =   840
         Width           =   4335
         Begin VB.Label lblFuel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1995
            TabIndex        =   44
            Tag             =   "1"
            Top             =   4245
            Width           =   60
         End
         Begin VB.Label lblSteering 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1995
            TabIndex        =   43
            Tag             =   "1"
            Top             =   3705
            Width           =   60
         End
         Begin VB.Label lblWheels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1995
            TabIndex        =   42
            Tag             =   "1"
            Top             =   3150
            Width           =   60
         End
         Begin VB.Label lblBreakes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1995
            TabIndex        =   41
            Tag             =   "1"
            Top             =   2610
            Width           =   60
         End
         Begin VB.Label lblGears 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1995
            TabIndex        =   40
            Tag             =   "1"
            Top             =   2070
            Width           =   60
         End
         Begin VB.Label lblMaxTorque 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1995
            TabIndex        =   39
            Tag             =   "1"
            Top             =   1530
            Width           =   60
         End
         Begin VB.Label lblMaxPower 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1995
            TabIndex        =   38
            Tag             =   "1"
            Top             =   975
            Width           =   60
         End
         Begin VB.Label lblEngine 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1995
            TabIndex        =   37
            Tag             =   "1"
            Top             =   435
            Width           =   60
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuel (ltrs)"
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
            Left            =   135
            TabIndex        =   35
            Tag             =   "1"
            Top             =   4245
            Width           =   975
         End
         Begin VB.Label Label7 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   135
            TabIndex        =   34
            Tag             =   "1"
            Top             =   3699
            Width           =   825
         End
         Begin VB.Label Label6 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   135
            TabIndex        =   33
            Tag             =   "1"
            Top             =   3155
            Width           =   735
         End
         Begin VB.Label Label5 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   135
            TabIndex        =   32
            Tag             =   "1"
            Top             =   2611
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gears"
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
            Left            =   135
            TabIndex        =   31
            Tag             =   "1"
            Top             =   2067
            Width           =   585
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   135
            TabIndex        =   30
            Tag             =   "1"
            Top             =   1523
            Width           =   1230
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   135
            TabIndex        =   29
            Tag             =   "1"
            Top             =   979
            Width           =   1125
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   135
            TabIndex        =   28
            Tag             =   "1"
            Top             =   435
            Width           =   690
         End
      End
      Begin VB.ComboBox cmbManufacturer 
         BeginProperty Font 
            Name            =   "Renault MN"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   450
         Left            =   1365
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "1"
         Top             =   405
         Width           =   3255
      End
      Begin VB.ComboBox cmbModel 
         BeginProperty Font 
            Name            =   "Renault MN"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   450
         Left            =   1365
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "1"
         Top             =   1560
         Width           =   3255
      End
      Begin VB.ComboBox cmbMake 
         BeginProperty Font 
            Name            =   "Renault MN"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   450
         Left            =   1365
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "1"
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblVendor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
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
         TabIndex        =   13
         Tag             =   "1"
         Top             =   420
         Width           =   720
      End
      Begin VB.Image imgCar 
         Height          =   3075
         Left            =   180
         Stretch         =   -1  'True
         Top             =   2565
         Width           =   4455
      End
      Begin VB.Label lblMake 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
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
         Left            =   105
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label lblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1650
         Width           =   645
      End
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "1"
      Top             =   2032
      Width           =   2220
   End
   Begin ManoharExtras.ManiExtras ManiExtras1 
      Height          =   165
      Left            =   780
      TabIndex        =   2
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
   Begin XPTEXTBOX.text txtRegNo 
      Height          =   465
      Left            =   1770
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Driver Id  [ NOT NULL ] "
      Top             =   1320
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   820
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
   Begin VB.Label lblYear 
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
      Left            =   225
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1987
      Width           =   465
   End
   Begin VB.Label lblRegNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RegNo"
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
      Left            =   225
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1335
      Width           =   690
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
      TabIndex        =   1
      Top             =   405
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cars DataBase"
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
      Width           =   1470
   End
End
Attribute VB_Name = "frmCars"
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
''On Error GoTo q
Select Case Index

Case 0
'add new record
Call CLEAR
Saved = False
LockTheControls (False)
Dim YEAR As Integer
For YEAR = 1900 To 2100
cmbYear.AddItem YEAR
Next
cmbYear.Text = Format(Now, "yyyy")
''Call CLEAR

cmbManufacturer.CLEAR
cmbMake.CLEAR
cmbModel.CLEAR
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

Case 1
'save record
On Error GoTo HandleErr

Cmd.ActiveConnection = Cn
Cmd.CommandText = "insert into cars values('" & txtRegNo.Text & "','" & cmbModel.Text & "','" & cmbYear.Text & "','" & txtDesc.Text & "','" & cmbMake.Text & "')"
MsgBox Cmd.CommandText
Cmd.Execute
Saved = True
Call btnEnable("10110111111")
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
Cstring = "delete from cars where reg_no='" & txtRegNo.Text & "'"
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
txtRegNo.Locked = True
txtRegNo.BackColor = vbWhite

Case 4
'Update
Cmd.ActiveConnection = Cn

Cstring = "update cars set MODEL ='" & cmbModel.Text & "',YEAR='" & cmbYear.Text & "',CARDESC='" & txtDesc.Text & "',MAKE='" & cmbMake.Text & "'"
Cmd.CommandText = Cstring
MsgBox Cmd.CommandText
Cmd.Execute

Saved = True
rs1.Requery
Call Load_Records
rs1.Requery

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

Private Sub cmbMake_Change()
Call DisplayCar
DispCarInfo
End Sub

Private Sub cmbMake_Click()
Call Load_Models(cmbMake.Text)
Call DisplayCar
DispCarInfo
End Sub

Private Sub cmbManufacturer_Change()
If Saved = False Then
Call DisplayCar
DispCarInfo
End If
End Sub

Private Sub cmbManufacturer_Click()
If Saved = False Then
Call Load_Makes(cmbManufacturer.Text)
Call DisplayCar
DispCarInfo
End If
End Sub

Private Sub cmbModel_Change()
If Saved = False Then
Call DisplayCar
DispCarInfo
End If
End Sub

Private Sub cmbModel_Click()
If Saved = False Then
Call DisplayCar
DispCarInfo
End If
End Sub



Private Sub Form_Activate()
'On Error GoTo q
Call Get_Theme
Call Apply_Theme(Me, 3)

Call Load_Records
loaded = True

If Saved = True Then
Call LockTheControls(True)
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
'On Error GoTo q
rs1.Open "select * from cars order by REG_NO", Cn, adOpenDynamic
Saved = True
loaded = False
Timer1.Enabled = True
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Cars Models"
Me.Icon = LoadPicture(App.Path & "\Themes\Green\Icons\Tree\car.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\Themes\Green\Icons\Tree\car.ico")
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
End If


End Sub
Sub Load_Manufacturers()
cmbManufacturer.CLEAR
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
Sub Load_Makes(VENDOR As String)
cmbMake.CLEAR
Dim rsz As New ADODB.Recordset
rsz.Open "select * from carmakes where vendor='" & VENDOR & "'", Cn, adOpenStatic
While Not rsz.EOF
cmbMake.AddItem rsz.Fields(0)
rsz.MoveNext
Wend
If cmbMake.ListCount <> 0 Then
cmbMake.Text = cmbMake.List(0)
End If
rsz.Close
End Sub
Sub Load_Models(MAKE As String)
cmbModel.CLEAR
Dim rsz As New ADODB.Recordset
rsz.Open "select * from carmodels where make='" & MAKE & "'", Cn, adOpenStatic
While Not rsz.EOF
cmbModel.AddItem rsz.Fields(0)
rsz.MoveNext
Wend
If cmbModel.ListCount <> 0 Then
cmbModel.Text = cmbModel.List(0)
End If
rsz.Close
End Sub

Sub Load_Records()
Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from cars"
Set rs = Cmd.Execute
If rs.BOF = True And rs.EOF = True Then
Call CLEAR
Call btnEnable("10000010000")
Else
rs.MoveFirst
If loaded = True Then
Call DISPLAY(rs)
End If
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

Sub CLEAR()
txtRegNo.Text = ""
txtDesc.Text = ""
cmbYear.CLEAR
cmbManufacturer.CLEAR
cmbMake.CLEAR
cmbModel.CLEAR
imgCar.Picture = Nothing

End Sub
Sub DisplayCar()

If MyFile.FileExists(App.Path & "\Common\Images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMake.Text & "\Models\" & cmbModel.Text & "\" & cmbModel.Text & ".jpg") Then
photoDisPath = App.Path & "\Common\Images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMake.Text & "\Models\" & cmbModel.Text & "\" & cmbModel.Text & ".jpg"
imgCar.Picture = LoadPicture(App.Path & "\Common\Images\CarManufacturers\" & cmbManufacturer.Text & "\Makes\" & cmbMake.Text & "\Models\" & cmbModel.Text & "\" & cmbModel.Text & ".jpg")
Else
imgCar.Picture = Nothing
photoDisPath = ""
End If

End Sub
Sub DispCarInfo()

Dim rsz As New ADODB.Recordset
rsz.Open "select * from carmodels where make='" & cmbMake.Text & "' and name='" & cmbModel.Text & "'", Cn, adOpenStatic
If rsz.BOF = True And rsz.EOF = True Then
Call ClearDesc
Else
If Not IsNull(rsz.Fields("E_LAYOUT")) Then
lblEngine.Caption = rsz.Fields("E_LAYOUT")
Else
lblEngine.Caption = ""
End If
If Not IsNull(rsz.Fields("E_MAXPOWER")) Then
lblMaxPower.Caption = rsz.Fields("E_MAXPOWER")
Else
lblMaxPower.Caption = ""
End If
If Not IsNull(rsz.Fields("E_MAXTORQUE")) Then
lblMaxTorque.Caption = rsz.Fields("E_MAXTORQUE")
Else
lblMaxTorque.Caption = ""
End If
If Not IsNull(rsz.Fields("GEARBOX")) Then
lblGears.Caption = rsz.Fields("GEARBOX")
Else
lblGears.Caption = ""
End If
If Not IsNull(rsz.Fields("B_FRONT")) Then
lblBreakes.Caption = rsz.Fields("B_FRONT")
Else
lblBreakes.Caption = ""
End If
If Not IsNull(rsz.Fields("B_REAR")) Then
lblBreakes.Caption = lblBreakes.Caption & " / " & rsz.Fields("B_REAR")
End If
If Not IsNull(rsz.Fields("W_SSIZE")) Then
lblWheels.Caption = rsz.Fields("W_SSIZE")
Else
lblWheels.Caption = ""
End If
If Not IsNull(rsz.Fields("W_RSIZE")) Then
lblWheels.Caption = lblWheels.Caption & " / " & rsz.Fields("W_RSIZE")
End If
If Not IsNull(rsz.Fields("STR_TYPE")) Then
lblSteering.Caption = rsz.Fields("STR_TYPE")
Else
lblSteering.Caption = ""
End If
If Not IsNull(rsz.Fields("F_CAPACITY")) Then
lblFuel.Caption = rsz.Fields("F_CAPACITY")
Else
lblFuel.Caption = ""
End If
End If
rsz.Close
End Sub
Sub ClearDesc()
lblEngine.Caption = ""
lblMaxPower.Caption = ""
lblMaxTorque.Caption = ""
lblGears.Caption = ""
lblBreakes.Caption = ""
lblWheels.Caption = ""
lblSteering.Caption = ""
lblFuel.Caption = ""

End Sub

Sub DISPLAY(DRS As ADODB.Recordset)

If Not IsNull("REG_NO") Then
txtRegNo.Text = DRS.Fields("REG_NO")
End If

cmbModel.CLEAR
If Not IsNull(DRS.Fields("MODEL")) Then
cmbModel.AddItem CStr(DRS.Fields("MODEL"))
cmbModel.Text = cmbModel.List(0)
End If

cmbYear.CLEAR
If Not IsNull(DRS.Fields("YEAR")) Then
cmbYear.AddItem CStr(DRS.Fields("YEAR"))
cmbYear.Text = cmbYear.List(0)
End If

If Not IsNull(DRS.Fields("CARDESC")) Then
txtDesc.Text = CStr(DRS.Fields("CARDESC"))
Else
txtDesc.Text = ""
End If

cmbMake.CLEAR
If Not IsNull(DRS.Fields("MAKE")) Then
cmbMake.AddItem CStr(DRS.Fields("MAKE"))
cmbMake.Text = cmbMake.List(0)
End If

Dim RSM As New ADODB.Recordset
RSM.Open "select * from carmakes where name='" & cmbMake.Text & "'", Cn, adOpenStatic
cmbManufacturer.CLEAR
cmbManufacturer.AddItem RSM.Fields("VENDOR")
cmbManufacturer.Text = cmbManufacturer.List(0)
RSM.Close
Call DisplayCar
DispCarInfo


LockTheControls (True)
End Sub



