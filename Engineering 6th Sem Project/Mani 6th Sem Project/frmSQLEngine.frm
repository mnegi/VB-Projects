VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D64F7BDA-9F14-4DD9-978A-BAAF91C9935B}#2.0#0"; "ManiExtras.ocx"
Object = "{E47F144F-B2CF-4858-AC24-5BA4CC3E1B6A}#4.0#0"; "ManoharButton.ocx"
Begin VB.Form frmSQLEngine 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Insurance DataBase Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00FFFFC0&
   Icon            =   "frmSQLEngine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3885
      TabIndex        =   37
      Top             =   150
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Frame fraLab 
      Height          =   2610
      Left            =   7440
      TabIndex        =   27
      Tag             =   "2"
      Top             =   4365
      Visible         =   0   'False
      Width           =   4020
      Begin VB.ComboBox cmbQ 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1140
         Width           =   1770
      End
      Begin VB.ComboBox cmbExp 
         Height          =   315
         Left            =   105
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1155
         Width           =   1770
      End
      Begin ManoharButton.MyButton Btns 
         Height          =   360
         Index           =   8
         Left            =   105
         TabIndex        =   34
         Tag             =   "0"
         Top             =   2085
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "&Load"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777152
         BCOLO           =   12640511
         FCOL            =   4194304
         FCOLO           =   4210688
         MCOL            =   255
         MPTR            =   1
         MICON           =   "frmSQLEngine.frx":000C
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
         Height          =   360
         Index           =   9
         Left            =   1425
         TabIndex        =   35
         Tag             =   "0"
         Top             =   2100
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "&Execute"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777152
         BCOLO           =   12640511
         FCOL            =   4194304
         FCOLO           =   4210688
         MCOL            =   255
         MPTR            =   1
         MICON           =   "frmSQLEngine.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ManoharButton.MyButton cmdOk 
         Height          =   360
         Index           =   10
         Left            =   2730
         TabIndex        =   36
         Tag             =   "0"
         Top             =   2085
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
         BTYPE           =   3
         TX              =   "&Ok"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777152
         BCOLO           =   12640511
         FCOL            =   4194304
         FCOLO           =   4210688
         MCOL            =   255
         MPTR            =   1
         MICON           =   "frmSQLEngine.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Queries"
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
         Index           =   5
         Left            =   2085
         TabIndex        =   33
         Tag             =   "1"
         Top             =   765
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Experiments"
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
         Index           =   4
         Left            =   90
         TabIndex        =   32
         Tag             =   "1"
         Top             =   765
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DBMS Lab Queries"
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
         Index           =   3
         Left            =   60
         TabIndex        =   29
         Tag             =   "1"
         Top             =   195
         Width           =   3915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   285
      TabIndex        =   22
      Top             =   735
      Width           =   11415
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Manohar SQL Engine 1.0"
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
         Height          =   330
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Tag             =   "1"
         Top             =   255
         Width           =   11310
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6690
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   465
      Top             =   30
   End
   Begin VB.Frame fraSQLEngine 
      BackColor       =   &H00FFFFFF&
      Height          =   7110
      Left            =   285
      TabIndex        =   3
      Top             =   1365
      Width           =   11415
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Describe Table/View"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5235
         TabIndex        =   7
         Top             =   270
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Execute SQL Command"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   3585
      End
      Begin VB.Frame fraSql 
         BackColor       =   &H00FFFFFF&
         Height          =   6360
         Left            =   30
         TabIndex        =   8
         Top             =   750
         Width           =   11355
         Begin RichTextLib.RichTextBox txtCmd 
            Height          =   1125
            Left            =   210
            TabIndex        =   28
            Tag             =   "2"
            Top             =   210
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   1984
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmSQLEngine.frx":0060
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
         Begin VB.CheckBox chkSelect 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Highlight First Row"
            Height          =   360
            Left            =   3900
            TabIndex        =   10
            Tag             =   "1"
            Top             =   1440
            Width           =   2640
         End
         Begin ManoharButton.MyButton Btns 
            Height          =   300
            Index           =   0
            Left            =   7020
            TabIndex        =   11
            Tag             =   "0"
            Top             =   1425
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "&Execute"
            ENAB            =   0   'False
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
            MICON           =   "frmSQLEngine.frx":00E6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CheckBox chkAuto 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Autosize Columns"
            Height          =   360
            Left            =   1170
            TabIndex        =   9
            Tag             =   "1"
            Top             =   1425
            Width           =   2610
         End
         Begin ManoharButton.MyButton Btns 
            Height          =   300
            Index           =   2
            Left            =   7020
            TabIndex        =   19
            Tag             =   "0"
            Top             =   1815
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
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
            MICON           =   "frmSQLEngine.frx":0102
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
            Height          =   300
            Index           =   3
            Left            =   8437
            TabIndex        =   20
            Tag             =   "0"
            Top             =   1410
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "&Open"
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
            MICON           =   "frmSQLEngine.frx":011E
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
            Height          =   300
            Index           =   4
            Left            =   8437
            TabIndex        =   21
            Tag             =   "0"
            Top             =   1830
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
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
            MICON           =   "frmSQLEngine.frx":013A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   4020
            Left            =   210
            TabIndex        =   12
            Top             =   2205
            Visible         =   0   'False
            Width           =   10950
            _ExtentX        =   19315
            _ExtentY        =   7091
            _Version        =   393216
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            SelectionMode   =   1
            AllowUserResizing=   1
         End
         Begin ManoharButton.MyButton Btns 
            Height          =   300
            Index           =   6
            Left            =   9870
            TabIndex        =   25
            Tag             =   "0"
            Top             =   1410
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "&Lab Queries"
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
            MICON           =   "frmSQLEngine.frx":0156
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
            Height          =   300
            Index           =   7
            Left            =   9855
            TabIndex        =   26
            Tag             =   "0"
            Top             =   1830
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "&About"
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
            MICON           =   "frmSQLEngine.frx":0172
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin RichTextLib.RichTextBox rtfTableDesc 
            Height          =   4005
            Left            =   240
            TabIndex        =   17
            Tag             =   "1"
            Top             =   2220
            Visible         =   0   'False
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   7064
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmSQLEngine.frx":018E
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Output"
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
            Index           =   1
            Left            =   225
            TabIndex        =   13
            Tag             =   "1"
            Top             =   1440
            Width           =   690
         End
      End
      Begin VB.Frame fraDesc 
         BackColor       =   &H00FFFFFF&
         Height          =   6360
         Left            =   15
         TabIndex        =   14
         Top             =   750
         Visible         =   0   'False
         Width           =   11355
         Begin VB.ComboBox cmbTabs 
            Height          =   315
            Left            =   5190
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "1"
            Top             =   345
            Width           =   4305
         End
         Begin ManoharButton.MyButton Btns 
            Height          =   390
            Index           =   5
            Left            =   9555
            TabIndex        =   24
            Tag             =   "0"
            Top             =   315
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   688
            BTYPE           =   3
            TX              =   "&Drop Table"
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
            MICON           =   "frmSQLEngine.frx":0214
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin RichTextLib.RichTextBox rtfDTabs 
            Height          =   5070
            Left            =   180
            TabIndex        =   18
            Tag             =   "1"
            Top             =   1125
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   8943
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"frmSQLEngine.frx":0230
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select a table or view from the list to describe"
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
            Index           =   0
            Left            =   210
            TabIndex        =   16
            Tag             =   "1"
            Top             =   375
            Width           =   4485
         End
      End
   End
   Begin VB.Frame fraOther 
      BackColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   285
      TabIndex        =   4
      Top             =   1365
      Visible         =   0   'False
      Width           =   11415
      Begin ManoharButton.MyButton Btns 
         Height          =   420
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Tag             =   "1"
         Top             =   870
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   741
         BTYPE           =   3
         TX              =   "&Insurance DataBase"
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
         MICON           =   "frmSQLEngine.frx":02B6
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
      Caption         =   "Manohar SQL Engine 1.0"
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
      Width           =   2505
   End
End
Attribute VB_Name = "frmSQLEngine"
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


Dim RsC As New ADODB.Recordset
Dim DoSave As Boolean
Dim SaveFile As String
Dim Filename As String
Private Sub Btns_Click(Index As Integer)
On Error GoTo q

Select Case Index

Case 0

fraLab.Visible = False
If Not txtCmd.Text = "" Then
Call ExecuteSQL
End If

txtCmd.SetFocus

Case 2
CommonDialog1.Filename = ""
rtfTableDesc.Visible = False
CommonDialog1.DialogTitle = "Enter the file name."
CommonDialog1.DefaultExt = "txt"
CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowSave
If CommonDialog1.Filename <> "" Then
DoSave = True
SaveFile = CommonDialog1.FileTitle
Filename = CommonDialog1.Filename
Else
'do nothing
End If

txtCmd.SetFocus

Case 3
CommonDialog1.Filename = ""
CommonDialog1.DialogTitle = "Select a file to open"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
If CommonDialog1.Filename <> "" Then

frmTextEditor.rtfFile.LoadFile (CommonDialog1.Filename)
frmTextEditor.Show vbModal

End If

Case 4

CommonDialog1.Filename = ""
CommonDialog1.DialogTitle = "Select a file to delete"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen

If CommonDialog1.Filename <> "" Then
Call Handle_Error("Confirm File Delete", "COFIRM", "Are you sure to delete this file ?", "Information1.jpg", "information.ico", 2, 0)
frmMsgbox.cmdCancel.Caption = "&No"
frmMsgbox.cmdok.Caption = "&Yes"
frmMsgbox.Show vbModal
If MsgBOx_R_Value Then
MyFile.DeleteFile (CommonDialog1.Filename)
End If
End If
txtCmd.SetFocus

Case 5

Call Handle_Error("Confirm Drop Table", "COFIRM", "Are you sure to drop this table ?", "Information1.jpg", "information.ico", 2, 0)
frmMsgbox.cmdCancel.Caption = "&No"
frmMsgbox.cmdok.Caption = "&Yes"
frmMsgbox.Show vbModal
If MsgBOx_R_Value Then
Cmd.ActiveConnection = Cn
Cmd.CommandText = "drop table " & cmbTabs.Text
On Error GoTo H
Cmd.Execute
If DoSave Then
    Open Filename For Append As #4
    Print #4, "OUTPUT : " & "EXECUTED  :   " & Cmd.CommandText
    Close #4
End If
rtfDTabs.Text = rtfDTabs.Text & "OUTPUT : " & "EXECUTED  :   " & Cmd.CommandText
If MyFile.FileExists(App.Path & "\Tables\" & cmbTabs.Text & ".txt") Then
MyFile.DeleteFile (App.Path & "\Tables\" & cmbTabs.Text & ".txt")
End If
End If

H:
If Err.Number <> 0 Then
If DoSave Then
    Open Filename For Append As #4
    Print #4, "OUTPUT : " & Err.Number & " :   " & Err.Description
    Close #4
End If
rtfDTabs.Text = rtfDTabs.Text & "OUTPUT : " & Err.Number & " :   " & Err.Description
End If
Call Load_Tabs



Case 6
'lab queries
fraLab.Visible = True
cmbExp.CLEAR
For i = 1 To 5
    cmbExp.AddItem i
Next
cmbExp.Text = cmbExp.List(0)


Case 7
frmSplashSE.Show vbModal

Case 8
fraLab.Visible = False
rtfTableDesc.Visible = False
grd.Visible = False
txtCmd.LoadFile (App.Path & "\Lab Queries\" & cmbExp.Text & "-" & cmbQ.Text & ".txt")

Case 9
fraLab.Visible = False
txtCmd.LoadFile (App.Path & "\Lab Queries\" & cmbExp.Text & "-" & cmbQ.Text & ".txt")
Call Btns_Click(0)

End Select

q:
If Err.Number <> 0 Then
Call Handle_Error(CStr("Error : " & Err.Number), CStr("Error : " & Err.Number), CStr(Err.Description), "Information1.jpg", "information.ico", 1, 0)
frmMsgbox.Show vbModal
rtfTableDesc.Visible = True
grd.Visible = False
rtfTableDesc.Text = "ERROR :   " & Err.Number & "   -   " & Err.Description

End If

End Sub

Private Sub cmbExp_Change()


cmbQ.CLEAR
File1.Path = App.Path & "\Lab Queries\"

For i = 0 To File1.ListCount - 1
    If Mid(File1.List(i), 1, 2) = cmbExp.Text Then
    cmbQ.AddItem Mid(File1.List(i), 3, Len(File1.List(i)) - 6)
    End If
Next
If cmbQ.ListCount <> 0 Then
    cmbQ.Text = cmbQ.List(0)
End If

End Sub
Private Sub cmbExp_click()
cmbQ.CLEAR
File1.Path = App.Path & "\Lab Queries\"
For i = 0 To File1.ListCount - 1
    If Mid(File1.List(i), 1, 1) = cmbExp.Text Then
    cmbQ.AddItem Mid(File1.List(i), 3, Len(File1.List(i)) - 6)
    End If
Next
If cmbQ.ListCount <> 0 Then
    cmbQ.Text = cmbQ.List(0)
End If
End Sub

Private Sub cmbTabs_Click()
If MyFile.FileExists(App.Path & "\Tables\" & cmbTabs.Text & ".txt") Then
rtfDTabs.Visible = True
rtfDTabs.LoadFile (App.Path & "\Tables\" & cmbTabs.Text & ".txt")
Else
rtfDTabs.Text = ""
End If
End Sub

Private Sub cmdok_Click(Index As Integer)
fraLab.Visible = False
End Sub

Private Sub Form_Activate()
On Error GoTo q
Call Get_Theme
Call Apply_Theme(Me, 3)
ManiExtras1.MinimizeAll
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


  
End Sub

Private Sub Form_Load()
On Error GoTo q
ManiExtras1.MinimizeAll
ManiExtras1.TaskBarHide
ManiExtras1.DesktopIconsHide

DoSave = False

Timer1.Enabled = True
Screen.MousePointer = 99
Screen.MouseIcon = LoadPicture(App.Path & "\Themes\" & theme & "\Cursors\Arrow.cur")
Me.Caption = "Manohar SQL Engine 1.0"
Me.Icon = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")
SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, 800, 600, 20, 20), True
imgIcon.Picture = LoadPicture(App.Path & "\common\Icons\App Icons\FrmInsuranceMain.ico")

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

Private Sub grd_Click()
txtCmd.Text = "select * from " & grd.Text
Call ExecuteSQL

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
Close #4
Unload Me
frmMain.Show
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

Private Sub Option3_Click()
fraSql.Visible = True
fraDesc.Visible = False
txtCmd.SetFocus
End Sub

Private Sub Option4_Click()

fraSql.Visible = False
fraDesc.Visible = True

Call Load_Tabs

End Sub
Sub Load_Tabs()
Dim rsT As New ADODB.Recordset

Cmd.ActiveConnection = Cn
Cmd.CommandText = "select * from tab"
Set rsT = Cmd.Execute
cmbTabs.CLEAR
While Not rsT.EOF
cmbTabs.AddItem rsT.Fields(0)
rsT.MoveNext
Wend
rsT.Close

If cmbTabs.ListCount <> 0 Then
cmbTabs.Text = cmbTabs.List(0)
cmbTabs.SetFocus
End If

End Sub
Private Sub Timer1_Timer()
lbltime.Caption = Format(Now, "hh:mm:ss", vbUseSystemDayOfWeek)
End Sub

Private Sub txtCmd_Change()
If txtCmd.Text = "" Then
    Btns(0).Enabled = False
Else
    Btns(0).Enabled = True
   
End If
x = ReturnCommandName(LTrim(txtCmd.Text))

Select Case LCase(x)
'Case "select"
'rtfTableDesc.Visible = False
'grd.Visible = True


Case "insert"
    rtfTableDesc.Visible = True
    grd.Visible = False
    rtfTableDesc.Text = ""
    Y = ReturnTableName(LTrim(txtCmd.Text))
    If Y <> "" Then
  
    If MyFile.FileExists(App.Path & "\Tables\" & Y & ".txt") Then
    rtfTableDesc.LoadFile (App.Path & "\Tables\" & Y & ".txt")
    Else
    rtfTableDesc.Text = ""
    End If

    
    End If
    
'Case Else
'    rtfTableDesc.Visible = True
'    grd.Visible = False

End Select


End Sub

Private Sub txtCmd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{BS}"
Call Btns_Click(0)
End If
End Sub
Sub ExecuteSQL()

On Error GoTo handle
If DoSave Then
Open Filename For Append As #4
End If

Cmd.ActiveConnection = Cn
Cmd.CommandText = txtCmd.Text
If DoSave Then
    
    Print #4, " "
    Print #4, "Query : " & Cmd.CommandText
    Print #4, " "
End If
'Dim i, j, Count, times As Integer
'j = 0
'times = 0
'Count = 0
'For i = 1 To Len(txtCmd.Text)
'If Mid(txtCmd.Text, i, 1) = " " Then
'If times = 0 Then
'Count = j
'End If
'times = times + 1
'Else
'j = j + 1
'End If
'Next
x = ReturnCommandName(LTrim(txtCmd.Text))

'If (Mid(txtCmd.Text, 1, Count) = StrConv("select", vbLowerCase) Or Mid(txtCmd.Text, 1, Count) = StrConv("select", vbProperCase) Or Mid(txtCmd.Text, 1, Count) = StrConv("select", vbUpperCase)) Then
If x = UCase("select") Or x = LCase("select") Or x = StrConv("select", vbProperCase) Then
    Set RsC = Cmd.Execute
    grd.Visible = True
    rtfTableDesc.Visible = False
        If chkAuto.Value = 1 Then
            If chkSelect.Value = 1 Then
                LoadRecordsetIntoGrid RsC, grd, True, True
            Else
                LoadRecordsetIntoGrid RsC, grd, True, False
            End If
        Else
            If chkSelect.Value = 1 Then
                LoadRecordsetIntoGrid RsC, grd, False, True
            Else
                LoadRecordsetIntoGrid RsC, grd, False, False
            End If
        End If
   
    If DoSave = True Then
    
    Dim Y As Long
    Dim WriteHeader As String
    Dim WriteString As String
    Print #4, "OUTPUT   : "
'    Print #4, " "
    For Y = 0 To RsC.Fields.Count - 1
        WriteHeader = WriteHeader & RsC.Fields(Y).Name & "   |   "
    Next
    Print #4, "---------------------------------------------------------------------------------------------------------------------------------------------"
    Print #4, " "
    Print #4, Mid(WriteHeader, 1, Len(WriteHeader) - 4)
    Print #4, " "
    Print #4, "---------------------------------------------------------------------------------------------------------------------------------------------"
    Print #4, " "
    RsC.MoveFirst
    While Not RsC.EOF
    For Y = 0 To RsC.Fields.Count - 1
        WriteString = WriteString & RsC.Fields(Y).Value & "   |   "
    Next
    Print #4, Mid(WriteString, 1, Len(WriteString) - 4)
    Print #4, " "
    WriteString = ""
    RsC.MoveNext
    Wend
    Print #4, "---------------------------------------------------------------------------------------------------------------------------------------------"
    End If
    
    
Else
    'If (Mid(txtCmd.Text, 1, Count) = StrConv("create", vbLowerCase) Or Mid(txtCmd.Text, 1, Count) = StrConv("create", vbProperCase) Or Mid(txtCmd.Text, 1, Count) = StrConv("create", vbUpperCase)) Then
     If x = UCase("create") Or x = LCase("create") Or x = StrConv("create", vbProperCase) Then
        r = CreateTable(txtCmd.Text)
        
        

        If Mid(r, 1, 1) = "1" Then
            rtfTableDesc.Visible = True
            grd.Visible = False
            
            rtfTableDesc.LoadFile (App.Path & "\Tables\" & Mid(r, 2, Len(r)) & ".txt")
            rtfTableDesc.Text = rtfTableDesc.Text & "OUTPUT : TABLE  " & UCase(Mid(r, 2, Len(r))) & "  CREATED"
            
            If DoSave Then
                Print #4, "OUTPUT : TABLE  ' " & UCase(Mid(r, 2, Len(r))) & " '  CREATED"
            End If
        Else
            rtfTableDesc.Visible = True
            grd.Visible = False
            rtfTableDesc.Text = Mid(r, 2, Len(r))
          
            If DoSave Then
            Print #4, "OUTPUT : " & rtfTableDesc.Text
            End If
        End If
    Else
   '     If (Mid(txtCmd.Text, 1, Count) = StrConv("alter", vbLowerCase) Or Mid(txtCmd.Text, 1, Count) = StrConv("alter", vbProperCase) Or Mid(txtCmd.Text, 1, Count) = StrConv("alter", vbUpperCase)) Then
            If x = UCase("alter") Or x = LCase("alter") Or x = StrConv("alter", vbProperCase) Then
            rtfTableDesc.Visible = True
            grd.Visible = False
        
            rtfTableDesc.Text = "SORRY  :   " & txtCmd.Text & "   -   CAN 'T BE EXECUTED FROM HERE"
            If DoSave Then
            Print #4, "OUTPUT : " & rtfTableDesc.Text
            End If
        Else
            'If (Mid(txtCmd.Text, 1, Count) = StrConv("insert", vbLowerCase) Or Mid(txtCmd.Text, 1, Count) = StrConv("insert", vbProperCase) Or Mid(txtCmd.Text, 1, Count) = StrConv("insert", vbUpperCase)) Then
            If x = UCase("insert") Or x = LCase("insert") Or x = StrConv("insert", vbProperCase) Then
            rtfTableDesc.Visible = True
            grd.Visible = False
            txtCmd.Text = ""
            Cmd.Execute
            rtfTableDesc.Text = "OUTPUT  : " & "EXECUTED  :   " & Cmd.CommandText
            If DoSave Then
            Print #4, "OUTPUT : " & "EXECUTED  :   " & Cmd.CommandText
            End If
            
            Else
                If x = UCase("drop") Or x = LCase("drop") Or x = StrConv("drop", vbProperCase) Then
                    rtfTableDesc.Visible = True
                    grd.Visible = False
                    d = ReturnTableName(txtCmd.Text)
                    If MyFile.FileExists(App.Path & "\Tables\" & d & ".txt") Then
                        MyFile.DeleteFile (App.Path & "\Tables\" & d & ".txt")
                    End If
                    Cmd.Execute
                    rtfTableDesc.Text = "OUTPUT  : " & "EXECUTED  :   " & Cmd.CommandText
                    If DoSave Then
                    Print #4, "OUTPUT : " & "EXECUTED  :   " & Cmd.CommandText
                    End If
                Else
                If x = UCase("desc") Or x = LCase("desc") Or x = StrConv("desc", vbProperCase) Or x = UCase("describe") Or x = LCase("describe") Or x = StrConv("describe", vbProperCase) Then
                        rtfTableDesc.Visible = True
                        grd.Visible = False
        
                        rtfTableDesc.Text = "HELP :   " & txtCmd.Text & "   -   PLEASE TRY THE DESCRIBER TABLES/VIEW OPTION  INSTREAD OF 'desc/describe' COMMAND"
                        If DoSave Then
                        Print #4, "OUTPUT : " & rtfTableDesc.Text
                        End If
                Else
                        grd.Visible = False
                        rtfTableDesc.Visible = True
                
                        Cmd.Execute
                        rtfTableDesc.Text = "OUTPUT  : " & "EXECUTED  :   " & Cmd.CommandText
                        If DoSave Then
                        Print #4, "OUTPUT : " & "EXECUTED  :   " & Cmd.CommandText
                        End If
                End If
                End If
            End If
        End If
    
    End If
End If
Close #4
handle:

If Err.Number <> 0 Then
grd.Visible = False
rtfTableDesc.Visible = True
rtfTableDesc.Text = "ERROR :   " & Err.Number & "   -   " & Err.Description

If DoSave Then
Print #4, "Output : " & "ERROR :   " & Err.Number & "   -   " & Err.Description
Close #4
End If
End If


End Sub
