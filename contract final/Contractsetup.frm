VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form Contractsys 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MC Setup"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9705
   Icon            =   "Contractsetup.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Contractsetup.frx":4F0A
   ScaleHeight     =   6660
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5910
      Left            =   0
      TabIndex        =   2
      Top             =   495
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10425
      _Version        =   196609
      PaneTree        =   "Contractsetup.frx":524C
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   3570
         Left            =   6975
         TabIndex        =   36
         Top             =   2310
         Width           =   2730
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "candcstaff"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   14
            Left            =   180
            TabIndex        =   44
            Top             =   3150
            Width           =   2355
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Candchead"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   13
            Left            =   180
            TabIndex        =   41
            Top             =   2340
            Width           =   2355
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Contract1AdvanceRent"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   12
            Left            =   1620
            TabIndex        =   39
            Top             =   720
            Width           =   870
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Contract1SecurityDep"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   11
            Left            =   1620
            TabIndex        =   37
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "C AND C Staff"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   135
            TabIndex        =   45
            Top             =   2880
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "C AND C HEAD"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   135
            TabIndex        =   43
            Top             =   2070
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Credit and Collection Setup"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   45
            TabIndex        =   42
            Top             =   1710
            Width           =   2640
         End
         Begin VB.Line Line1 
            X1              =   135
            X2              =   2610
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Advance Rent"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   135
            TabIndex        =   40
            Top             =   810
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "SecDep"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   38
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   30
         TabIndex        =   33
         Top             =   5175
         Width           =   6855
         Begin LVbuttons.LaVolpeButton cmd_op 
            Height          =   375
            Index           =   0
            Left            =   1980
            TabIndex        =   34
            Top             =   225
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Save"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            BCOL            =   14737632
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "Contractsetup.frx":52FE
            ALIGN           =   1
            IMGLST          =   "(None)"
            IMGICON         =   "(None)"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   0
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin LVbuttons.LaVolpeButton cmdDone 
            Height          =   375
            Left            =   3420
            TabIndex        =   35
            Top             =   225
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "&Done"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            BCOL            =   14737632
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "Contractsetup.frx":531A
            ALIGN           =   1
            IMGLST          =   "(None)"
            IMGICON         =   "(None)"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   0
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1410
         Left            =   30
         TabIndex        =   25
         Top             =   3675
         Width           =   6855
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "ControllerPosition"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   18
            Left            =   495
            TabIndex        =   53
            Top             =   135
            Width           =   1860
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "ControllerPlaceIssued"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   9
            Left            =   4815
            TabIndex        =   28
            Top             =   810
            Width           =   1770
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Controllerresnumber"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   8
            Left            =   495
            TabIndex        =   27
            Top             =   810
            Width           =   1770
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Controllername"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   7
            Left            =   2520
            TabIndex        =   26
            Top             =   135
            Width           =   3390
         End
         Begin MSComCtl2.DTPicker dtStart 
            DataField       =   "ControllerDateIssued"
            Height          =   375
            Index           =   2
            Left            =   2700
            TabIndex        =   29
            Top             =   810
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   21757953
            CurrentDate     =   36584
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Place Issued"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   5085
            TabIndex        =   32
            Top             =   585
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date Issued"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   2880
            TabIndex        =   31
            Top             =   585
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Manager Res. Cert No:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   270
            TabIndex        =   30
            Top             =   585
            Width           =   2145
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1275
         Left            =   30
         TabIndex        =   4
         Top             =   2310
         Width           =   6855
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MallPosition"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   17
            Left            =   540
            TabIndex        =   52
            Top             =   180
            Width           =   2310
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MallManagers"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   4
            Left            =   3285
            TabIndex        =   20
            Top             =   180
            Width           =   3345
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "managerresnumber"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   5
            Left            =   360
            TabIndex        =   19
            Top             =   765
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "ManagerPlaceIssued"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   6
            Left            =   4725
            TabIndex        =   18
            Top             =   765
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker dtStart 
            DataField       =   "ManagerDateIssued"
            Height          =   375
            Index           =   1
            Left            =   2610
            TabIndex        =   21
            Top             =   765
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CalendarForeColor=   0
            Format          =   21757953
            CurrentDate     =   36584
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Manager Res. Cert No:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   24
            Top             =   540
            Width           =   2145
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date Issued"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   2835
            TabIndex        =   23
            Top             =   540
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Place Issued"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   4995
            TabIndex        =   22
            Top             =   540
            Width           =   1245
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Company Information"
         Height          =   2190
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9675
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MallLocated"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   16
            Left            =   5760
            MaxLength       =   50
            TabIndex        =   48
            Top             =   630
            Width           =   2490
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MallName"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   15
            Left            =   5760
            MaxLength       =   16
            TabIndex        =   46
            Top             =   1035
            Width           =   2490
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "CompanyTIN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   1
            Left            =   1620
            MaxLength       =   16
            TabIndex        =   10
            Top             =   1035
            Width           =   2490
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Malladd"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   0
            Left            =   1620
            TabIndex        =   9
            Top             =   630
            Width           =   2490
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "CompanyRestNo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   2
            Left            =   6030
            TabIndex        =   8
            Top             =   1710
            Width           =   2310
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "CompanyPlace"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   3
            Left            =   1710
            TabIndex        =   7
            Top             =   1710
            Width           =   1905
         End
         Begin VB.ComboBox txtCombo 
            BackColor       =   &H00FFFFFF&
            DataField       =   "CompanyCode"
            Height          =   315
            Left            =   1620
            TabIndex        =   6
            Top             =   270
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "CompanyName"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   10
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   225
            Width           =   4425
         End
         Begin MSComCtl2.DTPicker dtStart 
            Bindings        =   "Contractsetup.frx":5336
            DataField       =   "companyDateIssued"
            Height          =   375
            Index           =   0
            Left            =   4005
            TabIndex        =   11
            Top             =   1710
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   21757953
            CurrentDate     =   36584
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mall Located"
            DataField       =   "CompanyTIN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   4500
            TabIndex        =   49
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mall Name"
            DataField       =   "CompanyTIN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   4680
            TabIndex        =   47
            Top             =   1125
            Width           =   1395
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Company TIN number"
            DataField       =   "CompanyTIN"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   17
            Top             =   1125
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Company ID"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   16
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mall Address "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   15
            Top             =   720
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Company Res. Cert. No :"
            DataField       =   "companyRestNo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1440
            TabIndex        =   14
            Top             =   1440
            Width           =   2400
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date Issued"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   4275
            TabIndex        =   13
            Top             =   1485
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Place Issued"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   6345
            TabIndex        =   12
            Top             =   1530
            Width           =   1245
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   6420
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "© 2005 MLV systems"
            TextSave        =   "© 2005 MLV systems"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      DataField       =   "Adv2word"
      Height          =   330
      Left            =   10305
      TabIndex        =   51
      Top             =   3870
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      DataField       =   "Sec2word"
      Height          =   330
      Left            =   10305
      TabIndex        =   50
      Top             =   3330
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Contract Setup Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   225
      TabIndex        =   1
      Top             =   45
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   0
      Picture         =   "Contractsetup.frx":5341
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9720
   End
End
Attribute VB_Name = "Contractsys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS3 As ADODB.Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub cmd_op_Click(Index As Integer)
If Index = 0 Then
  
   adoPrimaryRS2.UpdateBatch adAffectAll
End If
End Sub

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub Form_Load()
    ' STARTUP SUPPLIERS DATABASE CONNECTIONS
   ' Rights5_Add = 1
   ' Rights5_Edit = 1
   ' Rights5_Save = 1
    Reload_PrimaryRS
    strSQL3 = "SELECT * FROM Company ORDER BY [Company Code]"
    Database_Refresh 2
    If adoPrimaryRS3.RecordCount <> 0 Then
        adoPrimaryRS3.MoveFirst
        Do While Not adoPrimaryRS3.EOF
            txtCombo.AddItem IIf(IsNull(adoPrimaryRS3("Company Code")), "", adoPrimaryRS3("Company Code"))
            adoPrimaryRS3.MoveNext
        Loop
    End If
End Sub
Public Sub Database_Refresh(xMode As Integer)
    ' PRE-DATABASE CONNECTION WITH PARAMETERIZED SQL VARIABLES ATTACHED IN EVERY MODE
    'On Error Resume Next
        
    If xMode = 0 Then
        Set adoPrimaryRS = New ADODB.Recordset
        adoPrimaryRS.Open strSQL, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 1 Then
        Set adoPrimaryRS2 = New ADODB.Recordset
        adoPrimaryRS2.Open strSQL2, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 2 Then
        Set adoPrimaryRS3 = New ADODB.Recordset
        adoPrimaryRS3.Open strSQL3, db, adOpenStatic, adLockOptimistic
    End If
End Sub

Private Sub Reload_PrimaryRS()
    ' RELOADING DATA OBJECTS AND DATABASE CONNECTIONS
   ' On Error Resume Next
    Dim oText As TextBox, i
    Dim odate As DTPicker, e
    
    'strSQL2 = "SELECT [Company Code] AS CompanyCode, [Company TIN number] AS CompanyTIN, [Mall Managers] AS MallManagers," & _
    '         "Controller AS Controllername,[Mall Address] AS Malladd,[company resnumber] AS companyRestNo," & _
    '         "[Company Date Issued] AS CompanyDateissued, [Company Place Issued] AS CompanyPlace," & _
    '         "[manager resnumber] as managerresnumber ,[Manager Date Issued] as ManagerDateIssued," & _
    '         "[Manager Place Issued] as ManagerPlaceIssued, [Controller resnumber] as companyresnumber," & _
    '         "[Controller Date Issued] as ControllerDateIssued, [Controller Place Issued]as ControllerPlaceIssued FROM MCSetup "
    strSQL2 = "Select [Comp Code] as CompanyCode, " & _
                "[Mall Address] as Malladd, " & _
                "[Company TIN number] as CompanyTIN, " & _
                "[Mall Manager] as MallManagers," & _
                "[Company Place Issued] as CompanyPlace, " & _
                "[Manager resnumber] as managerresnumber, " & _
                "[Manager Date Issued] as ManagerDateIssued," & _
                "[Manager Place Issued] as ManagerPlaceIssued, " & _
                "[Controller] as Controllername, " & _
                "[Controller resnumber] as Controllerresnumber, " & _
                "[Controller Date Issued] as ControllerDateIssued, " & _
                "[Controller Place Issued] as ControllerPlaceIssued, " & _
                "[Company Date Issued] as companyDateIssued, " & _
                "[Contract1 Security Dep] as Contract1SecurityDep, [Contract1 Security Dep number to word] as Sec2word, " & _
                "[Contract1 Advance Rent] as Contract1AdvanceRent, [Contract1 Advance Rent number to word] as Adv2word, " & _
                "[Company Name] as CompanyName,[Mall Located] as Malllocated," & _
                "[C and C Head] as candchead, [Mall Position] as MallPosition, [Controller Position] as ControllerPosition," & _
                "[C and C staff] as candcstaff, [Mall Name] as MallName," & _
                "[company resnumber] as CompanyRestNo FROM MCSetup"
              
    Database_Refresh 1
      'Set Text1(0).DataSource = adoPrimaryRS2
      'Set Text1(1).DataSource = adoPrimaryRS2
    For Each oText In Me.Text1
        Set oText.DataSource = adoPrimaryRS2
    Next
'    For Each odate In dtStart
   
        Set dtStart(1).DataSource = adoPrimaryRS2
        Set dtStart(2).DataSource = adoPrimaryRS2
        Set dtStart(0).DataSource = adoPrimaryRS2
        Set Label3.DataSource = adoPrimaryRS2
        Set Label4.DataSource = adoPrimaryRS2
'    Next
    
    If adoPrimaryRS2.RecordCount <> 0 Then
        adoPrimaryRS2.MoveFirst
        Set txtCombo.DataSource = adoPrimaryRS2
        mbDataChanged = False
    End If
        
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
If Index = 11 Then
      pili = "ITO"
        Label3.Caption = Amt2Words(Text1(11).Text)
ElseIf Index = 12 Then
       pili = "ITO"
        Label4.Caption = Amt2Words(Text1(12).Text)
End If
End Sub

Private Sub txtCombo_Change()
    
    'If adoPrimaryRS3.RecordCount <> 0 Then
    '    adoPrimaryRS3.MoveFirst
    '    Do While Not adoPrimaryRS3.EOF
    '        txtCombo.AddItem IIf(IsNull(adoPrimaryRS3("Company Name")), "", adoPrimaryRS3("Company Name"))
    '        adoPrimaryRS3.MoveNext
    '    Loop
    'End If
End Sub

Private Sub txtCombo_Click()
strSQL = "SELECT [Company Name] AS companyname FROM Company where [Company Code] like '" & txtCombo.Text & "'"
    Database_Refresh 0
    Text1(10).Text = adoPrimaryRS.Fields(0)
End Sub

Private Sub txtCombo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Text1(0).SetFocus
End Sub
