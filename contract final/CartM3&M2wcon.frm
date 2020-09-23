VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form cartv2wcon 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cart with con"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      DataField       =   "rentofmonth1"
      Height          =   375
      Left            =   9990
      TabIndex        =   84
      Text            =   "Text9"
      Top             =   2565
      Width           =   1140
   End
   Begin VB.TextBox Text8 
      DataField       =   "AdvCompute2word1"
      Height          =   375
      Left            =   9990
      TabIndex        =   83
      Text            =   "Text8"
      Top             =   2025
      Width           =   870
   End
   Begin VB.TextBox Text7 
      DataField       =   "SecCompute2word1"
      Height          =   330
      Left            =   9990
      TabIndex        =   82
      Text            =   "Text7"
      Top             =   1575
      Width           =   825
   End
   Begin VB.TextBox Text6 
      DataField       =   "AdvCompute1"
      Height          =   285
      Left            =   9990
      TabIndex        =   81
      Text            =   "Text6"
      Top             =   1080
      Width           =   780
   End
   Begin VB.TextBox Text5 
      DataField       =   "SecCompute1"
      Height          =   330
      Left            =   9990
      TabIndex        =   80
      Text            =   "Text5"
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "mk1"
      Height          =   285
      Left            =   13005
      TabIndex        =   60
      Top             =   1935
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   6
      Left            =   14445
      TabIndex        =   59
      Top             =   4230
      Width           =   2760
   End
   Begin VB.TextBox Text2 
      DataField       =   "mk"
      Height          =   285
      Left            =   12780
      TabIndex        =   58
      Text            =   "01"
      Top             =   2295
      Visible         =   0   'False
      Width           =   1230
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   45
      TabIndex        =   30
      Top             =   405
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   8388608
      TabCaption(0)   =   "Tenant "
      TabPicture(0)   =   "CartM3&M2wcon.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Other "
      TabPicture(1)   =   "CartM3&M2wcon.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   3120
         Left            =   90
         TabIndex        =   50
         Top             =   375
         Width           =   9285
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Unitcode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   23
            Left            =   6165
            TabIndex        =   78
            Top             =   180
            Width           =   2310
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "PresentativePosition"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   15
            Left            =   2295
            TabIndex        =   4
            Top             =   1395
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "TenantName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   0
            Left            =   2295
            TabIndex        =   1
            Top             =   585
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "TenantPresentative"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   1
            Left            =   3285
            TabIndex        =   3
            Top             =   990
            Width           =   5550
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "PresentativeAddress"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   2
            Left            =   2295
            TabIndex        =   5
            Top             =   1800
            Width           =   6495
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "TenantTIN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   3
            Left            =   2295
            TabIndex        =   6
            Top             =   2160
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "TENANTidt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   14
            Left            =   2295
            TabIndex        =   0
            Top             =   180
            Width           =   2760
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "secondName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2295
            TabIndex        =   2
            Top             =   990
            Width           =   960
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2250
            TabIndex        =   7
            Top             =   2655
            Width           =   6090
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Unit Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   32
            Left            =   5310
            TabIndex        =   79
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative Position"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   23
            Left            =   270
            TabIndex        =   57
            Top             =   1440
            Width           =   1980
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tenant Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   5
            Left            =   1170
            TabIndex        =   56
            Top             =   630
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tenant Presentative"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   0
            Left            =   585
            TabIndex        =   55
            Top             =   990
            Width           =   1665
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   1
            Left            =   225
            TabIndex        =   54
            Top             =   1845
            Width           =   2025
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tenant TIN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   2
            Left            =   1305
            TabIndex        =   53
            Top             =   2205
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tenant Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   21
            Left            =   1170
            TabIndex        =   52
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Type of Contract"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   30
            Left            =   855
            TabIndex        =   51
            Top             =   2700
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   885
         Left            =   90
         TabIndex        =   46
         Top             =   3480
         Width           =   9285
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Presentativeresnumber"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   4
            Left            =   360
            TabIndex        =   8
            Top             =   405
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "PresentativePlaceissued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   5
            Left            =   6345
            TabIndex        =   10
            Top             =   405
            Width           =   2760
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Presentativedateissued"
            Height          =   375
            Left            =   3375
            TabIndex        =   9
            Top             =   405
            Width           =   2760
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative resnumber"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   3
            Left            =   315
            TabIndex        =   49
            Top             =   180
            Width           =   2250
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative Res. Date issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   4
            Left            =   3240
            TabIndex        =   48
            Top             =   180
            Width           =   2685
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative Place issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   6
            Left            =   6390
            TabIndex        =   47
            Top             =   180
            Width           =   2355
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   4485
         Left            =   -74910
         TabIndex        =   41
         Top             =   420
         Width           =   4590
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Amountofrent2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   22
            Left            =   1350
            TabIndex        =   20
            Top             =   3915
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Firstmonth"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   18
            Left            =   1350
            TabIndex        =   17
            Top             =   2565
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Secondmonth"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   17
            Left            =   1350
            TabIndex        =   19
            Top             =   3465
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Amountofrent"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   7
            Left            =   1350
            TabIndex        =   18
            Top             =   2970
            Width           =   2760
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Option1"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4140
            TabIndex        =   28
            Top             =   3015
            Width           =   285
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Option1"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4185
            TabIndex        =   29
            Top             =   3960
            Width           =   285
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "TermsYear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   16
            Left            =   180
            TabIndex        =   16
            Top             =   1800
            Width           =   4200
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "leasedpremisesamt"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   15
            Top             =   1170
            Width           =   4200
         End
         Begin VB.ComboBox txtCombo 
            BackColor       =   &H00FFFFFF&
            DataField       =   "LeasedPremises"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   990
            TabIndex        =   14
            Top             =   405
            Width           =   2715
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "1st Months To"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   31
            Left            =   90
            TabIndex        =   75
            Top             =   2610
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Months"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   27
            Left            =   585
            TabIndex        =   74
            Top             =   3555
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Terms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   24
            Left            =   1890
            TabIndex        =   45
            Top             =   1575
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Location Leased"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   7
            Left            =   1485
            TabIndex        =   44
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Amount of Rent"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   8
            Left            =   1575
            TabIndex        =   43
            Top             =   2295
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Area Square Meter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   14
            Left            =   1440
            TabIndex        =   42
            Top             =   900
            Width           =   1560
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   4485
         Left            =   -70275
         TabIndex        =   35
         Top             =   420
         Width           =   4590
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "CusaCharges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   8
            Left            =   585
            TabIndex        =   21
            Top             =   405
            Width           =   3570
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "AirconCharges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   12
            Left            =   585
            TabIndex        =   22
            Top             =   990
            Width           =   3570
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "usageofPremises"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   9
            Left            =   585
            TabIndex        =   25
            Top             =   2790
            Width           =   3525
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "RentalCommenDate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   10
            Left            =   585
            TabIndex        =   26
            Top             =   3375
            Width           =   3570
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "LeaseExpiryDate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   11
            Left            =   585
            TabIndex        =   27
            Top             =   4005
            Width           =   3570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cusa Charges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   9
            Left            =   1710
            TabIndex        =   77
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aircon Charges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   13
            Left            =   1710
            TabIndex        =   76
            Top             =   810
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "usage of Premises"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   10
            Left            =   1665
            TabIndex        =   40
            Top             =   2610
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lease Commencenmen Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   11
            Left            =   1260
            TabIndex        =   39
            Top             =   3195
            Width           =   2400
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lease Expiry Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   12
            Left            =   1665
            TabIndex        =   38
            Top             =   3780
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Security Deposit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   15
            Left            =   1755
            TabIndex        =   37
            Top             =   1395
            Width           =   1350
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Sectotalamt1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Index           =   16
            Left            =   585
            TabIndex        =   23
            Top             =   1620
            Width           =   3540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "advancerent1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Index           =   17
            Left            =   585
            TabIndex        =   24
            Top             =   2205
            Width           =   3540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Advance Rental"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   18
            Left            =   1755
            TabIndex        =   36
            Top             =   2025
            Width           =   1260
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item to be Display"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1005
         Left            =   90
         TabIndex        =   31
         Top             =   4365
         Width           =   9285
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "fieldC"
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
            Index           =   21
            Left            =   6075
            TabIndex        =   13
            Top             =   540
            Width           =   2940
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "fieldB"
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
            Index           =   20
            Left            =   3105
            TabIndex        =   12
            Top             =   540
            Width           =   2940
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "fieldA"
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
            Index           =   19
            Left            =   135
            TabIndex        =   11
            Top             =   540
            Width           =   2940
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Item A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   26
            Left            =   180
            TabIndex        =   34
            Top             =   315
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Item B"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   28
            Left            =   3150
            TabIndex        =   33
            Top             =   315
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Item C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   29
            Left            =   6165
            TabIndex        =   32
            Top             =   315
            Width           =   540
         End
      End
   End
   Begin LVbuttons.LaVolpeButton cmdedit 
      Height          =   330
      Left            =   2565
      TabIndex        =   61
      Top             =   5895
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&Edit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "CartM3&M2wcon.frx":0038
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
   Begin LVbuttons.LaVolpeButton cmdcancel 
      Height          =   330
      Left            =   4905
      TabIndex        =   62
      Top             =   5895
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "CartM3&M2wcon.frx":0054
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
   Begin LVbuttons.LaVolpeButton cmd_op 
      Height          =   330
      Index           =   0
      Left            =   3375
      TabIndex        =   63
      Top             =   5895
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "CartM3&M2wcon.frx":0070
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
      Height          =   330
      Left            =   5805
      TabIndex        =   64
      Top             =   5895
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&Done"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "CartM3&M2wcon.frx":008C
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   65
      Top             =   6285
      Width           =   9525
      _ExtentX        =   16801
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
   Begin MSComCtl2.DTPicker dtStart 
      Bindings        =   "CartM3&M2wcon.frx":00A8
      Height          =   375
      Index           =   1
      Left            =   12915
      TabIndex        =   66
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21757953
      CurrentDate     =   36584
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   330
      Left            =   4140
      TabIndex        =   67
      Top             =   5895
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "CartM3&M2wcon.frx":00B3
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
   Begin VB.Label Label3 
      Caption         =   "Label3"
      DataField       =   "termstoword"
      Height          =   285
      Left            =   9810
      TabIndex        =   73
      Top             =   6795
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Tenants"
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
      Left            =   180
      TabIndex        =   72
      Top             =   45
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cusa Charges"
      DataField       =   "TenantName2"
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
      Index           =   22
      Left            =   14130
      TabIndex        =   71
      Top             =   3105
      Width           =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cusa Charges"
      DataField       =   "Settotalrenttotext1"
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
      Left            =   14580
      TabIndex        =   70
      Top             =   4950
      Width           =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cusa Charges"
      DataField       =   "settotalamttotext1"
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
      Left            =   14580
      TabIndex        =   69
      Top             =   4590
      Width           =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cusa Charges"
      DataField       =   "settotalamtrenttext1"
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
      Index           =   25
      Left            =   14580
      TabIndex        =   68
      Top             =   3915
      Width           =   4515
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   0
      Picture         =   "CartM3&M2wcon.frx":00CF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12285
   End
End
Attribute VB_Name = "cartv2wcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim strSQL4 As String
Dim strSQL5 As String
Dim WithEvents adoPrimaryRS5 As ADODB.Recordset
Attribute adoPrimaryRS5.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS3 As ADODB.Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS4 As ADODB.Recordset
Attribute adoPrimaryRS4.VB_VarHelpID = -1
Dim strSQL6 As String
Dim WithEvents adoPrimaryRS6 As ADODB.Recordset
Attribute adoPrimaryRS6.VB_VarHelpID = -1
Dim strSQL7 As String
Dim WithEvents adoPrimaryRS7 As ADODB.Recordset
Attribute adoPrimaryRS7.VB_VarHelpID = -1
Dim strSQL8 As String
Dim WithEvents adoPrimaryRS8 As ADODB.Recordset
Attribute adoPrimaryRS8.VB_VarHelpID = -1
Dim strSQL9 As String
Dim WithEvents adoPrimaryRS9 As ADODB.Recordset
Attribute adoPrimaryRS9.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim SQLcopy As String
Dim WithEvents adoPrimaryRS10 As ADODB.Recordset
Attribute adoPrimaryRS10.VB_VarHelpID = -1

Private Sub cmd_Click()
Rights1_Add = 1
If Rights1_Add = 1 Then
    adoPrimaryRS2.AddNew
    clearing
    cmd.Visible = False
    cmd_op(0).Visible = True
    cmdcancel.Visible = True
    cmdDone.Visible = False
    cmdedit.Visible = False
    unlocking
    Text1(14).SetFocus
Else
    MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End If
End Sub

Private Sub cmd_op_Click(Index As Integer)
'adoPrimaryRS2.Update
adoPrimaryRS2.UpdateBatch adAffectCurrent
cmd.Visible = True
cmd_op(0).Visible = False
cmdDone.Visible = True
cmdcancel.Visible = False
cmdedit.Visible = True
locking
End Sub

Private Sub cmdCancel_Click()
cmd.Visible = True
cmd_op(0).Visible = False
cmdDone.Visible = True
cmdcancel.Visible = False
cmdedit.Visible = True
Call Form_Load
End Sub

Private Sub cmdDelete_Click()
        If MsgBox("Are you sure you want to remove  " & Chr(10) & Chr(10) & StrConv(lvprod.SelectedItem.SubItems(2), vbUpperCase), vbYesNo + vbQuestion, "Remove Item") = vbYes Then
            lvprod.ListItems.Remove lvprod.SelectedItem.Index
        Else
            Exit Sub
        End If
End Sub

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
Rights1_Edit = 1
If Rights1_Edit = 1 Then
Dim oText As TextBox, i
    Dim odate As DTPicker, e
xcode = InputBox("Please Enter Tenants Code:", " Suppliers Information - Edit Mode")
If xcode <> "" Then
           strSQL2 = "Select [Tenant Name] as TenantName, " & _
                "[Tenant Presentative] as TenantPresentative, " & _
                "[Presentative Address] as PresentativeAddress, " & _
                "[Presentative resnumber] as Presentativeresnumber, " & _
                "[Presentative Place issued] as PresentativePlaceissued," & _
                "[leased premises amt] as leasedpremisesamt, [SecCompute] as SecCompute1, [AdvCompute] as AdvCompute1, [SecCompute2word] as SecCompute2word1, [AdvCompute2word] as AdvCompute2word1, " & _
                "[Leased Premises] as LeasedPremises,[Second month] as Secondmonth, " & _
                "[Amount of rent] as Amountofrent,[First month] as Firstmonth, " & _
                "[Cusa Charges] as CusaCharges, [Amount of rent2] as Amountofrent2, " & _
                "[usage of Premises] as usageofPremises, " & _
                "[Rental Commen Date] as RentalCommenDate, " & _
                "[Lease Expiry Date] as LeaseExpiryDate, " & _
                "[Presentative date issued] as Presentativedateissued, " & _
                "[Aircon Charges] as AirconCharges, " & _
                "[Sectotalamt] as Sectotalamt1, " & _
                "[advancerent] as advancerent1 ,[Unit code] as unitcode, " & _
                "[settotalamttotext] as settotalamttotext1," & _
                "[Settotalrenttotext] as Settotalrenttotext1," & _
                "[settotalamtrenttext] as settotalamtrenttext1, rentofmonth as rentofmonth1," & _
                "[Presentative Position] as PresentativePosition , " & _
                "[TENANTid] as TENANTidt,[Field A] as fieldA,[Field B] as fieldB,[Field C] as fieldC," & _
                "[Terms Year] as TermsYear," & _
                "[Company Code] as mk, [Terms to word] as termstoword," & _
                "[second Name] as secondName, [Typeofbis] as mk1," & _
                "[Tenant TIN] as TenantTIN from [CONTRACT LEASE] where [TENANTid] = '" & xcode & "'"
                mbEditFlag = True
                Database_Refresh 1
                If adoPrimaryRS2.RecordCount = 0 Then
                    MsgBox "No record!, ", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
                Else
                 clearing
               For Each oText In Me.Text1
                    Set oText.DataSource = adoPrimaryRS2
                Next
                Set Text3.DataSource = adoPrimaryRS2
                Set Label3.DataSource = adoPrimaryRS2
                Set Label1(16).DataSource = adoPrimaryRS2
                Set Label1(17).DataSource = adoPrimaryRS2
                Set Label1(19).DataSource = adoPrimaryRS2
                Set Label1(20).DataSource = adoPrimaryRS2
                Set Label1(25).DataSource = adoPrimaryRS2
                Set dtStart(1).DataSource = adoPrimaryRS2
                Set Text4.DataSource = adoPrimaryRS2
                Set Text5.DataSource = adoPrimaryRS2
                Set Text6.DataSource = adoPrimaryRS2
                Set Text7.DataSource = adoPrimaryRS2
                Set Text8.DataSource = adoPrimaryRS2
                Set Text9.DataSource = adoPrimaryRS2
'                Set Label1(22).DataSource = adoPrimaryRS2
                Set Text2.DataSource = adoPrimaryRS2
                Set txtCombo.DataSource = adoPrimaryRS2
                Set Combo1.DataSource = adoPrimaryRS2
                Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
                Label1(17).Caption = Format(Label1(17).Caption, "##,##0.00")
                Text1(7).Text = Format(Text1(7).Text, "##,##0.00")
                Text1(8).Text = Format(Text1(8).Text, "##,##0.00")
                Text1(12).Text = Format(Text1(12).Text, "##,##0.00")
                unlocking
                cmd.Visible = False
                cmd_op(0).Visible = True
                cmdcancel.Visible = True
                cmdDone.Visible = False
                cmdedit.Visible = False
                  Call hanap_bistype
                End If
            Else
                Beep
            End If
  Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
EditErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
                Case 13
                    Text1(1).SetFocus
                Case Else
                     KeyAscii = 0
      End Select
End Sub



Private Sub Combo2_Click()
    strSQL8 = "SELECT *  FROM typeOfContract where [Contract Name]= '" & Combo2.Text & "'"
    Database_Refresh 7
    Text4.Text = adoPrimaryRS8.Fields(0)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1(4).SetFocus
End Sub

Private Sub dtStart_LostFocus(Index As Integer)
Text3.Text = dtStart(1).Value
End Sub

Private Sub Form_Load()
    ' STARTUP SUPPLIERS DATABASE CONNECTIONS
   ' Rights5_Add = 1
   ' Rights5_Edit = 1
   ' Rights5_Save = 1
    locking
    Reload_PrimaryRS

    strSQL = "SELECT [Comp Code], [Contract1 Security Dep], [Contract1 Advance Rent] FROM MCSetup"
    Database_Refresh 0
    Text2.Text = adoPrimaryRS("Comp Code")
    secdep = adoPrimaryRS("Contract1 Security Dep")
    adrent = adoPrimaryRS("Contract1 Advance Rent")
    Text5.Text = secdep
    Text6.Text = adrent
     strSQL4 = "SELECT [Location Name]  FROM location ORDER BY [Location code]"
    Database_Refresh 3
    If adoPrimaryRS4.RecordCount <> 0 Then
        adoPrimaryRS4.MoveFirst
        Do While Not adoPrimaryRS4.EOF
            txtCombo.AddItem IIf(IsNull(adoPrimaryRS4("Location Name")), "", adoPrimaryRS4("Location Name"))
            adoPrimaryRS4.MoveNext
        Loop
    End If
    
    strSQL5 = "SELECT [sex Name]  FROM sex ORDER BY [sex code]"
    Database_Refresh 4
    If adoPrimaryRS5.RecordCount <> 0 Then
        adoPrimaryRS5.MoveFirst
        Do While Not adoPrimaryRS5.EOF
            Combo1.AddItem IIf(IsNull(adoPrimaryRS5("Sex Name")), "", adoPrimaryRS5("Sex Name"))
            adoPrimaryRS5.MoveNext
        Loop
    End If
    
    strSQL8 = "SELECT [Contract Name]  FROM typeOfContract WHERE [TypeS] like '" & "C" & "'"
    Database_Refresh 7
    If adoPrimaryRS8.RecordCount <> 0 Then
        adoPrimaryRS8.MoveFirst
        Do While Not adoPrimaryRS8.EOF
            Combo2.AddItem IIf(IsNull(adoPrimaryRS8("Contract Name")), "", adoPrimaryRS8("Contract Name"))
            adoPrimaryRS8.MoveNext
        Loop
    End If
    adoPrimaryRS8.Close
End Sub
Private Sub dview()
Do While Not adoPrimaryRS7.EOF
        Set GroupGrid.DataSource = adoPrimaryRS7
        adoPrimaryRS7.MoveNext
Loop
End Sub


Public Sub Name_supp()
strSQL7 = "SELECT * from [ItemsContract]"
  Database_Refresh 6
  GroupGrid.ClearSelCols
     dview
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
    ElseIf xMode = 3 Then
        Set adoPrimaryRS4 = New ADODB.Recordset
        adoPrimaryRS4.Open strSQL4, db, adOpenStatic, adLockOptimistic
     ElseIf xMode = 4 Then
        Set adoPrimaryRS5 = New ADODB.Recordset
        adoPrimaryRS5.Open strSQL5, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 5 Then
        Set adoPrimaryRS6 = New ADODB.Recordset
        adoPrimaryRS6.Open strSQL6, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 6 Then
        Set adoPrimaryRS7 = New ADODB.Recordset
        adoPrimaryRS7.Open strSQL7, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 7 Then
        Set adoPrimaryRS8 = New ADODB.Recordset
        adoPrimaryRS8.Open strSQL8, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 8 Then
        Set adoPrimaryRS9 = New ADODB.Recordset
        adoPrimaryRS9.Open strSQL9, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 10 Then
        Set adoPrimaryRS10 = New ADODB.Recordset
        adoPrimaryRS10.Open SQLcopy, db, adOpenStatic, adLockOptimistic
    End If
End Sub

Private Sub Reload_PrimaryRS()
    ' RELOADING DATA OBJECTS AND DATABASE CONNECTIONS
    
    Dim oText As TextBox, i
    Dim odate As DTPicker, e
    
    strSQL2 = "Select [Tenant Name] as TenantName, " & _
                "[Tenant Presentative] as TenantPresentative, " & _
                "[Presentative Address] as PresentativeAddress, " & _
                "[Presentative resnumber] as Presentativeresnumber, " & _
                "[Presentative Place issued] as PresentativePlaceissued," & _
                "[leased premises amt] as leasedpremisesamt, [SecCompute] as SecCompute1, [AdvCompute] as AdvCompute1, [SecCompute2word] as SecCompute2word1, [AdvCompute2word] as AdvCompute2word1," & _
                "[Leased Premises] as LeasedPremises,[Second month] as Secondmonth, " & _
                "[Amount of rent] as Amountofrent,[First month] as Firstmonth, " & _
                "[Cusa Charges] as CusaCharges, [Amount of rent2] as Amountofrent2, " & _
                "[usage of Premises] as usageofPremises, " & _
                "[Rental Commen Date] as RentalCommenDate, " & _
                "[Lease Expiry Date] as LeaseExpiryDate, " & _
                "[Presentative date issued] as Presentativedateissued, " & _
                "[Aircon Charges] as AirconCharges, " & _
                "[Sectotalamt] as Sectotalamt1, " & _
                "[advancerent] as advancerent1 , [Unit code] as unitcode," & _
                "[settotalamttotext] as settotalamttotext1,rentofmonth as rentofmonth1," & _
                "[Settotalrenttotext] as Settotalrenttotext1," & _
                "[settotalamtrenttext] as settotalamtrenttext1," & _
                "[Presentative Position] as PresentativePosition , " & _
                "[TENANTid] as TENANTidt,[Field A] as fieldA,[Field B] as fieldB,[Field C] as fieldC," & _
                "[Terms Year] as TermsYear," & _
                "[Company Code] as mk, [Terms to word] as termstoword," & _
                "[second Name] as secondName, [Typeofbis] as mk1," & _
                "[Tenant TIN] as TenantTIN from [CONTRACT LEASE]"
                Database_Refresh 1
                 '"[Tenant Name 2] as TenantName2," & _
                '"[Leased approximately] as Leasedapproximately, "
                For Each oText In Me.Text1
                    Set oText.DataSource = adoPrimaryRS2
                Next
                Set Text3.DataSource = adoPrimaryRS2
                Set Text5.DataSource = adoPrimaryRS2
                Set Text6.DataSource = adoPrimaryRS2
                Set Text7.DataSource = adoPrimaryRS2
                Set Text8.DataSource = adoPrimaryRS2
                Set Text9.DataSource = adoPrimaryRS2
                Set Label3.DataSource = adoPrimaryRS2
                Set Label1(16).DataSource = adoPrimaryRS2
                Set Label1(17).DataSource = adoPrimaryRS2
                Set Label1(19).DataSource = adoPrimaryRS2
                Set Label1(20).DataSource = adoPrimaryRS2
                Set Label1(25).DataSource = adoPrimaryRS2
                'Set dtStart(1).DataSource = adoPrimaryRS2
'                Set Label1(22).DataSource = adoPrimaryRS2
                Set Text2.DataSource = adoPrimaryRS2
                Set Text4.DataSource = adoPrimaryRS2
                  'If adoPrimaryRS2.RecordCount <> 0 Then
                  '  adoPrimaryRS2.MoveFirst
                    Set txtCombo.DataSource = adoPrimaryRS2
                  'End If
'                 If adoPrimaryRS2.RecordCount <> 0 Then
 '                   adoPrimaryRS2.MoveFirst
                    Set Combo1.DataSource = adoPrimaryRS2
                    Call hanap_bistype
  '               End If
End Sub
Function hanap_bistype()
On Error Resume Next
strSQL6 = "SELECT *  FROM typeOfContract where [Contract Code]= '" & Text4.Text & "'"
                Database_Refresh 5
'                Text4.Text = adoPrimaryRS6.Fields(0)
                Combo2.Text = adoPrimaryRS6.Fields(1)
End Function

Private Sub LaVolpeButton1_Click()
'  strSQL6 = "SELECT * FROM ItemsContract"
'                        Database_Refresh 5
'                        With adoPrimaryRS6
'                                .AddNew
'                                 .Fields(0) = Text1(14).Text
'                                 .Fields(1) = Text5.Text
'                                .Update
'                                .Requery
'                                .Close
'                        End With
'                        Text5.Text = ""
'                        Call Name_supp
                    If Text1(20).Text = "" Or Text1(21).Text = "" Then
                    Else
                        Set lst1 = lvprod.ListItems.Add(, , Text1(20).Text) 'code
                        With lst1
                            lst1.SubItems(1) = Text1(21).Text
                            'lst1.SubItems(0) = Text1(21).Text
                        End With
                    End If
                        Text1(20).Text = ""
                        Text1(21).Text = ""
End Sub


Private Sub lvprod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then lvprod.ListItems.Remove lvprod.SelectedItem.Index
End Sub


Private Sub Option1_Click()
Label1(16).Caption = Val(Text1(7).Text * secdep)
Label1(17).Caption = Val(Text1(7).Text * adrent)
 Text5.Text = secdep
                           Text6.Text = adrent
                          pili = "ITO"
                          Text7.Text = Amt2Words(Text5.Text)
                          pili = "ITO"
                          Text8.Text = Amt2Words(Text6.Text)
Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
Label1(17).Caption = Format(Label1(17).Caption, "##,##0.00")
Label1(19).Caption = Amt2Words(Label1(16).Caption)
Label1(20).Caption = Amt2Words(Label1(17).Caption)
Label1(25).Caption = Amt2Words(Text1(7).Text)

End Sub

Private Sub Option2_Click()
Label1(16).Caption = Val(Text1(22).Text * secdep)
Label1(17).Caption = Val(Text1(22).Text * adrent)
Text9.Text = Val(Text1(22).Text * Text1(13).Text)
 Text5.Text = secdep
                           Text6.Text = adrent
                          pili = "ITO"
                          Text7.Text = Amt2Words(Text5.Text)
                          pili = "ITO"
                          Text8.Text = Amt2Words(Text6.Text)
Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
Label1(17).Caption = Format(Label1(17).Caption, "##,##0.00")
Label1(19).Caption = Amt2Words(Label1(16).Caption)
Label1(20).Caption = Amt2Words(Label1(17).Caption)
Label1(25).Caption = Amt2Words(Text1(7).Text)
End Sub

Private Sub txtCombo_Click()
Text1(6).Text = txtCombo.Text
End Sub

Private Sub txtCombo_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case 13
             Text1(13).SetFocus
        Case Else
             KeyAscii = 0
End Select
End Sub
Function clearing()
For i = 0 To 23
   Text1(i).Text = ""
   Me.Label1(16).Caption = ""
   Text2.Text = "01"
   Combo2.Text = ""
Next i
End Function

Public Sub find_existingtenants()
xcode = Text1(14).Text
If xcode <> "" Then
            strSQL9 = "Select [Tenant Name] as TenantName, " & _
                "[Tenant Presentative] as TenantPresentative, " & _
                "[Presentative Address] as PresentativeAddress, " & _
                "[Presentative resnumber] as Presentativeresnumber, " & _
                "[Presentative Place issued] as PresentativePlaceissued," & _
                "[leased premises amt] as leasedpremisesamt, [SecCompute] as SecCompute1, [AdvCompute] as AdvCompute1, [SecCompute2word] as SecCompute2word1, [AdvCompute2word] as AdvCompute2word1," & _
                "[Leased Premises] as LeasedPremises,[Second month] as Secondmonth, " & _
                "[Amount of rent] as Amountofrent,[First month] as Firstmonth, " & _
                "[Cusa Charges] as CusaCharges, [Amount of rent2] as Amountofrent2, " & _
                "[usage of Premises] as usageofPremises, " & _
                "[Rental Commen Date] as RentalCommenDate, " & _
                "[Lease Expiry Date] as LeaseExpiryDate, " & _
                "[Presentative date issued] as Presentativedateissued, " & _
                "[Aircon Charges] as AirconCharges, " & _
                "[Sectotalamt] as Sectotalamt1, rentofmonth as rentofmonth1," & _
                "[advancerent] as advancerent1 , [Unit code] as unitcode," & _
                "[settotalamttotext] as settotalamttotext1," & _
                "[Settotalrenttotext] as Settotalrenttotext1," & _
                "[settotalamtrenttext] as settotalamtrenttext1," & _
                "[Presentative Position] as PresentativePosition , " & _
                "[TENANTid] as TENANTidt,[Field A] as fieldA,[Field B] as fieldB,[Field C] as fieldC," & _
                "[Terms Year] as TermsYear," & _
                "[Company Code] as mk, [Terms to word] as termstoword," & _
                "[second Name] as secondName, [Typeofbis] as mk1," & _
                "[Tenant TIN] as TenantTIN  from [CONTRACT LEASE] where [TENANTid] = '" & xcode & "'"
                mbEditFlag = True
                Database_Refresh 8
                If adoPrimaryRS9.RecordCount = 0 Then
                    'MsgBox "No record!, ", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
                    Text1(23).SetFocus
                Else
                   msgs = MsgBox("RENEW CONTRACT!, ", vbYesNo + vbCritical, "Warning:End-User:" + UserName)
                    If msgs = vbYes Then
                        renew
                        saveOLdrec
                    End If
                End If
            Else
                Beep
            End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 16 Or Index = 7 Or Index = 8 Or Index = 12 Or Index = 13 Then
        Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                Case 13
                     If Index = 13 Then Text1(16).SetFocus
                     If Index = 16 Then
                        pili = "ITO"
                        Label3.Caption = Amt2Words(Text1(16).Text)
                        Text1(18).SetFocus
                     End If
                Case x8
                
                Case Else
                     KeyAscii = 0
        End Select
 End If
If KeyAscii = 13 Then
    If Index = 14 Then
        If Text1(14).Text = "" Then
           Text1(14).SetFocus
        Else
           Call find_existingtenants
        End If
    End If
    If Index = 0 Then Combo1.SetFocus
    If Index = 1 Then Text1(15).SetFocus
    If Index = 15 Then Text1(2).SetFocus
    If Index = 2 Then Text1(3).SetFocus
    If Index = 3 Then Combo2.SetFocus
    If Index = 4 Then Text3.SetFocus
    If Index = 5 Then Text1(19).SetFocus
    If Index = 19 Then Text1(20).SetFocus
    If Index = 20 Then Text1(21).SetFocus
    If Index = 21 Then
        SSTab1.Tab = 1
        txtCombo.SetFocus
    End If
    If Index = 6 Then Text1(13).SetFocus
    'If Index = 13 Then Text1(16).SetFocus
    
    If Index = 7 Then Text1(17).SetFocus
    If Index = 17 Then Text1(22).SetFocus
    If Index = 22 Then Text1(8).SetFocus
    If Index = 8 Then Text1(12).SetFocus
    If Index = 12 Then Text1(9).SetFocus
    If Index = 9 Then Text1(10).SetFocus
    If Index = 10 Then Text1(11).SetFocus
    If Index = 18 Then Text1(7).SetFocus
    If Index = 23 Then Text1(0).SetFocus
End If
End Sub



Private Sub Text1_LostFocus(Index As Integer)
    'If Index = 7 Then
    '    Dim sn As Long
    '    Dim tsn As Long
    '    Dim sn1 As Long
    '    sn = CLng(Text1(13).Text * Text1(7).Text)
    '    tsn = CLng(sn * 4)
    '    sn1 = CLng(Val(Text1(13).Text * Text1(7).Text) * 2)
    '    Label1(16).Caption = sn
    '    Label1(17).Caption = sn1
    '    Label1(19).Caption = Amt2Words(Label1(16).Caption)
    '    Label1(20).Caption = Amt2Words(Label1(17).Caption)
    'ElseIf Index = 0 Then
    '    Text1(0).Text = UCase(Text1(0).Text)
    '    Label1(22).Caption = Text1(0).Text
    'End If
    'If Index = 5 Then Text1(5).Text = Format(Text1(5).Text, "##,##0.00")
    If Index = 13 Then Text1(13).Text = Format(Text1(13).Text, "##,##0.00")
    If Index = 7 Then Text1(7).Text = Format(Text1(7).Text, "##,##0.00")
    If Index = 8 Then Text1(8).Text = Format(Text1(8).Text, "##,##0.00")
    If Index = 12 Then Text1(12).Text = Format(Text1(12).Text, "##,##0.00")
    If Index = 22 Then Text1(22).Text = Format(Text1(22).Text, "##,##0.00")
    'If Index = 9 Then Text1(9).Text = Format(Text1(9).Text, "##,##0.00")
    If Index = 16 Then
                        pili = "ITO"
                        If Text1(16).Text = "" Then
                        Else
                        Label3.Caption = Amt2Words(Text1(16).Text)
                        End If
    End If
End Sub
Function locking()
For i = 0 To 23
    Text1(i).Enabled = False
Next i
  Label1(16).Enabled = False
  Label1(17).Enabled = False
  dtStart(1).Enabled = False
  Combo1.Enabled = False
  Combo2.Enabled = False
  txtCombo.Enabled = False
  Text3.Enabled = False
End Function
Function unlocking()
For i = 0 To 23
    Text1(i).Enabled = True
Next i
Label1(16).Enabled = True
  Label1(17).Enabled = True
  dtStart(1).Enabled = True
  Combo1.Enabled = True
  Combo2.Enabled = True
  Text3.Enabled = True
  txtCombo.Enabled = True
End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1(5).SetFocus
End Sub




Function saveOLdrec()
On Error Resume Next
SQLcopy = "Select * from [CONTRACT LEASE] where TENANTid ='" & Text1(14).Text & "'"
Database_Refresh 10

strSQL7 = "Select * from [Contract Lease OLD]"
                Database_Refresh 6
                
                adoPrimaryRS7.AddNew
                adoPrimaryRS7.Fields(1) = adoPrimaryRS10.Fields(0)
                adoPrimaryRS7.Fields(2) = adoPrimaryRS10.Fields(1)
                adoPrimaryRS7.Fields(3) = adoPrimaryRS10.Fields(2)
                adoPrimaryRS7.Fields(4) = adoPrimaryRS10.Fields(3)
                adoPrimaryRS7.Fields(5) = adoPrimaryRS10.Fields(4)
                adoPrimaryRS7.Fields(6) = adoPrimaryRS10.Fields(5)
                adoPrimaryRS7.Fields(7) = adoPrimaryRS10.Fields(6)
                adoPrimaryRS7.Fields(8) = adoPrimaryRS10.Fields(7)
                adoPrimaryRS7.Fields(9) = adoPrimaryRS10.Fields(8)
                adoPrimaryRS7.Fields(10) = adoPrimaryRS10.Fields(9)
                adoPrimaryRS7.Fields(11) = adoPrimaryRS10.Fields(10)
                adoPrimaryRS7.Fields(12) = adoPrimaryRS10.Fields(11)
                adoPrimaryRS7.Fields(13) = adoPrimaryRS10.Fields(12)
                adoPrimaryRS7.Fields(14) = adoPrimaryRS10.Fields(13)
                adoPrimaryRS7.Fields(15) = adoPrimaryRS10.Fields(14)
                adoPrimaryRS7.Fields(16) = adoPrimaryRS10.Fields(15)
                adoPrimaryRS7.Fields(17) = adoPrimaryRS10.Fields(16)
                adoPrimaryRS7.Fields(18) = adoPrimaryRS10.Fields(17)
                adoPrimaryRS7.Fields(19) = adoPrimaryRS10.Fields(18)
                adoPrimaryRS7.Fields(20) = adoPrimaryRS10.Fields(19)
                adoPrimaryRS7.Fields(21) = adoPrimaryRS10.Fields(20)
                adoPrimaryRS7.Fields(22) = adoPrimaryRS10.Fields(21)
                adoPrimaryRS7.Fields(23) = adoPrimaryRS10.Fields(22)
                adoPrimaryRS7.Fields(24) = adoPrimaryRS10.Fields(23)
                adoPrimaryRS7.Fields(25) = adoPrimaryRS10.Fields(24)
                adoPrimaryRS7.Fields(26) = adoPrimaryRS10.Fields(25)
                adoPrimaryRS7.Fields(27) = adoPrimaryRS10.Fields(26)
                adoPrimaryRS7.Fields(28) = adoPrimaryRS10.Fields(27)
                adoPrimaryRS7.Fields(29) = adoPrimaryRS10.Fields(28)
                adoPrimaryRS7.Fields(30) = adoPrimaryRS10.Fields(29)
                adoPrimaryRS7.Fields(31) = adoPrimaryRS10.Fields(30)
                adoPrimaryRS7.Fields(32) = adoPrimaryRS10.Fields(31)
                adoPrimaryRS7.Fields(33) = adoPrimaryRS10.Fields(32)
                adoPrimaryRS7.Fields(34) = adoPrimaryRS10.Fields(33)
                adoPrimaryRS7.Fields(35) = adoPrimaryRS10.Fields(34)
                adoPrimaryRS7.Fields(36) = adoPrimaryRS10.Fields(35)
                adoPrimaryRS7.Fields(37) = adoPrimaryRS10.Fields(36)
                adoPrimaryRS7.Fields(38) = adoPrimaryRS10.Fields(37)
                adoPrimaryRS7.Fields(39) = adoPrimaryRS10.Fields(38)
                adoPrimaryRS7.Fields(40) = adoPrimaryRS10.Fields(39)
                adoPrimaryRS7.Fields(41) = adoPrimaryRS10.Fields(39)
                adoPrimaryRS7.Fields(41) = adoPrimaryRS10.Fields(40)
                adoPrimaryRS7.Fields(42) = adoPrimaryRS10.Fields(41)
                adoPrimaryRS7.Fields(43) = adoPrimaryRS10.Fields(42)
                adoPrimaryRS7.Fields(44) = adoPrimaryRS10.Fields(43)
                adoPrimaryRS7.Fields(45) = adoPrimaryRS10.Fields(44)
                adoPrimaryRS7.Fields(46) = adoPrimaryRS10.Fields(45)
                adoPrimaryRS7.Fields(47) = adoPrimaryRS10.Fields(46)
                adoPrimaryRS7.Fields(48) = adoPrimaryRS10.Fields(47)
                adoPrimaryRS7.Fields(49) = adoPrimaryRS10.Fields(48)
                adoPrimaryRS7.Update
    End Function

Function renew()
strSQL2 = "Select [Tenant Name] as TenantName, " & _
                "[Tenant Presentative] as TenantPresentative, " & _
                "[Presentative Address] as PresentativeAddress, " & _
                "[Presentative resnumber] as Presentativeresnumber, " & _
                "[Presentative Place issued] as PresentativePlaceissued," & _
                "[leased premises amt] as leasedpremisesamt, [SecCompute] as SecCompute1, [AdvCompute] as AdvCompute1, [SecCompute2word] as SecCompute2word1, [AdvCompute2word] as AdvCompute2word1, " & _
                "[Leased Premises] as LeasedPremises,[Second month] as Secondmonth, " & _
                "[Amount of rent] as Amountofrent,[First month] as Firstmonth, " & _
                "[Cusa Charges] as CusaCharges, [Amount of rent2] as Amountofrent2, " & _
                "[usage of Premises] as usageofPremises, " & _
                "[Rental Commen Date] as RentalCommenDate, " & _
                "[Lease Expiry Date] as LeaseExpiryDate, " & _
                "[Presentative date issued] as Presentativedateissued, " & _
                "[Aircon Charges] as AirconCharges, " & _
                "[Sectotalamt] as Sectotalamt1, " & _
                "[advancerent] as advancerent1 ,[Unit code] as unitcode, " & _
                "[settotalamttotext] as settotalamttotext1," & _
                "[Settotalrenttotext] as Settotalrenttotext1," & _
                "[settotalamtrenttext] as settotalamtrenttext1,rentofmonth as rentofmonth1" & _
                "[Presentative Position] as PresentativePosition , " & _
                "[TENANTid] as TENANTidt,[Field A] as fieldA,[Field B] as fieldB,[Field C] as fieldC," & _
                "[Terms Year] as TermsYear," & _
                "[Company Code] as mk, [Terms to word] as termstoword," & _
                "[second Name] as secondName, [Typeofbis] as mk1," & _
                "[Tenant TIN] as TenantTIN from [CONTRACT LEASE] where [TENANTid] = '" & Text1(14) & "'"
                mbEditFlag = True
                Database_Refresh 1
                If adoPrimaryRS2.RecordCount = 0 Then
                    MsgBox "No record!, ", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
                Else
                 clearing
               For Each oText In Me.Text1
                    Set oText.DataSource = adoPrimaryRS2
                Next
                Set Text3.DataSource = adoPrimaryRS2
                Set Label3.DataSource = adoPrimaryRS2
                Set Label1(16).DataSource = adoPrimaryRS2
                Set Label1(17).DataSource = adoPrimaryRS2
                Set Label1(19).DataSource = adoPrimaryRS2
                Set Label1(20).DataSource = adoPrimaryRS2
                Set Label1(25).DataSource = adoPrimaryRS2
                Set dtStart(1).DataSource = adoPrimaryRS2
                Set Text4.DataSource = adoPrimaryRS2
                Set Text5.DataSource = adoPrimaryRS2
                Set Text6.DataSource = adoPrimaryRS2
                Set Text7.DataSource = adoPrimaryRS2
                Set Text8.DataSource = adoPrimaryRS2
                Set Text9.DataSource = adoPrimaryRS2
'                Set Label1(22).DataSource = adoPrimaryRS2
                Set Text2.DataSource = adoPrimaryRS2
                Set txtCombo.DataSource = adoPrimaryRS2
                Set Combo1.DataSource = adoPrimaryRS2
                Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
                Label1(17).Caption = Format(Label1(17).Caption, "##,##0.00")
                Text1(7).Text = Format(Text1(7).Text, "##,##0.00")
                Text1(8).Text = Format(Text1(8).Text, "##,##0.00")
                Text1(12).Text = Format(Text1(12).Text, "##,##0.00")
                unlocking
                cmd.Visible = False
                cmd_op(0).Visible = True
                cmdcancel.Visible = True
                cmdDone.Visible = False
                cmdedit.Visible = False
                  Call hanap_bistype
                End If
            
End Function









