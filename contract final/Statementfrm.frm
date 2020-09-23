VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form statementfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Statement of Account"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7725
   Icon            =   "Statementfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CURRENT ACCOUNTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3210
      Left            =   45
      TabIndex        =   14
      Top             =   2115
      Width           =   7620
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   5
         Left            =   4095
         TabIndex        =   27
         Top             =   630
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   4
         Left            =   900
         TabIndex        =   15
         Top             =   630
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Entry 2 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   6
         Left            =   4095
         TabIndex        =   32
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Entry 1 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   5
         Left            =   900
         TabIndex        =   31
         Top             =   405
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "RENTAL DEPOSIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   7
         Left            =   1260
         TabIndex        =   25
         Top             =   1170
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADVANCE RENTAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   8
         Left            =   1125
         TabIndex        =   24
         Top             =   1575
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   9
         Left            =   2025
         TabIndex        =   23
         Top             =   2070
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         DataField       =   "RentalDeposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   10
         Left            =   3015
         TabIndex        =   22
         Top             =   1125
         Width           =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         DataField       =   "AdvanceRental"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   11
         Left            =   3015
         TabIndex        =   21
         Top             =   1530
         Width           =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Total1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   12
         Left            =   3015
         TabIndex        =   20
         Top             =   2025
         Width           =   2880
      End
      Begin VB.Line Line2 
         X1              =   2655
         X2              =   6300
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "LESS: EXISTING DEPOSIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   13
         Left            =   720
         TabIndex        =   19
         Top             =   2475
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Lessexisting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   14
         Left            =   3015
         TabIndex        =   18
         Top             =   2385
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTAL AMOUNT DUE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   15
         Left            =   990
         TabIndex        =   17
         Top             =   2835
         Width           =   1650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TotalAmountDue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   16
         Left            =   3015
         TabIndex        =   16
         Top             =   2790
         Width           =   2880
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   45
      TabIndex        =   1
      Top             =   585
      Width           =   7665
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "PeriodBilling"
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
         Index           =   6
         Left            =   5175
         TabIndex        =   28
         Top             =   450
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Stalllocation"
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
         Left            =   5175
         TabIndex        =   26
         Top             =   1035
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Telephone1"
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
         Left            =   1305
         TabIndex        =   7
         Top             =   1035
         Width           =   3570
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Address1"
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
         Left            =   1305
         TabIndex        =   6
         Top             =   675
         Width           =   3570
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "StallName"
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
         Left            =   1305
         TabIndex        =   4
         Top             =   315
         Width           =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DATE PERIOD:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   4
         Left            =   5175
         TabIndex        =   30
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "STALL LOCATION:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   2
         Left            =   5175
         TabIndex        =   29
         Top             =   810
         Width           =   1470
      End
      Begin VB.Line Line1 
         X1              =   315
         X2              =   7650
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TELEPHONE    :       "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1035
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADDRESS       :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "STALL NAME :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   1140
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   5790
      Width           =   7725
      _ExtentX        =   13626
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
   Begin LVbuttons.LaVolpeButton cmdedit 
      Height          =   330
      Left            =   1530
      TabIndex        =   9
      Top             =   5400
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
      MICON           =   "Statementfrm.frx":4F0A
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
      Left            =   3870
      TabIndex        =   10
      Top             =   5400
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
      MICON           =   "Statementfrm.frx":4F26
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
      Left            =   2340
      TabIndex        =   11
      Top             =   5400
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
      MICON           =   "Statementfrm.frx":4F42
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
      Left            =   4770
      TabIndex        =   12
      Top             =   5400
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
      MICON           =   "Statementfrm.frx":4F5E
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
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   330
      Left            =   3105
      TabIndex        =   13
      Top             =   5400
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
      MICON           =   "Statementfrm.frx":4F7A
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Statement of Account"
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
      TabIndex        =   0
      Top             =   45
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   0
      Picture         =   "Statementfrm.frx":4F96
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12285
   End
End
Attribute VB_Name = "statementfrm"
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
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

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

Else
    MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End If
End Sub

Private Sub cmd_op_Click(Index As Integer)
adoPrimaryRS2.Update


' adoPrimaryRS2.UpdateBatch adAffectAll
cmd.Visible = True
cmd_op(0).Visible = False
cmdDone.Visible = True
cmdcancel.Visible = False
cmdedit.Visible = True
End Sub

Private Sub cmdCancel_Click()
cmd.Visible = True
cmd_op(0).Visible = False
cmdDone.Visible = True
cmdcancel.Visible = False
cmdedit.Visible = True
Call Form_Load
End Sub

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
Rights1_Edit = 1
If Rights1_Edit = 1 Then
Dim oText As TextBox, i
    Dim odate As DTPicker, e
xcode = InputBox("Please Enter Supplier Code:", " Suppliers Information - Edit Mode")
If xcode <> "" Then
    strSQL2 = "Select [Stall Name] as StallName,[Addresssd] as Address1," & _
              "[Telephone] as Telephone1,[Period Billing] as PeriodBilling," & _
              "[Stall location] as Stalllocation,[Rental Deposit] as RentalDeposit," & _
              "[Advance Rental] as AdvanceRental,[Total] as Total1," & _
              "[Less existing] as Lessexisting, [Total Amount Due] as TotalAmountDue from [Statement of Account] where [Stall Name] = '" & xcode & "'"
                mbEditFlag = True
                Database_Refresh 1
                If adoPrimaryRS2.RecordCount = 0 Then
                    MsgBox "No record!, ", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
                Else
                    clearing
                    For Each oText In Me.Text1
                        Set oText.DataSource = adoPrimaryRS2
                    Next
                    
                    Set Label1(10).DataSource = adoPrimaryRS2
                    Set Label1(11).DataSource = adoPrimaryRS2
                    Set Label1(12).DataSource = adoPrimaryRS2
                    Set Label1(13).DataSource = adoPrimaryRS2
                    Set Label1(14).DataSource = adoPrimaryRS2
                    Set Label1(16).DataSource = adoPrimaryRS2
                    Label1(10).Caption = Format(Label1(10).Caption, "##,##0.00")
                    Label1(11).Caption = Format(Label1(11).Caption, "##,##0.00")
                    Label1(12).Caption = Format(Label1(12).Caption, "##,##0.00")
                    Label1(13).Caption = Format(Label1(13).Caption, "##,##0.00")
                    Label1(14).Caption = Format(Label1(14).Caption, "##,##0.00")
                    Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
                    Text1(6).Text = Format(Text1(6).Text, "MMMM yyyy")
                    unlocking
                    cmd.Visible = False
                    cmd_op(0).Visible = True
                    cmdcancel.Visible = True
                    cmdDone.Visible = False
                    cmdedit.Visible = False
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


Private Sub Form_Load()
    ' STARTUP SUPPLIERS DATABASE CONNECTIONS
   ' Rights5_Add = 1
   ' Rights5_Edit = 1
   ' Rights5_Save = 1
    
    locking
    Reload_PrimaryRS

    strSQL = "SELECT [Comp Code], [Contract1 Security Dep], [Contract1 Advance Rent] FROM MCSetup"
    Database_Refresh 0
    
    secdep = adoPrimaryRS("Contract1 Security Dep")
    adrent = adoPrimaryRS("Contract1 Advance Rent")
    Label1(10).Caption = Format(Label1(10).Caption, "##,##0.00")
    Label1(11).Caption = Format(Label1(11).Caption, "##,##0.00")
    Label1(12).Caption = Format(Label1(12).Caption, "##,##0.00")
    Label1(13).Caption = Format(Label1(13).Caption, "##,##0.00")
    Label1(14).Caption = Format(Label1(14).Caption, "##,##0.00")
    Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
    Text1(6).Text = Format(Text1(6).Text, "MMMM yyyy")
End Sub
Public Sub Database_Refresh(xMode As Integer)
    ' PRE-DATABASE CONNECTION WITH PARAMETERIZED SQL VARIABLES ATTACHED IN EVERY MODE
    
    
        
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
    End If
End Sub

Private Sub Reload_PrimaryRS()
    ' RELOADING DATA OBJECTS AND DATABASE CONNECTIONS
    On Error Resume Next
    Dim oText As TextBox, i
    Dim odate As DTPicker, e

    strSQL2 = "Select [Stall Name] as StallName,[Addresssd] as Address1," & _
              "[Telephone] as Telephone1,[Period Billing] as PeriodBilling," & _
              "[Stall location] as Stalllocation,[Rental Deposit] as RentalDeposit," & _
              "[Advance Rental] as AdvanceRental,[Total] as Total1," & _
              "[Less existing] as Lessexisting,[Total Amount Due] as TotalAmountDue from [Statement of Account]"
               Database_Refresh 1
                For Each oText In Me.Text1
                    Set oText.DataSource = adoPrimaryRS2
                Next
               'Set Me.dtStart.DataSource = adoPrimaryRS2
               Set Label1(10).DataSource = adoPrimaryRS2
               Set Label1(11).DataSource = adoPrimaryRS2
               Set Label1(12).DataSource = adoPrimaryRS2
               Set Label1(13).DataSource = adoPrimaryRS2
               Set Label1(14).DataSource = adoPrimaryRS2
               Set Label1(16).DataSource = adoPrimaryRS2
               
                
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
For i = 0 To 6
   Text1(i).Text = ""
   Me.Label1(16).Caption = ""
'   Text2.Text = "01"
'   Text4.Text = "02"
Next i
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 4 Or Index = 5 Then
        Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                Case 13
                     If Index = 3 Then Text1(4).SetFocus
                     
                Case x8
                
                Case Else
                     KeyAscii = 0
        End Select
 End If
If KeyAscii = 13 Then
    If Index = 6 Then
    Text1(6).Text = Format(Text1(6).Text, "MMM yyyy")
    Text1(1).SetFocus
    End If
    If Index = 4 Then
        Text1(5).SetFocus
        Dim sn As Currency
        sn = Val(Text1(4).Text * secdep)
        Label1(10).Caption = Format(sn, "##,##0.00")
        Dim sn1 As Currency
        sn1 = Val(Text1(4).Text * adrent)
        Label1(11).Caption = Format(sn1, "##,##0.00")
        Dim tot As Currency
        
        tot = Val(Val(sn) + Val(sn1))
        
        Label1(12).Caption = Format(tot, "##,##0.00")
    End If
    If Index = 5 Then
       Dim ttot As String
       Dim ttot1 As Currency
       Dim ttot2 As Currency
       
        Label1(14).Caption = Format(Text1(5).Text, "##,##0.00")
        ttot1 = Label1(12).Caption
        ttot2 = Text1(5).Text
        
        ttot = Val(ttot1 + ttot2)
        
        Label1(16).Caption = Format(ttot, "##,##0.00")
    End If
    If Index = 0 Then Text1(1).SetFocus
    If Index = 1 Then Text1(2).SetFocus
    If Index = 2 Then Text1(3).SetFocus
    If Index = 3 Then Text1(4).SetFocus
    If Index = 4 Then
       If Text1(4).Text = "" Then Text1(4).SetFocus
       Text1(5).SetFocus
    End If
       
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
    If Index = 9 Then Text1(9).Text = Format(Text1(9).Text, "##,##0.00")
End Sub
Function locking()
For i = 0 To 6
    Text1(i).Enabled = False
Next i
  
'  dtStart.Enabled = False
Label1(10).Enabled = False
Label1(11).Enabled = False
Label1(12).Enabled = False
Label1(14).Enabled = False
Label1(16).Enabled = False

  
End Function
Function unlocking()
For i = 0 To 6
    Text1(i).Enabled = True
Next i
Label1(10).Enabled = True
Label1(11).Enabled = True
Label1(12).Enabled = True
Label1(14).Enabled = True
Label1(16).Enabled = True


End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1(5).SetFocus
End Sub








