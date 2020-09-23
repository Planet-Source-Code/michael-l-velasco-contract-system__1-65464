VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form CartFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "For Cart Contract"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9660
   Icon            =   "Fixfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      DataField       =   "mk1"
      Height          =   285
      Left            =   10170
      TabIndex        =   36
      Top             =   3735
      Width           =   1230
   End
   Begin LVbuttons.LaVolpeButton cmdedit 
      Height          =   330
      Left            =   2565
      TabIndex        =   33
      Top             =   6705
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
      MICON           =   "Fixfrm.frx":4F0A
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
   Begin VB.TextBox Text2 
      DataField       =   "mk"
      Height          =   285
      Left            =   10170
      TabIndex        =   28
      Text            =   "01"
      Top             =   3420
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
      Left            =   13005
      TabIndex        =   23
      Top             =   4050
      Width           =   2760
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6135
      Left            =   90
      TabIndex        =   21
      Top             =   540
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   10821
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      BackColor       =   16777215
      PaneTree        =   "Fixfrm.frx":4F26
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   2460
         Left            =   4755
         TabIndex        =   53
         Top             =   3675
         Width           =   4710
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1620
            MaxLength       =   100
            TabIndex        =   18
            Top             =   1035
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1620
            TabIndex        =   19
            Top             =   1440
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1620
            TabIndex        =   20
            Top             =   1845
            Width           =   2760
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "usage of Premises"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   10
            Left            =   225
            TabIndex        =   58
            Top             =   1125
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Rental Commen Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   11
            Left            =   90
            TabIndex        =   57
            Top             =   1530
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lease Expiry Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   12
            Left            =   270
            TabIndex        =   56
            Top             =   1935
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Security Deposit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   15
            Left            =   405
            TabIndex        =   55
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Left            =   1620
            TabIndex        =   16
            Top             =   270
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Left            =   1620
            TabIndex        =   17
            Top             =   630
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Advance Rental"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   18
            Left            =   360
            TabIndex        =   54
            Top             =   675
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   2460
         Left            =   0
         TabIndex        =   48
         Top             =   3675
         Width           =   4695
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1665
            TabIndex        =   61
            Top             =   2925
            Visible         =   0   'False
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1530
            TabIndex        =   13
            Top             =   1575
            Width           =   690
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2340
            TabIndex        =   14
            Top             =   1575
            Width           =   1995
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1575
            TabIndex        =   12
            Top             =   855
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1530
            TabIndex        =   15
            Top             =   1980
            Width           =   2805
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1575
            TabIndex        =   11
            Top             =   450
            Width           =   2760
         End
         Begin VB.ComboBox txtCombo 
            DataField       =   "LeasedPremises"
            Height          =   315
            Left            =   1575
            TabIndex        =   10
            Top             =   90
            Width           =   2715
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aircon Charges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   13
            Left            =   450
            TabIndex        =   62
            Top             =   3015
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "1st Months To"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   26
            Left            =   360
            TabIndex        =   60
            Top             =   1665
            Width           =   1020
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
            Left            =   2160
            TabIndex        =   59
            Top             =   1350
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Terms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   24
            Left            =   1035
            TabIndex        =   52
            Top             =   945
            Width           =   450
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   315
            TabIndex        =   51
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cusa Charges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   450
            TabIndex        =   50
            Top             =   2070
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Area Square Meter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   14
            Left            =   135
            TabIndex        =   49
            Top             =   540
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   870
         Left            =   0
         TabIndex        =   44
         Top             =   2745
         Width           =   9465
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   270
            TabIndex        =   7
            Top             =   360
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   6255
            TabIndex        =   9
            Top             =   360
            Width           =   2760
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            DataField       =   "Presentativedateissued"
            Height          =   375
            Left            =   3285
            TabIndex        =   8
            Top             =   360
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   225
            TabIndex        =   47
            Top             =   135
            Width           =   1920
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative Res. Date issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   3150
            TabIndex        =   46
            Top             =   135
            Width           =   2370
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative Place issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   6300
            TabIndex        =   45
            Top             =   135
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   2685
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   9465
         Begin VB.ComboBox Combo2 
            DataField       =   "mk1"
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
            Left            =   5670
            TabIndex        =   64
            Top             =   2295
            Width           =   870
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   2340
            TabIndex        =   4
            Top             =   1350
            Width           =   3480
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1530
            TabIndex        =   1
            Top             =   585
            Width           =   6315
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   2925
            TabIndex        =   3
            Top             =   990
            Width           =   5235
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   2340
            TabIndex        =   5
            Top             =   1710
            Width           =   6405
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1575
            TabIndex        =   6
            Top             =   2250
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1530
            TabIndex        =   0
            Top             =   180
            Width           =   2760
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "secondName"
            Height          =   315
            Left            =   1935
            TabIndex        =   2
            Top             =   990
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Type of Contract"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   28
            Left            =   4410
            TabIndex        =   66
            Top             =   2340
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   27
            Left            =   6570
            TabIndex        =   65
            Top             =   2295
            Width           =   2760
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative Position"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   23
            Left            =   495
            TabIndex        =   43
            Top             =   1395
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tenant Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   495
            TabIndex        =   42
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tenant Presentative"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   450
            TabIndex        =   41
            Top             =   990
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Representative Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   495
            TabIndex        =   40
            Top             =   1755
            Width           =   1770
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tenant TIN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   765
            TabIndex        =   39
            Top             =   2295
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tenant Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   21
            Left            =   540
            TabIndex        =   38
            Top             =   270
            Width           =   915
         End
      End
   End
   Begin LVbuttons.LaVolpeButton cmdcancel 
      Height          =   330
      Left            =   4860
      TabIndex        =   22
      Top             =   6705
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
      MICON           =   "Fixfrm.frx":4FB8
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
      TabIndex        =   24
      Top             =   6705
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
      MICON           =   "Fixfrm.frx":4FD4
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
      Left            =   5760
      TabIndex        =   25
      Top             =   6705
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
      MICON           =   "Fixfrm.frx":4FF0
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
      TabIndex        =   26
      Top             =   7065
      Width           =   9660
      _ExtentX        =   17039
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
      Bindings        =   "Fixfrm.frx":500C
      Height          =   375
      Index           =   1
      Left            =   12915
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   63832065
      CurrentDate     =   36584
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   330
      Left            =   4140
      TabIndex        =   34
      Top             =   6705
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
      MICON           =   "Fixfrm.frx":5017
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
      Height          =   195
      Left            =   10125
      TabIndex        =   63
      Top             =   6615
      Visible         =   0   'False
      Width           =   960
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
      Left            =   12465
      TabIndex        =   35
      Top             =   4500
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
      Left            =   12690
      TabIndex        =   32
      Top             =   2610
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
      Left            =   12465
      TabIndex        =   31
      Top             =   3330
      Width           =   4515
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
      Left            =   12690
      TabIndex        =   30
      Top             =   2925
      Width           =   4515
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
      TabIndex        =   29
      Top             =   45
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   0
      Picture         =   "Fixfrm.frx":5033
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12285
   End
End
Attribute VB_Name = "CartFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim strSQL4 As String
Dim strSQL5 As String
Dim strSQL6 As String
Dim WithEvents adoPrimaryRS6 As ADODB.Recordset
Attribute adoPrimaryRS6.VB_VarHelpID = -1
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
    Text1(14).SetFocus
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
Rights1_Edit = 1
If Rights1_Edit = 1 Then
Dim oText As TextBox, i
    Dim odate As DTPicker, e
xcode = InputBox("Please Enter Supplier Code:", " Suppliers Information - Edit Mode")
If xcode <> "" Then
           strSQL2 = "Select [Tenant Name] as TenantName, " & _
                "[Tenant Presentative] as TenantPresentative, " & _
                "[Presentative Address] as PresentativeAddress, " & _
                "[Presentative resnumber] as Presentativeresnumber, " & _
                "[Presentative Place issued] as PresentativePlaceissued," & _
                "[leased premises amt] as leasedpremisesamt, " & _
                "[Leased Premises] as LeasedPremises, " & _
                "[Amount of rent] as Amountofrent,[First month] as Firstmonth, " & _
                "[Cusa Charges] as CusaCharges, " & _
                "[usage of Premises] as usageofPremises, " & _
                "[Rental Commen Date] as RentalCommenDate, " & _
                "[Lease Expiry Date] as LeaseExpiryDate, " & _
                "[Presentative date issued] as Presentativedateissued, " & _
                "[Aircon Charges] as AirconCharges, " & _
                "[Sectotalamt] as Sectotalamt1, " & _
                "[advancerent] as advancerent1 , " & _
                "[settotalamttotext] as settotalamttotext1," & _
                "[Settotalrenttotext] as Settotalrenttotext1," & _
                "[settotalamtrenttext] as settotalamtrenttext1," & _
                "[Presentative Position] as PresentativePosition , " & _
                "[TENANTid] as TENANTidt," & _
                "[Terms Year] as TermsYear," & _
                "[Company Code] as mk,[Terms to word] as termstoword," & _
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
'                Set Label1(22).DataSource = adoPrimaryRS2
                Set Text2.DataSource = adoPrimaryRS2
                Set Combo2.DataSource = adoPrimaryRS2
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
    strSQL6 = "SELECT *  FROM typeOfContract where [Contract Code] = '" & Combo2.Text & "'"
    Database_Refresh 5
    Label1(27).Caption = adoPrimaryRS6.Fields(1)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")


                Case 13
                    Text1(4).SetFocus
                Case x8
                
                Case Else
                     KeyAscii = 0
End Select
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
    Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
    Label1(17).Caption = Format(Label1(17).Caption, "##,##0.00")
    Text1(7).Text = Format(Text1(7).Text, "##,##0.00")
    Text1(8).Text = Format(Text1(8).Text, "##,##0.00")
    Text1(12).Text = Format(Text1(12).Text, "##,##0.00")
    
    
    strSQL6 = "SELECT [Contract Code]  FROM typeOfContract ORDER BY [Contract Code]"
    Database_Refresh 5
    If adoPrimaryRS6.RecordCount <> 0 Then
        adoPrimaryRS6.MoveFirst
        Do While Not adoPrimaryRS6.EOF
            Combo2.AddItem IIf(IsNull(adoPrimaryRS6("Contract Code")), "", adoPrimaryRS6("Contract Code"))
            adoPrimaryRS6.MoveNext
        Loop
    End If
    
    hanap_bistype
    
End Sub
Function hanap_bistype()
'On Error Resume Next
strSQL6 = "SELECT *  FROM typeOfContract where [Contract Code]= '" & Combo2.Text & "'"
                Database_Refresh 5
'                Text4.Text = adoPrimaryRS6.Fields(0)
                Label1(27).Caption = adoPrimaryRS6.Fields(1)
    
End Function
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
    ElseIf xMode = 5 Then
        Set adoPrimaryRS6 = New ADODB.Recordset
        adoPrimaryRS6.Open strSQL6, db, adOpenStatic, adLockOptimistic
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
                "[leased premises amt] as leasedpremisesamt, " & _
                "[Leased Premises] as LeasedPremises, " & _
                "[Amount of rent] as Amountofrent,[First month] as Firstmonth, " & _
                "[Cusa Charges] as CusaCharges, " & _
                "[Aircon Charges] as AirconCharges, " & _
                "[usage of Premises] as usageofPremises, " & _
                "[Rental Commen Date] as RentalCommenDate, " & _
                "[Lease Expiry Date] as LeaseExpiryDate, " & _
                "[Presentative date issued] as Presentativedateissued, " & _
                "[Sectotalamt] as Sectotalamt1, " & _
                "[advancerent] as advancerent1 , " & _
                "[settotalamttotext] as settotalamttotext1," & _
                "[Settotalrenttotext] as Settotalrenttotext1," & _
                "[settotalamtrenttext] as settotalamtrenttext1," & _
                "[Presentative Position] as PresentativePosition , " & _
                "[TENANTid] as TENANTidt," & _
                "[Terms Year] as TermsYear," & _
                "[Company Code] as mk,[Terms to word] as termstoword," & _
                "[second Name] as secondName, [Typeofbis] as mk1," & _
                "[Tenant TIN] as TenantTIN from [CONTRACT LEASE]  where [typeofbis] like '" & "02" & "'"
                Database_Refresh 1
                 '"[Tenant Name 2] as TenantName2," & _
                '"[Leased approximately] as Leasedapproximately, "
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
                Set Text2.DataSource = adoPrimaryRS2
                Set Combo2.DataSource = adoPrimaryRS2
                
                
                'Set dtStart(1).DataSource = adoPrimaryRS2
'                Set Label1(22).DataSource = adoPrimaryRS2
                Set Text2.DataSource = adoPrimaryRS2
                  'If adoPrimaryRS2.RecordCount <> 0 Then
                  '  adoPrimaryRS2.MoveFirst
                    Set txtCombo.DataSource = adoPrimaryRS2
                  'End If
'                 If adoPrimaryRS2.RecordCount <> 0 Then
 '                   adoPrimaryRS2.MoveFirst
                    Set Combo1.DataSource = adoPrimaryRS2
  '               End If
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
For i = 0 To 13
   Text1(i).Text = ""
   Me.Label1(16).Caption = ""
   Text2.Text = "01"
   Text4.Text = "02"
Next i
End Function


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
                        Text1(7).SetFocus
                     End If
                Case x8
                
                Case Else
                     KeyAscii = 0
        End Select
 End If
If KeyAscii = 13 Then
    If Index = 14 Then Text1(0).SetFocus
    If Index = 0 Then Combo1.SetFocus
    If Index = 1 Then Text1(15).SetFocus
    If Index = 15 Then Text1(2).SetFocus
    If Index = 2 Then Text1(3).SetFocus
    If Index = 3 Then Combo2.SetFocus
    If Index = 4 Then Text3.SetFocus
    If Index = 5 Then txtCombo.SetFocus
    If Index = 6 Then Text1(13).SetFocus
    'If Index = 13 Then Text1(16).SetFocus
    If Index = 16 Then Text1(18).SetFocus
    If Index = 18 Then Text1(7).SetFocus
    If Index = 7 Then
        Text1(8).SetFocus
        Dim sn As String
        Dim sn1 As String
        sn = Val(Text1(7).Text * secdep)
        sn1 = Val(Text1(7).Text * adrent)
        Label1(16).Caption = Format(sn, "##,##0.00")
        Label1(17).Caption = Format(sn1, "##,##0.00")
        Label1(19).Caption = Amt2Words(Label1(16).Caption)
        Label1(20).Caption = Amt2Words(Label1(17).Caption)
        Label1(25).Caption = Amt2Words(Text1(7).Text)
    End If
    If Index = 8 Then Text1(9).SetFocus
    'If Index = 12 Then Text1(9).SetFocus
    If Index = 9 Then Text1(10).SetFocus
    If Index = 10 Then Text1(11).SetFocus
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
For i = 0 To 16
    Text1(i).Enabled = False
Next i
  Label1(16).Enabled = False
  Label1(17).Enabled = False
  dtStart(1).Enabled = False
  Label1(27).Enabled = False
  Combo1.Enabled = False
  Combo2.Enabled = False
  txtCombo.Enabled = False
  Text3.Enabled = False
  Text1(18).Enabled = False
End Function
Function unlocking()
For i = 0 To 16
    Text1(i).Enabled = True
Next i
Label1(16).Enabled = True
Text1(18).Enabled = True
  Label1(17).Enabled = True
  dtStart(1).Enabled = True
  Label1(27).Enabled = True
  Combo2.Enabled = True
  Combo1.Enabled = True
  Text3.Enabled = True
  txtCombo.Enabled = True
End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1(5).SetFocus
End Sub






