VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fPasswordSecurity 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password Security"
   ClientHeight    =   6915
   ClientLeft      =   690
   ClientTop       =   780
   ClientWidth     =   8700
   Icon            =   "FrmPasswordSecurity.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H000000C0&
      Height          =   2775
      Left            =   45
      TabIndex        =   19
      Top             =   855
      Width           =   8610
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "User_Description"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Index           =   3
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   1320
         Width           =   6255
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "User_Password"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2160
         TabIndex        =   21
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "User_Name"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   27
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Enter Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   26
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1170
         TabIndex        =   25
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   24
         Top             =   255
         Width           =   945
      End
      Begin VB.Image Image2 
         Height          =   1365
         Left            =   180
         Picture         =   "FrmPasswordSecurity.frx":000C
         Stretch         =   -1  'True
         Top             =   1350
         Width           =   1365
      End
   End
   Begin VB.PictureBox picButtons 
      BackColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   45
      ScaleHeight     =   405
      ScaleWidth      =   8535
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6180
      Width           =   8595
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   18
         Top             =   50
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         TabIndex        =   17
         Top             =   225
         Width           =   1335
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         TabIndex        =   16
         Top             =   50
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   15
         Top             =   50
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         TabIndex        =   14
         Top             =   50
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Undo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4320
         TabIndex        =   13
         Top             =   50
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   300
      Left            =   45
      MouseIcon       =   "FrmPasswordSecurity.frx":059E
      MousePointer    =   99  'Custom
      Picture         =   "FrmPasswordSecurity.frx":08A8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5865
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   300
      Left            =   390
      MouseIcon       =   "FrmPasswordSecurity.frx":0BEA
      MousePointer    =   99  'Custom
      Picture         =   "FrmPasswordSecurity.frx":0EF4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5865
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      Height          =   300
      Left            =   7980
      MouseIcon       =   "FrmPasswordSecurity.frx":1236
      MousePointer    =   99  'Custom
      Picture         =   "FrmPasswordSecurity.frx":1540
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5865
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdLast 
      Height          =   300
      Left            =   8325
      MouseIcon       =   "FrmPasswordSecurity.frx":1882
      MousePointer    =   99  'Custom
      Picture         =   "FrmPasswordSecurity.frx":1B8C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5865
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H000000C0&
      Height          =   2175
      Left            =   45
      TabIndex        =   1
      Top             =   3600
      Width           =   8610
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cart with Cons."
         DataField       =   "User_Rights3_CarwtConse"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   6570
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   43
         Top             =   675
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Statement of Account"
         DataField       =   "User_Rights2_Sales_Report"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   6570
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   42
         Top             =   270
         Width           =   1455
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Percentage With Cons."
         DataField       =   "User_Rights2_Inventory_Report"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   4005
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   41
         Top             =   1755
         Width           =   2355
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cart with Cons."
         DataField       =   "User_Rights3_CarwtConse"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   4005
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   40
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shop with Cons."
         DataField       =   "User_Rights2_Post_ReceivingOrders"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   19
         Left            =   4005
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   39
         Top             =   1125
         Width           =   1755
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Percentage"
         DataField       =   "User_Rights2_Post_SalesOrders"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   4005
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   38
         Top             =   810
         Width           =   1485
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cart Only"
         DataField       =   "User_Rights2_ReceivingOrders"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   4005
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   37
         Top             =   495
         Width           =   1455
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shop Only"
         DataField       =   "User_Rights2_PurchaseOrders"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   4005
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   36
         Top             =   225
         Width           =   1815
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete Module (List)"
         DataField       =   "User_Rights2_SalesOrders"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   35
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Location Module"
         DataField       =   "User_Rights2_Supplier"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   34
         Top             =   1530
         Width           =   2055
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sex Module"
         DataField       =   "User_Rights2_Menu"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   33
         Top             =   1260
         Width           =   1935
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Company Module"
         DataField       =   "User_Rights2_Ingredients"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   32
         Top             =   945
         Width           =   2415
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Company"
         DataField       =   "User_Rights2_Tables"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   225
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   31
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report Module"
         DataField       =   "User_Rights2_Service_Crew"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   225
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   30
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Record"
         DataField       =   "User_Rights1_Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   7
         Top             =   225
         Width           =   1455
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit Record"
         DataField       =   "User_Rights1_Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   225
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   6
         Top             =   495
         Width           =   1455
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete Record"
         DataField       =   "User_Rights1_Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   5
         Top             =   810
         Width           =   1575
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Backup"
         DataField       =   "User_Rights3_Backup"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   225
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   4
         Top             =   1755
         Width           =   1095
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Restore"
         DataField       =   "User_Rights3_Restore"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   17
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   3
         Top             =   225
         Width           =   1095
      End
      Begin VB.CheckBox chkFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password Security"
         DataField       =   "User_Rights3_Password_Security"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   18
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   2
         Top             =   585
         Width           =   2055
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   6675
      Width           =   8700
      _ExtentX        =   15346
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
   Begin MSComDlg.CommonDialog CDlgExcel 
      Left            =   1800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Password Code"
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
      Left            =   135
      TabIndex        =   29
      Top             =   90
      Width           =   5175
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   735
      TabIndex        =   28
      Top             =   5865
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "FrmPasswordSecurity.frx":1ECE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9090
   End
End
Attribute VB_Name = "fPasswordSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: PASSWORD SECURITY UTILITIES
' VERSION: VB6

Dim strSQL As String

Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Public strSQL2 As String
Public WithEvents adoPrimaryRS2 As Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim oText As TextBox
Dim xDeleteLogic As Boolean
Dim intColIdx As Integer
Dim blnListShow As Boolean
Dim intKeyCode As Integer
Dim xButton As Integer

Private Sub Form_Load()
    ' STARTUP MODULE FOR PASSWORD SECURITY
    Dim oText As TextBox, oCheckBox As CheckBox
    blnListShow = False
    strDB = App.Path + "\data.MDB;Jet OLEDB:Database Password=;"
    strSQL = "SELECT * FROM Password_Security ORDER BY User_Name"
    Database_Refresh 0
    For Each oText In Me.txtFields
        Set oText.DataSource = adoPrimaryRS
    Next
    For Each oCheckBox In Me.chkFields
        Set oCheckBox.DataSource = adoPrimaryRS
    Next
    Call DisplayRestrictions
    SetButtons True
    xDeleteLogic = False
    EditClicked = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Controlling the Buttons for adoPrimaryRS Recordset.
    If mbEditFlag Or mbAddNewFlag Then Exit Sub
    Select Case KeyCode
            Case vbKeyEscape
                    cmdClose_Click
            Case vbKeyEnd
                    cmdLast_Click
            Case vbKeyHome
                    cmdFirst_Click
            Case vbKeyUp, vbKeyPageUp
                    If Shift = vbCtrlMask Then
                        cmdFirst_Click
                    Else
                        cmdPrevious_Click
                    End If
            Case vbKeyDown, vbKeyPageDown
                    If Shift = vbCtrlMask Then
                        cmdLast_Click
                    Else
                        cmdNext_Click
                    End If
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' MOUSE NATURE DEFAULT
    Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' adoPrimaryRS RECORD NUMBER
    On Error Resume Next
    lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' ADDITIONAL BUT OPTIONAL VALIDATIONS
    Dim bCancel As Boolean
    Select Case adReason
            Case adRsnAddNew
            Case adRsnClose
            Case adRsnDelete
            Case adRsnFirstChange
            Case adRsnMove
            Case adRsnRequery
            Case adRsnResynch
            Case adRsnUndoAddNew
            Case adRsnUndoDelete
            Case adRsnUndoUpdate
            Case adRsnUpdate
    End Select
    If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
    ' ADD/NEW BUTTON
    On Error GoTo AddErr
    Rights1_Add = 1
    If Rights1_Add = 1 Then
        With adoPrimaryRS
            If Not (.BOF And .EOF) Then
                mvBookMark = .Bookmark
            End If
            .AddNew
            mbAddNewFlag = True
            SetButtons False
        End With
        xDeleteLogic = True
        EditClicked = True
        For i = 0 To 3
            txtFields(i).Enabled = True
        Next i
        txtFields(0).SetFocus
        For i = 0 To 19
            chkFields(i).Enabled = True
        Next i
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
AddErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdEdit_Click()
    ' EDIT BUTTON
    On Error GoTo EditErr
    If Rights1_Edit = 1 Then
        mbEditFlag = True
        SetButtons False
        xDeleteLogic = False
        EditClicked = True
        txtFields(1) = UnCode_Pass(txtFields(1))
        For i = 1 To 3
            txtFields(i).Enabled = True
        Next i
        txtFields(1).SetFocus
        For i = 0 To 19
            chkFields(i).Enabled = True
        Next i
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
EditErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdDelete_Click()
    ' DELETE BUTTON
    On Error GoTo DeleteErr
    If Rights1_Delete = 1 Then
        Msg = MsgBox("Do you want to delete User Name " & txtFields(0) & "?", vbYesNo + vbExclamation + vbDefaultButton2, _
                      "Warning:End-User:" + UserName)
        If Msg = vbYes Then
            With adoPrimaryRS
                .Delete
                .MoveNext
                If .EOF Then .MoveLast
            End With
        End If
    Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
    End If
    Exit Sub
DeleteErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdRefresh_Click()
    ' REFRESH BUTTON
    On Error GoTo RefreshErr
    adoPrimaryRS.Requery
    Call DisplayRestrictions
    Exit Sub
RefreshErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdCancel_Click()
    ' UNDO BUTTON
    On Error Resume Next
    SetButtons True
    mbEditFlag = False
    mbAddNewFlag = False
    adoPrimaryRS.CancelUpdate
    If mvBookMark > 0 Then
        adoPrimaryRS.Bookmark = mvBookMark
    Else
        adoPrimaryRS.MoveFirst
    End If
    For i = 0 To 3
        txtFields(i).Enabled = False
    Next i
    For i = 0 To 19
        chkFields(i).Enabled = False
    Next i
    mbDataChanged = False
    EditClicked = False
End Sub

Private Sub cmdUpdate_Click()
    ' SAVE BUTTON
    'On Error GoTo UpdateErr
    Dim ActiveBlankFields As String
    If txtFields(0) = "" Then
        ActiveBlankFields = ActiveBlankFields + "User Name"
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(1) = "" Then
        If txtFields(0) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Password"
        Else
            ActiveBlankFields = ActiveBlankFields + "Password"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If txtFields(3) = "" Then
        If txtFields(1) = "" Then
            ActiveBlankFields = ActiveBlankFields + ",Description"
        Else
            ActiveBlankFields = ActiveBlankFields + "Description"
        End If
    Else
        ActiveBlankFields = ""
    End If
    If ActiveBlankFields = "" Then
        If EditClicked = True Then
            txtFields(1) = Decode_Pass(txtFields(1))
        End If
        adoPrimaryRS.UpdateBatch adAffectAll
        If mbAddNewFlag Then
            adoPrimaryRS.MoveLast              'move to the new record
        End If
        mbEditFlag = False
        mbAddNewFlag = False
        SetButtons True
        mbDataChanged = False
        If xGridLogic = True Then
            cmdRefresh_Click
            xGridLogic = False
        End If
        For i = 0 To 3
            txtFields(i).Enabled = False
        Next i
        For i = 0 To 19
            chkFields(i).Enabled = False
        Next i
        EditClicked = False
    Else
        MsgBox ActiveBlankFields & " is empty!!", vbOKOnly + vbCritical, " Warning:End-User" + UserName
    End If
    Exit Sub
UpdateErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Private Sub cmdClose_Click()
    ' CLOSE BUTTON - EXIT
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    ' TOP/FIRST BUTTON
    On Error Resume Next
    Dim Msg
    adoPrimaryRS.MoveFirst
    mbDataChanged = False
    Exit Sub
GoFirstError:
    Msg = MsgBox(Err.Description, vbOKOnly, "Validation:End-User")
End Sub

Private Sub cmdLast_Click()
    ' BOTTOM/LAST BUTTON
    On Error Resume Next
    Dim Msg
    adoPrimaryRS.MoveLast
    mbDataChanged = False
    Exit Sub
GoLastError:
    Msg = MsgBox(Err.Description, vbOKOnly, "Validation:End-User")
End Sub

Private Sub cmdNext_Click()
    ' NEXT BUTTON
    On Error Resume Next
    Dim Msg
    If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
    If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveLast
    End If
    mbDataChanged = False
    Exit Sub
GoNextError:
    Msg = MsgBox(Err.Description, vbOKOnly, "Validation:End-User")
End Sub

Private Sub cmdPrevious_Click()
    ' PREVIOUS BUTTON
    On Error Resume Next
    Dim Msg
    If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
    If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveFirst
    End If
    mbDataChanged = False
    Exit Sub
GoPrevError:
    Msg = MsgBox(Err.Description, vbOKOnly, "Validation:End-User")
End Sub

Private Sub SetButtons(bVal As Boolean)
    ' COMMAND BUTTONS ENABLED MODES
    On Error GoTo ErrorSetButtons
    cmdAdd.Visible = bVal
    cmdedit.Visible = bVal
    cmdUpdate.Visible = Not bVal
    cmdcancel.Visible = Not bVal
    cmdDelete.Visible = bVal
    cmdClose.Visible = bVal
    cmdNext.Visible = bVal
    cmdFirst.Visible = bVal
    cmdLast.Visible = bVal
    cmdPrevious.Visible = bVal
    Exit Sub
ErrorSetButtons:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Public Sub Database_Refresh(xMode As Integer)
    ' DATABASE CONNECTIVITY SETTINGS
    'On Error Resume Next
    'Set db = New Connection
    '    db.CursorLocation = adUseClient
        'db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & strDB
    '    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB
    If xMode = 0 Then
        Set adoPrimaryRS = New Recordset
        adoPrimaryRS.Open strSQL, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 1 Then
        Set adoPrimaryRS2 = New Recordset
        adoPrimaryRS2.Open strSQL2, db, adOpenStatic, adLockOptimistic
    End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
    ' txtFields(0) VALIDATION KEY INPUTTED
    On Error GoTo ErrorTxtFieldsFocus
    txtFields(0).Text = UCase(txtFields(0).Text)
    If Index = 0 Then
        If Get_User_Name Then
                Msg = MsgBox("User Name already exist!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
                txtFields(0) = ""
                txtFields(0).SetFocus
        ElseIf txtFields(0) = "" Then
                Msg = MsgBox("User Name cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
                cmdCancel_Click
        End If
    ElseIf Index = 2 Then
        If txtFields(1) <> txtFields(2) Then
            Msg = MsgBox("Password does not match!!" & vbCrLf & "Please Re-Enter your Password.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
            txtFields(1) = ""
            txtFields(2) = ""
            txtFields(1).SetFocus
        ElseIf txtFields(1) = "" Then
            Msg = MsgBox("Password cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
            txtFields(1).SetFocus
        ElseIf txtFields(2) = "" Then
            Msg = MsgBox("Re-Entered Password cannot be empty!!", vbOKOnly + vbCritical, "Warning:End-User:" + UserName)
            txtFields(2).SetFocus
        End If
    End If
    Exit Sub
ErrorTxtFieldsFocus:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub

Function Get_User_Name() As Boolean
    ' USER NAME VALIDATION ON txtFields(0)
    On Error Resume Next
    strSQL2 = "SELECT * FROM Password_Security WHERE User_Name = '" & txtFields(0) & "'"
    Database_Refresh 1
    If adoPrimaryRS2.AbsolutePosition <> -1 Then
        Get_User_Name = True
    Else
        Get_User_Name = False
    End If
End Function

Function FindTag_PasswordSecurity()
    ' FOR FIND BUTTON FUNCTION USE
    On Error Resume Next
    Dim oText As TextBox, oCheckBox As CheckBox
    Do While adoPrimaryRS.Fields("User_Name") <> Trim(Finder.txtWord)
        adoPrimaryRS.MoveNext
    Loop
    For Each oText In Me.txtFields
        Set oText.DataSource = adoPrimaryRS
    Next
        For Each oCheckBox In Me.chkFields
        Set oCheckBox.DataSource = adoPrimaryRS
    Next
    Call DisplayRestrictions
End Function

Function DisplayRestrictions()
    ' CHECKBOX DISPLAY RESTRICTIONS/VALUES
    If adoPrimaryRS.RecordCount <> 0 Then
        chkFields(0) = IIf(IsNull(adoPrimaryRS("User_Rights1_Add")), 0, IIf(adoPrimaryRS("User_Rights1_Add") = 0, 0, 1))
        chkFields(1) = IIf(IsNull(adoPrimaryRS("User_Rights1_Edit")), 0, IIf(adoPrimaryRS("User_Rights1_Edit") = 0, 0, 1))
        chkFields(2) = IIf(IsNull(adoPrimaryRS("User_Rights1_Delete")), 0, IIf(adoPrimaryRS("User_Rights1_Delete") = 0, 0, 1))
        chkFields(3) = IIf(IsNull(adoPrimaryRS("User_Rights2_Tables")), 0, IIf(adoPrimaryRS("User_Rights2_Tables") = 0, 0, 1))
        chkFields(4) = IIf(IsNull(adoPrimaryRS("User_Rights2_Service_Crew")), 0, IIf(adoPrimaryRS("User_Rights2_Service_Crew") = 0, 0, 1))
        chkFields(5) = IIf(IsNull(adoPrimaryRS("User_Rights2_Ingredients")), 0, IIf(adoPrimaryRS("User_Rights2_Ingredients") = 0, 0, 1))
        chkFields(6) = IIf(IsNull(adoPrimaryRS("User_Rights2_Menu")), 0, IIf(adoPrimaryRS("User_Rights2_Menu") = 0, 0, 1))
        chkFields(7) = IIf(IsNull(adoPrimaryRS("User_Rights2_Supplier")), 0, IIf(adoPrimaryRS("User_Rights2_Supplier") = 0, 0, 1))
        chkFields(8) = IIf(IsNull(adoPrimaryRS("User_Rights2_SalesOrders")), 0, IIf(adoPrimaryRS("User_Rights2_SalesOrders") = 0, 0, 1))
        chkFields(9) = IIf(IsNull(adoPrimaryRS("User_Rights2_PurchaseOrders")), 0, IIf(adoPrimaryRS("User_Rights2_PurchaseOrders") = 0, 0, 1))
        chkFields(10) = IIf(IsNull(adoPrimaryRS("User_Rights2_ReceivingOrders")), 0, IIf(adoPrimaryRS("User_Rights2_ReceivingOrders") = 0, 0, 1))
        chkFields(11) = IIf(IsNull(adoPrimaryRS("User_Rights2_Post_SalesOrders")), 0, IIf(adoPrimaryRS("User_Rights2_Post_SalesOrders") = 0, 0, 1))
        chkFields(12) = IIf(IsNull(adoPrimaryRS("User_Rights2_Post_SalesOrders")), 0, IIf(adoPrimaryRS("User_Rights2_ReceivingOrders") = 0, 0, 1))
        chkFields(13) = IIf(IsNull(adoPrimaryRS("User_Rights2_Inventory_Report")), 0, IIf(adoPrimaryRS("User_Rights2_Inventory_Report") = 0, 0, 1))
        chkFields(14) = IIf(IsNull(adoPrimaryRS("User_Rights2_Sales_Report")), 0, IIf(adoPrimaryRS("User_Rights2_Sales_Report") = 0, 0, 1))
        chkFields(15) = IIf(IsNull(adoPrimaryRS("User_Rights2_Critical_Report")), 0, IIf(adoPrimaryRS("User_Rights2_Critical_Report") = 0, 0, 1))
        chkFields(16) = IIf(IsNull(adoPrimaryRS("User_Rights3_Backup")), 0, IIf(adoPrimaryRS("User_Rights3_Backup") = 0, 0, 1))
        chkFields(17) = IIf(IsNull(adoPrimaryRS("User_Rights3_Restore")), 0, IIf(adoPrimaryRS("User_Rights3_Restore") = 0, 0, 1))
        chkFields(18) = IIf(IsNull(adoPrimaryRS("User_Rights3_Password_Security")), 0, IIf(adoPrimaryRS("User_Rights3_Password_Security") = 0, 0, 1))
        chkFields(19) = IIf(IsNull(adoPrimaryRS("User_Rights3_CarwtConse")), 0, IIf(adoPrimaryRS("User_Rights3_CarwtConse") = 0, 0, 1))
    End If
End Function


