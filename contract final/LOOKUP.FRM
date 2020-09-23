VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Begin VB.Form lookup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LookuP"
   ClientHeight    =   3225
   ClientLeft      =   7410
   ClientTop       =   2475
   ClientWidth     =   5205
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   1845
      Top             =   1215
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196609
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   347
      DesignHeight    =   215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "TENANTidt"
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
      Left            =   1845
      TabIndex        =   2
      Top             =   2565
      Width           =   3300
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   2985
      Width           =   5205
      _ExtentX        =   9181
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
   Begin MSDataGridLib.DataGrid GroupGrid 
      Height          =   1635
      Left            =   45
      TabIndex        =   4
      Top             =   900
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   2884
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "List"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tenants Search"
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
      Index           =   21
      Left            =   180
      TabIndex        =   3
      Top             =   2655
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "List of Tenants"
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
      TabIndex        =   0
      Top             =   90
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "lookup.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9090
   End
End
Attribute VB_Name = "lookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strDB1 As String
Dim SQL1 As String
Dim WithEvents adoPriRS As ADODB.Recordset
Attribute adoPriRS.VB_VarHelpID = -1
'Dim db As String

Dim i As Integer
Private Sub Form_Load()
  strDB1 = App.Path + "\data.MDB;Jet OLEDB:Database Password=;"
      Call Name_supp12
      'Call lucking
End Sub
Public Sub Database_Refresh(xMode As Integer)
    ' PRE-DATABASE CONNECTION WITH PARAMETERIZED SQL VARIABLES ATTACHED IN EVERY MODE
    'On Error Resume Next
    Set db = New Connection
        db.CursorLocation = adUseClient
        db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB1
    If xMode = 0 Then
        Set adoPriRS = New ADODB.Recordset
        adoPriRS.Open SQL1, db, adOpenStatic, adLockOptimistic
    End If
End Sub
Private Sub dview12()
Do While Not adoPriRS.EOF
        Set GroupGrid.DataSource = adoPriRS
        adoPriRS.MoveNext
Loop
End Sub


Public Sub Name_supp12()

    SQL1 = "SELECT [TENANTid], [Tenant Name], [Presentative Position], [Tenant Presentative] from [CONTRACT LEASE] order by [TENANTid]"

  Database_Refresh 0
  GroupGrid.ClearSelCols
     dview12
End Sub



Private Sub Form_Unload(Cancel As Integer)
If choice = "Items" Then
Me.Hide
xCode = InputBox("Please Enter Quantity:", " Enter Quantity")
        If xCode <> "" Then
                ProQTY = xCode
                ReqdataFrm.SaveEntryItem
                Unload Me
        End If
End If
End Sub

Private Sub dview112()
    Do While Not adoPriRS.EOF
            Set GroupGrid.DataSource = adoPriRS
            adoPriRS.MoveNext
    Loop
End Sub


Private Sub GroupGrid_Click()
 t2 = GroupGrid.Columns(0)
 Viewtelfrm.Combo1.Text = t2
 Unload Me
End Sub

Private Sub Text1_Change(Index As Integer)

   SQL1 = "SELECT [TENANTid], [Tenant Name], [Presentative Position], [Tenant Presentative] from [CONTRACT LEASE] where [Tenant Name] like '" & Trim(Text2) & "%'"
     Database_Refresh 0
     Call dview112
End Sub
