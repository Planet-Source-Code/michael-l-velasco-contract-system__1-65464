VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "Lvbuttons.ocx"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List of Tenants"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9135
   Icon            =   "ListofTenants.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1755
      TabIndex        =   0
      Top             =   675
      Width           =   3300
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   1845
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196609
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   609
      DesignHeight    =   463
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   6705
      Width           =   9135
      _ExtentX        =   16113
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
   Begin LVbuttons.LaVolpeButton cmd_op 
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   2
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancel"
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
      MICON           =   "ListofTenants.frx":144A
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
   Begin MSDataGridLib.DataGrid GroupGrid 
      Height          =   5490
      Left            =   45
      TabIndex        =   3
      Top             =   1185
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   9684
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "TENANTid"
         Caption         =   "TENANT Code"
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
         DataField       =   "Tenant Name"
         Caption         =   "Tenant Name"
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
      BeginProperty Column02 
         DataField       =   "Presentative Position"
         Caption         =   "Presentative Position"
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
      BeginProperty Column03 
         DataField       =   "Tenant Presentative"
         Caption         =   "Tenant Presentative"
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
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2759.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2174.74
         EndProperty
         BeginProperty Column03 
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
      Left            =   90
      TabIndex        =   5
      Top             =   765
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
      TabIndex        =   4
      Top             =   90
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "ListofTenants.frx":1466
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9315
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SQL1 As String
Dim WithEvents adoPriRS As ADODB.Recordset
Attribute adoPriRS.VB_VarHelpID = -1
'Dim db As String

Dim strSQL As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS3 As ADODB.Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Dim i As Integer

Private Sub cmd_op_Click(Index As Integer)
If Index = 4 Then
Unload Me
ElseIf Index = 0 Then
     msgs = MsgBox("Are you sure want to delete this record", vbInformation + vbOKCancel)
'     X = Label1(11).Caption
If t1 = "" Then
   Exit Sub
Else

     If msgs = vbOK Then
        strSQL2 = "select *  from [CONTRACT LEASE] WHERE [TENANTid] = '" & t1 & "'"
           Database_Refresh 1
                With adoPrimaryRS2
                     .Delete
                     .Update
                     .Requery
                     .Close
                End With
     End If
     Call Name_supp12
End If
End If
End Sub

Private Sub Form_Load()
  
      Call Name_supp12
      'Call lucking
End Sub
Public Sub Database_Refresh(xMode As Integer)
    ' PRE-DATABASE CONNECTION WITH PARAMETERIZED SQL VARIABLES ATTACHED IN EVERY MODE
    'On Error Resume Next
   
    If xMode = 0 Then
        Set adoPriRS = New ADODB.Recordset
        adoPriRS.Open SQL1, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 1 Then
        Set adoPrimaryRS2 = New ADODB.Recordset
        adoPrimaryRS2.Open strSQL2, db, adOpenStatic, adLockOptimistic
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
  GroupGrid.Refresh
  GroupGrid.ReBind
  dview12
End Sub



Private Sub Form_Unload(Cancel As Integer)
If choice = "Items" Then
Me.Hide
xcode = InputBox("Please Enter Quantity:", " Enter Quantity")
        If xcode <> "" Then
                ProQTY = xcode
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
On Error Resume Next
 t1 = GroupGrid.Columns(0)
End Sub


Private Sub Text1_Change(Index As Integer)
If Index = 14 Then
SQL1 = "SELECT [TENANTid], [Tenant Name], [Presentative Position], [Tenant Presentative] from [CONTRACT LEASE] where [TENANTid] like '" & Trim(Text1(14).Text) & "%'"
     Database_Refresh 0
     Call dview112
End If
End Sub


