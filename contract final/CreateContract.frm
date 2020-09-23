VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Createfrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Type of Contract Module"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5985
   Icon            =   "CreateContract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "CreateContract.frx":4F0A
      Left            =   2340
      List            =   "CreateContract.frx":4F0C
      TabIndex        =   13
      Top             =   1395
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Company Name"
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
      Height          =   285
      Index           =   1
      Left            =   2340
      TabIndex        =   1
      Top             =   990
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Company Code"
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
      Height          =   285
      Index           =   0
      Left            =   2340
      TabIndex        =   0
      Top             =   630
      Width           =   2940
   End
   Begin MSDataGridLib.DataGrid GroupGrid 
      Height          =   1635
      Left            =   90
      TabIndex        =   2
      Top             =   2565
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   2884
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   3
      FormatLocked    =   -1  'True
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "List"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Contract Code"
         Caption         =   "Code"
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
         DataField       =   "Contract Name"
         Caption         =   "Name"
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
         DataField       =   "TypeS"
         Caption         =   "Type"
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
            DividerStyle    =   3
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2819.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   659.906
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   45
      TabIndex        =   3
      Top             =   2025
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Add Group"
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
      MICON           =   "CreateContract.frx":4F0E
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
   Begin LVbuttons.LaVolpeButton cmdEdit 
      Height          =   375
      Left            =   1530
      TabIndex        =   4
      Top             =   2025
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Edit"
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
      MICON           =   "CreateContract.frx":4F2A
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
      Height          =   375
      Index           =   0
      Left            =   1530
      TabIndex        =   5
      Top             =   2025
      Visible         =   0   'False
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
      MICON           =   "CreateContract.frx":4F46
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
      Height          =   375
      Left            =   2970
      TabIndex        =   6
      Top             =   2025
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
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
      MICON           =   "CreateContract.frx":4F62
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
   Begin LVbuttons.LaVolpeButton cmdDelete 
      Height          =   375
      Left            =   2970
      TabIndex        =   7
      Top             =   2025
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Delete"
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
      MICON           =   "CreateContract.frx":4F7E
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
      Left            =   4455
      TabIndex        =   8
      Top             =   2025
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
      MICON           =   "CreateContract.frx":4F9A
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
      TabIndex        =   9
      Top             =   4305
      Width           =   5985
      _ExtentX        =   10557
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   3330
      TabIndex        =   15
      Top             =   1395
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type contract"
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
      Index           =   2
      Left            =   855
      TabIndex        =   14
      Top             =   1485
      Width           =   1350
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Type of Contract"
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
      TabIndex        =   12
      Top             =   45
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description:"
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
      Left            =   1080
      TabIndex        =   11
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Code:"
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
      Index           =   0
      Left            =   1530
      TabIndex        =   10
      Top             =   630
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   0
      Picture         =   "CreateContract.frx":4FB6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9090
   End
End
Attribute VB_Name = "Createfrm"
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
Dim choice As String
Dim i As Integer

Private Sub cmd_Click()
On Error GoTo AddErr
 If Rights1_Add = 1 Then
    Call unlocking
    choice = "Add"
    cmdDelete.Visible = False
    cmd.Visible = False
    cmdedit.Visible = False
    cmdDone.Visible = False
    cmd_op(0).Visible = True
    cmdcancel.Visible = True
Else
        MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End If
    Exit Sub
AddErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Warning:End-User:" + UserName
End Sub
Private Sub dview()
Do While Not adoPrimaryRS.EOF
        Set GroupGrid.DataSource = adoPrimaryRS
        adoPrimaryRS.MoveNext
Loop
End Sub


Public Sub Name_supp()
strSQL = "SELECT * from [typeOfContract] order by [Contract Code]"
  Database_Refresh 0
  GroupGrid.ClearSelCols
     dview
End Sub

Private Sub cmd_op_Click(Index As Integer)
       If choice = "Edit" Then
           strSQL3 = "SELECT [Contract Name] from [typeOfContract] where [Contract Code] = '" & Trim(Text1(0).Text) & "'"
           Database_Refresh 2
           With adoPrimaryRS3
                .Fields(0) = Text1(1).Text
                .Update
                .Requery
                .Close
           End With
              Call Name_supp
              Call clearing
              Call locking
              cmdDelete.Visible = True
              cmd.Visible = True
              cmdedit.Visible = True
              cmdDone.Visible = True
              cmd_op(0).Visible = False
              cmdcancel.Visible = False
        ElseIf choice = "Add" Then
                If Len(Text1(0).Text) <> 0 Then
                        'On Error GoTo A1:
                         strSQL2 = "SELECT * FROM [typeOfContract]"
                        Database_Refresh 1
                        With adoPrimaryRS2
                                .AddNew
                                 .Fields(0) = Text1(0).Text
                                 .Fields(1) = Text1(1).Text
                                 .Fields(2) = Combo1.Text
                                .Update
                                .Requery
                                .Close
                        End With
                        Call Name_supp
                        Call clearing
                         Call locking
                        cmdDelete.Visible = True
                        cmd.Visible = True
                        cmdedit.Visible = True
                        cmdDone.Visible = True
                        cmd_op(0).Visible = False
                        cmdcancel.Visible = False
                Else
                        MsgBox "Enter Item type ...", vbInformation, "You can not save Zero length Item name ..."
                End If
                

        End If
Exit Sub
A1:
MsgBox "Duplicate Item name Found ..." & vbCrLf & "Enter Another name of Close this form ...", vbCritical, "Duplicate Entry Found ..."
End Sub

Private Sub cmdCancel_Click()
cmdDelete.Visible = True
cmd.Visible = True
cmdedit.Visible = True
cmdDone.Visible = True
cmd_op(0).Visible = False
cmdcancel.Visible = False
End Sub

Private Sub cmdDelete_Click()
strSQL3 = "SELECT * from [typeOfContract] where [Contract Code] = '" & Trim(Text1(0).Text) & "'"
   Database_Refresh 2
   With adoPrimaryRS3
        .Delete
        .Update
        .Requery
        .Close
   End With
      Call clearing
      Call Name_supp
'      adoPrimaryRS.Close
'      GroupGrid.ClearFields
'strSQL1 = "SELECT * from [Group Item] order by [Group Code]"
'  Database_Refresh 1
  'GroupGrid.ClearSelCols
  
'     dview1
End Sub

Private Sub cmdDone_Click()
Unload Me
GroupGrid.ClearSelCols
End Sub

Private Sub cmdEdit_Click()
 If Rights1_Edit = 1 Then
    Call unlocking
    choice = "Edit"
    cmdDelete.Visible = False
    cmd.Visible = False
    cmdedit.Visible = False
    cmdDone.Visible = False
    cmd_op(0).Visible = True
    cmdcancel.Visible = True
 Else
    MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
 End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "S" Then Label1(3).Caption = "SHOP"
If Combo1.Text = "P" Then Label1(3).Caption = "PERCENTAGES"
If Combo1.Text = "C" Then Label1(3).Caption = "CART"
End Sub

Private Sub Form_Load()
  
      Call Name_supp
      Call locking
      Combo1.AddItem "S"
      Combo1.AddItem "P"
      Combo1.AddItem "C"
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
Function locking()
For i = 0 To 1
    Text1(i).Locked = True
Next i
End Function
Function unlocking()
For i = 0 To 1
    Text1(i).Locked = False
Next i
End Function
Function clearing()
For i = 0 To 1
   Text1(i).Text = ""
Next i
End Function

Private Sub GroupGrid_Click()
On Error Resume Next
Text1(1).Text = GroupGrid.Columns(1)
Text1(0).Text = GroupGrid.Columns(0)
Combo1.Text = GroupGrid.Columns(2)
Call locking
End Sub

Private Sub LaVolpeButton1_Click()
If LaVolpeButton1.Caption = "Edit" Then
    cmd.Enabled = False
    cmd_op(2).Enabled = False
    Call unlocking
    LaVolpeButton1.Caption = "Save"
    LaVolpeButton1.Visible = True
End If
End Sub



