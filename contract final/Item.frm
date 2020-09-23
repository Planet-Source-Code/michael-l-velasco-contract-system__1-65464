VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Display"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   2295
      TabIndex        =   1
      Top             =   630
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   2295
      TabIndex        =   0
      Top             =   990
      Width           =   2940
   End
   Begin MSDataGridLib.DataGrid GroupGrid 
      Height          =   1635
      Left            =   45
      TabIndex        =   2
      Top             =   2115
      Width           =   5865
      _ExtentX        =   10345
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
      Caption         =   "Company List"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Company Code"
         Caption         =   "Company ID"
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
         DataField       =   "Company Name"
         Caption         =   "Company Name"
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
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   90
      TabIndex        =   3
      Top             =   1575
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
      MICON           =   "Item.frx":0000
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
      Left            =   1575
      TabIndex        =   4
      Top             =   1575
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
      MICON           =   "Item.frx":001C
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
      Left            =   1575
      TabIndex        =   5
      Top             =   1575
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
      MICON           =   "Item.frx":0038
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
      Left            =   3015
      TabIndex        =   6
      Top             =   1575
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
      MICON           =   "Item.frx":0054
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
      Left            =   3015
      TabIndex        =   7
      Top             =   1575
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
      MICON           =   "Item.frx":0070
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
      Left            =   4500
      TabIndex        =   8
      Top             =   1575
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
      MICON           =   "Item.frx":008C
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
      Top             =   3765
      Width           =   6000
      _ExtentX        =   10583
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Company ID:"
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
      Left            =   765
      TabIndex        =   12
      Top             =   630
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Company Name:"
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
      Index           =   1
      Left            =   450
      TabIndex        =   11
      Top             =   990
      Width           =   1680
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Item Display"
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
      TabIndex        =   10
      Top             =   90
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   0
      Picture         =   "Item.frx":00A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9090
   End
End
Attribute VB_Name = "Form2"
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
    cmdEdit.Visible = False
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
strSQL = "SELECT * from [Company] order by [Company Code]"
  Database_Refresh 0
  GroupGrid.ClearSelCols
     dview
End Sub

Private Sub cmd_op_Click(Index As Integer)
       If choice = "Edit" Then
           strSQL3 = "SELECT [Company Name] from [Company] where [Company Code] = '" & Trim(Text1(0).Text) & "'"
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
              cmdEdit.Visible = True
              cmdDone.Visible = True
              cmd_op(0).Visible = False
              cmdcancel.Visible = False
        ElseIf choice = "Add" Then
                If Len(Text1(0).Text) <> 0 Then
                        'On Error GoTo A1:
                         strSQL2 = "SELECT * FROM [Company]"
                        Database_Refresh 1
                        With adoPrimaryRS2
                                .AddNew
                                 .Fields(0) = Text1(0).Text
                                 .Fields(1) = Text1(1).Text
                                .Update
                                .Requery
                                .Close
                        End With
                        Call Name_supp
                        Call clearing
                         Call locking
                        cmdDelete.Visible = True
                        cmd.Visible = True
                        cmdEdit.Visible = True
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
cmdEdit.Visible = True
cmdDone.Visible = True
cmd_op(0).Visible = False
cmdcancel.Visible = False
End Sub

Private Sub cmdDelete_Click()
strSQL3 = "SELECT * from [Company] where [Company Code] = '" & Trim(Text1(0).Text) & "'"
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
    cmdEdit.Visible = False
    cmdDone.Visible = False
    cmd_op(0).Visible = True
    cmdcancel.Visible = True
 Else
    MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
 End If
End Sub

Private Sub Form_Load()
  
      Call Name_supp
      Call locking
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



