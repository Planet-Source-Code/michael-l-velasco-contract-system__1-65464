VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Other Representative"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9270
   Icon            =   "OtherRepresentative.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "TenantPresentative2"
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
      Left            =   3375
      TabIndex        =   17
      Top             =   990
      Width           =   4020
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
      Index           =   1
      Left            =   2385
      TabIndex        =   16
      Top             =   585
      Width           =   5010
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   45
      TabIndex        =   6
      Top             =   1890
      Width           =   9195
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PresentativePlaceissued2"
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
         Left            =   6255
         TabIndex        =   21
         Top             =   405
         Width           =   2760
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Presentativedateissued2"
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
         Left            =   3195
         TabIndex        =   20
         Top             =   405
         Width           =   2760
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Presentativeresnumber2"
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
         Left            =   135
         TabIndex        =   19
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "PresentativePosition2"
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
         Left            =   2340
         TabIndex        =   18
         Top             =   1350
         Width           =   5010
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
         Index           =   0
         Left            =   2340
         TabIndex        =   15
         Top             =   180
         Width           =   2760
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "SecondName2"
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
         Height          =   330
         Left            =   2340
         TabIndex        =   1
         Top             =   990
         Width           =   960
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
         Left            =   1215
         TabIndex        =   5
         Top             =   270
         Width           =   1050
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
         Left            =   630
         TabIndex        =   4
         Top             =   990
         Width           =   1665
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
         TabIndex        =   3
         Top             =   630
         Width           =   1080
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
         Left            =   315
         TabIndex        =   2
         Top             =   1395
         Width           =   1980
      End
   End
   Begin LVbuttons.LaVolpeButton cmdedit 
      Height          =   330
      Left            =   2565
      TabIndex        =   10
      Top             =   2835
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
      MICON           =   "OtherRepresentative.frx":144A
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
      TabIndex        =   11
      Top             =   2835
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
      MICON           =   "OtherRepresentative.frx":1466
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
      TabIndex        =   12
      Top             =   2835
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
      MICON           =   "OtherRepresentative.frx":1482
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
      TabIndex        =   13
      Top             =   2835
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
      MICON           =   "OtherRepresentative.frx":149E
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
      Left            =   4140
      TabIndex        =   14
      Top             =   2835
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
      MICON           =   "OtherRepresentative.frx":14BA
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
Attribute VB_Name = "Form1"
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
'    Text1(14).SetFocus
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
           strSQL2 = "Select [TENANTid] as TENANTidt, [Tenant Name] as TenantName, [Second Name2] as SecondName2," & _
              "[Tenant Presentative2] as TenantPresentative2, [Presentative Position2] as PresentativePosition2," & _
              "[Presentative resnumber2] as Presentativeresnumber2, [Presentative date issued2] as Presentativedateissued2," & _
              "[Presentative Place issued2] as PresentativePlaceissued2 from [CONTRACT LEASE] where [TENANTid] = '" & xcode & "'"
                mbEditFlag = True
                Database_Refresh 1
                If adoPrimaryRS2.RecordCount = 0 Then
                    MsgBox "No record!, ", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
                Else
                 clearing
               For Each oText In Me.Text1
                    Set oText.DataSource = adoPrimaryRS2
                Next
'                Set Text3.DataSource = adoPrimaryRS2
'                Set Label3.DataSource = adoPrimaryRS2
'                Set Label1(16).DataSource = adoPrimaryRS2
 '               Set Label1(17).DataSource = adoPrimaryRS2
 '               Set Label1(19).DataSource = adoPrimaryRS2
 '               Set Label1(20).DataSource = adoPrimaryRS2
  '              Set Label1(25).DataSource = adoPrimaryRS2
                
'                Set Label1(22).DataSource = adoPrimaryRS2
'                Set Text2.DataSource = adoPrimaryRS2
'                Set Combo2.DataSource = adoPrimaryRS2
'                Set txtCombo.DataSource = adoPrimaryRS2
                Set Combo1.DataSource = adoPrimaryRS2
'                Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
 '               Label1(17).Caption = Format(Label1(17).Caption, "##,##0.00")
  '              Text1(7).Text = Format(Text1(7).Text, "##,##0.00")
  '              Text1(8).Text = Format(Text1(8).Text, "##,##0.00")
   '             Text1(12).Text = Format(Text1(12).Text, "##,##0.00")
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
                    Text1(2).SetFocus
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
'Text3.Text = dtStart(1).Value
End Sub

Private Sub Form_Load()
    ' STARTUP SUPPLIERS DATABASE CONNECTIONS
   ' Rights5_Add = 1
   ' Rights5_Edit = 1
   ' Rights5_Save = 1
    locking
    Reload_PrimaryRS

 '   strSQL = "SELECT [Comp Code], [Contract1 Security Dep], [Contract1 Advance Rent] FROM MCSetup"
 '   Database_Refresh 0
'    Text2.Text = adoPrimaryRS("Comp Code")
 '   secdep = adoPrimaryRS("Contract1 Security Dep")
 '   adrent = adoPrimaryRS("Contract1 Advance Rent")
       
'    strSQL4 = "SELECT [Location Name]  FROM location ORDER BY [Location code]"
'    Database_Refresh 3
'    If adoPrimaryRS4.RecordCount <> 0 Then
'        adoPrimaryRS4.MoveFirst
'        Do While Not adoPrimaryRS4.EOF
'            txtCombo.AddItem IIf(IsNull(adoPrimaryRS4("Location Name")), "", adoPrimaryRS4("Location Name"))
'            adoPrimaryRS4.MoveNext
'        Loop
'    End If
    
    strSQL5 = "SELECT [sex Name]  FROM sex ORDER BY [sex code]"
    Database_Refresh 4
    If adoPrimaryRS5.RecordCount <> 0 Then
        adoPrimaryRS5.MoveFirst
        Do While Not adoPrimaryRS5.EOF
            Combo1.AddItem IIf(IsNull(adoPrimaryRS5("Sex Name")), "", adoPrimaryRS5("Sex Name"))
            adoPrimaryRS5.MoveNext
        Loop
    End If
'    Label1(16).Caption = Format(Label1(16).Caption, "##,##0.00")
'    Label1(17).Caption = Format(Label1(17).Caption, "##,##0.00")
'    Text1(7).Text = Format(Text1(7).Text, "##,##0.00")
'    Text1(8).Text = Format(Text1(8).Text, "##,##0.00")
'    Text1(12).Text = Format(Text1(12).Text, "##,##0.00")
    
    
'    strSQL6 = "SELECT [Contract Code]  FROM typeOfContract ORDER BY [Contract Code]"
'    Database_Refresh 5
'    If adoPrimaryRS6.RecordCount <> 0 Then
'        adoPrimaryRS6.MoveFirst
'        Do While Not adoPrimaryRS6.EOF
'            Combo2.AddItem IIf(IsNull(adoPrimaryRS6("Contract Code")), "", adoPrimaryRS6("Contract Code"))
'            adoPrimaryRS6.MoveNext
'        Loop
'    End If
    
    hanap_bistype
    
End Sub
Function hanap_bistype()
'On Error Resume Next
'strSQL6 = "SELECT *  FROM typeOfContract where [Contract Code]= '" & Combo2.Text & "'"
'                Database_Refresh 5
''                Text4.Text = adoPrimaryRS6.Fields(0)
'                Label1(27).Caption = adoPrimaryRS6.Fields(1)
    
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

    strSQL2 = "Select [TENANTid] as TENANTidt, [Tenant Name] as TenantName, [Second Name2] as SecondName2," & _
              "[Tenant Presentative2] as TenantPresentative2, [Presentative Position2] as PresentativePosition2," & _
              "[Presentative resnumber2] as Presentativeresnumber2, [Presentative date issued2] as Presentativedateissued2," & _
              "[Presentative Place issued2] as PresentativePlaceissued2 From [CONTRACT LEASE]"
              Database_Refresh 1
                 '"[Tenant Name 2] as TenantName2," & _
                '"[Leased approximately] as Leasedapproximately, "
                For Each oText In Me.Text1
                    Set oText.DataSource = adoPrimaryRS2
                    
                Next
'                Set Text3.DataSource = adoPrimaryRS2
'                Set Label3.DataSource = adoPrimaryRS2
'                Set Label1(16).DataSource = adoPrimaryRS2
'                Set Label1(17).DataSource = adoPrimaryRS2
'                Set Label1(19).DataSource = adoPrimaryRS2
'                Set Label1(20).DataSource = adoPrimaryRS2
'                Set Label1(25).DataSource = adoPrimaryRS2
'                Set Text2.DataSource = adoPrimaryRS2
'                Set Combo2.DataSource = adoPrimaryRS2
                
                
                'Set dtStart(1).DataSource = adoPrimaryRS2
'                Set Label1(22).DataSource = adoPrimaryRS2
'                Set Text2.DataSource = adoPrimaryRS2
                  'If adoPrimaryRS2.RecordCount <> 0 Then
                  '  adoPrimaryRS2.MoveFirst
'                    Set txtCombo.DataSource = adoPrimaryRS2
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
For i = 0 To 6
   Text1(i).Text = ""
  ' Me.Label1(16).Caption = ""
  ' Text2.Text = "01"
  ' Text4.Text = "02"
Next i
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then Text1(1).SetFocus
    If Index = 1 Then Combo1.SetFocus
    If Index = 2 Then Text1(3).SetFocus
    If Index = 3 Then Text1(4).SetFocus
    If Index = 4 Then Text1(5).SetFocus
    If Index = 5 Then Text1(6).SetFocus
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
'  Label1(16).Enabled = False
'  Label1(17).Enabled = False

'  Label1(27).Enabled = False
  Combo1.Enabled = False
'  Combo2.Enabled = False
'  txtCombo.Enabled = False
'  Text3.Enabled = False
'  Text1(18).Enabled = False
End Function
Function unlocking()
For i = 0 To 6
    Text1(i).Enabled = True
Next i
'Label1(16).Enabled = True
'Text1(18).Enabled = True
'  Label1(17).Enabled = True
  'dtStart(1).Enabled = True
'  Label1(27).Enabled = True
'  Combo2.Enabled = True
  Combo1.Enabled = True
'  Text3.Enabled = True
'  txtCombo.Enabled = True
End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1(5).SetFocus
End Sub








