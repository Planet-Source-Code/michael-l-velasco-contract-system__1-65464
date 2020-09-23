VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Viewtelfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View Report"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4395
   Icon            =   "ViewTetFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   810
      TabIndex        =   3
      Top             =   1530
      Width           =   3075
   End
   Begin LVbuttons.LaVolpeButton cmdDone 
      Height          =   375
      Left            =   765
      TabIndex        =   1
      Top             =   2385
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&View"
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
      MICON           =   "ViewTetFrm.frx":4F0A
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
      TabIndex        =   2
      Top             =   3810
      Width           =   4395
      _ExtentX        =   7752
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   2475
      TabIndex        =   5
      Top             =   2385
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
      MICON           =   "ViewTetFrm.frx":4F26
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tenant Name"
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
      TabIndex        =   7
      Top             =   2820
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "codeten"
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
      Height          =   555
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   3045
      Width           =   4050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tenant Code"
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
      TabIndex        =   4
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "View Tenants"
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
      Top             =   45
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "ViewTetFrm.frx":4F42
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12285
   End
End
Attribute VB_Name = "Viewtelfrm"
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
Dim strSQL6 As String
Dim WithEvents adoPrimaryRS6 As ADODB.Recordset
Attribute adoPrimaryRS6.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS4 As ADODB.Recordset
Attribute adoPrimaryRS4.VB_VarHelpID = -1
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
 adoPrimaryRS2.UpdateBatch adAffectAll
End Sub

Private Sub cmdDone_Click()

        Dim qw As String
    
        'If Option1.Value = True Then qw = "01"
        'If Option2.Value = True Then qw = "02"
            
        
        xcode = Combo1.Text
        'strSQL4 = "SELECT * FROM [MCSetup], [CONTRACT LEASE]" & _
                  '"WHERE [MCSetup].[Comp Code]=[CONTRACT LEASE].[Company Code]" & _
                  '"AND [CONTRACT LEASE].[TENANTid] like '" & xCode & "' AND  where [typeofbis] like '" & dd & "'"
        'strSQL4 = "SELECT * FROM [MCSetup], [CONTRACT LEASE]" & _
        '          "WHERE [TENANTid]='" & xCode & "' AND [typeofbis] ='" & dd & "'"
        strSQL4 = "SELECT * FROM [MCSetup], [CONTRACT LEASE]" & _
                  "WHERE [TENANTid]='" & xcode & "'"
          
        Database_Refresh 3
        dd = adoPrimaryRS4("Typeofbis")
        Set CRReport1 = CRApp.OpenReport(App.Path & "\report\" & dd & ".rpt")
            CRReport1.Database.Tables(1).Location = FileNameTXT

        Dim P As New clsPrintDialog     ' Set p = the printer class\
        P.Min = 1                       'The first page
        P.Max = 1 ' find the number of pages
        P.ToPage = P.Max                ' In the printer dialog show the to page as the last page
        CRReport1.DiscardSavedData
        CRReport1.Database.Tables(1).SetLogOnInfo "", "", "", "mykpogi"
        CRReport1.Database.SetDataSource adoPrimaryRS4, 3, 1
        Viewfrm.Show 1

    
End Sub

Private Sub Combo1_Change()
'strSQL2 = "SELECT [TENANTid]as codeten FROM [CONTRACT LEASE] WHERE [TENANTid] = " & Trim(Combo1.Text)
'    Database_Refresh 1
'    Set Label1(0).DataSource = adoPrimaryRS2
Call Combo1_Click
End Sub

Private Sub Combo1_Click()
strSQL2 = "SELECT [Tenant Name] as codeten FROM [CONTRACT LEASE] WHERE [TENANTid] like '" & Trim(Combo1.Text) & "'"
    Database_Refresh 1
    Set Label1(0).DataSource = adoPrimaryRS2
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
        lookup.Show 1
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
      Select Case KeyAscii
                Case 13
                    Combo1.SetFocus
                Case Else
                     KeyAscii = 0
      End Select

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    xcode = Combo1.Text
    If xcode <> "" Then
           strSQL2 = "Select * from [CONTRACT LEASE] where [TENANTid] = '" & xcode & "'"
                Database_Refresh 1
                If adoPrimaryRS2.RecordCount = 0 Then
                    MsgBox "No record!, ", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
                    Combo1.SetFocus
                    
                Else
                 Call Combo1_Click
                End If
            Else
                Beep
                
            End If
  cmdDone.SetFocus
End If
End Sub

Private Sub Form_Load()
    ' STARTUP SUPPLIERS DATABASE CONNECTIONS
   ' Rights5_Add = 1
   ' Rights5_Edit = 1
   ' Rights5_Save = 1
    'Reload_PrimaryRS
   
    
'    strSQL5 = "SELECT * FROM [typeOfContract] ORDER BY [Contract Code]"
'    Database_Refresh 4
'    If adoPrimaryRS5.RecordCount <> 0 Then
'        adoPrimaryRS5.MoveFirst
'        Do While Not adoPrimaryRS5.EOF
'            Combo2.AddItem IIf(IsNull(adoPrimaryRS5("Contract Name")), " ", adoPrimaryRS5("Contract Name"))
'            adoPrimaryRS5.MoveNext
'        Loop
'    End If

'    strSQL6 = "SELECT [Contract Code],[Contract Name] FROM [typeOfContract] where [Contract Name]= '" & Combo2.Text & "'"
 '   Database_Refresh 5
    'dd = adoPrimaryRS6("Contract Code")
    Combo1.Clear
    'strSQL3 = "SELECT * FROM [CONTRACT LEASE] where [Typeofbis] like '" & dd & "' ORDER BY [TENANTid]"
    strSQL3 = "SELECT * FROM [CONTRACT LEASE] where [Typeofbis]"
    Database_Refresh 2
    If adoPrimaryRS3.RecordCount <> 0 Then
        adoPrimaryRS3.MoveFirst
        Do While Not adoPrimaryRS3.EOF
            Combo1.AddItem IIf(IsNull(adoPrimaryRS3("TENANTid")), " ", adoPrimaryRS3("TENANTid"))
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
   ' On Error Resume Next
    Dim oText As TextBox, i
    Dim odate As DTPicker, e
    
    'strSQL2 = "SELECT [Company Code] AS CompanyCode, [Company TIN number] AS CompanyTIN, [Mall Managers] AS MallManagers," & _
    '         "Controller AS Controllername,[Mall Address] AS Malladd,[company resnumber] AS companyRestNo," & _
    '         "[Company Date Issued] AS CompanyDateissued, [Company Place Issued] AS CompanyPlace," & _
    '         "[manager resnumber] as managerresnumber ,[Manager Date Issued] as ManagerDateIssued," & _
    '         "[Manager Place Issued] as ManagerPlaceIssued, [Controller resnumber] as companyresnumber," & _
    '         "[Controller Date Issued] as ControllerDateIssued, [Controller Place Issued]as ControllerPlaceIssued FROM MCSetup "
    strSQL2 = "Select [Company Code] as CompanyCode, " & _
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
                "[company resnumber] as CompanyRestNo FROM MCSetup"
                
    Database_Refresh 1
   
    If adoPrimaryRS2.RecordCount <> 0 Then
        adoPrimaryRS2.MoveFirst
        Set Me.Combo1.DataSource = adoPrimaryRS2
        mbDataChanged = False
    End If
        
End Sub

Private Sub txtCombo_Change()
    strSQL = "SELECT [Company Name] FROM Company where [Company Code] = '" & txtCombo & "'"
    Database_Refresh 0
    Label1(12).Caption = adoPrimaryRS("Company Name")
    'If adoPrimaryRS3.RecordCount <> 0 Then
    '    adoPrimaryRS3.MoveFirst
    '    Do While Not adoPrimaryRS3.EOF
    '        txtCombo.AddItem IIf(IsNull(adoPrimaryRS3("Company Name")), "", adoPrimaryRS3("Company Name"))
    '        adoPrimaryRS3.MoveNext
    '    Loop
    'End If
End Sub



Private Sub LaVolpeButton1_Click()
Unload Me
End Sub


