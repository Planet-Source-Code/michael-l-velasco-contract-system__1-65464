VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox txtCombo 
      DataField       =   "CompanyCode"
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      Top             =   945
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strDB As String
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

Private Sub cmd_Click()
clearing
End Sub

Private Sub cmd_op_Click(Index As Integer)
 adoPrimaryRS2.UpdateBatch adAffectAll
End Sub

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub Form_Load()
    ' STARTUP SUPPLIERS DATABASE CONNECTIONS
   ' Rights5_Add = 1
   ' Rights5_Edit = 1
   ' Rights5_Save = 1
    Reload_PrimaryRS
    
End Sub
Public Sub Database_Refresh(xMode As Integer)
    ' PRE-DATABASE CONNECTION WITH PARAMETERIZED SQL VARIABLES ATTACHED IN EVERY MODE
    'On Error Resume Next
    strDB = App.Path + "\DATABASE"
    Set db = New Connection
        db.CursorLocation = adUseClient
        db.Open "Driver={Microsoft FoxPro VFP Driver (*.dbf)};SourceDB=" & strDB & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine"
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

Private Sub Reload_PrimaryRS()
    ' RELOADING DATA OBJECTS AND DATABASE CONNECTIONS
   ' On Error Resume Next
    Dim oText As TextBox, i
    Dim odate As DTPicker, e
    strSQL2 = "select * from tenants"
'      If adoPrimaryRS2.RecordCount <> 0 Then
'        adoPrimaryRS2.MoveFirst
        Set txtCombo.DataSource = adoPrimaryRS2
'        mbDataChanged = False
'    End If
    'strSQL2 = "SELECT [Company Code] AS CompanyCode, [Company TIN number] AS CompanyTIN, [Mall Managers] AS MallManagers," & _
    '         "Controller AS Controllername,[Mall Address] AS Malladd,[company resnumber] AS companyRestNo," & _
    '         "[Company Date Issued] AS CompanyDateissued, [Company Place Issued] AS CompanyPlace," & _
    '         "[manager resnumber] as managerresnumber ,[Manager Date Issued] as ManagerDateIssued," & _
    '         "[Manager Place Issued] as ManagerPlaceIssued, [Controller resnumber] as companyresnumber," & _
    '         "[Controller Date Issued] as ControllerDateIssued, [Controller Place Issued]as ControllerPlaceIssued FROM MCSetup "
    'strSQL2 = "Select [Tenant Name] as TenantName, " & _
    '            "[Tenant Presentative] as TenantPresentative, " & _
    '            "[Presentative Address] as PresentativeAddress, " & _
    '            "[Presentative resnumber] as Presentativeresnumber, " & _
    '            "[Presentative Place issued] as PresentativePlaceissued," & _
    '            "[leased premises amt] as leasedpremisesamt, " & _
    '            "[Leased Premises] as LeasedPremises, " & _
    '            "[Amount of rent] as Amountofrent, " & _
    '            "[Cusa Charges] as CusaCharges, " & _
    '            "[usage of Premises] as usageofPremises, " & _
    '            "[Rental Commen Date] as RentalCommenDate, " & _
    '            "[Lease Expiry Date] as LeaseExpiryDate, " & _
    '            "[Presentative date issued] as Presentativedateissued, " & _
    '            "[Aircon Charges] as AirconCharges, " & _
    '            "[Leased approximately] as Leasedapproximately, " & _
    '            "[Sectotalamt] as Sectotalamt1, " & _
    '            "[Tenant TIN] as TenantTIN from [CONTRACT LEASE]"
    '            '
    
                '
                '
                '
                '
                '
                '
                '
                '
                '
                ' from [CONTRACT LEASE]"
                
                
    Database_Refresh 1
      'Set Text1(0).DataSource = adoPrimaryRS2
      'Set Text1(1).DataSource = adoPrimaryRS2
         
   
        
'        Set dtStart(2).DataSource = adoPrimaryRS2
'             Set dtStart(0).DataSource = adoPrimaryRS2
'    Next
    
'    If adoPrimaryRS2.RecordCount <> 0 Then
'        adoPrimaryRS2.MoveFirst
'        Set txtCombo.DataSource = adoPrimaryRS2
'        mbDataChanged = False
'    End If
        
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

Private Sub txtCombo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Text1(0).SetFocus
End Sub
Function clearing()
For i = 0 To 13
   Text1(i).Text = ""
   Me.Label1(16).Caption = ""
Next i
End Function




