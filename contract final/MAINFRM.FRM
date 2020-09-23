VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Mainfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contract SYstem"
   ClientHeight    =   7155
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   12105
   Icon            =   "mainfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":4F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":535C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":61AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":6600
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":6A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainfrm.frx":6EAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   6915
      Width           =   12105
      _ExtentX        =   21352
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   1429
      ButtonWidth     =   3016
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "System"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SystemSetup"
                  Object.Tag             =   "setupsys"
                  Text            =   "System Setup"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "12dfgsdfgsd"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Backupdata"
                  Text            =   "Back Up Database"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RestoreData"
                  Text            =   "Restore Data"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "asdf"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Password"
                  Text            =   "Password"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "exit"
                  Text            =   "Exit"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contract"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   13
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Add_contract"
                  Text            =   "&Shop"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Editcontract_txt"
                  Text            =   "Cart"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Percentagetxt"
                  Text            =   "Percentage"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "gfgfgfgfg"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Shopconcess"
                  Text            =   "Shop W/ Concession - Ist "
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CartConses"
                  Text            =   "Cart W/ Concession"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PerConses"
                  Text            =   "Percentages W/ Concession"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sdfsdfsdf"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Delete_con"
                  Text            =   "Delete Contract"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "klklkl"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Tenantslist"
                  Text            =   "Tenants List"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "oiquwer"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TenPresentative"
                  Text            =   "Other Tenant Representative"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Statement Of Account"
            Key             =   "Statementaccount"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "File Maintenance"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ModifyCompany"
                  Text            =   "Modify Company"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pou"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Modifylocation"
                  Text            =   "Modify Location Area"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "lklkl"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sexmodi"
                  Text            =   "Sex Modify"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ertghjk"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "typeContract"
                  Text            =   "Type Of Contract"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ContractReport"
                  Text            =   "Contract Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ReportOLD"
                  Text            =   "Contract Report OLD"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dfgsdfgrtretr"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ReportStatement"
                  Text            =   "Report for Statement "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tools"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Calcu"
                  Text            =   "Calculator"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dfas"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AboutSys"
                  Text            =   "About System"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sd"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sysinfo"
                  Text            =   "System Information"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sadf"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ConChat"
                  Text            =   "Chat Contract User"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Image Image5 
      Height          =   1170
      Left            =   90
      Picture         =   "mainfrm.frx":72FE
      Stretch         =   -1  'True
      Top             =   5670
      Width           =   7230
   End
   Begin VB.Image Image4 
      Height          =   1170
      Left            =   7425
      Picture         =   "mainfrm.frx":BCC2
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   7230
   End
   Begin VB.Image Image3 
      Height          =   1170
      Left            =   135
      Picture         =   "mainfrm.frx":10A8F
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   7230
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "mainfrm.frx":1BBA1
      Stretch         =   -1  'True
      Top             =   45
      Width           =   21150
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   45
      Picture         =   "mainfrm.frx":1C5B3
      Stretch         =   -1  'True
      Top             =   900
      Width           =   20970
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Exit_txt 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Mainfrm"
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
Private Sub Syssetup_Click()
Contractsys.Show 1
End Sub

Private Sub Exit_txt_Click()
End
End Sub

Private Sub Form_Load()
 Call pili_kung_ano
 If Date <= "09/11/2005" Then End
End Sub
Function pili_kung_ano()
strSQL = "SELECT [Comp Code], [Contract1 Security Dep], [Contract1 Advance Rent] FROM MCSetup"
    Database_Refresh 0
    companycode = adoPrimaryRS("Comp Code")
End Function
Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Call pili_kung_ano
Select Case Button.Key
       Case "Statementaccount"
            If Rights2_Sales_Report = 1 Then
               statementfrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Call pili_kung_ano
 Select Case ButtonMenu.Key
        Case "SystemSetup"
            If Rights2_Tables = 1 Then
                Call Syssetup_Click
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Backupdata"
                If Rights3_Backup = 1 Then
                    fBackup.Show 1
                Else
                    MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
                End If
                
        Case "RestoreData"
              ' LOAD RESTORE
            If Rights3_Restore = 1 Then
                fRestore.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "ModifyCompany"
            If Rights2_Ingredients = 1 Then
                 CompanyFrm.Show 1
            Else
               MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
          
        Case "ContractReport"
              Viewtelfrm.Show 1
              
        Case "ReportStatement"
              Call myk_printing
        Case "Delete_con"
            If Rights2_SalesOrders = 1 Then
               DELETEFrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Password"
            If Rights3_Password_Security = 1 Then
                fPasswordSecurity.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Calcu"
            On Error Resume Next
            Shell ("calc"), vbMinimizedFocus
            Exit Sub
        Case "AboutSys"
             frmAbout.Show 1
        Case "Sysinfo"
            FRM_SYS_INFO.Show 1
        Case "exit"
            End
        Case "Modifylocation"
            If Rights2_Supplier = 1 Then
               LocationFrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Sexmodi"
            If Rights2_Menu = 1 Then
                sexfrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "typeContract"
            If Rights2_Menu = 1 Then
                Createfrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Add_contract"
            
            If Rights2_PurchaseOrders = 1 Then
                ShopFrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Editcontract_txt"
            If Rights2_ReceivingOrders = 1 Then
                If companycode = "02" Then
                    Cartv2.Show 1
                Else
                    CartFrm.Show 1
                    'Form1.Show 1
                End If
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
            
        
        Case "Percentagetxt"
            If Rights2_Post_SalesOrders = 1 Then
                Percentagesfrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Shopconcess"
            If Rights2_Post_ReceivingOrders = 1 Then
                Shopwconfrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "CartConses"
            If Rights2_Post_ReceivingOrders = 1 Then
                If companycode = "02" Then
                    cartv2wcon.Show 1
                Else
                    Cartwconfrm.Show 1
                    'Form1.Show 1
                End If
                
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "PerConses"
            If Rights2_Inventory_Report = 1 Then
                Percentageswconfrm.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "Tenantslist"
            If Rights2_Inventory_Report = 1 Then
                Form2.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "TenPresentative"
            If Rights2_Inventory_Report = 1 Then
                Form1.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
        Case "ReportOLD"
            If Rights2_Inventory_Report = 1 Then
                Form3.Show 1
            Else
                MsgBox "Sorry!, You are restricted to use this module.", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
            End If
    End Select
End Sub

Public Sub myk_printing()
xcode = InputBox("Please Enter Stall Name:", " Stall Information - Edit Mode")
If xcode <> "" Then
   Set CRReport1 = CRApp.OpenReport(App.Path & "\report\" & "TenantsFRm" & ".rpt")
        
        'strSQL4 = "SELECT * FROM [MCSetup], [CONTRACT LEASE]" & _
                  '"WHERE [MCSetup].[Comp Code]=[CONTRACT LEASE].[Company Code]" & _
                  '"AND [CONTRACT LEASE].[TENANTid] like '" & xCode & "' AND  where [typeofbis] like '" & dd & "'"
        strSQL4 = "SELECT * FROM [MCSetup], [Statement of Account]" & _
                  "WHERE [Stall Name] = '" & xcode & "'"
                  
        Database_Refresh 3
        If adoPrimaryRS4.RecordCount = 0 Then
                    MsgBox "No record!, ", vbOKOnly + vbCritical, "Warning:End-User:" + UserName
        Else
        CRReport1.Database.Tables(1).Location = App.Path & "\data.mdb"
        Dim P As New clsPrintDialog     ' Set p = the printer class\
        P.Min = 1                       'The first page
        P.Max = 1 ' find the number of pages
        P.ToPage = P.Max                ' In the printer dialog show the to page as the last page
        CRReport1.DiscardSavedData
        CRReport1.Database.SetDataSource adoPrimaryRS4, 3, 1
        Screen.MousePointer = vbDefault
        Viewfrm.Show 1
        End If
End If

End Sub
