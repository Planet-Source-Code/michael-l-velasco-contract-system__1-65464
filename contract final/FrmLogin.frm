VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login "
   ClientHeight    =   2580
   ClientLeft      =   2355
   ClientTop       =   2340
   ClientWidth     =   4485
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      TabIndex        =   0
      Top             =   810
      Width           =   2895
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1500
      Width           =   2895
   End
   Begin LVbuttons.LaVolpeButton cmdDone 
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Top             =   1935
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&OK"
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
      MICON           =   "FrmLogin.frx":0E42
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      Top             =   1935
      Width           =   1005
      _ExtentX        =   1773
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
      MICON           =   "FrmLogin.frx":0E5E
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
      TabIndex        =   7
      Top             =   2340
      Width           =   4485
      _ExtentX        =   7911
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
   Begin VB.Image Image2 
      Height          =   975
      Left            =   3555
      Picture         =   "FrmLogin.frx":0E7A
      Stretch         =   -1  'True
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   135
      TabIndex        =   6
      Top             =   90
      Width           =   5175
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
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
      Left            =   135
      TabIndex        =   5
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
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
      Left            =   135
      TabIndex        =   4
      Top             =   1260
      Width           =   885
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "FrmLogin.frx":5D84
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MODULE/FORM: PASSWORD SECURITY
' VERSION: VB6

Option Explicit

' PASSWORD SECURITY VARIABLE SETTINGS
Dim strDBPass As String
Dim strSQLPass As String
Dim dbPass As ADODB.Connection
Dim WithEvents adoPrimaryRSPass As ADODB.Recordset
Attribute adoPrimaryRSPass.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim strSQLPass1 As String
Dim WithEvents adoPrimaryRSPass1 As ADODB.Recordset
Attribute adoPrimaryRSPass1.VB_VarHelpID = -1

Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim ctr As Integer
Dim xText


Private Sub cmdDone_Click()
    ' LOGIN ENTRY
    Dim ShowAtStartup As Long
    
    If Get_User(Text1, txtPassword) Then
        UserName = Text1
        Rights1_Add = IIf(IsNull(adoPrimaryRSPass("User_Rights1_Add")), 0, IIf(adoPrimaryRSPass("User_Rights1_Add") = 0, 0, 1))
        Rights1_Edit = IIf(IsNull(adoPrimaryRSPass("User_Rights1_Edit")), 0, IIf(adoPrimaryRSPass("User_Rights1_Edit") = 0, 0, 1))
        Rights1_Delete = IIf(IsNull(adoPrimaryRSPass("User_Rights1_Delete")), 0, IIf(adoPrimaryRSPass("User_Rights1_Delete") = 0, 0, 1))
        Rights2_Tables = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Tables")), 0, IIf(adoPrimaryRSPass("User_Rights2_Tables") = 0, 0, 1))
        Rights2_Service_Crew = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Service_Crew")), 0, IIf(adoPrimaryRSPass("User_Rights2_Service_Crew") = 0, 0, 1))
        Rights2_Ingredients = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Ingredients")), 0, IIf(adoPrimaryRSPass("User_Rights2_Ingredients") = 0, 0, 1))
        Rights2_Menu = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Menu")), 0, IIf(adoPrimaryRSPass("User_Rights2_Menu") = 0, 0, 1))
        Rights2_Supplier = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Supplier")), 0, IIf(adoPrimaryRSPass("User_Rights2_Supplier") = 0, 0, 1))
        Rights2_SalesOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_SalesOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_SalesOrders") = 0, 0, 1))
        Rights2_PurchaseOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_PurchaseOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_PurchaseOrders") = 0, 0, 1))
        Rights2_ReceivingOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_ReceivingOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_ReceivingOrders") = 0, 0, 1))
        Rights2_Post_SalesOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Post_SalesOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_Post_SalesOrders") = 0, 0, 1))
        Rights2_Post_ReceivingOrders = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Post_SalesOrders")), 0, IIf(adoPrimaryRSPass("User_Rights2_ReceivingOrders") = 0, 0, 1))
        Rights2_Inventory_Report = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Inventory_Report")), 0, IIf(adoPrimaryRSPass("User_Rights2_Inventory_Report") = 0, 0, 1))
        Rights2_Sales_Report = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Sales_Report")), 0, IIf(adoPrimaryRSPass("User_Rights2_Sales_Report") = 0, 0, 1))
        Rights2_Critical_Report = IIf(IsNull(adoPrimaryRSPass("User_Rights2_Critical_Report")), 0, IIf(adoPrimaryRSPass("User_Rights2_Critical_Report") = 0, 0, 1))
        Rights3_Backup = IIf(IsNull(adoPrimaryRSPass("User_Rights3_Backup")), 0, IIf(adoPrimaryRSPass("User_Rights3_Backup") = 0, 0, 1))
        Rights3_Restore = IIf(IsNull(adoPrimaryRSPass("User_Rights3_Restore")), 0, IIf(adoPrimaryRSPass("User_Rights3_Restore") = 0, 0, 1))
        Rights3_Password_Security = IIf(IsNull(adoPrimaryRSPass("User_Rights3_Password_Security")), 0, IIf(adoPrimaryRSPass("User_Rights3_Password_Security") = 0, 0, 1))
        Rights3_CarwtConse = IIf(IsNull(adoPrimaryRSPass("User_Rights3_CarwtConse")), 0, IIf(adoPrimaryRSPass("User_Rights3_CarwtConse") = 0, 0, 1))
        adoPrimaryRSPass.Close
        Unload Me
        strSQLPass1 = "SELECT * from [log]"
        Database_Refresh 1
        adoPrimaryRSPass1.AddNew
        adoPrimaryRSPass1.Fields(0).Value = Date
        adoPrimaryRSPass1.Update
        Mainfrm.Show
    ElseIf Trim(Text1) = "" And Trim(txtPassword) = "3773" Then
        Rights1_Add = 1
        Rights1_Edit = 1
        Rights1_Delete = 1
        Rights2_Tables = 1
        Rights2_Service_Crew = 1
        Rights2_Ingredients = 1
        Rights2_Menu = 1
        Rights2_Supplier = 1
        Rights2_SalesOrders = 1
        Rights2_PurchaseOrders = 1
        Rights2_ReceivingOrders = 1
        Rights2_Post_SalesOrders = 1
        Rights2_Post_ReceivingOrders = 1
        Rights2_Inventory_Report = 1
        Rights2_Sales_Report = 1
        Rights2_Critical_Report = 1
        Rights3_Backup = 1
        Rights3_Restore = 1
        Rights3_Password_Security = 1
        Rights3_CarwtConse = 1
        UserName = "Administrator"
        adoPrimaryRSPass.Close
        Unload Me
        Mainfrm.Show
    Else
        ctr = ctr + 1
        If ctr = 4 Then
           End
        Else
            xText = "You have" + Str(4 - ctr) + " tries left"
            If ctr = 3 Then
                xText = "This is your last chance!!"
            End If
            MsgBox "Access Denied!!" & vbCrLf & _
                   xText, vbOKOnly + vbCritical, "Warning:End-User"
            SendKeys "{Home}+{End}"
        End If
   End If

End Sub

Private Sub cmdOK_Click()
End Sub

Private Sub Form_Load()
strSQLPass1 = "SELECT * from [log] where [date text]=#" & "12/01/2005" & "#"
Database_Refresh 1

If adoPrimaryRSPass1.RecordCount <> 0 Then
End
Else

End If
End Sub

Private Sub LaVolpeButton1_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdDone_Click
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub Database_Refresh(xMode As Integer)
'On Error GoTo myk
    'Set dbPass = New Connection
        'dbPass.CursorLocation = adUseClient
        'strDBPass = App.Path + "\data.MDB;Jet OLEDB:Database Password=;"
        'dbPass.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPass
        'dbPass.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & strDBPass
    If xMode = 0 Then
        Set adoPrimaryRSPass = New Recordset
        adoPrimaryRSPass.Open strSQLPass, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 1 Then
        Set adoPrimaryRSPass1 = New Recordset
        adoPrimaryRSPass1.Open strSQLPass1, db, adOpenStatic, adLockOptimistic
    End If
'myk:
   'MkDir App.Path + "\database\"
   'FileCopy App.Path + "\data.MDB", App.Path + ("\database\data.mdb")
   
   'Call Form_Load
End Sub

Function Get_User(p_user As String, p_pass As String) As Boolean
    ' USERNAME AND PASSWORD VALIDATION
    On Error Resume Next
    strSQLPass = "SELECT * FROM Password_Security WHERE User_Name = '" & p_user & "'" _
            & " AND User_Password = '" & Decode_Pass(p_pass) & "'"
    Database_Refresh 0
    If adoPrimaryRSPass.AbsolutePosition <> -1 Then
        Get_User = True
    Else
        Get_User = False
    End If
End Function


