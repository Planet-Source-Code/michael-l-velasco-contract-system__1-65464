VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tenants OLD"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10845
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1080
      Width           =   3705
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "BrowseOLD.frx":0000
      Left            =   1170
      List            =   "BrowseOLD.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   2355
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":0051
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":056B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":0C4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":13E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":1985
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":2333
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":29BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":3277
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":36BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":3BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":410C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":46E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":4B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":58A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BrowseOLD.frx":60F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   1852
      ButtonWidth     =   1455
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Report"
            Key             =   "Reports_txt_1"
            Description     =   "Reports"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   6105
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   503
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
   Begin LVbuttons.LaVolpeButton cmdDone 
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Top             =   5715
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&Close"
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
      MICON           =   "BrowseOLD.frx":6E48
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2310
      Left            =   12600
      TabIndex        =   5
      Top             =   4410
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   4075
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Supplier"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Supplier_Code"
         Caption         =   "Supplier Code"
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
         DataField       =   "Supplier_Name"
         Caption         =   "Supplier Name"
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
         DataField       =   "Supplier_Address"
         Caption         =   "Supplier Address"
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
         DataField       =   "Supplier_Terms"
         Caption         =   "Supplier Terms"
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
            Locked          =   -1  'True
            ColumnWidth     =   2340.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3899.906
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView plv1 
      Height          =   4185
      Left            =   45
      TabIndex        =   6
      Top             =   1485
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   7382
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CTRL #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tenants Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tenants Name"
         Object.Width           =   4023
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Representative / Owner"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   5468
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "BrowseOLD.frx":6E64
      Stretch         =   -1  'True
      Top             =   990
      Width           =   10890
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9270
      TabIndex        =   8
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      TabIndex        =   7
      Top             =   1125
      Width           =   6975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strDB           As String
Dim strSQL          As String
Dim strSQL2         As String
Dim strSQL3         As String
Dim strSQL4         As String
Dim strSQL5         As String
Dim strSQL6         As String
Dim strSQL7         As String
Dim strSQL1         As String
Dim mbChangedByCode As Boolean
Dim mvBookMark      As Variant
Dim mbEditFlag      As Boolean
Dim mbAddNewFlag    As Boolean
Dim mbDataChanged   As Boolean
Dim p1              As String
Dim p2              As String
Dim p3              As String
Dim p4              As String
Dim p5              As String
Dim p6              As String
Dim p7              As String
Dim p8              As String
Dim p9              As String
Dim p10              As String
Dim p11              As String
Dim p12              As String
Dim x1              As String
Dim x2              As String
Dim x3              As String
Dim x4              As String
Dim x5              As String
Dim xt As String
Dim yt As String
Dim WithEvents adoPrimaryRS1 As ADODB.Recordset
Attribute adoPrimaryRS1.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS2 As ADODB.Recordset
Attribute adoPrimaryRS2.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS3 As ADODB.Recordset
Attribute adoPrimaryRS3.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS4 As ADODB.Recordset
Attribute adoPrimaryRS4.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS5 As ADODB.Recordset
Attribute adoPrimaryRS5.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS6 As ADODB.Recordset
Attribute adoPrimaryRS6.VB_VarHelpID = -1
Dim WithEvents adoPrimaryRS7 As ADODB.Recordset
Attribute adoPrimaryRS7.VB_VarHelpID = -1

Public Sub Database_Refresh(xMode As Integer)
   ' PRE-DATABASE CONNECTION WITH PARAMETERIZED SQL VARIABLES ATTACHED IN EVERY MODE
    
    If xMode = 0 Then
        Set adoPrimaryRS = New ADODB.Recordset
        adoPrimaryRS.Open strSQL, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 1 Then
        Set adoPrimaryRS1 = New ADODB.Recordset
        adoPrimaryRS1.Open strSQL1, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 2 Then
        Set adoPrimaryRS2 = New ADODB.Recordset
        adoPrimaryRS2.Open strSQL2, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 3 Then
        Set adoPrimaryRS3 = New ADODB.Recordset
        adoPrimaryRS3.Open strSQL3, db, adOpenStatic, adLockPessimistic
    ElseIf xMode = 4 Then
        Set adoPrimaryRS4 = New ADODB.Recordset
        adoPrimaryRS4.Open strSQL4, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 5 Then
        Set adoPrimaryRS5 = New ADODB.Recordset
        adoPrimaryRS5.Open strSQL5, db, adOpenStatic, adLockOptimistic
    ElseIf xMode = 6 Then
        Set adoPrimaryRS6 = New ADODB.Recordset
        adoPrimaryRS6.Open strSQL6, db, adOpenStatic, adLockOptimistic
    End If
End Sub
Public Sub Reload_PrimaryRS()
'On Error Resume Next
    ' RELOADING DATA OBJECTS AND DATABASE CONNECTIONS
        strSQL = "SELECT * from [Contract LEase OLD] "
        Database_Refresh 0
  callitem
End Sub

Function callitem()
On Error Resume Next
plv1.ListItems.Clear
            With adoPrimaryRS
                Do While Not .EOF
                
                Set lst = plv1.ListItems.Add(, , .Fields(0))
                    lst.SubItems(1) = .Fields(1)
                    lst.SubItems(2) = .Fields(4)
                    lst.SubItems(3) = .Fields(6)
                    lst.SubItems(4) = .Fields(7)
                    'lst.SubItems(5) = .Fields(5)
                    .MoveNext
                
                Loop
        End With
End Function



Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Reload_PrimaryRS
End Sub

Private Sub plv1_ItemClick(ByVal Item As MSComctlLib.ListItem)
xt = plv1.SelectedItem.Text
yt = plv1.SelectedItem.SubItems(1)
End Sub

Private Sub Text1_Change()
If Combo1.Text = "Tenants Code" Then
         strSQL = "SELECT * from [Contract LEase OLD] where [TENANTid] like '" & Trim(Text1) & "%'"
            Database_Refresh 0
            callitem
ElseIf Combo1.Text = "Tenants Name" Then
        strSQL = "SELECT * from [Contract LEase OLD] where [Tenant Name] like '" & Trim(Text1) & "%'"
            Database_Refresh 0
            callitem
ElseIf Combo1.Text = "Tenants Representative" Then
        strSQL = "SELECT * from [Contract LEase OLD] where [Tenant Presentative] like '" & Trim(Text1) & "%'"
            Database_Refresh 0
            callitem
ElseIf Combo1.Text = "Address" Then
        strSQL = "SELECT * from [Contract LEase OLD] where [Presentative Address] like '" & Trim(Text1) & "%'"
            Database_Refresh 0
            callitem

End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
       Case "Reports_txt_1"
            Dim qw As String
    
        'If Option1.Value = True Then qw = "01"
        'If Option2.Value = True Then qw = "02"
            
        
        xcode = Combo1.Text
        'strSQL4 = "SELECT * FROM [MCSetup], [CONTRACT LEASE]" & _
                  '"WHERE [MCSetup].[Comp Code]=[CONTRACT LEASE].[Company Code]" & _
                  '"AND [CONTRACT LEASE].[TENANTid] like '" & xCode & "' AND  where [typeofbis] like '" & dd & "'"
        'strSQL4 = "SELECT * FROM [MCSetup], [CONTRACT LEASE]" & _
        '          "WHERE [TENANTid]='" & xCode & "' AND [typeofbis] ='" & dd & "'"
        strSQL4 = "SELECT * FROM [MCSetup], [CONTRACT LEASE OLD]" & _
                  "WHERE [TENANTid] ='" & yt & "' and  [CTRL NO] like '" & xt & "'"
          
        Database_Refresh 4
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

End Select
End Sub

