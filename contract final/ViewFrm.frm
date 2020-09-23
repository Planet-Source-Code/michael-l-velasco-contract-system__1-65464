VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Begin VB.Form Viewfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "View Tenants"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10185
   Icon            =   "ViewFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   4365
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196609
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   679
      DesignHeight    =   478
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      lastProp        =   500
      _cx             =   17965
      _cy             =   12647
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "Viewfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
             ' In the printer dialog show the to page as the last page
'printing = P.ShowPrinter        ' Show printer

With CRViewer91

        .ReportSource = CRReport1
        .ViewReport
        .Refresh
End With
Set CRApp = Nothing
Set CRReport = Nothing
End Sub

