VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   Caption         =   "Crystal Report Viewer 10 Example by Tan Jiunn Yann"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Group 2"
      Height          =   735
      Left            =   2280
      TabIndex        =   11
      Top             =   1560
      Width           =   6015
      Begin VB.ComboBox cmbField 
         Height          =   315
         Index           =   1
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbOpr 
         Height          =   315
         Index           =   1
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtField 
         Height          =   285
         Index           =   1
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group 1"
      Height          =   735
      Left            =   2280
      TabIndex        =   7
      Top             =   600
      Width           =   6015
      Begin VB.ComboBox cmbField 
         Height          =   315
         Index           =   0
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbOpr 
         Height          =   315
         Index           =   0
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtField 
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "No filtering"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.ComboBox cmbOpr 
      Height          =   315
      Index           =   2
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.OptionButton optFilter 
      Caption         =   "Filter the recports by:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "View Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtAuth 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "Your Name"
      Top             =   240
      Width           =   2415
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRView 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   10695
      lastProp        =   600
      _cx             =   18865
      _cy             =   8705
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "This report is prepared by:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRecordsFormula As String
Private Sub InitCrystalReport()
    ' Make sure the "Crystal Report ActiveX Designer Run Time Library 10"
    ' is referenced in the current project
    
    Dim CRReport As CRAXDRT.Report
    Dim CRApp As New CRAXDRT.Application
    
    'Open the Crystal Report
    Set CRReport = CRApp.OpenReport(App.Path & "\report1.rpt")
    
    With CRReport
        'Set the Access database path
        .Database.Tables(1).Location = App.Path & "\data.mdb"
        
        'This will only retrieve the records that filtered by user
        If optFilter(0).Value = True And strRecordsFormula <> "" Then
            .RecordSelectionFormula = strRecordsFormula
        End If
        
        'The following is passing a value inside the crystal report's text field
        'You can open the report file in the Crystal Report, and you'll see a
        'text field name : @txtAuth
        
        If Len(Trim(txtAuth.Text)) > 0 Then
            .FormulaFields(1).Text = "stringvar strAuth; strAuth:='" & txtAuth & "'"
        End If
        
    End With
    

    With CRView
        .ReportSource = CRReport
        .ViewReport
        While .IsBusy
            DoEvents
        Wend
        .Visible = False
        .Refresh
    End With
    
    
    Set CRApp = Nothing
    Set CRReport = Nothing

End Sub

Private Sub cmdRefresh_Click()
    Dim strCriteria1 As String
    Dim strCriteria2 As String
    
    strRecordsFormula = ""
    
    If cmbField(0).Text <> "" And cmbOpr(0).Text <> "" And txtField(0).Text <> "" Then
        
        If cmbField(0).Text = "ID" Or cmbField(0).Text = "Description" Then
            strCriteria1 = "'" & txtField(0).Text & "'"
        End If
        
        If cmbField(0).Text = "Quantity" Or cmbField(0).Text = "Price" Then
            If IsNumeric(txtField(0).Text) = False Then
                MsgBox "Please put a number!"
                txtField(0).SetFocus
                txtField(0).SelStart = 0
                txtField(0).SelLength = Len(txtField(0).Text)
                Exit Sub
            Else
                strCriteria1 = Val(txtField(0).Text)
            End If
        End If
        
        strRecordsFormula = "{Stock." & cmbField(0).Text & "} " & cmbOpr(0).Text & " " & strCriteria1
    End If
    
    If cmbField(1).Text <> "" And cmbOpr(1).Text <> "" And cmbOpr(2).Text <> "" And txtField(1).Text <> "" Then
        
        If cmbField(1).Text = "ID" Or cmbField(1).Text = "Description" Then
            strCriteria2 = "'" & txtField(1).Text & "'"
        End If
        
        If cmbField(1).Text = "Quantity" Or cmbField(1).Text = "Price" Then
            If IsNumeric(txtField(1).Text) = False Then
                MsgBox "Please put a number!"
                txtField(1).SetFocus
                txtField(1).SelStart = 0
                txtField(1).SelLength = Len(txtField(1).Text)
                Exit Sub
            Else
                strCriteria2 = Val(txtField(1).Text)
            End If
        End If
        
        strRecordsFormula = strRecordsFormula & " " & cmbOpr(2).Text & " {Stock." & cmbField(1).Text & "} " & cmbOpr(1).Text & " " & strCriteria2
    End If
    
    
    
    InitCrystalReport
End Sub

Private Sub Form_Load()
    For x = 0 To 1
        cmbField(x).AddItem "ID"
        cmbField(x).AddItem "Description"
        cmbField(x).AddItem "Quantity"
        cmbField(x).AddItem "Price"
    Next
    
    For x = 0 To 1
        cmbOpr(x).AddItem "="
        cmbOpr(x).AddItem ">"
        cmbOpr(x).AddItem "<"
        cmbOpr(x).AddItem ">="
        cmbOpr(x).AddItem "<="
        cmbOpr(x).AddItem "LIKE"
    Next
    
    cmbOpr(2).AddItem "AND"
    cmbOpr(2).AddItem "OR"
    
    DisableGroup
    'InitCrystalReport
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    CRView.Width = Me.ScaleWidth - 50
    CRView.Height = Me.ScaleHeight - 2400
End Sub
Private Sub EnableGroup()
    Frame1.Enabled = True
    Frame2.Enabled = True
    For x = 0 To 1
        cmbField(x).Enabled = True
        cmbOpr(x).Enabled = True
        txtField(x).Enabled = True
    Next
    
    cmbOpr(2).Enabled = True
End Sub
Private Sub DisableGroup()
    Frame1.Enabled = False
    Frame2.Enabled = False
    For x = 0 To 1
        cmbField(x).Enabled = False
        cmbOpr(x).Enabled = False
        txtField(x).Enabled = False
    Next
    
    cmbOpr(2).Enabled = False
End Sub

Private Sub optFilter_Click(Index As Integer)
    If Index = 0 Then
        EnableGroup
    Else
        DisableGroup
    End If
End Sub
