VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptStatementOfCandFBills 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C and F Statement Report"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRptStatementOfCandFBills.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optForwardDate 
      BackColor       =   &H00EAF5F5&
      Height          =   345
      Left            =   4785
      TabIndex        =   23
      Top             =   1215
      Value           =   -1  'True
      Width           =   210
   End
   Begin VB.OptionButton optBENo 
      BackColor       =   &H00EAF5F5&
      Height          =   345
      Left            =   1830
      TabIndex        =   22
      Top             =   3570
      Width           =   285
   End
   Begin VB.OptionButton opthawb 
      BackColor       =   &H00EAF5F5&
      Height          =   345
      Left            =   4410
      TabIndex        =   21
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optBillDate 
      BackColor       =   &H00EAF5F5&
      Height          =   345
      Left            =   4200
      TabIndex        =   20
      Top             =   3600
      Width           =   285
   End
   Begin VB.OptionButton optBillNo 
      BackColor       =   &H00EAF5F5&
      Height          =   345
      Left            =   1650
      TabIndex        =   19
      Top             =   2880
      Width           =   285
   End
   Begin VB.OptionButton optVocherNo 
      BackColor       =   &H00EAF5F5&
      Height          =   345
      Left            =   4020
      TabIndex        =   18
      Top             =   2190
      Width           =   255
   End
   Begin VB.ComboBox cmbPartyName 
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1230
      Width           =   3105
   End
   Begin VB.TextBox txtVoucherNo 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   150
      MaxLength       =   25
      TabIndex        =   2
      Top             =   2190
      Width           =   3795
   End
   Begin VB.TextBox txtBENO 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   150
      MaxLength       =   20
      TabIndex        =   6
      Top             =   3600
      Width           =   1665
   End
   Begin VB.TextBox txtBillNo 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   120
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtBillDate 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   2340
      MaxLength       =   15
      TabIndex        =   4
      Top             =   3630
      Width           =   1785
   End
   Begin VB.TextBox txtForwardDate 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   3300
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1230
      Width           =   1485
   End
   Begin VB.TextBox txtHAWB 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   2340
      MaxLength       =   20
      TabIndex        =   5
      Top             =   2880
      Width           =   2025
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3525
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1185
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   28
      Top             =   4200
      Width           =   8895
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "C && F Statement"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   -15
      TabIndex        =   25
      Top             =   4785
      Width           =   5685
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   -15
      TabIndex        =   24
      Top             =   555
      Width           =   5685
   End
   Begin VB.Label lblbeno 
      BackStyle       =   0  'Transparent
      Caption         =   "B/E No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   17
      Left            =   180
      TabIndex        =   17
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Vocher No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   14
      Left            =   150
      TabIndex        =   16
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   13
      Left            =   2340
      TabIndex        =   15
      Top             =   3390
      Width           =   1365
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   12
      Left            =   150
      TabIndex        =   14
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
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
      Height          =   45
      Index           =   10
      Left            =   150
      TabIndex        =   13
      Top             =   1740
      Width           =   5655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "HAWB :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2340
      TabIndex        =   12
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Forward Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3255
      TabIndex        =   10
      Top             =   960
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   45
      Left            =   0
      TabIndex        =   9
      Top             =   4185
      Width           =   6435
   End
End
Attribute VB_Name = "frmRptStatementOfCandFBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmprst As Recordset
Dim bytRowNo As Byte

Private Sub cmdPreview_Click()
    If Len(cmbPartyName.Text) = 0 Then
        MsgBox "You should provide party name.", vbInformation, App.Title
        cmbPartyName.SetFocus
        Exit Sub
    End If
    If Len(txtForwardDate.Text) = 0 Then
        MsgBox "You should provide forward date.", vbInformation, App.Title
        txtForwardDate.SetFocus
        Exit Sub
    End If
    


     If optForwardDate.Value = True Then
'            With CrystalReport1
                Set Report = Appl.OpenReport(ReportPath + "\statement_of_c_and_f_bills.rpt")
                Set Report = Appl.OpenReport(ReportPath + "\statement_of_c_and_f_bills.rpt")
                Mysql = "Select * from Query_C_And_F_Out Where  compid = " & bytCompID & " and PartyName = '" & cmbPartyName.Text & "' and ForwardDate= '" & txtForwardDate.Text & "';"
'                .PrintReport
'            End With
        
        ElseIf optVocherNo.Value = True Then
'            With CrystalReport1
                Set Report = Appl.OpenReport(ReportPath + "\statement_of_c_and_f_bills_voucher.rpt")
                Set Report = Appl.OpenReport(ReportPath + "\statement_of_c_and_f_bills_voucher.rpt")
                Mysql = "Select * from Query_C_And_F_Out Where  compid = " & bytCompID & " and PartyName = '" & cmbPartyName.Text & "' and ForwardDate= '" & txtForwardDate.Text & "' and VoucherNo = '" & txtVoucherNo.Text & "';"
'                .PrintReport
'            End With
        ElseIf optBillNo.Value = True Then
'            With CrystalReport1
                Set Report = Appl.OpenReport(ReportPath + "\statement_of_c_and_f_bills_billsno.rpt")
                Set Report = Appl.OpenReport(ReportPath + "\statement_of_c_and_f_bills_billsno.rpt")
                
                Mysql = "Select * from Query_C_And_F_Out Where  compid = " & bytCompID & " and PartyName = '" & cmbPartyName.Text & "' and ForwardDate= '" & txtForwardDate.Text & "' and BillNo = '" & txtBillNo.Text & "';"
'                .PrintReport
'            End With

        
        ElseIf opthawb.Value = True Then

                Set Report = Appl.OpenReport(ReportPath + "\statement_of_c_and_f_bills_hawb.rpt")
                Set Report = Appl.OpenReport(ReportPath + "\statement_of_c_and_f_bills_hawb.rpt")
                
                Mysql = "Select * from Query_C_And_F_Out Where  compid = " & bytCompID & " and PartyName = '" & cmbPartyName.Text & "' and ForwardDate= '" & txtForwardDate.Text & "' and hawb = '" & txtHAWB.Text & "';"

        ElseIf optBENo.Value = True Then

        ElseIf optBillDate.Value = True Then

        
        End If
        
        frmMainReport.CRVIEWER.ReportSource = Report
        Report.SQLQueryString = Mysql
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
        Screen.MousePointer = vbDefault
        
        
        '        If optForwardDate.Value = True Then
'            With CrystalReport1
'                .ReportFileName = ReportPath + "\statement_of_c_and_f_bills.rpt"
'                .PrintFileName = ReportPath + "\statement_of_c_and_f_bills.rpt"
'                .SQLQuery = "Select * from Query_C_And_F_Out Where  compid = " & bytCompID & " and PartyName = '" & cmbPartyName.Text & "' and ForwardDate= '" & txtForwardDate.Text & "';"
'                .PrintReport
'            End With
'
'        ElseIf optVocherNo.Value = True Then
'            With CrystalReport1
'                .ReportFileName = ReportPath + "\statement_of_c_and_f_bills_voucher.rpt"
'                .PrintFileName = ReportPath + "\statement_of_c_and_f_bills_voucher.rpt"
'                .SQLQuery = "Select * from Query_C_And_F_Out Where  compid = " & bytCompID & " and PartyName = '" & cmbPartyName.Text & "' and ForwardDate= '" & txtForwardDate.Text & "' and VoucherNo = '" & txtVoucherNo.Text & "';"
'                .PrintReport
'            End With
'        ElseIf optBillNo.Value = True Then
'            With CrystalReport1
'                .ReportFileName = ReportPath + "\statement_of_c_and_f_bills_billsno.rpt"
'                .PrintFileName = ReportPath + "\statement_of_c_and_f_bills_billsno.rpt"
'
'                .SQLQuery = "Select * from Query_C_And_F_Out Where  compid = " & bytCompID & " and PartyName = '" & cmbPartyName.Text & "' and ForwardDate= '" & txtForwardDate.Text & "' and BillNo = '" & txtBillNo.Text & "';"
'                .PrintReport
'            End With
'
'
'        ElseIf opthawb.Value = True Then
'            With CrystalReport1
'                .ReportFileName = ReportPath + "\statement_of_c_and_f_bills_hawb.rpt"
'                .PrintFileName = ReportPath + "\statement_of_c_and_f_bills_hawb.rpt"
'
'                .SQLQuery = "Select * from Query_C_And_F_Out Where  compid = " & bytCompID & " and PartyName = '" & cmbPartyName.Text & "' and ForwardDate= '" & txtForwardDate.Text & "' and hawb = '" & txtHAWB.Text & "';"
'                .PrintReport
'            End With
'        ElseIf optBENo.Value = True Then
''            .ReportFileName = App.Path + "\statement_of_c_and_f_bills.rpt"
''            .PrintFileName = App.Path + "\statement_of_c_and_f_bills.rpt"
''
''            .SQLQuery = "Select * from Query_C_And_F_Out Where M_ID = " & flexDetails.Tag & ";"
''            .PrintReport
'        ElseIf optBillDate.Value = True Then
''            .ReportFileName = App.Path + "\statement_of_c_and_f_bills_billdate.rpt"
''            .PrintFileName = App.Path + "\statement_of_c_and_f_bills_billdate.rpt"
''
''            .SQLQuery = "Select * from Query_C_And_F_Out Where M_ID = " & flexDetails.Tag & ";"
''            .PrintReport
'
'        End If

   
    
End Sub

Private Sub txtBillDate_GotFocus()
    Call txtGotFocus(txtBillDate)
End Sub

Private Sub txtBillDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBillDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtBillDate_LostFocus()
    Call FormatDate(txtBillDate)
End Sub


Private Sub txtBillNo_GotFocus()
    Call txtGotFocus(txtBillNo)
End Sub

Private Sub cmbPartyName_Click()
    cmbPartyName.Tag = 0
    If cmbPartyName.ListIndex >= 0 Then
        cmbPartyName.Tag = cmbPartyName.ItemData(cmbPartyName.ListIndex)
        cmbPartyName.ToolTipText = cmbPartyName.Text
    End If
End Sub

Private Sub cmbPartyName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
        Call prcAddParty(cmbPartyName)
'    Call prcGridSetting
End Sub

Public Function fncBlank(param As Integer) As Boolean
fncBlank = True
    If param = 1 Then
        If Len(txtBillNo.Text) = 0 Then
            MsgBox "You should provide bill no.", vbCritical, App.Title
            txtBillNo.SetFocus
            Exit Function
        End If

        If Len(txtBillDate.Text) = 0 Then
            MsgBox "You should bill date.", vbCritical, App.Title
            txtBillDate.SetFocus
            Exit Function
        End If
    
    ElseIf param = 2 Then
        If Len(cmbPartyName.Text) = 0 Then
            MsgBox "You should provide party name.", vbCritical, App.Title
            cmbPartyName.SetFocus
            Exit Function
        End If
    End If
fncBlank = False
End Function

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBillNo_LostFocus()
    txtBillNo.Text = UCase(Trim(txtBillNo.Text))
End Sub

Private Sub txtForwardDate_GotFocus()
    Call txtGotFocus(txtForwardDate)
End Sub

Private Sub txtForwardDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtForwardDate_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 46 Then Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtForwardDate_LostFocus()
    Call FormatDate(txtForwardDate)
    
    If Len(cmbPartyName.Text) = 0 Or Len(txtForwardDate.Text) = 0 Then
        txtVoucherNo.Text = ""
'        flexDetails.Tag = 0: flexDetails.Rows = 1
    Else
        txtVoucherNo.Text = ""
'        flexDetails.Tag = 0: flexDetails.Rows = 1
        
'        Call cmdFind_Click
    End If
End Sub

Private Sub txtVoucherNo_GotFocus()
    Call txtGotFocus(txtVoucherNo)
End Sub

Private Sub txtVoucherNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVoucherNo_LostFocus()
    txtVoucherNo.Text = Trim(UCase(txtVoucherNo.Text))
End Sub

Private Sub txtHAWB_GotFocus()
    Call txtGotFocus(txtHAWB)
End Sub

Private Sub txtHAWB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHAWB_LostFocus()
    txtHAWB.Text = UCase(Trim(txtHAWB.Text))
End Sub

Private Sub txtbeno_GotFocus()
    Call txtGotFocus(txtBENO)
End Sub

Private Sub txtbeno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBEno_LostFocus()
    txtBENO.Text = UCase(Trim(txtBENO.Text))
End Sub
