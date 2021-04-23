VERSION 5.00
Begin VB.Form frmRptStatementForwarderDate 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C and F Statement Report Forwarding Date"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
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
   Icon            =   "frmRptStatementForwarderDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox txtVoucherNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2880
      Width           =   2355
   End
   Begin VB.ComboBox cmbForwardFromDate 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2205
      Width           =   2355
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
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1275
      Width           =   3345
   End
   Begin VB.TextBox txtForwardDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
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
      Left            =   2910
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2205
      Width           =   1755
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
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3540
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
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3540
      Width           =   1185
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No :"
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
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   1365
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Width           =   8895
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "C && F Statement Forwarding Date"
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
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label3 
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
      TabIndex        =   10
      Top             =   3885
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
      TabIndex        =   9
      Top             =   795
      Width           =   5685
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
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
      Left            =   2940
      TabIndex        =   8
      Top             =   1965
      Width           =   1365
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
      Left            =   360
      TabIndex        =   7
      Top             =   1005
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Forward From Date :"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1935
      Width           =   2115
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
      Left            =   -15
      TabIndex        =   5
      Top             =   3300
      Width           =   6435
   End
End
Attribute VB_Name = "frmRptStatementForwarderDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmprst As Recordset
Dim bytRowNo As Byte

Private Sub cmbForwardFromDate_Click()
    txtForwardDate.Text = cmbForwardFromDate.Text
    Set R = New ADODB.Recordset
'        cmdSave.Caption = "&Save"
'        cmdDelete.Enabled = False
'        cmdPreview.Enabled = False
        
'        Frame1.Visible = False
    
        strRecord = "SELECT DISTINCT VoucherNo FROM Query_C_And_F_Out WHERE CompID = " & bytCompID & " and PartyID = " & cmbPartyName.Tag & " And  ForwardDate = '" & Format(cmbForwardFromDate.Text, cnstDtFrmtI) & "' ;"
        R.Open strRecord, MainConn, adOpenStatic
        txtVoucherNo.Clear
            
    If R.EOF = False And R.BOF = False Then
        Do While Not R.EOF
        txtVoucherNo.AddItem R![VoucherNo]
        If param = 1 Then txtVoucherNo.ItemData(txtVoucherNo.NewIndex) = R![VoucherNo]

        R.MoveNext
        Loop
    End If
    Set R = Nothing
End Sub

Private Sub cmbForwardFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

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

         Set Report = Appl.OpenReport(ReportPath & "statement_of_c_and_f_bills_ForwardingDate.rpt")

        
         Mysql = "Select * from Query_C_And_F_Out Where compid = " & bytCompID & " and PartyID = " & cmbPartyName.Tag & " And ((ForwardDate) Between '" & cmbForwardFromDate.Text & "' And '" & txtForwardDate.Text & "') Order By slno;"
        
         
         Report.FormulaFields(1).Text = "'Date From : " & Format(cmbForwardFromDate.Text, "DD-MMM-YYYY") & "  To  " & Format(txtForwardDate.Text, "DD-MMM-YYYY") & "'"


        frmMainReport.CRVIEWER.Refresh
        frmMainReport.CRVIEWER.ReportSource = Report
        Report.SQLQueryString = Mysql
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
        Screen.MousePointer = vbDefault


End Sub


Private Sub cmbPartyName_Click()
    cmbPartyName.Tag = 0
    If cmbPartyName.ListIndex >= 0 Then
        cmbPartyName.Tag = cmbPartyName.ItemData(cmbPartyName.ListIndex)
        cmbPartyName.ToolTipText = cmbPartyName.Text
    End If
    
    Set R = New ADODB.Recordset
        strRecord = "SELECT PartyID, ForwardDate From C_And_F_Out_Main Where  compid = " & bytCompID & "  GROUP BY PartyID, ForwardDate Having ((PartyID)=" & cmbPartyName.Tag & ") Order By ForwardDate;"
        R.Open strRecord, MainConn, adOpenStatic
        
        cmbForwardFromDate.Clear
        If R.EOF = False And R.BOF = False Then
            Do While Not R.EOF
                cmbForwardFromDate.AddItem R.Fields(1).Value
            R.MoveNext
            Loop
        End If
    Set R = Nothing
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
End Sub

