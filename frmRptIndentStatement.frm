VERSION 5.00
Begin VB.Form frmRptIndentStatement 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report - Indent Statement"
   ClientHeight    =   4455
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "frmRptIndentStatement.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkBank 
      BackColor       =   &H00EAF5F5&
      Caption         =   "BankDate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2850
      TabIndex        =   14
      Top             =   2055
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   525
      Left            =   690
      TabIndex        =   8
      Top             =   960
      Width           =   4035
      Begin VB.OptionButton optExport 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2190
         TabIndex        =   1
         Top             =   180
         Width           =   975
      End
      Begin VB.OptionButton optImport 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   540
         TabIndex        =   0
         Top             =   180
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAF5F5&
      Height          =   525
      Left            =   690
      TabIndex        =   11
      Top             =   1380
      Width           =   4035
      Begin VB.OptionButton optDetails 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   510
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optSummary 
         BackColor       =   &H00EAF5F5&
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2220
         TabIndex        =   3
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.TextBox txtToDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2850
      TabIndex        =   5
      Top             =   2865
      Width           =   1875
   End
   Begin VB.TextBox txtFromDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2850
      TabIndex        =   4
      Top             =   2415
      Width           =   1875
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3930
      Width           =   1185
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   6
      ToolTipText     =   "Save Record"
      Top             =   3930
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   3720
      Width           =   8895
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indent Statement"
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
      Left            =   360
      TabIndex        =   17
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   -240
      TabIndex        =   16
      Top             =   0
      Width           =   8895
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
      Left            =   0
      TabIndex        =   15
      Top             =   795
      Width           =   5685
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
      Left            =   0
      TabIndex        =   13
      Top             =   3690
      Width           =   5685
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1290
      TabIndex        =   12
      Top             =   2070
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Date  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   10
      Top             =   2895
      Width           =   1365
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Date  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1320
      TabIndex        =   9
      Top             =   2445
      Width           =   1365
   End
End
Attribute VB_Name = "frmRptIndentStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
    Dim strStatus As String

    If Len(txtFromDate) = 0 Then
        MsgBox "Enter From Date of Payment", vbInformation, cnstMsgInfo
        txtFromDate.SetFocus
    Exit Sub
    End If
    
    If Len(txtToDate) = 0 Then
        MsgBox "Enter To Date of Payment", vbInformation, cnstMsgInfo
        txtToDate.SetFocus
    Exit Sub
    End If

    varBeginDate = Format(txtFromDate.Text, cnstDtFrmtI)
    varEndDate = Format(txtToDate.Text, cnstDtFrmtI)

    If optImport.Value = True Then strStatus = "I" Else strStatus = "E"

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "DELETE  FROM Temp_Indent_Statement;"
        MainComm.Execute

        If optDetails.Value = True And ChkBank.Value = 0 Then
            MainComm.CommandText = "INSERT INTO Temp_Indent_Statement(JobNo,InvNo,PayDate,IndAmount,BPayDate,BPayValue,TotBAmount,TotIndValue,AccName,Remarks,AdjDate,AdjAmount,BAdjAmount,TotAdjust,AdjTax,AdjVat,compName,compAdd)" _
                & " SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotBAmount,new_Indent_Information.TotIndValue,new_Acc_Head.AccName," _
                & " new_Indent_Information.Remarks,new_Indent_Information.AdjDate,new_Indent_Information.AdjAmount,new_Indent_Information.BAdjAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat," _
                & " Company_Information.compName,Company_Information.compAdd FROM (Company_Information INNER JOIN new_Indent_Information" _
                & " ON Company_Information.compID = new_Indent_Information.compID) INNER JOIN new_Acc_Head ON new_Indent_Information.AccID = new_Acc_Head.AccID" _
                & " WHERE (((new_Indent_Information.PayDate) Between '" & varBeginDate & "' And '" & varEndDate & "') AND ((new_Indent_Information.compID)=" & bytCompID & ")" _
                & " AND ((new_Indent_Information.Status)='" & strStatus & "'));"
            MainComm.Execute

            'varNullDate = "#" + "Dec-30-1899" + "#"
             varNullDate = "30-DEC-1900"
            MainComm.CommandText = "UPDATE Temp_Indent_Statement SET AdjDate=NULL WHERE AdjDate='" & varNullDate & "';"
            MainComm.Execute

        ElseIf optSummary.Value = True And ChkBank.Value = 0 Then

            MainComm.CommandText = "INSERT INTO Temp_Indent_Statement(PayDate,TotIndValue,compName,compAdd) SELECT new_Indent_Information.PayDate,Sum(new_Indent_Information.TotIndValue) AS TIndAmount," _
                & " Company_Information.compName,Company_Information.compAdd FROM Company_Information INNER JOIN new_Indent_Information ON Company_Information.compID = new_Indent_Information.compID" _
                & " WHERE (((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='" & strStatus & "')) GROUP BY new_Indent_Information.PayDate," _
                & " Company_Information.compName, Company_Information.compAdd HAVING (((new_Indent_Information.PayDate) Between '" & varBeginDate & "' And '" & varEndDate & "'));"
            MainComm.Execute
        End If
      
        '''Added By Shemul
        If optDetails.Value = True And ChkBank.Value = 1 Then
                    MainComm.CommandText = "INSERT INTO Temp_Indent_Statement(JobNo,InvNo,PayDate,IndAmount,BPayDate,BPayValue,TotBAmount,TotIndValue,AccName,Remarks,AdjDate,AdjAmount,BAdjAmount,TotAdjust,AdjTax,AdjVat,compName,compAdd)" _
                & " SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotBAmount,new_Indent_Information.TotIndValue,new_Acc_Head.AccName," _
                & " new_Indent_Information.Remarks,new_Indent_Information.AdjDate,new_Indent_Information.AdjAmount,new_Indent_Information.BAdjAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat," _
                & " Company_Information.compName,Company_Information.compAdd FROM (Company_Information INNER JOIN new_Indent_Information" _
                & " ON Company_Information.compID = new_Indent_Information.compID) INNER JOIN new_Acc_Head ON new_Indent_Information.AccID = new_Acc_Head.AccID" _
                & " WHERE (((new_Indent_Information.BPayDate) Between '" & varBeginDate & "' And '" & varEndDate & "') AND ((new_Indent_Information.compID)=" & bytCompID & ")" _
                & " AND ((new_Indent_Information.Status)='" & strStatus & "'));"
                MainComm.Execute

                'varNullDate = "#" + "Dec-30-1899" + "#"
                  varNullDate = "30-DEC-1900"
                  
                MainComm.CommandText = "UPDATE Temp_Indent_Statement SET AdjDate=NULL WHERE AdjDate='" & varNullDate & "';"
            MainComm.Execute
                
         ElseIf optSummary.Value = True And ChkBank.Value = 1 Then
                MainComm.CommandText = "INSERT INTO Temp_Indent_Statement(BPayDate,TotIndValue,compName,compAdd) SELECT new_Indent_Information.BPayDate,Sum(new_Indent_Information.TotIndValue) AS TIndAmount," _
                & " Company_Information.compName,Company_Information.compAdd FROM Company_Information INNER JOIN new_Indent_Information ON Company_Information.compID = new_Indent_Information.compID" _
                & " WHERE (((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='" & strStatus & "')) GROUP BY new_Indent_Information.BPayDate," _
                & " Company_Information.compName, Company_Information.compAdd HAVING (((new_Indent_Information.BPayDate) Between '" & varBeginDate & "' And '" & varEndDate & "'));"
            MainComm.Execute
        End If
        
      
        
'        strRecord = "SELECT PayDate FROM Temp_Indent_Statement GROUP BY PayDate;"
        ''Added By Shemul
        If ChkBank.Value = 1 Then
            strRecord = "SELECT BPayDate FROM Temp_Indent_Statement GROUP BY BPayDate;"
            Else
            strRecord = "SELECT PayDate FROM Temp_Indent_Statement GROUP BY PayDate;"
        End If
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount = 0 Then
            MsgBox "Record not Found", vbInformation, cnstMsgInfo
            Screen.MousePointer = vbDefault
            MainConn.RollbackTrans
        Exit Sub
        End If
        R.Close
        Set R = Nothing
    
        MainComm.CommandText = "UPDATE Temp_Indent_Statement SET BeginDate='" & varBeginDate & "',EndDate='" & varEndDate & "';"
        MainComm.Execute
    MainConn.CommitTrans
    
    If optDetails.Value = True Then
'    Set Report = Appl.OpenReport(ReportPath + "\indent statement.rpt")
     Set Report = Appl.OpenReport(ReportPath & "indent statement.rpt")
Else
'    Set Report = Appl.OpenReport(ReportPath + "\indent st summary.rpt")
      Set Report = Appl.OpenReport(ReportPath & "indent st summary.rpt")
End If
'       Set Report = Appl.OpenReport(
        frmMainReport.CRVIEWER.ReportSource = Report
        frmMainReport.CRVIEWER.ViewReport
        frmMainReport.Show 1
        Screen.MousePointer = vbDefault

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        
        MainConn.RollbackTrans
    Exit Sub
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRptIndentStatement = Nothing
End Sub

Private Sub optDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optExport_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optImport_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optSummary_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFromDate_GotFocus()
    Call txtGotFocus(txtFromDate)
End Sub

Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtFromDate_LostFocus()
    If Len(txtFromDate) = 0 Then Exit Sub
    txtFromDate.Text = FormatDate(txtFromDate)
End Sub

Private Sub txtToDate_GotFocus()
    Call txtGotFocus(txtToDate)
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub

Private Sub txtToDate_LostFocus()
    If Len(txtToDate) = 0 Then Exit Sub
    txtToDate.Text = FormatDate(txtToDate)
End Sub
