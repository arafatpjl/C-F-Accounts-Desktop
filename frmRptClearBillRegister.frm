VERSION 5.00
Begin VB.Form frmRptClearBillRegister 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report - Clearing Bill Register"
   ClientHeight    =   4440
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "frmRptClearBillRegister.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   465
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   1995
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
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   1275
      End
   End
   Begin VB.CheckBox chkPSA 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Show PSA Job # Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1875
      TabIndex        =   8
      Top             =   3300
      Width           =   1965
   End
   Begin VB.CheckBox chkMonth 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Show Month Name as Report Header"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1875
      TabIndex        =   7
      Top             =   2790
      Width           =   1965
   End
   Begin VB.CheckBox chkUnAdj 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Show Unadjust Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1875
      TabIndex        =   6
      Top             =   2460
      Width           =   1965
   End
   Begin VB.CheckBox chkJobSL 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Show Job # Ascending"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1875
      TabIndex        =   5
      Top             =   2130
      Width           =   1965
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
      Left            =   1875
      TabIndex        =   1
      Top             =   1755
      Width           =   2025
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
      Left            =   1875
      TabIndex        =   0
      Top             =   1380
      Width           =   2025
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3915
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
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save Record"
      Top             =   3915
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   3720
      Width           =   8895
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clearing Bill Register"
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
      TabIndex        =   15
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   14
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
      Index           =   3
      Left            =   -30
      TabIndex        =   13
      Top             =   570
      Width           =   5685
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2910
   End
   Begin VB.Label Label3 
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
      Index           =   0
      Left            =   960
      TabIndex        =   12
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   9
      Top             =   3690
      Width           =   6405
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date From  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   645
      TabIndex        =   4
      Top             =   1335
      Width           =   1125
   End
End
Attribute VB_Name = "frmRptClearBillRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdShow_Click()
On Error GoTo ErrorHandler

    Dim strHeader As String

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
    'varNullDate = "30-DEC-1899"
    varNullDate = "30-DEC-1900"

    Screen.MousePointer = vbHourglass

    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "DELETE FROM Temp_Indent_Statement;"
        MainComm.Execute
        If ChkBank.Value = 0 Then
        MainComm.CommandText = "INSERT INTO Temp_Indent_Statement(JobNo,InvNo,PayDate,IndAmount,BPayDate,BPayValue,TotIndValue,TotBAmount,AccName,Remarks,AdjDate,AdjAmount,BAdjAmount,TotAdjust,AdjTax,AdjVat,compName,compAdd)" _
            & " SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.TotBAmount,new_Acc_Head.AccName," _
            & " new_Indent_Information.Remarks,new_Indent_Information.AdjDate,new_Indent_Information.AdjAmount,new_Indent_Information.BAdjAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat," _
            & " Company_Information.compName,Company_Information.compAdd FROM (Company_Information INNER JOIN new_Indent_Information" _
            & " ON Company_Information.compID = new_Indent_Information.compID) INNER JOIN new_Acc_Head ON new_Indent_Information.AccID = new_Acc_Head.AccID" _
            & " WHERE (((new_Indent_Information.PayDate) Between '" & varBeginDate & "' And '" & varEndDate & "') AND ((new_Indent_Information.compID)=" & bytCompID & ")" _
            & " AND ((new_Indent_Information.Status)='I'));"
        MainComm.Execute
        Else
        MainComm.CommandText = "INSERT INTO Temp_Indent_Statement(JobNo,InvNo,PayDate,IndAmount,BPayDate,BPayValue,TotIndValue,TotBAmount,AccName,Remarks,AdjDate,AdjAmount,BAdjAmount,TotAdjust,AdjTax,AdjVat,compName,compAdd)" _
            & " SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.TotBAmount,new_Acc_Head.AccName," _
            & " new_Indent_Information.Remarks,new_Indent_Information.AdjDate,new_Indent_Information.AdjAmount,new_Indent_Information.BAdjAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat," _
            & " Company_Information.compName,Company_Information.compAdd FROM (Company_Information INNER JOIN new_Indent_Information" _
            & " ON Company_Information.compID = new_Indent_Information.compID) INNER JOIN new_Acc_Head ON new_Indent_Information.AccID = new_Acc_Head.AccID" _
            & " WHERE (((new_Indent_Information.BPayDate) Between '" & varBeginDate & "' And '" & varEndDate & "') AND ((new_Indent_Information.compID)=" & bytCompID & ")" _
            & " AND ((new_Indent_Information.Status)='I'));"
        MainComm.Execute
        End If
        strHeader = "Clearing Bill Register"
        
        If chkPSA.Value = 1 Then
            strHeader = strHeader & " - PSA "
            
            MainComm.CommandText = "DELETE  FROM Temp_Indent_Statement WHERE RIGHT([JobNo],3) <> 'PSA';"
            MainComm.Execute
        End If

        strRecord = "SELECT JobNo FROM Temp_Indent_Statement GROUP BY JobNo;"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount = 0 Then
            MsgBox "Record not Found", vbInformation, cnstMsgInfo
            Screen.MousePointer = vbDefault
            
            MainConn.RollbackTrans
        Exit Sub
        End If
        
        R.Close
        Set R = Nothing

        If chkUnAdj.Value = 1 Then
            strHeader = strHeader & " (Unadjusted Only)"
            
            MainComm.CommandText = "DELETE  FROM Temp_Indent_Statement WHERE AdjAmount<>0;"
            MainComm.Execute
        End If

        MainComm.CommandText = "UPDATE Temp_Indent_Statement SET Header='" & UCase(strHeader) & "';"
        MainComm.Execute

        strHeader = "Period from payment date " & txtFromDate.Text & " To " & txtToDate.Text
        If chkMonth.Value = 1 Then
            strHeader = "Period for the payment month of " & Format(txtToDate.Text, "MMMM") & ", " & Format(txtToDate.Text, "YYYY")
        End If
        MainComm.CommandText = "UPDATE Temp_Indent_Statement SET SubHeader='" & UCase(strHeader) & "';"
        MainComm.Execute

        MainComm.CommandText = "UPDATE Temp_Indent_Statement SET AdjDate=NULL WHERE AdjDate='" & varNullDate & "';"
        MainComm.Execute
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        
        MainConn.RollbackTrans
    Exit Sub
    End If

     If chkJobSL.Value = 1 Then
       Set Report = Appl.OpenReport(ReportPath + "\clear bill job sl.rpt")
    Else
        Set Report = Appl.OpenReport(ReportPath + "\clear bill register.rpt")
    End If
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.Show 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRptClearBillRegister = Nothing
End Sub
Private Sub txtFromDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtToDate.SetFocus
End Sub
Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub
Private Sub txtFromDate_LostFocus()
    If Len(txtFromDate) = 0 Then Exit Sub
    txtFromDate.Text = FormatDate(txtFromDate)
End Sub
Private Sub txtToDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdShow.SetFocus
End Sub
Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub
Private Sub txtToDate_LostFocus()
    If Len(txtToDate) = 0 Then Exit Sub
    txtToDate.Text = FormatDate(txtToDate)
End Sub
