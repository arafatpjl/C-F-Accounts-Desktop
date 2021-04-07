VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808000&
   Caption         =   "C & F Indent and Adjustment System"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "frmMDIMain.frx":0000
   LinkTopic       =   "PictureClip1"
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Banner1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10665
      Left            =   0
      ScaleHeight     =   10605
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.PictureBox Banner2 
         BackColor       =   &H00000080&
         Height          =   735
         Left            =   0
         ScaleHeight     =   675
         ScaleMode       =   0  'User
         ScaleWidth      =   15225
         TabIndex        =   3
         Top             =   0
         Width           =   15290
         Begin VB.Label lblCompName 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000018&
            Height          =   375
            Left            =   0
            TabIndex        =   4
            Top             =   120
            Width           =   15135
         End
      End
      Begin VB.TextBox txtUserPassword 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   0
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "C && F                     ACCOUNTS"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3855
         Left            =   3240
         TabIndex        =   2
         Top             =   3360
         Width           =   9375
      End
   End
   Begin VB.Menu mnuCompInformation 
      Caption         =   "   &Company Information"
      Begin VB.Menu mnuNewComp 
         Caption         =   "New Company"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNext10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpenComp 
         Caption         =   "Open Company"
      End
   End
   Begin VB.Menu mnuHouseKeeping 
      Caption         =   "   &House Keeping"
      Begin VB.Menu mnuNewAccountsHead 
         Caption         =   "New &Accounts Head ..."
      End
      Begin VB.Menu mnuNext1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewExpenses 
         Caption         =   "New &Expenses ..."
      End
      Begin VB.Menu mnuNewParty 
         Caption         =   "New &Party ..."
      End
      Begin VB.Menu mnuSupplier 
         Caption         =   "New &Supplier ..."
      End
      Begin VB.Menu mnuNewBuyer 
         Caption         =   "New &Buyer ..."
      End
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "   Data E&ntry"
      Begin VB.Menu mnuNewIndentInfo 
         Caption         =   "New Indent Information"
      End
      Begin VB.Menu mnuNext4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuInvoice 
         Caption         =   "Invoic Entry"
      End
      Begin VB.Menu mnuNewIndentFund 
         Caption         =   "New Indent Fund"
      End
      Begin VB.Menu mnuCandFBillOut 
         Caption         =   "C&&F Bill [OUT]"
      End
   End
   Begin VB.Menu mnuDataEdit 
      Caption         =   "   Data E&dit"
      Begin VB.Menu mnuEditIndentInfo 
         Caption         =   "Edit Indent Information"
      End
      Begin VB.Menu mnuNext5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditIndentFund 
         Caption         =   "Edit Indent Fund"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "   Data &Report   "
      Begin VB.Menu mnuNext6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptIndentSt 
         Caption         =   "Indent Statement"
      End
      Begin VB.Menu mnuRptAdjustSt 
         Caption         =   "Adjustment Statement"
      End
      Begin VB.Menu mnuNext2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptClearBill 
         Caption         =   "Clearing Bill Register"
      End
      Begin VB.Menu mnuRptForwardBill 
         Caption         =   "Forwarding Bill Register"
      End
      Begin VB.Menu mnustatementofcandfbills 
         Caption         =   "Statement of C and F Bills"
      End
      Begin VB.Menu mnustatementofforwardingdate 
         Caption         =   "Statement of C and F Bills Forwarding Date"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "   Op&tion   "
      Begin VB.Menu mnuModifyComp 
         Caption         =   "Modify Company"
      End
      Begin VB.Menu mnuDeleteComp 
         Caption         =   "Delete Company"
      End
      Begin VB.Menu mnuNext3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateUser 
         Caption         =   "Create or Modify User"
      End
      Begin VB.Menu mnuSetupPermission 
         Caption         =   "Setup Permission "
         Begin VB.Menu mnuUserMenu 
            Caption         =   "User -> Menu"
         End
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "   E&xit   "
      Begin VB.Menu mnuQuit 
         Caption         =   "&Exit C && F ACCOUNT"
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
  'ConnectDatabase
  mnuCompInformation.Visible = False
  mnuHouseKeeping.Visible = True
  mnuDataEntry.Visible = True
  mnuDataEdit.Visible = True
  mnuReport.Visible = True
  mnuOption.Visible = True
  'mnuCloseCompany.Visible = True
  'mnuNext4.Visible = True
End Sub

Private Sub MDIForm_Resize()
    Banner1.Height = Me.Height
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Set frmMDIMain = Nothing
End Sub
Private Sub mnuAbout_Click()
'''  frmAbout.Show 1
End Sub

'Private Sub mnrSupplier_Click()
'frmSupplier.Show 1, Me
'End Sub

Private Sub mnuCandFBillOut_Click()
    frmCandFOut.Show 1, Me
End Sub

Private Sub mnuCloseCompany_Click()
  mnuCompInformation.Visible = True
  mnuDataEntry.Visible = False
  mnuDataEdit.Visible = False
  mnuReport.Visible = False
  mnuCloseCompany.Visible = False
  mnuNext4.Visible = False
  
  frmMDIMain.Caption = Mid(frmMDIMain.Caption, 1, InStr(1, frmMDIMain.Caption, "-") - 1)
End Sub
Private Sub mnuCreateUser_Click()
'  frmCreateUser.Show 1
End Sub
Private Sub mnuDeleteComp_Click()
'  frmDeleteCompany.Show 1
End Sub

Private Sub mnuEditIndentFund_Click()
    frmEditIndentFund.Show 1, Me
End Sub

Private Sub mnuEditIndentInfo_Click()
  frmEditIndentEntry.Show 1
End Sub

Private Sub MnuInvoice_Click()
frmVatInvoice.Show 1
End Sub

Private Sub mnuModifyComp_Click()
  frmModifyCompany.Show 1
End Sub

Private Sub mnuNewAccountsHead_Click()
    frmAccountsHead.Show 1, Me
End Sub

Private Sub mnuNewBuyer_Click()
    frmBuyer.Show 1, Me
End Sub

Private Sub mnuNewComp_Click()
  frmCreateCompany.Show 1
End Sub

Private Sub mnuNewExpenses_Click()
    frmExpense.Show 1, Me
End Sub

Private Sub mnuNewIndentFund_Click()
    frmNewIndentFund.Show 1, Me
End Sub

Private Sub mnuNewIndentInfo_Click()
  frmNewIndentEntry.Show 1
End Sub

Private Sub mnuNewParty_Click()
    frmParty.Show 1, Me
End Sub

Private Sub mnuOpenComp_Click()
  frmOpenCompany.Show 1
End Sub
Private Sub mnuQuit_Click()
  If MsgBox("Are you sure to Exit", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  End
End Sub
Private Sub mnuRptAdjustSt_Click()
  frmRptAdjustmentSt.Show 1
End Sub
Private Sub mnuRptClearBill_Click()
  frmRptClearBillRegister.Show 1
End Sub
Private Sub mnuRptForwardBill_Click()
  frmRptFordBillRegister.Show 1
End Sub
Private Sub mnuRptIndentSt_Click()
  frmRptIndentStatement.Show 1
End Sub

Private Sub mnustatementofcandfbills_Click()
    frmRptStatementOfCandFBills.Show 1
End Sub

Private Sub mnustatementofforwardingdate_Click()
    frmRptStatementForwarderDate.Show 1
End Sub

Private Sub mnuSupplier_Click()
frmSupplier.Show 1, Me
End Sub

Private Sub mnuUserMenu_Click()
frmPermissionMenu.Show 1
End Sub
