VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3390
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   7425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00B37102&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSplash.frx":0000
   ScaleHeight     =   3390
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      Begin VB.Label lblHUAL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H.U.ACCESSORIES LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   2640
         MouseIcon       =   "frmSplash.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   2040
         Width           =   1950
      End
      Begin VB.Label lblNHTF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NHT FASHIONS LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   480
         MouseIcon       =   "frmSplash.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2040
         Width           =   1710
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C && F ACCOUNT"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   2385
         TabIndex        =   7
         Top             =   120
         Width           =   2685
      End
      Begin VB.Label lblALL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALL COMPANY"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   5670
         MouseIcon       =   "frmSplash.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1635
         Width           =   1245
      End
      Begin VB.Label lblNZN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NZN GARMENTS LTD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   3630
         MouseIcon       =   "frmSplash.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1635
         Width           =   1770
      End
      Begin VB.Label lblPAL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PACIFIC ACCESSORIES LIMITED"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   510
         MouseIcon       =   "frmSplash.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1635
         Width           =   2640
      End
      Begin VB.Label lblUJL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIVERSAL JEANS LIMITED"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   480
         MouseIcon       =   "frmSplash.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1275
         Width           =   2295
      End
      Begin VB.Label lblJ2L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JEANS 2000 LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   5670
         MouseIcon       =   "frmSplash.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1275
         Width           =   1350
      End
      Begin VB.Label lblPJL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PACIFIC JEANS LIMITED"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   3270
         MouseIcon       =   "frmSplash.frx":1850
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1275
         Width           =   1995
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         X1              =   0
         X2              =   8520
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         X1              =   -870
         X2              =   7650
         Y1              =   675
         Y2              =   675
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Unload Me
End Sub
Private Sub Form_Load()
Dim bytyear As Integer
Call ConnectDatabase
Set R1 = New ADODB.Recordset
strRecord = "Select ComId,UJL,J2L,PAL,NHTF,HUAL from sys_Configure_Status Where IDNo='" & getComVolID & "' ;"
R1.Open strRecord, MainConn, adOpenStatic
If R1.RecordCount > 0 Then
Dim logID As Integer
For i = 1 To 6
    If i = 1 Then
    logID = R1!ComID
    ElseIf i = 2 Then
    logID = R1!UJL
    ElseIf i = 3 Then
    logID = R1!J2L
    ElseIf i = 4 Then
    logID = R1!PAL
    ElseIf i = 5 Then
    logID = R1!NHTF
    ElseIf i = 6 Then
    logID = R1!HUAL
    End If
    
    
    If logID = 1 Then
        lblPJL.Enabled = True

    ElseIf logID = 3 Then
        lblPAL.Enabled = True

    ElseIf logID = 4 Then
        lblNZN.Enabled = True

    ElseIf logID = 5 Then
        lblJ2L.Enabled = True

    ElseIf logID = 7 Then
        lblUJL.Enabled = True

    ElseIf logID = 8 Then
        lblALL.Enabled = True
        
    ElseIf logID = 13 Then
        lblNHTF.Enabled = True
        
    ElseIf logID = 20 Then
        lblHUAL.Enabled = True
        
End If
Next i
Else
MsgBox "Please Contact Software Vendor..", vbCritical, "ERROR"
Exit Sub
End If
R1.Close
Set R1 = Nothing

'txtyear.Text = Year(Date)


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSplash = Nothing
End Sub

Private Sub Frame1_DblClick()
    Unload Me
End Sub

Private Sub lblALL_Click()
    Frame1.Tag = lblALL.Caption
    Call PrcCompanyLogIn
    frmUserLogin.Show
    Unload frmSplash

End Sub

Private Sub lblHUAL_Click()
Frame1.Tag = lblHUAL.Caption
    strComNameTxt = Frame1.Tag
    Call PrcCompanyLogIn
    frmUserLogin.Show
    Unload frmSplash
End Sub

Private Sub lblJ2L_Click()
    Frame1.Tag = lblJ2L.Caption
    strComNameTxt = Frame1.Tag
    Call PrcCompanyLogIn
    frmUserLogin.Show
    Unload frmSplash
End Sub

Private Sub lblNHTF_Click()
Frame1.Tag = lblNHTF.Caption
    strComNameTxt = Frame1.Tag
    Call PrcCompanyLogIn
    frmUserLogin.Show
    Unload frmSplash
End Sub

Private Sub lblNZN_Click()
    Frame1.Tag = lblNZN.Caption
    Call PrcCompanyLogIn
    frmUserLogin.Show
    Unload frmSplash
End Sub

Private Sub lblPAL_Click()
    Frame1.Tag = lblPAL.Caption
    strComNameTxt = Frame1.Tag
    Call PrcCompanyLogIn
    frmUserLogin.Show
    Unload frmSplash
End Sub

Private Sub lblPJL_Click()
    Frame1.Tag = lblPJL.Caption
    strComNameTxt = Frame1.Tag
    Call PrcCompanyLogIn
    Load frmUserLogin
    frmUserLogin.Refresh
    frmUserLogin.Show
    Unload frmSplash
    
End Sub

Private Sub lblUJL_Click()
    Frame1.Tag = lblUJL.Caption
    strComNameTxt = Frame1.Tag
    Call PrcCompanyLogIn
    frmUserLogin.Show
    Unload frmSplash
End Sub
Public Function PrcCompanyLogIn()
On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set R = New ADODB.Recordset
strRecord = "SELECT * FROM Company_Information WHERE CompName='" & Frame1.Tag & "';"
R.Open strRecord, MainConn, adOpenStatic
        WSName = R("ServerName").Value
        WDBName = R("DBName").Value
        WDBPrefix = R("DBPrefix").Value
        ReportPath = R("ReportPath").Value
        bytCompID = R("CompID").Value

    isPrimaryDB = False
    Call ConnectDatabase
'    frmUserLogin.Show
'    Unload frmSplash
Exit Function
ErrorHandler:
    MsgBox Err.Description, vbCritical, cnstMsgErDB
    Screen.MousePointer = vbDefault
    MainConn.RollbackTrans
R.Close
Set R = Nothing
End Function
