VERSION 5.00
Begin VB.Form frmNewSupplier 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " New Supplier"
   ClientHeight    =   2775
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   3825
   Icon            =   "frmSupplier.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   3585
      Begin VB.TextBox txtSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
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
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   3285
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Buyer Name :"
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
         Left            =   150
         TabIndex        =   4
         Top             =   210
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Cancel"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C000&
      Caption         =   "&OK"
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
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Create New Company"
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   8895
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Supplier"
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
      TabIndex        =   6
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmNewSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdOK_Click()
On Error GoTo x
    Dim NewID As Integer
    
    If Len(txtSupplier) = 0 Then
        If strMark = "BUYER" Then
            MsgBox "Enter name of Supplier", vbInformation, AppTitle
        
        End If
        txtSupplier.SetFocus
    Exit Sub
    End If

        If fncDuplicate = True Then Exit Sub
        MainConn.BeginTrans
            If flagMark = "INSERT" Then
                NewID = fncNewID("BUYER")
                MainComm.CommandText = "INSERT INTO New_Supplier( SupplierID, SupplierName) VALUES(" & NewID & ", '" & txtSupplier.Text & "');"
                MainComm.Execute
            Else
                MainComm.CommandText = "UPDATE New_Supplier SET SupplierName='" & txtSupplier.Text & "' WHERE SupplierID=" & Tag & ";"
                MainComm.Execute
            End If
        MainConn.CommitTrans
    strValue = txtSupplier.Text
    Unload Me
Exit Sub
x:
    MsgBox Err.Description, vbCritical, cnstMsgErDB
    MainConn.RollbackTrans

''If strMark = "DN" Then
''    With frmBuyer.flexDetails
''        .Rows = .Rows + 1
''        .Row = .Rows - 1
''        .Col = 0:   .Text = txtSupplier.Text
''
''        .Row = 0
''        For intI = 1 To .Rows
''            If intI Mod 2 = 0 Then
''                .CellBackColor = &HF0F1F2
''            Else
''                .CellBackColor = &H80000005
''            End If
''
''            If .Row < (.Rows - 1) Then .Row = .Row + 1
''        Next intI
''
''        If .Rows = 1 And boUserEdit = True Then
''            frmBuyer.cmdEdit.Enabled = True
''            frmBuyer.cmdDelete.Enabled = True
''        End If
''    End With
''    frmBuyer.txtTotal.Text = CLng(frmBuyer.txtTotal.Text) + 1
''    txtSupplier.SetFocus
''
''ElseIf strMark = "DE" Then
''    With frmBuyer.flexDetails
''        .Col = 0:    .Text = txtSupplier.Text
''    End With
''    Unload Me
''ElseIf strMark = "DISTPRE" Then
''    frmNewEmpInfo.cboDistPre.AddItem txtSupplier.Text
''    frmNewEmpInfo.cboDistPer.AddItem txtSupplier.Text
''    frmNewEmpInfo.cboDistPre.Text = txtSupplier.Text
''
''    Unload Me
''
''    frmNewEmpInfo.cboDistPre.SetFocus
''ElseIf strMark = "DISTPER" Then
''    frmNewEmpInfo.cboDistPer.AddItem txtSupplier.Text
''    frmNewEmpInfo.cboDistPre.AddItem txtSupplier.Text
''    frmNewEmpInfo.cboDistPer.Text = txtSupplier.Text
''
''    Unload Me
''
''    frmNewEmpInfo.cboDistPer.SetFocus
''End If
''txtSupplier.Text = ""
End Sub

Private Sub Form_Load()
    If strMark = "BUYER" Then
        Me.Caption = "New Buyer"
        Label1.Caption = "Buyer"
    Else
        Me.Caption = "New City"
        Label1.Caption = "City"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNewBuyer = Nothing
End Sub
Private Sub txtSupplier_GotFocus()
    Call txtGotFocus(txtSupplier)
End Sub
Private Sub txtSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub txtSupplier_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtSupplier_LostFocus()
    txtSupplier.Text = UCase(Trim(txtSupplier.Text))
End Sub

Private Function fncDuplicate() As Boolean
fncDuplicate = True
    Set R = New ADODB.Recordset
    If flagMark = "INSERT" Or (Trim(txtSupplier.Text) = Trim(txtSupplier.Tag)) Then
        strRecord = "SELECT Count(*) FROM New_Supplier WHERE (SupplierName='" & txtSupplier.Text & "');"
        MainComm.CommandText = strRecord
        Set R = MainComm.Execute
        If R.Fields(0) > 0 Then
            MsgBox "Supplier '" & txtSupplier.Text & "' already exist." & vbCrLf & "Change the Buyer name.", vbInformation, AppTitle
            txtSupplier.SetFocus
        Exit Function
        End If
    End If
    Set R = Nothing
fncDuplicate = False
End Function
