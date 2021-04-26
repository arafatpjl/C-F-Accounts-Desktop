VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSupplier 
   Caption         =   "New SUpplier"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "New SUpplier"
   ScaleHeight     =   5145
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
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
      Left            =   3720
      TabIndex        =   7
      Top             =   4680
      Width           =   1065
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   4680
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   3645
      Left            =   120
      TabIndex        =   2
      Top             =   750
      Width           =   4695
      Begin MSFlexGridLib.MSFlexGrid flexDetails 
         Height          =   2655
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4683
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483626
         BackColorFixed  =   -2147483632
         BackColorSel    =   8388608
         BackColorBkg    =   15132390
         AllowBigSelection=   0   'False
         FocusRect       =   0
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00789C20&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   3270
         Width           =   675
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00789C20&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   4395
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00789C20&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entries  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   3240
         Width           =   4395
      End
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   1
      Top             =   4680
      Width           =   1065
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1065
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "List of the Supplier"
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
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EAF5F5&
      BorderStyle     =   1  'Fixed Single
      Height          =   3555
      Left            =   4830
      TabIndex        =   11
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "List of the Buyer"
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
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   4440
      Width           =   6135
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  strMark = "Supplier"
  flagMark = "INSERT"
  strValue = ""
  
  frmNewSupplier.Show 1
  
  Call prcLoadData
End Sub
Private Sub cmdDelete_Click()
On Error GoTo x
    If MsgBox("Are you sure to delete" & vbCrLf & " Buyer '" & flexDetails.Text & "' ?", vbQuestion + vbYesNo, AppTitle) = vbNo Then Exit Sub
        MainConn.BeginTrans
            MainComm.CommandText = "DELETE  FROM New_Supplier WHERE SupplierName='" & flexDetails.Text & "';"
            MainComm.Execute
        MainConn.CommitTrans
    Call prcLoadData
    Call prcGridFormat(Me, flexDetails)
Exit Sub
x:
MsgBox Err.Description, vbCritical, cnstMsgErDB
MainConn.RollbackTrans
Exit Sub
End Sub
Private Sub cmdEdit_Click()
    strMark = "Supplier"
    flagMark = "EDIT"
    
    With frmNewSupplier
        flexDetails.Col = 0:    .Tag = flexDetails.Text
        flexDetails.Col = 1:
            .txtSupplier.Text = flexDetails.Text
            .txtSupplier.Tag = flexDetails.Text
        .Caption = "Edit Buyer"
        .Show 1
    End With
    Call prcLoadData
    Call prcGridFormat(Me, flexDetails)
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo x
    Call prcMakeCenter(Me)
    
    Call prcGridInitialize
    Call prcLoadData
Exit Sub
x:
MsgBox Err.Description, vbCritical, AppTitle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBuyer = Nothing
End Sub

Private Sub prcLoadData()
On Error GoTo x
    Set R = New ADODB.Recordset
    
    Dim inti As Integer
    strRecord = "SELECT SupplierID, SupplierName FROM New_Supplier ORDER BY SupplierName;"
    'MainComm.CommandText = strRecord
    'Set R = MainComm.Execute
      R.Open strRecord, MainConn, adOpenStatic
    inti = 0
    flexDetails.Rows = 0
    If R.EOF = False And R.BOF = False Then
        With flexDetails
            Do While Not R.EOF
                inti = inti + 1
                
                .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .Col = 0:  .Text = R(0).Value
                .Col = 1:  .Text = R(1).Value
            R.MoveNext
            Loop
        End With
        txtTotal.Text = inti
    End If
    Call prcGridFormat(Me, flexDetails)
    Set R = Nothing
Exit Sub
x:
MsgBox Err.Description, vbCritical, AppTitle
End Sub

Private Sub prcGridInitialize()
    With flexDetails
        .Row = 0
        .Cols = 2
        
        .ColWidth(0) = 0: .ColWidth(1) = 4000
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
    End With
End Sub


