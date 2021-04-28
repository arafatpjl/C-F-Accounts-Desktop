VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmVatInvoice 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invoice Entry"
   ClientHeight    =   6945
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   6945
   ControlBox      =   0   'False
   DrawMode        =   4  'Mask Not Pen
   Icon            =   "frmVatInvoice.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDate 
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
      ItemData        =   "frmVatInvoice.frx":0442
      Left            =   2400
      List            =   "frmVatInvoice.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox cboInvoiceNo 
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
      ItemData        =   "frmVatInvoice.frx":0446
      Left            =   2400
      List            =   "frmVatInvoice.frx":0448
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3120
      TabIndex        =   47
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   46
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtMode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   240
      Width           =   825
   End
   Begin VB.TextBox TxtEntryDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   2400
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox VatInvAmt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4440
      TabIndex        =   12
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox VatAmtBdt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4440
      TabIndex        =   42
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   3960
      TabIndex        =   41
      Text            =   "%"
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Index           =   0
      Left            =   3960
      TabIndex        =   40
      Text            =   "%"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox TxtAIT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox TxtTtlExp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox TxtPayble 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2400
      TabIndex        =   13
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox VatBdt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox AmtBdt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2385
      TabIndex        =   8
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Amtusd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox TxtExc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   3180
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
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
      Left            =   1800
      TabIndex        =   26
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete"
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
      Left            =   4440
      TabIndex        =   25
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "&Save"
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
      Left            =   5760
      TabIndex        =   14
      ToolTipText     =   "Update Record"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtInvDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox cboSuppName 
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
      ItemData        =   "frmVatInvoice.frx":044A
      Left            =   2400
      List            =   "frmVatInvoice.frx":044C
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "cboSuppName"
      Top             =   915
      Width           =   3255
   End
   Begin VB.TextBox txtInvNo 
      Appearance      =   0  'Flat
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
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   44
      Top             =   2280
      Width           =   225
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Entry Date"
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
      Left            =   360
      TabIndex        =   43
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "AIT(%On Inv Amt)"
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
      Left            =   360
      TabIndex        =   39
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   38
      Top             =   5040
      Width           =   225
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   37
      Top             =   4560
      Width           =   225
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Exp."
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
      Left            =   360
      TabIndex        =   36
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   35
      Top             =   5520
      Width           =   225
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Paybale Amt"
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
      Left            =   360
      TabIndex        =   34
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Vat(BDT)"
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
      Left            =   360
      TabIndex        =   33
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount(BDT)"
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
      Left            =   360
      TabIndex        =   32
      Top             =   3630
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   31
      Top             =   3600
      Width           =   225
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount(USD)"
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
      Left            =   360
      TabIndex        =   30
      Top             =   2790
      Width           =   1575
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   29
      Top             =   2760
      Width           =   225
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   28
      Top             =   3135
      Width           =   225
   End
   Begin VB.Label TxtExchange 
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange Rate"
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
      Left            =   360
      TabIndex        =   27
      Top             =   3195
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   24
      Top             =   6120
      Width           =   7455
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Entry"
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
      TabIndex        =   23
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   21
      Top             =   4080
      Width           =   225
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice  Date"
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
      Left            =   360
      TabIndex        =   20
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   19
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
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
      Left            =   360
      TabIndex        =   18
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   960
      Width           =   225
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   1320
      Width           =   225
   End
End
Attribute VB_Name = "frmVatInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Amtusd_KeyPress(KeyAscii As Integer)
 CurrFunc KeyAscii
End Sub

Private Sub cmbDate_Click()
cboInvoiceNo.Clear
 Set R1 = New ADODB.Recordset
        strRecord = "SELECT InvNo FROM new_Invoice_Information where compID=" & bytCompID & " AND EntryDate= '" & cmbDate.Text & "' and SuppID=" & findSuppID(cboSuppName.Text) & " order By InvNo;"
        R1.Open strRecord, MainConn, adOpenStatic
        cboInvoiceNo.Clear
        If R1.RecordCount > 0 Then
            Do Until R1.EOF
                cboInvoiceNo.AddItem R1![InvNo]
'                If param = 1 Then cboMain.ItemData(cboMain.NewIndex) = R![SupplierID]
            R1.MoveNext
            Loop
        End If
    Set R1 = Nothing
End Sub

Private Sub Form_Load()
    Call addsuppName(cboSuppName)
    cmdDelete.Enabled = False
    cmdPreview.Enabled = False
    cboInvoiceNo.Visible = False
    cmbDate.Visible = False
     txtMode.Text = "NEW"
 
End Sub

Private Sub cmdSave_Click()

    
    If Len(cboSuppName) = 0 Then
        MsgBox "Select Supplier Name", vbInformation, cnstMsgInfo
        cboSuppName.SetFocus
    Exit Sub
    End If

    If Len(txtInvDate) = 0 Then
        MsgBox "Enter Invoice Date", vbInformation, cnstMsgInfo
        txtInvDate.SetFocus
        Exit Sub
    End If
    
    If Len(TxtEntryDate) = 0 Then
        MsgBox "Enter Entry Date", vbInformation, cnstMsgInfo
        TxtEntryDate.SetFocus
    Exit Sub
    End If

    If Len(AmtBdt) = 0 Then
        MsgBox "Enter Invoice Amount", vbInformation, cnstMsgInfo
        AmtBdt.SetFocus
    Exit Sub
    End If
    
    If Len(VatAmtBdt) = 0 Then
        MsgBox "Enter Vat Amount", vbInformation, cnstMsgInfo
        VatAmtBdt.SetFocus
    Exit Sub
    End If
    
   If Len(TxtTtlExp) = 0 Then
        MsgBox "Enter Total Expense", vbInformation, cnstMsgInfo
        TxtTtlExp.SetFocus
    Exit Sub
    End If
    If Len(VatInvAmt) = 0 Then
        MsgBox "Enter Total Expense", vbInformation, cnstMsgInfo
        VatInvAmt.SetFocus
    Exit Sub
    End If
    If Len(TxtPayble) = 0 Then
        MsgBox "Enter Payble Amount", vbInformation, cnstMsgInfo
        TxtPayble.SetFocus
    Exit Sub
    End If

    
   If cmdSave.Caption = "&Save" Then
   If Len(txtInvNo) = 0 Then
        MsgBox "Enter Invoice No", vbInformation, cnstMsgInfo
        txtInvNo.SetFocus
    Exit Sub
    End If
   
   Set R1 = New ADODB.Recordset

        strRecord = "SELECT InvNo FROM  new_Invoice_Information WHERE ((InvNo='" & txtInvNo.Text & "') AND (SuppID=" & findSuppID(cboSuppName.Text) & ") AND (compID=" & bytCompID & "));"

    R1.Open strRecord, MainConn, adOpenStatic
    If R1.RecordCount > 0 Then
        MsgBox "Duplicate Invoice No", vbInformation, cnstMsgInfo
        txtInvNo.SetFocus
    Exit Sub
    End If
    If MsgBox("Are you sure to Save ", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
   Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainConn.BeginTrans
    

        MainComm.CommandText = "Insert into dbo. new_Invoice_Information(compID, InvNo, SuppID, InvDate, AmtUSD, AmtBDT, InvAmount, InvVat, InvVatAMT, TotalEXP, AitVat, AitVatBDT, AmtPayble, EntryDate, EntryTime , CompName, Userid)" _
                                & " VALUES(" & bytCompID & ", '" & txtInvNo.Text & "', " & findSuppID(cboSuppName.Text) & ", '" & txtInvDate.Text & "',  " & CCur(Amtusd.Text) & ", " & CCur(TxtExc.Text) & ", " & CCur(AmtBdt.Text) & ", " & CCur(VatBdt.Text) & ", " & CCur(VatAmtBdt.Text) & ", " & CCur(TxtTtlExp.Text) & "," _
                                & " " & CCur(TxtAIT.Text) & ", " & CCur(VatInvAmt.Text) & ", " & CCur(TxtPayble.Text) & ", '" & TxtEntryDate.Text & "', '" & Time & "', '" & getComName & "', '" & struser & "' );"
        MainComm.Execute
          
        MainComm.CommandText = "Insert into dbo. new_Invoice_Information_bak(compID, InvNo, SuppID, InvDate, AmtUSD, AmtBDT, InvAmount, InvVat, InvVatAMT, TotalEXP, AitVat, AitVatBDT, AmtPayble, EntryDate, EntryTime , CompName, Userid)" _
                                & " VALUES(" & bytCompID & ", '" & txtInvNo.Text & "', " & findSuppID(cboSuppName.Text) & ", '" & txtInvDate.Text & "',  " & CCur(Amtusd.Text) & ", " & CCur(TxtExc.Text) & ", " & CCur(AmtBdt.Text) & ", " & CCur(VatBdt.Text) & ", " & CCur(VatAmtBdt.Text) & ", " & CCur(TxtTtlExp.Text) & "," _
                                & " " & CCur(TxtAIT.Text) & ", " & CCur(VatInvAmt.Text) & ", " & CCur(TxtPayble.Text) & ", '" & TxtEntryDate.Text & "', '" & Time & "', '" & getComName & "', '" & struser & "' );"
       '                   MainComm.Execute
    MainConn.CommitTrans
    MsgBox "Invoice Information Successfully Saved.", vbInformation, App.Title
    Else
''------------------- UPDATE STATEMENT -----------------------
        If MsgBox("Do you want to update Invoice information.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
          
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
        MainConn.BeginTrans
        
        MainComm.CommandText = "UPDATE new_Invoice_Information SET InvDate='" & Format(txtInvDate.Text, cnstDtFrmtI) & "'," _
            & " AmtUSD=" & CCur(Amtusd.Text) & ",AmtBDT=" & CCur(TxtExc.Text) & ",InvAmount=" & CCur(AmtBdt.Text) & ",InvVat=" & CCur(VatBdt.Text) & "," _
            & " InvVatAMT=" & CCur(VatAmtBdt.Text) & ",TotalEXP=" & CCur(TxtTtlExp.Text) & ",AitVat=" & CCur(TxtAIT.Text) & ",AitVatBDT=" & CCur(VatInvAmt.Text) & ",AmtPayble=" & CCur(TxtPayble.Text) & " WHERE ((compID=" & bytCompID & ")" _
            & " AND (InvNo='" & cboInvoiceNo.Text & "') AND (SuppID=" & findSuppID(cboSuppName.Text) & "));"
        MainComm.Execute
        MainConn.CommitTrans
        
        MsgBox "Invoice Information Successfully Updated.", vbInformation, App.Title
        
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

  
txtclear
    
'    txtInvNo.SetFocus
    Screen.MousePointer = vbDefault
End Sub
Private Sub cboInvoiceNo_Click()
  If cboInvoiceNo.Text = "" Then Exit Sub
    Set R1 = New ADODB.Recordset
    
'    If optImport.Value = True Then
'        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.OtherPayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.AdjDate," _
'              & " new_Indent_Information.TotBAmount,new_Indent_Information.OtherAdjAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat,new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
'              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & cboJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='I'));"
'    Else
        strRecord = "select new_Invoice_Information.InvDate, new_Invoice_Information.AmtUSD, new_Invoice_Information.AmtBDT, new_Invoice_Information.InvAmount, " _
                     & " new_Invoice_Information.InvVat , new_Invoice_Information.InvVatAMT, new_Invoice_Information.TotalEXP, new_Invoice_Information.AitVat, new_Invoice_Information.AitVatBDT, new_Invoice_Information.AmtPayble, new_Invoice_Information.EntryDate" _
                     & "  from   new_Invoice_Information WHERE (((new_Invoice_Information.InvNo)='" & cboInvoiceNo.Text & "') AND ((new_Invoice_Information.compID)=" & bytCompID & ") AND ((new_Invoice_Information.SuppID)=" & findSuppID(cboSuppName.Text) & "));"
'    End If
    R1.Open strRecord, MainConn, adOpenStatic
    
    If R1.RecordCount > 0 Then
    
'        txtJobNo.Text = R![JobNo]
'        txtInvNo.Text = R![InvNo]
        txtInvDate.Text = Format(R1![InvDate], cnstDtFrmtI)
        TxtEntryDate.Text = Format(R1![EntryDate], cnstDtFrmtI)
'        OtherAdjDate.Text = Format(R![OtherPayDate], cnstDtFrmtI)
        Amtusd.Text = Format(R1![Amtusd])
        TxtExc.Text = Format(R1![AmtBdt])
        AmtBdt.Text = Format(R1![InvAmount])
        VatBdt.Text = Format(R1![InvVat])
        VatAmtBdt.Text = Format(R1![InvVatAMT])
        TxtTtlExp.Text = Format(R1![TotalEXP])
        TxtAIT.Text = Format(R1![AitVat])
        VatInvAmt.Text = Format(R1![AitVatBDT])
        TxtPayble.Text = Format(R1![AmtPayble])
        
        
    End If
    R1.Close
    Set R1 = Nothing
End Sub

Private Sub cboSuppName_Click()
cmbDate.Clear
 Set R1 = New ADODB.Recordset
        strRecord = "SELECT DISTINCT EntryDate FROM new_Invoice_Information where compID=" & bytCompID & " AND SuppID=" & findSuppID(cboSuppName.Text) & " order By EntryDate;"
        R1.Open strRecord, MainConn, adOpenStatic
        cmbDate.Clear
        If R1.RecordCount > 0 Then
            Do Until R1.EOF
                cmbDate.AddItem R1![EntryDate]
'                If param = 1 Then cboMain.ItemData(cboMain.NewIndex) = R![SupplierID]
            R1.MoveNext
            Loop
        End If
    Set R1 = Nothing
End Sub
Private Sub AmtBdt_GotFocus()
Call txtGotFocus(AmtBdt)
End Sub

Private Sub AmtBdt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Amtusd_GotFocus()
 Call txtGotFocus(Amtusd)
End Sub

Private Sub Amtusd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Amtusd_LostFocus()
AmtBdt.Text = Format(CCur(Amtusd.Text)) * Format(CCur(TxtExc.Text))
'Call txtGotFocus(Amtusd.Text)
End Sub

Sub txtGotFocus(ByRef txtMain As TextBox)
    With txtMain
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub btnEdit_Click()
If txtMode.Text = "NEW" Then
        cmdSave.Caption = "&Update"
        cmdSave.Enabled = True
        cmdPreview.Enabled = True
        cmdDelete.Enabled = True
        btnEdit.Caption = "&New"
        txtMode.Text = "EDIT"
        cboInvoiceNo.Visible = True
        cboSuppName.Text = ""
        cboInvoiceNo.Clear
        cmbDate.Visible = True
        cmbDate.Clear
        TxtEntryDate.Enabled = False
'        txtVoucherNo.Clear
'        txtForwardDate.Text = ""
'        Call prcLoadData
'        flexDetails.Clear


    Else
        cmbDate.Visible = False
        cboInvoiceNo.Visible = False
        cmdSave.Caption = "&Save"
        cmdSave.Enabled = True
        cmdDelete.Enabled = False
        cmdPreview.Enabled = False
        btnEdit.Caption = "&Edit"
        txtMode.Text = "NEW"
        txtclear
        txtInvDate.Text = ""
        TxtEntryDate.Text = ""
        cboSuppName.Text = ""
        TxtEntryDate.Enabled = True
        
    End If
End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdPreview_Click()
   If Len(cboSuppName.Text) = 0 Then
        MsgBox "You should provide Supplier name.", vbInformation, App.Title
        cboSuppName.SetFocus
        Exit Sub
    End If
    If Len(cmbDate.Text) = 0 Then
        MsgBox "You should provide Invoice dATE", vbInformation, App.Title
        cmbDate.SetFocus
        Exit Sub
    End If
       Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "DELETE FROM Temp_Invoice_Information;"
        MainComm.Execute
    MainComm.CommandText = "Insert into dbo.Temp_Invoice_Information(compName, InvNo, SuppName, InvDate,AmtUSD, AmtBDT, InvAmount, InvVatAMT, TotalEXP,  AitVatBDT, AmtPayble, EntryDate)" _
                                & " SELECT  Company_Information.compName, new_Invoice_Information.InvNo,New_Supplier.SupplierName ,new_Invoice_Information.InvDate,new_Invoice_Information.AmtUSD, " _
                                & " new_Invoice_Information.AmtBDT, new_Invoice_Information.InvAmount,new_Invoice_Information.InvVatAMT , new_Invoice_Information.TotalEXP, new_Invoice_Information.AitVatBDT, " _
                                & " new_Invoice_Information.AmtPayble, new_Invoice_Information.EntryDate FROM  Company_Information INNER JOIN dbo.new_Invoice_Information on Company_Information.compID = new_Invoice_Information.compID " _
                                & "  inner join New_Supplier on new_Invoice_Information.SuppID=New_Supplier.SupplierID WHERE (new_Invoice_Information.compID=" & bytCompID & ")AND (new_Invoice_Information.SuppID=" & findSuppID(cboSuppName.Text) & ") " _
                                & " AND (new_Invoice_Information.EntryDate='" & (cmbDate.Text) & "');"
                                
     MainComm.Execute
     MainConn.CommitTrans
 
   
End Sub


Private Sub cmdDelete_Click()
    If Len(cboInvoiceNo) = 0 Then
        MsgBox "Select Invoice No", vbInformation, cnstMsgInfo
        cboInvoiceNo.SetFocus
    Exit Sub
    End If

    If MsgBox("Are you sure to Delete", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
                    
        MainComm.CommandText = "Insert into dbo.new_Invoice_Information_BAK(compID, InvNo, SuppID, InvDate, AmtUSD, AmtBDT, InvAmount, InvVat, InvVatAMT, TotalEXP, AitVat, AitVatBDT, AmtPayble, EntryDate, EntryTime , CompName, Userid)" _
                                & " SELECT  compID, InvNo, SuppID, InvDate, AmtUSD, AmtBDT, InvAmount, InvVat, InvVatAMT, TotalEXP, AitVat, AitVatBDT, AmtPayble, EntryDate,'" & Time & "','" & getComName & "','" & struser & "' FROM dbo.new_Invoice_Information WHERE (compID=" & bytCompID & ")" _
                                & " AND (InvNo='" & cboInvoiceNo.Text & "') AND (SuppID=" & findSuppID(cboSuppName.Text) & ");"
        MainComm.Execute
        
        MainComm.CommandText = "DELETE FROM new_Invoice_Information WHERE ((compID=" & bytCompID & ") AND (InvNo='" & cboInvoiceNo.Text & "')) AND (SuppID=" & findSuppID(cboSuppName.Text) & ");"
        MainComm.Execute
        
        
    MainConn.CommitTrans
     MsgBox "Invoice Delete Successfully .", vbInformation, App.Title

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

cboInvoiceNo.Clear
cmbDate.Clear
TxtEntryDate.Text = ""
txtclear
cboInvoiceNo.SetFocus
Screen.MousePointer = vbDefault

End Sub

Private Sub txtclear()


  txtInvNo.Text = "":         Amtusd.Text = "0.00"
    TxtExc.Text = "0.00":           AmtBdt.Text = ""
    VatBdt.Text = "0.00":           VatAmtBdt.Text = ""
    TxtTtlExp.Text = "":    TxtAIT.Text = "0.00"
    VatInvAmt.Text = "":       TxtPayble.Text = ""
    txtInvDate.Text = "":
    End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEditIndentEntry = Nothing
End Sub


Private Sub TxtAIT_GotFocus()
Call txtGotFocus(TxtAIT)
End Sub

Private Sub TxtAIT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtAIT_KeyPress(KeyAscii As Integer)
 CurrFunc KeyAscii
End Sub

Private Sub TxtAIT_LostFocus()
VatInvAmt = ((Format(CCur(TxtAIT.Text)) * Format(CCur(AmtBdt.Text))) / 100)
TxtPayble = CCur(Val(TxtTtlExp.Text)) - (CCur(Val(VatAmtBdt.Text)) + CCur(Val(VatInvAmt.Text)))
'TxtTtlExp.Text = CCur(Val(VatAmtBdt.Text)) + CCur(Val(AmtBdt.Text))
End Sub



Private Sub TxtEntryDate_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub TxtEntryDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtEntryDate_KeyPress(KeyAscii As Integer)
DateFunc KeyAscii
End Sub

Private Sub TxtEntryDate_LostFocus()
 If Len(TxtEntryDate) = 0 Then Exit Sub
    TxtEntryDate.Text = FormatDate(TxtEntryDate)
End Sub

Private Sub TxtExc_GotFocus()
Call txtGotFocus(TxtExc)
End Sub

Private Sub TxtExc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtExc_KeyPress(KeyAscii As Integer)
 CurrFunc KeyAscii
End Sub

Private Sub TxtExc_LostFocus()
AmtBdt.Text = Format(CCur(Amtusd.Text)) * Format(CCur(TxtExc.Text))
End Sub



Private Sub txtInvDate_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtInvDate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtInvDate_KeyPress(KeyAscii As Integer)
 DateFunc KeyAscii
End Sub

Private Sub txtInvDate_LostFocus()
   If Len(txtInvDate) = 0 Then Exit Sub
    txtInvDate.Text = FormatDate(txtInvDate)
End Sub

Private Sub txtInvNo_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

'Private Sub txtInvNo_GotFocus()
'    SendKeys "{Home}+{End}"
'End Sub
Private Sub txtInvNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtInvNo_LostFocus()
    If Len(txtInvNo) = 0 Then Exit Sub
    txtInvNo.Text = Trim(UCase(txtInvNo.Text))
End Sub








Private Sub TxtPayble_GotFocus()
Call txtGotFocus(TxtPayble)
End Sub

Private Sub TxtTtlExp_GotFocus()
Call txtGotFocus(TxtTtlExp)
End Sub

Private Sub TxtTtlExp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub VatAmtBdt_GotFocus()
Call txtGotFocus(VatAmtBdt)
End Sub

Private Sub VatAmtBdt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub VatBdt_GotFocus()
Call txtGotFocus(VatBdt)
End Sub

Private Sub VatBdt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub VatBdt_KeyPress(KeyAscii As Integer)
 CurrFunc KeyAscii
End Sub

Private Sub VatBdt_LostFocus()
VatAmtBdt.Text = (Format(CCur(AmtBdt.Text)) * Format(CCur(VatBdt.Text))) / 100
'TxtTtlExp.Text = (Format(CCur(VatAmtBdt.Text))) + (Format(CCur(AmtBdt.Text)))
TxtTtlExp.Text = CCur(Val(VatAmtBdt.Text)) + CCur(Val(AmtBdt.Text))
   
End Sub

Private Sub VatInvAmt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
