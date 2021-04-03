VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEditIndentEntry 
   BackColor       =   &H00EAF5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Indent Information"
   ClientHeight    =   10440
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   8895
   ControlBox      =   0   'False
   DrawMode        =   4  'Mask Not Pen
   Icon            =   "frmEditIndentEntry.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Others"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   210
      TabIndex        =   71
      Top             =   4890
      Width           =   4230
      Begin VB.TextBox OtherTotalAdjAmount 
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
         Left            =   2160
         TabIndex        =   76
         Top             =   690
         Width           =   1935
      End
      Begin VB.TextBox OtherAdjDate 
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
         Left            =   2160
         TabIndex        =   4
         Top             =   255
         Width           =   1935
      End
      Begin VB.TextBox txtOtherAdjAmount 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Text            =   "0.00"
         ToolTipText     =   "Type Bank Adjust Payment"
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust. Amount"
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
         Left            =   120
         TabIndex        =   78
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label Label28 
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
         Left            =   1890
         TabIndex        =   77
         Top             =   1170
         Width           =   225
      End
      Begin VB.Label Label38 
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
         Left            =   1890
         TabIndex        =   75
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label37 
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
         Left            =   1920
         TabIndex        =   74
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
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
         Left            =   120
         TabIndex        =   73
         Top             =   255
         Width           =   1380
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Other.Amount"
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
         Left            =   90
         TabIndex        =   72
         Top             =   750
         Width           =   1785
      End
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
      Left            =   2520
      TabIndex        =   70
      ToolTipText     =   "Close Window"
      Top             =   9900
      Width           =   1455
   End
   Begin VB.CommandButton CmdAdjUpdate 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit Adjust"
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
      Left            =   4080
      TabIndex        =   69
      ToolTipText     =   "Cash/Bank Adjustment Edit"
      Top             =   9900
      Width           =   1455
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
      Left            =   5640
      TabIndex        =   68
      ToolTipText     =   "Delete Record"
      Top             =   9900
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "&Update"
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
      Left            =   7200
      TabIndex        =   7
      ToolTipText     =   "Update Record"
      Top             =   9900
      Width           =   1455
   End
   Begin VB.TextBox txtTotAdjust 
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   62
      ToolTipText     =   "Total Adjust Value"
      Top             =   6450
      Width           =   2085
   End
   Begin VB.TextBox txtAdjDate 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2475
      TabIndex        =   1
      ToolTipText     =   "Type Adjust Date "
      Top             =   2790
      Width           =   1935
   End
   Begin VB.TextBox txtJobNo 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2475
      MaxLength       =   30
      TabIndex        =   51
      ToolTipText     =   "Chage Job No"
      Top             =   2010
      Width           =   2790
   End
   Begin VB.ComboBox cboJobNo1 
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
      ItemData        =   "frmEditIndentEntry.frx":0442
      Left            =   5160
      List            =   "frmEditIndentEntry.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   9
      Text            =   "cboJobNo"
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtInvNo 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2475
      MaxLength       =   30
      TabIndex        =   52
      ToolTipText     =   "Change Invoice No"
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Job"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   180
      TabIndex        =   46
      Top             =   6840
      Width           =   6450
      Begin VB.TextBox txtBalance 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "0.00"
         ToolTipText     =   "Total Balance Value"
         Top             =   570
         Width           =   2025
      End
      Begin VB.TextBox txtIndAmount 
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
         Left            =   2175
         TabIndex        =   15
         Text            =   "0.00"
         ToolTipText     =   "Total Indent Value"
         Top             =   225
         Width           =   2025
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balance Amount"
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
         Left            =   105
         TabIndex        =   60
         Top             =   615
         Width           =   1935
      End
      Begin VB.Label Label33 
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
         Left            =   1965
         TabIndex        =   59
         Top             =   600
         Width           =   225
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
         Left            =   1965
         TabIndex        =   48
         Top             =   255
         Width           =   225
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Indent Amount"
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
         Left            =   105
         TabIndex        =   47
         Top             =   255
         Width           =   1875
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAF5F5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   180
      TabIndex        =   37
      Top             =   7890
      Width           =   6450
      Begin VB.ComboBox cboAccNo 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Choose Nature of Expenses"
         Top             =   960
         Width           =   3585
      End
      Begin VB.TextBox txtRemarks 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Type Remarks For Job No."
         Top             =   1320
         Width           =   3585
      End
      Begin VB.TextBox txtTax 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         ToolTipText     =   "Type Income Tax Value"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtVat 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         ToolTipText     =   "Type Vat Value"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Vat"
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
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Income Tax"
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
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nature of Expenses"
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
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label45 
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
         Left            =   1980
         TabIndex        =   41
         Top             =   1000
         Width           =   225
      End
      Begin VB.Label Label44 
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
         Left            =   1980
         TabIndex        =   40
         Top             =   1380
         Width           =   225
      End
      Begin VB.Label Label43 
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
         Left            =   1980
         TabIndex        =   39
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label42 
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
         Left            =   1980
         TabIndex        =   38
         Top             =   620
         Width           =   225
      End
   End
   Begin VB.Frame BFrame 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   4500
      TabIndex        =   29
      Top             =   3300
      Width           =   4230
      Begin VB.TextBox txtBAdjAmount 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Text            =   "0.00"
         ToolTipText     =   "Type Bank Adjust Payment"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtChequePayDate 
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
         Left            =   2160
         TabIndex        =   13
         Top             =   255
         Width           =   1935
      End
      Begin VB.TextBox txtChequePayValue 
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
         Left            =   2160
         TabIndex        =   14
         Top             =   652
         Width           =   1935
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust. Amount"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
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
         Left            =   120
         TabIndex        =   34
         Top             =   255
         Width           =   1380
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Indent Value"
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
         Left            =   120
         TabIndex        =   33
         Top             =   652
         Width           =   1380
      End
      Begin VB.Label Label36 
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
         Left            =   1920
         TabIndex        =   32
         Top             =   630
         Width           =   225
      End
      Begin VB.Label Label35 
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
         Left            =   1920
         TabIndex        =   31
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label30 
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
         Left            =   1920
         TabIndex        =   30
         Top             =   1020
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   180
      TabIndex        =   21
      Top             =   3300
      Width           =   4230
      Begin VB.TextBox txtAdjAmount 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2145
         TabIndex        =   2
         Text            =   "0.00"
         ToolTipText     =   "Type Cash Adjust Paymet"
         Top             =   1020
         Width           =   1935
      End
      Begin VB.TextBox txtPayDate 
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
         Left            =   2145
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCashPayValue 
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
         Left            =   2145
         TabIndex        =   22
         Top             =   630
         Width           =   1935
      End
      Begin VB.Label lblAdjCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust. Amount"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Label Label17 
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
         Left            =   1920
         TabIndex        =   28
         Top             =   975
         Width           =   225
      End
      Begin VB.Label Label9 
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
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label26 
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
         Left            =   1920
         TabIndex        =   26
         Top             =   607
         Width           =   225
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Indent Value"
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
         Left            =   120
         TabIndex        =   24
         Top             =   637
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAF5F5&
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   7110
      TabIndex        =   18
      Top             =   1485
      Width           =   1545
      Begin VB.OptionButton optEditIm 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   20
         Top             =   390
         Width           =   915
      End
      Begin VB.OptionButton optEditEx 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   19
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   525
      Left            =   2700
      TabIndex        =   8
      Top             =   840
      Width           =   2295
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
         Left            =   1230
         TabIndex        =   17
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
         Left            =   180
         TabIndex        =   16
         Top             =   180
         Width           =   975
      End
   End
   Begin MSDataListLib.DataCombo cboJobNo 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   1590
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   67
      Top             =   9690
      Width           =   8895
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Indent Information"
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
      TabIndex        =   66
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   8895
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
      Left            =   2100
      TabIndex        =   64
      Top             =   6480
      Width           =   225
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Adust  Amount"
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
      Left            =   210
      TabIndex        =   63
      Top             =   6510
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust. Date"
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
      Left            =   255
      TabIndex        =   58
      Top             =   2790
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
      Left            =   2205
      TabIndex        =   57
      Top             =   2775
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
      Left            =   255
      TabIndex        =   56
      Top             =   2325
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find Job No."
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
      Left            =   255
      TabIndex        =   55
      Top             =   1635
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "New Job No."
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
      Left            =   255
      TabIndex        =   54
      Top             =   1980
      Width           =   1185
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
      Left            =   2205
      TabIndex        =   53
      Top             =   2010
      Width           =   225
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
      Left            =   2205
      TabIndex        =   50
      Top             =   1635
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
      Left            =   2205
      TabIndex        =   49
      Top             =   2400
      Width           =   225
   End
End
Attribute VB_Name = "frmEditIndentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cboJobNo_Change()
  If cboJobNo.Text = "" Then Exit Sub
    Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.OtherPayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.AdjDate," _
              & " new_Indent_Information.TotBAmount,new_Indent_Information.OtherAdjAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat,new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & cboJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='I'));"
    Else
        strRecord = "SELECT new_Indent_Information.JobNo,new_Indent_Information.InvNo,new_Indent_Information.PayDate,new_Indent_Information.OtherPayDate,new_Indent_Information.IndAmount,new_Indent_Information.BPayDate,new_Indent_Information.BPayValue,new_Indent_Information.TotIndValue,new_Indent_Information.AdjDate," _
              & " new_Indent_Information.TotBAmount,new_Indent_Information.OtherAdjAmount,new_Indent_Information.TotAdjust,new_Indent_Information.AdjTax,new_Indent_Information.AdjVat,new_Acc_Head.AccName,new_Indent_Information.Remarks,new_Indent_Information.Status FROM new_Indent_Information INNER JOIN new_Acc_Head" _
              & " ON new_Indent_Information.AccID = new_Acc_Head.AccID WHERE (((new_Indent_Information.JobNo)='" & cboJobNo.Text & "') AND ((new_Indent_Information.compID)=" & bytCompID & ") AND ((new_Indent_Information.Status)='E'));"
    End If
    R.Open strRecord, MainConn, adOpenStatic
    
    If R.RecordCount > 0 Then
    
        txtJobNo.Text = R![JobNo]
        txtInvNo.Text = R![InvNo]
        txtPayDate.Text = Format(R![PayDate], cnstDtFrmtI)
        OtherAdjDate.Text = Format(R![OtherPayDate], cnstDtFrmtI)
        txtCashPayValue.Text = Format(R![IndAmount], "##,##0.00")
        txtChequePayDate.Text = Format(R![BPayDate], cnstDtFrmtI)
        txtChequePayValue.Text = Format(R![BPayValue], "0.00")
'        txtChequePayValue.Text = Format(R![BPayValue], "##,##0.00")
        txtIndAmount.Text = Format(R![TotIndValue], "##,##0.00")
        txtAdjDate.Text = IIf(Format(R![AdjDate], cnstDtFrmtI) = "30-Dec-1899", "", Format(R![AdjDate], cnstDtFrmtI))
        txtTax.Text = Format(R![AdjTax], "##,##0.00")
        txtVAT.Text = Format(R![AdjVat], "##,##0.00")
        txtAdjAmount.Text = "0.00"
        txtBAdjAmount.Text = "0.00"
        txtOtherAdjAmount.Text = "0.00"
        OtherTotalAdjAmount.Text = Format(R![OtherAdjAmount], "##,##0.00")
        txtBalance.Text = Format(R![TotBAmount], "##,##0.00")
        txtTotAdjust.Text = Format(R![TotAdjust], "##,##0.00")
       '' txtBalance.Text = Format(R![IndAmount] - R![AdjAmount], "##,##0.00")
        cboAccNo.Text = R![AccName]
        txtRemarks.Text = R![Remarks]
        If R![Status] = "I" Then optEditIm.Value = True Else optEditEx.Value = True
    End If
    R.Close
    Set R = Nothing
'    txtAdjDate.SetFocus
End Sub
Private Sub cboJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
'Private Sub cboJobNo_GotFocus()
' On Error Resume Next
'  cboJobNo.SetListContainer jobSearch
'  setListContainerPosition cboJobNo
'End Sub
Private Sub CmdAdjUpdate_Click()
    If Len(frmEditIndentEntry.cboJobNo.Text) <> 0 Then
        Set R1 = New ADODB.Recordset
            If frmEditIndentEntry.optImport.Value = True Then
                strRecord = "Select * from new_Indent_information Where  CompID=" & bytCompID & " And JobNo = '" & frmEditIndentEntry.cboJobNo.Text & "'And Status='I' ;"
            Else
                strRecord = "Select * from new_Indent_information Where  CompID=" & bytCompID & " And JobNo = '" & frmEditIndentEntry.cboJobNo.Text & "' And Status='E' ;"
            End If
            R1.Open strRecord, MainConn, adOpenStatic
            If R1.RecordCount <> 0 Then
                frmAdjustmentDetails.txtJobNo.Text = R1![JobNo]
                frmAdjustmentDetails.txtAdjDate.Text = Format(R1![AdjDate], cnstDtFrmtI)
                frmAdjustmentDetails.txtAdjAmount = Format(R1![AdjAmount], "##,##0.00")
                frmAdjustmentDetails.txtBAdjAmount = Format(R1![BAdjAmount], "##,##0.00")
                frmAdjustmentDetails.txtOtherAdjAmount = Format(R1![OtherAdjAmount], "##,##0.00")
                frmAdjustmentDetails.Show 1
            End If
            R1.Close
        Set R1 = Nothing
    Else
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    If Len(cboJobNo) = 0 Then
        MsgBox "Select Job No", vbInformation, cnstMsgInfo
        cboJobNo.SetFocus
    Exit Sub
    End If

    If MsgBox("Are you sure to Delete", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
    
        MainComm.CommandText = "Insert into dbo.new_Indent_Information_bak(compID, JobNo, InvNo, PayDate,OtherPayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount, BAdjAmount,OtherAdjAmount, TotAdjust, AdjTax, AdjVat, Deleterow, Username,Userpwd, Comname, Entrydate, Entrytime)" _
                                & " SELECT  compID, JobNo, InvNo, PayDate,OtherPayDate, TotIndValue, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount ,OtherAdjAmount, TotAdjust, AdjTax, AdjVat, 2, '" & struser & "','" & strpass & "','" & getComName & "','" & Date & "', '" & Time & "' FROM dbo.new_Indent_Information WHERE (compID=" & bytCompID & ")" _
                                & " AND (JobNo='" & cboJobNo.Text & "')  ;"
        MainComm.Execute
        
        MainComm.CommandText = "DELETE  FROM new_Indent_Information WHERE ((compID=" & bytCompID & ") AND (JobNo='" & cboJobNo.Text & "'));"
        MainComm.Execute
        
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

    optImport_Click

txtJobNo.Text = "":     txtInvNo.Text = ""
txtPayDate.Text = "":   txtIndAmount.Text = ""
txtAdjDate.Text = "":   txtAdjAmount.Text = ""
txtTax.Text = "":       txtVAT.Text = ""
txtBalance.Text = "":   cboAccNo.ListIndex = -1
txtRemarks.Text = "": txtCashPayValue.Text = ""
txtChequePayDate.Text = "": txtChequePayValue.Text = ""
txtBAdjAmount.Text = "": txtTotAdjust.Text = ""
txtOtherAdjAmount.Text = "0.00"
OtherAdjDate.Text = ""
optEditIm.Value = False
optEditEx.Value = False

cboJobNo.SetFocus
Screen.MousePointer = vbDefault

End Sub
Private Sub cmdSave_Click()
    Dim strNewStatus  As String, strOldStatus As String
    Dim StrTotAdjust As Currency, StrCashAdjust As Currency, StrBankAdjust As Currency, StrOtherAdjust As Currency
    
    If cboJobNo.Text = vbNullString Then
        MsgBox "Select Job No", vbInformation, cnstMsgInfo
        cboJobNo.SetFocus
    Exit Sub
    End If

    If Len(txtJobNo) = 0 Then
        MsgBox "Enter Job No", vbInformation, cnstMsgInfo
        txtJobNo.SetFocus
    Exit Sub
    End If

    If Len(txtInvNo) = 0 Then
        MsgBox "Enter Invoice No", vbInformation, cnstMsgInfo
        txtInvNo.SetFocus
    Exit Sub
    End If
    
    If Len(txtAdjDate) = 0 Then
        MsgBox "Enter Adjust Date", vbInformation, cnstMsgInfo
        txtAdjDate.SetFocus
        Exit Sub
    End If
    
    If Len(txtPayDate) = 0 Then
        MsgBox "Enter Date of Payment", vbInformation, cnstMsgInfo
        txtPayDate.SetFocus
    Exit Sub
    End If

    If Len(txtIndAmount) = 0 Then
        MsgBox "Enter Indent Amount", vbInformation, cnstMsgInfo
        txtIndAmount.SetFocus
    Exit Sub
    End If
    
    If CCur(txtIndAmount) = 0 Then
        MsgBox "Invalid Indent Amount", vbInformation, cnstMsgInfo
        txtIndAmount.SetFocus
    Exit Sub
    End If
    
    If Len(cboAccNo) = 0 Then
        MsgBox "Select Head of Account", vbInformation, cnstMsgInfo
        cboAccNo.SetFocus
    Exit Sub
    End If
    
    Set R = New ADODB.Recordset
    If txtJobNo.Text <> cboJobNo.Text Then
      If optEditIm.Value = True Then
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
      ElseIf optEditEx.Value = True Then
        strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & txtJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
      End If
    
      R.Open strRecord, MainConn, adOpenStatic
      If R.RecordCount > 0 Then
        MsgBox "Duplicate Job No", vbInformation, cnstMsgInfo
        txtJobNo.SetFocus
        Exit Sub
      End If
    End If
    
''Added By Shemul
StrTotAdjust = TotAdjustAmt
StrCashAdjust = TotCashAdjustAmt
StrBankAdjust = TotBankAdjustAmt
StrOtherAdjust = TotOtherAdjustAmt

    If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    
    If Len(txtChequePayValue) = 0 Then txtChequePayValue.Text = "0"
    If optImport.Value = True Then strOldStatus = "I" Else strOldStatus = "E"
    If optEditIm.Value = True Then strNewStatus = "I" Else strNewStatus = "E"

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass

    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn

    MainConn.BeginTrans
        MainComm.CommandText = "UPDATE new_Indent_Information SET JobNo='" & txtJobNo.Text & "',InvNo='" & txtInvNo.Text & "',PayDate='" & Format(txtPayDate.Text, cnstDtFrmtI) & "',IndAmount=" & CCur(txtCashPayValue.Text) & ",BPayDate='" & Format(IIf(txtChequePayDate.Text <> "", txtChequePayDate.Text, "01-JAN-1900"), cnstDtFrmtI) & "'" _
            & " ,OtherPayDate='" & Format(OtherAdjDate.Text, cnstDtFrmtI) & "',OtherAdjAmount=" & CCur(txtOtherAdjAmount.Text) + StrOtherAdjust & ",BPayValue=" & CCur(txtChequePayValue.Text) & ",TotIndValue=" & CCur(txtIndAmount.Text) & ",AdjDate='" & Format(txtAdjDate.Text, cnstDtFrmtI) & "',AdjAmount=" & CCur(txtAdjAmount.Text) + StrCashAdjust & ",BAdjAmount=" & CCur(txtBAdjAmount.Text) + StrBankAdjust & ",TotAdjust=" & CCur(txtAdjAmount.Text) + CCur(txtBAdjAmount.Text) + StrTotAdjust & ",TotBAmount=" & CCur(txtIndAmount.Text) - CCur(txtTotAdjust.Text) & ",AdjTax=" & CCur(txtTax.Text) & "," _
            & " AdjVat=" & CCur(txtVAT.Text) & ",AccID=" & findAccID(cboAccNo.Text) & ",Remarks='" & txtRemarks.Text & "',Status='" & strNewStatus & "' WHERE ((compID=" & bytCompID & ")" _
            & " AND (JobNo='" & cboJobNo.Text & "') AND (Status='" & strOldStatus & "'));"
        MainComm.Execute
        
        MainComm.CommandText = "Insert into dbo.new_Indent_Information_bak(compID, JobNo, InvNo, PayDate, TotIndValue,OtherPayDate, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount, BAdjAmount,OtherAdjAmount, TotAdjust, AdjTax, AdjVat, Deleterow, Username,Userpwd, Comname, Entrydate, Entrytime)" _
                                & " SELECT  compID, JobNo, InvNo, PayDate, TotIndValue,OtherPayDate, BPayDate, BPayValue, IndAmount, AccID, Remarks, Status, AdjDate, TotBAmount, AdjAmount,BAdjAmount ,OtherAdjAmount, TotAdjust, AdjTax, AdjVat, 1, '" & struser & "','" & strpass & "','" & getComName & "', '" & Date & "', '" & Time & "' FROM dbo.new_Indent_Information WHERE (compID=" & bytCompID & ")" _
                                & " AND (JobNo='" & cboJobNo.Text & "') AND (Status='" & strOldStatus & "') ;"
'            MainComm.CommandText = "UPDATE new_Indent_Information SET JobNo='" & txtJobNo.Text & "',InvNo='" & txtInvNo.Text & "',PayDate='" & Format(txtPayDate.Text, cnstDtFrmtI) & "'," _
'            & " IndAmount=" & CCur(txtIndAmount.Text) & ",AdjDate='" & Format(txtAdjDate.Text, cnstDtFrmtI) & "',AdjAmount=" & CCur(txtAdjAmount.Text) & ",AdjTax=" & CCur(txtTax.Text) & "," _
'            & " AdjVat=" & CCur(txtVat.Text) & ",AccID=" & findAccID(cboAccNo.Text) & ",Remarks='" & txtRemarks.Text & "',Status='" & strNewStatus & "' WHERE ((compID=" & bytCompID & ")" _
'            & " AND (JobNo='" & cboJobNo.Text & "') AND (Status='" & strOldStatus & "'));"
        MainComm.Execute
    MainConn.CommitTrans

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
    Exit Sub
    End If

    txtJobNo.Text = "":         txtInvNo.Text = ""
    txtPayDate.Text = "":       txtCashPayValue.Text = "0.00"
    txtChequePayDate.Text = "": txtChequePayValue.Text = "0.00"
    txtIndAmount.Text = "0.00": txtAdjAmount.Text = "0.00"
    txtTax.Text = "0.00":       txtVAT.Text = "0.00"
    txtBalance.Text = "0.00":   cboAccNo.ListIndex = -1
    txtRemarks.Text = ""
    txtTotAdjust.Text = "0.00"
    txtOtherAdjAmount.Text = "0.00"
    OtherAdjDate.Text = ""
    optEditIm.Value = False:    optEditEx.Value = False
    
    cboJobNo.SetFocus
    Screen.MousePointer = vbDefault
     Set R = Nothing
End Sub
Private Sub Form_Load()
    Call addAccName(cboAccNo)
'      strRecord = "SELECT AccID, AccName FROM new_Acc_Head order By AccName;"
    
    Call BFrameDisable
    optImport.Value = True
    txtIndAmount.Text = "0.00"
    txtBalance.Text = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEditIndentEntry = Nothing
End Sub

Private Sub optExport_Click()
    optImport_Click
End Sub
Private Sub optImport_Click()
    cboJobNo.Text = ""
    txtJobNo.Text = "":         txtInvNo.Text = ""
    txtPayDate.Text = "":       txtIndAmount.Text = ""
    txtAdjDate.Text = "":       txtAdjAmount.Text = ""
    txtTax.Text = "":           txtVAT.Text = ""
    txtBalance.Text = "":       cboAccNo.ListIndex = -1
    txtRemarks.Text = ""
    
      Set R = New ADODB.Recordset
    If optImport.Value = True Then
        strRecord = "SELECT JobNo  FROM new_Indent_Information WHERE compID=" & bytCompID & " AND Status='I';"
    Else
        strRecord = "SELECT JobNo  FROM new_Indent_Information WHERE compID=" & bytCompID & " AND Status='E';"
    End If
    
'    cboJobNo.Clear
    
         R.Open strRecord, MainConn, adOpenStatic
             Set cboJobNo.RowSource = R
            cboJobNo.ListField = "JobNo"
             
End Sub



Private Sub OtherAdjDate_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub OtherAdjDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub OtherAdjDate_KeyPress(KeyAscii As Integer)
DateFunc KeyAscii
End Sub

Private Sub OtherAdjDate_LostFocus()
If Len(OtherAdjDate) = 0 Then Exit Sub
    OtherAdjDate.Text = FormatDate(OtherAdjDate)
End Sub

'Private Sub optImport_Click()
'    cboJobNo.Text = ""
'    txtJobNo.Text = "":         txtInvNo.Text = ""
'    txtPayDate.Text = "":       txtIndAmount.Text = ""
'    txtAdjDate.Text = "":       txtAdjAmount.Text = ""
'    txtTax.Text = "":           txtVat.Text = ""
'    txtBalance.Text = "":       cboAccNo.ListIndex = -1
'    txtRemarks.Text = ""
'
'    Set R = New ADODB.Recordset
'    If optImport.Value = True Then
'        prcSetData "SELECT JobNo J,JobNo ID FROM new_Indent_Information WHERE compID=" & bytCompID & " AND Status='I';", arrsearch, cboJobNo
'    Else
'        prcSetData "SELECT JobNo J,JobNo ID  FROM new_Indent_Information WHERE compID=" & bytCompID & " AND Status='E';", arrsearch, cboJobNo
'    End If
'
''   cboJobNo.Clear
''
''
''
''
''    R.Open strRecord, MainConn, adOpenStatic
''    If R.RecordCount > 0 Then
''        Do Until R.EOF
''            cboJobNo.AddItem R![JobNo]
''        R.MoveNext
''        Loop
''    End If
''    R.Close
'
'
''    Set R = Nothing
'End Sub


Private Sub txtAdjAmount_GotFocus()
    Call txtGotFocus(txtAdjAmount)
End Sub
Private Sub txtAdjAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtAdjAmount_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtAdjAmount_LostFocus()
    txtTotAdjust.Text = "0.00"
    If Len(txtAdjAmount) = 0 Then txtAdjAmount.Text = "0.00"
'    txtAdjAmount.Text = Format(txtAdjAmount.Text, "##,##0.00")
    txtTotAdjust.Text = Format(CCur(TotAdjustAmt) + CCur(Val(txtAdjAmount.Text)) + CCur(Val(txtBAdjAmount.Text)), "##,##0.00")
    txtBalance.Text = Format(CCur(txtIndAmount.Text) - CCur(txtTotAdjust.Text), "##,##0.00")
'    txtBalance.Text = Format(findTotal - Val(txtTotAdjust.Text), "##,##0.00")
   End Sub
Private Sub txtAdjDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtAdjDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtAdjDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub
Private Sub txtAdjDate_LostFocus()
    If Len(txtAdjDate) = 0 Then Exit Sub
    txtAdjDate.Text = FormatDate(txtAdjDate)
    
End Sub

Private Sub txtBAdjAmount_GotFocus()
Call txtGotFocus(txtBAdjAmount)
End Sub
Private Sub txtBAdjAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtBAdjAmount_KeyPress(KeyAscii As Integer)
 CurrFunc KeyAscii
End Sub
Private Sub txtBAdjAmount_LostFocus()
  txtTotAdjust.Text = "0.00"
    If Len(txtBAdjAmount) = 0 Then txtBAdjAmount.Text = "0.00"
'    txtBAdjAmount.Text = Format(txtBAdjAmount.Text, "##,##0.00")
      If txtChequePayValue.Text <> "0.00" Then
         If txtChequePayValue.Text <> txtBAdjAmount.Text Then
         MsgBox "Adjustment Amount Not match ", vbInformation, cnstMsgInfo
        txtBAdjAmount.Text = "0.00"
'        txtBAdjAmount.SetFocus
         End If
      End If
'      txtBAdjAmount.Text=
'    txtTotAdjust.Text = TotAdjustAmt + Format(CCur(Val(txtAdjAmount.Text)) + CCur(Val(txtBAdjAmount.Text)), "##,##0.00")
txtTotAdjust.Text = Format(CCur(TotAdjustAmt) + CCur(Val(txtAdjAmount.Text)) + CCur(Val(txtBAdjAmount.Text)), "##,##0.00")
txtBalance.Text = Format(CCur(txtIndAmount.Text) - CCur(txtTotAdjust.Text), "##,##0.00")
End Sub

Private Sub txtBalance_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBalance_LostFocus()
    If Len(txtBalance) = 0 Then txtBalance.Text = "0.00"
    txtBalance.Text = Format(txtBalance.Text, "##,##0.00")
End Sub

Private Sub txtCashPayValue_GotFocus()
 SendKeys "{Home}+{End}"
End Sub
Private Sub txtCashPayValue_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCashPayValue_LostFocus()
    If Len(txtCashPayValue) = 0 Then txtCashPayValue.Text = "0.00"
    txtCashPayValue.Text = Format(txtCashPayValue.Text, "##,##0.00")
End Sub

Private Sub txtChequePayDate_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtChequePayDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChequePayDate_KeyPress(KeyAscii As Integer)
DateFunc KeyAscii
End Sub

Private Sub txtChequePayDate_LostFocus()
    If Len(txtChequePayDate) = 0 Then Exit Sub
    txtChequePayDate.Text = FormatDate(txtChequePayDate)
End Sub

Private Sub txtChequePayValue_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtChequePayValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChequePayValue_KeyPress(KeyAscii As Integer)
CurrFunc KeyAscii
End Sub

Private Sub txtChequePayValue_LostFocus()
    If Len(txtChequePayValue) = 0 Then txtChequePayValue.Text = "0.00"
    txtChequePayValue.Text = Format(txtChequePayValue.Text, "##,##0.00")
End Sub

Private Sub txtIndAmount_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtIndAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtIndAmount_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtIndAmount_LostFocus()
    If Len(txtIndAmount) = 0 Then txtIndAmount.Text = "0.00"
    txtIndAmount.Text = Format(txtIndAmount.Text, "##,##0.00")
   '' txtBalance.Text = Format(CCur(txtIndAmount.Text) - CCur(txtAdjAmount.Text), "##,##0.00")
End Sub
Private Sub txtInvNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
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
Private Sub txtJobNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtJobNo_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtJobNo_LostFocus()
    If Len(txtJobNo) = 0 Then Exit Sub
    txtJobNo.Text = Trim(UCase(txtJobNo.Text))
End Sub



Private Sub txtOtherAdjAmount_GotFocus()
Call txtGotFocus(txtOtherAdjAmount)
End Sub

Private Sub txtOtherAdjAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtOtherAdjAmount_KeyPress(KeyAscii As Integer)
CurrFunc KeyAscii
End Sub

Private Sub txtPayDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtPayDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtPayDate_KeyPress(KeyAscii As Integer)
'    DateFunc KeyAscii
End Sub
Private Sub txtPayDate_LostFocus()
    If Len(txtPayDate) = 0 Then Exit Sub
    txtPayDate.Text = FormatDate(txtPayDate)

End Sub
Private Sub txtRemarks_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtTax_GotFocus()
    Call txtGotFocus(txtTax)
End Sub
Private Sub txtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTax_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtTax_LostFocus()
    If Len(txtTax) = 0 Then txtTax.Text = "0.00"
    txtTax.Text = Format(txtTax.Text, "##,##0.00")
End Sub

Private Sub txtTotAdjust_LostFocus()
    If Len(txtTotAdjust) = 0 Then txtTotAdjust.Text = "0.00"
    txtTotAdjust.Text = Format(CCur(txtTotAdjust.Text), "##,##0.00")
End Sub

Private Sub txtVat_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtVat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtVat_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtVat_LostFocus()
    If Len(txtVAT) = 0 Then txtVAT.Text = "0.00"
    txtVAT.Text = Format(txtVAT.Text, "##,##0.00")
End Sub
Private Sub BFrameDisable()
txtPayDate.Enabled = False
txtCashPayValue.Enabled = False
txtChequePayDate.Enabled = False
txtChequePayValue.Enabled = False
txtIndAmount.Enabled = False
txtBalance.Enabled = False
OtherTotalAdjAmount.Enabled = False
End Sub

Function TotAdjustAmt() As Currency
    TotAdjustAmt = 0
    Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT TotAdjust FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT TotAdjust FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![TotAdjust]) Then TotAdjustAmt = R![TotAdjust]
    End If
    R.Close
    Set R = Nothing
End Function

Function TotCashAdjustAmt() As Currency
TotCashAdjustAmt = 0
Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT AdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT AdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![AdjAmount]) Then TotCashAdjustAmt = R![AdjAmount]
    End If
    R.Close
    Set R = Nothing
End Function
Function TotBankAdjustAmt() As Currency
TotBankAdjustAmt = 0
Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT BAdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT BAdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![BAdjAmount]) Then TotBankAdjustAmt = R![BAdjAmount]
    End If
    R.Close
    Set R = Nothing
End Function


Function TotOtherAdjustAmt() As Currency
TotOtherAdjustAmt = 0
Set R = New ADODB.Recordset
    
    If optImport.Value = True Then
        strRecord = "SELECT OtherAdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
    Else
        strRecord = "SELECT OtherAdjAmount FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
    End If
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        If Not IsNull(R![OtherAdjAmount]) Then TotOtherAdjustAmt = R![OtherAdjAmount]
    End If
    R.Close
    Set R = Nothing
End Function

Public Sub CheckRecord()
    Set R = New ADODB.Recordset
        If optImport.Value = True Then
            strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='I') AND (compID=" & bytCompID & "));"
        Else
            strRecord = "SELECT JobNo FROM new_Indent_Information WHERE ((JobNo='" & cboJobNo.Text & "') AND (Status='E') AND (compID=" & bytCompID & "));"
        End If
    
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount = 0 Then
            MsgBox "Record Not Found", vbInformation, cnstMsgInfo
            cboJobNo.SetFocus
        cboJobNo.Text = ""
        Exit Sub
        End If
    
        R.Close
    Set R = Nothing
End Sub


