VERSION 5.00
Begin VB.MDIForm MDIfrmMain 
   BackColor       =   &H00808000&
   Caption         =   " New Project"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "MDIfrmMain.frx":0000
   LinkTopic       =   "MDIfrmMain"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Banner1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11850
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.Timer Timer2 
         Left            =   10800
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   11280
         Top             =   240
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   10560
         TabIndex        =   1
         Top             =   520
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Pacific"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label lblPeriod 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   8835
      End
      Begin VB.Label lblComName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   60
         Width           =   8775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   10560
         TabIndex        =   3
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   10560
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Menu mnuComInformation 
      Caption         =   "Company &Information"
      Begin VB.Menu mnuNew 
         Caption         =   "New Company"
      End
      Begin VB.Menu mnuNext1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Company"
      End
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "   Data E&ntry   "
   End
   Begin VB.Menu mnuDataEdit 
      Caption         =   "   Data E&dit   "
   End
   Begin VB.Menu mnuReport 
      Caption         =   "   Data &Report   "
   End
   Begin VB.Menu mnuTools 
      Caption         =   "   &Option   "
      Begin VB.Menu mnuModifyCom 
         Caption         =   "Modify Company"
      End
      Begin VB.Menu mnuDeleteCompany 
         Caption         =   "Delete Company"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "   &About   "
   End
   Begin VB.Menu mnuExit 
      Caption         =   "   E&xit   "
   End
End
Attribute VB_Name = "MDIfrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
  ConnectDatabase
  lblComName.Caption = "New Project"
  mnuComInformation.Visible = True
  mnuDataEntry.Visible = False
  mnuDataEdit.Visible = False
  mnuReport.Visible = False
  mnuTools.Visible = False
  Timer1_Timer
End Sub
Private Sub MDIForm_Resize()
  If WindowState = 0 Then WindowState = 2
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
  Set MDIfrmMain = Nothing
End Sub
Private Sub mnuAbout_Click()
  frmAbout.Show 1
End Sub
Private Sub mnuDeleteCompany_Click()
  frmDeleteCompany.Show 1
End Sub
Private Sub mnuExit_Click()
  End
End Sub
Private Sub mnuModifyCom_Click()
  frmModifyCompany.Show 1
End Sub
Private Sub mnuNew_Click()
  frmCreateCompany.Show 1
End Sub
Private Sub mnuOpen_Click()
  frmOpenCompany.Show 1
End Sub
Private Sub Timer1_Timer()
  Dim dtMyDate As Date
  dtMyDate = Now
  lblNumber = Format(dtMyDate, "DDDD")
  lblDate = Format(dtMyDate, cnstDtFrmtI)
  lblTime = Format(dtMyDate, "HH-MM-SS AMPM")
End Sub
