VERSION 5.00
Begin VB.Form frmQuickLaunch 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Quick Launch"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmQuickLaunch.frx":0000
   ScaleHeight     =   9060
   ScaleWidth      =   12795
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTimer 
      Interval        =   100
      Left            =   7200
      Top             =   960
   End
   Begin VB.CommandButton cmdAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Caption         =   "About HMS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblDisplayPresentTime 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblDisplayTimeIn 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblCurrentDateTime 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   7920
      TabIndex        =   6
      Top             =   3650
      Width           =   4815
   End
   Begin VB.Label lblPresentTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Present Time :"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblTimeIn 
      BackStyle       =   0  'Transparent
      Caption         =   "Time In : "
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblDesignation 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome, "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Colombo, Sri Lanka"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Durdans Hospitals (Pvt) Ltd."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmQuickLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub Form_Load()
    lblDisplayTimeIn.Caption = DateTime.Time
    lblCurrentDateTime.Caption = "Today is " & FormatDateTime(Now, vbLongDate)
End Sub


Private Sub tmrTimer_Timer()
    lblDisplayPresentTime.Caption = DateTime.Time
End Sub
