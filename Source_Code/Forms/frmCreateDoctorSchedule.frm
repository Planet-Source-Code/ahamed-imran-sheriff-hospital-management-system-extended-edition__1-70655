VERSION 5.00
Begin VB.Form frmCreateDoctorSchedule 
   Caption         =   "Create Doctor's Schedule Module"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCreateDoctorSchedule.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   22
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton cmdCusWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      ToolTipText     =   "Click Here to select Customer"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8760
      Picture         =   "frmCreateDoctorSchedule.frx":1F5A1
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8760
      Picture         =   "frmCreateDoctorSchedule.frx":222E5
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   8760
      Picture         =   "frmCreateDoctorSchedule.frx":25029
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   8760
      Picture         =   "frmCreateDoctorSchedule.frx":27D6D
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   9
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   3
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   1
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Saturday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   4050
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   4560
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   5175
      Left            =   8520
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   6525
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   6045
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   3525
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Channeling Charges"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   5085
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Time In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   5565
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3045
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   4005
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   8160
      Y1              =   7680
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   1800
      Y1              =   2520
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   8160
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   2160
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Schedule Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   5520
      Y1              =   2520
      Y2              =   2520
   End
End
Attribute VB_Name = "frmCreateDoctorSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblClearanceNo_Click()

End Sub

