VERSION 5.00
Begin VB.Form frmAddNewHospitalService 
   Caption         =   "Add New Hospital Service Module"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddNewHospitalService.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8880
      Picture         =   "frmAddNewHospitalService.frx":1FA05
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8880
      Picture         =   "frmAddNewHospitalService.frx":22749
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   8880
      Picture         =   "frmAddNewHospitalService.frx":2548D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   8880
      Picture         =   "frmAddNewHospitalService.frx":281D1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   4920
      TabIndex        =   0
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   4815
      Left            =   8640
      Top             =   2880
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
      Left            =   2880
      TabIndex        =   14
      Top             =   5925
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Name"
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
      Left            =   2880
      TabIndex        =   13
      Top             =   4125
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   5325
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Service ID"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   3525
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount / Rate"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   4725
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   8280
      X2              =   8280
      Y1              =   7680
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   1920
      Y1              =   2880
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   8280
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   2280
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Service Information"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   8280
      X2              =   5400
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmAddNewHospitalService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
