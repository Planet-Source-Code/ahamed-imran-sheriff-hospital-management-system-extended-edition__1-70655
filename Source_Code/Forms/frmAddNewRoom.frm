VERSION 5.00
Begin VB.Form frmAddNewRoom 
   Caption         =   "Add New Room Module"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddNewRoom.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   4920
      TabIndex        =   20
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      ToolTipText     =   "Click Here to select Customer"
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   4920
      TabIndex        =   18
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCusWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   17
      ToolTipText     =   "Click Here to select Customer"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   15
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8880
      Picture         =   "frmAddNewRoom.frx":1D1EE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8880
      Picture         =   "frmAddNewRoom.frx":1FF32
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   8880
      Picture         =   "frmAddNewRoom.frx":22C76
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   8880
      Picture         =   "frmAddNewRoom.frx":259BA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   4920
      TabIndex        =   0
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Cost"
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
      TabIndex        =   16
      Top             =   5685
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   4695
      Left            =   8640
      Top             =   2760
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
      Top             =   6165
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID"
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
      Top             =   3765
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward ID"
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
      Top             =   4725
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward Number"
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
      TabIndex        =   11
      Top             =   5205
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Room ID"
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
      Top             =   3285
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
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
      TabIndex        =   9
      Top             =   4245
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   8280
      X2              =   8280
      Y1              =   7440
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   1920
      Y1              =   2760
      Y2              =   7440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   8280
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   2280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Information"
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
      TabIndex        =   8
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   8280
      X2              =   4320
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "frmAddNewRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
