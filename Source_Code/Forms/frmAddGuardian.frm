VERSION 5.00
Begin VB.Form frmAddGuardian 
   Caption         =   "Add Guardian Details Module"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddGuardian.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   23
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   21
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   19
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   17
      Top             =   6120
      Width           =   2295
   End
   Begin VB.ComboBox cboSearchType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmAddGuardian.frx":1EA66
      Left            =   4560
      List            =   "frmAddGuardian.frx":1EA70
      TabIndex        =   16
      Text            =   "----------SELECT-----------"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddGuardian.frx":1EA82
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddGuardian.frx":217C6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddGuardian.frx":2450A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddGuardian.frx":2724E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   4560
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Relation To Patient"
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
      Left            =   2520
      TabIndex        =   24
      Top             =   7605
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
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
      Left            =   2520
      TabIndex        =   22
      Top             =   7125
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No. (Mob)"
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
      Left            =   2520
      TabIndex        =   20
      Top             =   6645
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No. (Home)"
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
      Left            =   2520
      TabIndex        =   18
      Top             =   6165
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   6015
      Left            =   8280
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   2520
      TabIndex        =   15
      Top             =   5205
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   3285
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   2520
      TabIndex        =   13
      Top             =   4245
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "NIC Number"
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
      Left            =   2520
      TabIndex        =   12
      Top             =   4725
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian ID"
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
      Left            =   2520
      TabIndex        =   11
      Top             =   2805
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   3765
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   7920
      Y1              =   8280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1560
      X2              =   1560
      Y1              =   2280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1560
      X2              =   7920
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1560
      X2              =   1920
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian Information"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   4320
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "frmAddGuardian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

