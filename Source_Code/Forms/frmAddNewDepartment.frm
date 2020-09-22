VERSION 5.00
Begin VB.Form frmAddNewDepartment 
   Caption         =   "Add New Department Module"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddNewDepartment.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8760
      Picture         =   "frmAddNewDepartment.frx":1E717
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8760
      Picture         =   "frmAddNewDepartment.frx":2145B
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   8760
      Picture         =   "frmAddNewDepartment.frx":2419F
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   8760
      Picture         =   "frmAddNewDepartment.frx":26EE3
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   4815
      Left            =   8520
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
      Left            =   2760
      TabIndex        =   8
      Top             =   5085
      Width           =   1695
   End
   Begin VB.Label Label12 
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
      Left            =   2760
      TabIndex        =   7
      Top             =   3885
      Width           =   1575
   End
   Begin VB.Label Label16 
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
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   4485
      Width           =   1335
   End
   Begin VB.Label Label13 
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
      Left            =   2760
      TabIndex        =   5
      Top             =   3285
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   8160
      Y1              =   7560
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   1800
      Y1              =   2760
      Y2              =   7560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   8160
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   2160
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Information"
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
      TabIndex        =   4
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   4800
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "frmAddNewDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
