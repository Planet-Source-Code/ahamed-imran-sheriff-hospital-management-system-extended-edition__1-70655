VERSION 5.00
Begin VB.Form frmAddNewWard 
   Caption         =   "Add New Ward Module"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddNewWard.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCusWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   6720
      TabIndex        =   17
      ToolTipText     =   "Click Here to select Customer"
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8760
      Picture         =   "frmAddNewWard.frx":1D307
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8760
      Picture         =   "frmAddNewWard.frx":2004B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   8760
      Picture         =   "frmAddNewWard.frx":22D8F
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   8760
      Picture         =   "frmAddNewWard.frx":25AD3
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   4800
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   4815
      Left            =   8520
      Top             =   2640
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
      TabIndex        =   16
      Top             =   6165
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
      TabIndex        =   15
      Top             =   3765
      Width           =   1575
   End
   Begin VB.Label Label15 
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
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   4965
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Rooms"
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
      TabIndex        =   13
      Top             =   5565
      Width           =   1815
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
      TabIndex        =   12
      Top             =   3165
      Width           =   1575
   End
   Begin VB.Label Label11 
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
      Left            =   2760
      TabIndex        =   11
      Top             =   4365
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   8160
      Y1              =   7440
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   1800
      Y1              =   2640
      Y2              =   7440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   8160
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1800
      X2              =   2160
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward Information"
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
      TabIndex        =   10
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   4200
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmAddNewWard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
