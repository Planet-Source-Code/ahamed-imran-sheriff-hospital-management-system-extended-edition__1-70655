VERSION 5.00
Begin VB.Form frmAddNewMedicine 
   Caption         =   "Add New Medicine Module"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddNewMedicine.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   4680
      TabIndex        =   16
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   6
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddNewMedicine.frx":1E118
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddNewMedicine.frx":20E5C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddNewMedicine.frx":23BA0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddNewMedicine.frx":268E4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   975
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   8040
      X2              =   4440
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine Information"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   2040
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   8040
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   1680
      Y1              =   2760
      Y2              =   7560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   8040
      X2              =   8040
      Y1              =   7560
      Y2              =   2760
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
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
      Left            =   2640
      TabIndex        =   14
      Top             =   4485
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine ID"
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
      Left            =   2640
      TabIndex        =   13
      Top             =   3285
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-order Level"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   5685
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Units In Stock"
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
      Left            =   2640
      TabIndex        =   11
      Top             =   5085
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine Name"
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
      Left            =   2640
      TabIndex        =   10
      Top             =   3885
      Width           =   1575
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
      Left            =   2640
      TabIndex        =   9
      Top             =   6285
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   4815
      Left            =   8400
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddNewMedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
