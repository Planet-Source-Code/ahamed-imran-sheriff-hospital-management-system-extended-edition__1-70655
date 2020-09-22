VERSION 5.00
Begin VB.Form frmAddNewDoctor 
   Caption         =   "Add New Doctor Module"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddNewDoctor.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   11745
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   6360
      Picture         =   "frmAddNewDoctor.frx":1D48F
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Height          =   855
      Left            =   9600
      Picture         =   "frmAddNewDoctor.frx":201D3
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8520
      Picture         =   "frmAddNewDoctor.frx":22F17
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   7440
      Picture         =   "frmAddNewDoctor.frx":25C5B
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   30
      Top             =   4080
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmAddNewDoctor.frx":2899F
      Left            =   8520
      List            =   "frmAddNewDoctor.frx":289AC
      TabIndex        =   29
      Text            =   "----------SELECT-----------"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   885
      Left            =   8520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   24
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   22
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   19
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   17
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   15
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   885
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   5520
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
      ItemData        =   "frmAddNewDoctor.frx":289D0
      Left            =   3000
      List            =   "frmAddNewDoctor.frx":289DA
      TabIndex        =   12
      Text            =   "----------SELECT-----------"
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtClearanceNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   11
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000001&
      X1              =   6000
      X2              =   11040
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000001&
      X1              =   6000
      X2              =   11040
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000001&
      X1              =   6000
      X2              =   6000
      Y1              =   6600
      Y2              =   8160
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000001&
      X1              =   11040
      X2              =   11040
      Y1              =   6600
      Y2              =   8160
   End
   Begin VB.Label Label11 
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
      Left            =   6480
      TabIndex        =   31
      Top             =   4125
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Category"
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
      Left            =   6480
      TabIndex        =   28
      Top             =   3165
      Width           =   1575
   End
   Begin VB.Label Label16 
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
      Left            =   6600
      TabIndex        =   27
      Top             =   5085
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Referring Charges"
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
      Left            =   6480
      TabIndex        =   25
      Top             =   4605
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Visiting Charges"
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
      Left            =   6480
      TabIndex        =   23
      Top             =   3645
      Width           =   1575
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   6000
      X2              =   6000
      Y1              =   2760
      Y2              =   6240
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   6000
      X2              =   11040
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   11040
      X2              =   11040
      Y1              =   2760
      Y2              =   6240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   8400
      X2              =   11040
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   6000
      X2              =   6360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Details"
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
      Left            =   6480
      TabIndex        =   21
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "License No."
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
      Left            =   1200
      TabIndex        =   20
      Top             =   7605
      Width           =   1695
   End
   Begin VB.Label Label8 
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
      Left            =   1200
      TabIndex        =   18
      Top             =   7125
      Width           =   1695
   End
   Begin VB.Label Label7 
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
      Left            =   1200
      TabIndex        =   16
      Top             =   6645
      Width           =   1695
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   5565
      Width           =   1335
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   5085
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   4605
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   1200
      TabIndex        =   6
      Top             =   4125
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   3645
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Left            =   1200
      TabIndex        =   3
      Top             =   3165
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   720
      X2              =   720
      Y1              =   2760
      Y2              =   8160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   720
      X2              =   5640
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   5640
      Y1              =   2760
      Y2              =   8160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2880
      X2              =   5640
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   720
      X2              =   960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblFrameTitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Details"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label lblClearanceNo 
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
      Left            =   4080
      TabIndex        =   0
      Top             =   2205
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddNewDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
