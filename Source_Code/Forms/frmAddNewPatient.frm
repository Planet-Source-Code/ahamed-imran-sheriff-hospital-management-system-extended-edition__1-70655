VERSION 5.00
Begin VB.Form frmAddNewInPatient 
   Caption         =   "Add New Patient Details Module"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddNewPatient.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   8520
      TabIndex        =   32
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      Height          =   285
      Left            =   8520
      TabIndex        =   30
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdCusWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   10440
      TabIndex        =   29
      ToolTipText     =   "Click Here to select Customer"
      Top             =   5640
      Width           =   375
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "frmAddNewPatient.frx":1F14A
      Left            =   8520
      List            =   "frmAddNewPatient.frx":1F157
      TabIndex        =   28
      Text            =   "----------SELECT-----------"
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtClearanceNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   10
      Top             =   2280
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
      ItemData        =   "frmAddNewPatient.frx":1F17B
      Left            =   3000
      List            =   "frmAddNewPatient.frx":1F185
      TabIndex        =   9
      Text            =   "----------SELECT-----------"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   885
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   7
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   5
      Top             =   4200
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
      ItemData        =   "frmAddNewPatient.frx":1F197
      Left            =   8520
      List            =   "frmAddNewPatient.frx":1F1A4
      TabIndex        =   4
      Text            =   "----------SELECT-----------"
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   4800
      Picture         =   "frmAddNewPatient.frx":1F1C8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   5880
      Picture         =   "frmAddNewPatient.frx":21F0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Height          =   855
      Left            =   6960
      Picture         =   "frmAddNewPatient.frx":24C50
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   3720
      Picture         =   "frmAddNewPatient.frx":27994
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Enabled         =   0   'False
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
      TabIndex        =   33
      Top             =   6165
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Corporate ID"
      Enabled         =   0   'False
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
      Top             =   5685
      Width           =   1455
   End
   Begin VB.Label lblClearanceNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
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
      TabIndex        =   27
      Top             =   2325
      Width           =   1335
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
      TabIndex        =   26
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   720
      X2              =   960
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2880
      X2              =   11280
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11280
      X2              =   11280
      Y1              =   2880
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   720
      X2              =   11280
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   720
      X2              =   720
      Y1              =   2880
      Y2              =   6840
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
      TabIndex        =   25
      Top             =   3285
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
      TabIndex        =   24
      Top             =   3765
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
      TabIndex        =   23
      Top             =   4245
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
      TabIndex        =   22
      Top             =   4725
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
      TabIndex        =   21
      Top             =   5205
      Width           =   1335
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
      TabIndex        =   20
      Top             =   5685
      Width           =   1335
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
      Left            =   6480
      TabIndex        =   19
      Top             =   3285
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
      Left            =   6480
      TabIndex        =   18
      Top             =   3765
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Occupation"
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
      TabIndex        =   17
      Top             =   4245
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type"
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
      TabIndex        =   16
      Top             =   5205
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Civil Status"
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
      TabIndex        =   15
      Top             =   4725
      Width           =   1575
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000001&
      X1              =   8400
      X2              =   8400
      Y1              =   7200
      Y2              =   8520
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000001&
      X1              =   3360
      X2              =   3360
      Y1              =   7200
      Y2              =   8520
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000001&
      X1              =   3360
      X2              =   8400
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000001&
      X1              =   3360
      X2              =   8400
      Y1              =   7200
      Y2              =   7200
   End
End
Attribute VB_Name = "frmAddNewInPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
