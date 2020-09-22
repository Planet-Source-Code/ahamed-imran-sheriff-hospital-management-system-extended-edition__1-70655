VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInpatientsMaintenance 
   Caption         =   "Inpatients Maintenance Module"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditPatientDetails.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInvalidTypingMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3720
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   50
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   51
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.PictureBox picInvalidKeyMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8640
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   48
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Digits Here! Only Alphabets Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   49
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.PictureBox picInvalidDataMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3720
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Digits Here! Only Alphabets Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   240
      Top             =   4320
   End
   Begin VB.PictureBox picInvalidKeypressMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8640
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   44
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdLaunchInpatientSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Launch Inpatient Search Wizard"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click here to launch the Search Wizard"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.CommandButton cmdStep3 
      BackColor       =   &H80000013&
      Caption         =   "Step 3"
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
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdStep2 
      BackColor       =   &H80000013&
      Caption         =   "Step 2"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdStep1 
      BackColor       =   &H80000013&
      Caption         =   "Step 1"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmEditPatientDetails.frx":1F99C
      Height          =   855
      Left            =   7920
      Picture         =   "frmEditPatientDetails.frx":1FE82
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmEditPatientDetails.frx":22BC6
      Height          =   855
      Left            =   6840
      Picture         =   "frmEditPatientDetails.frx":23044
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      DisabledPicture =   "frmEditPatientDetails.frx":25D88
      Height          =   855
      Left            =   5760
      Picture         =   "frmEditPatientDetails.frx":2618A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmEditPatientDetails.frx":28ECE
      Height          =   855
      Left            =   9000
      Picture         =   "frmEditPatientDetails.frx":29397
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmEditPatientDetails.frx":2C0DB
      Height          =   855
      Left            =   10080
      Picture         =   "frmEditPatientDetails.frx":2C59A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtCompanyName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox txtCompanyID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdCompanySearchWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   10320
      TabIndex        =   17
      ToolTipText     =   "Click Here to select a Company"
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox txtPatientID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.ComboBox cboCivilStatus 
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
      ItemData        =   "frmEditPatientDetails.frx":2F2DE
      Left            =   8400
      List            =   "frmEditPatientDetails.frx":2F2E8
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtPatientOccupation 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   30
      TabIndex        =   13
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox txtPhoneMob 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   15
      TabIndex        =   12
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtPhoneHome 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   15
      TabIndex        =   11
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   6240
      Width           =   2295
   End
   Begin VB.ComboBox cboGender 
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
      ItemData        =   "frmEditPatientDetails.frx":2F300
      Left            =   2880
      List            =   "frmEditPatientDetails.frx":2F30A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtNICNumber 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   9
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtSurname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox txtFirstName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   3840
      Width           =   2295
   End
   Begin VB.ComboBox cboAccountType 
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
      ItemData        =   "frmEditPatientDetails.frx":2F31C
      Left            =   8400
      List            =   "frmEditPatientDetails.frx":2F326
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmEditPatientDetails.frx":2F341
      Height          =   750
      Left            =   3960
      Picture         =   "frmEditPatientDetails.frx":2F71B
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7740
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmEditPatientDetails.frx":318D7
      Height          =   750
      Left            =   3000
      Picture         =   "frmEditPatientDetails.frx":31CAD
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7740
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmEditPatientDetails.frx":33E69
      Height          =   750
      Left            =   1080
      Picture         =   "frmEditPatientDetails.frx":34245
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7740
      Width           =   890
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmEditPatientDetails.frx":36401
      Height          =   750
      Left            =   2040
      Picture         =   "frmEditPatientDetails.frx":36816
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7740
      Width           =   890
   End
   Begin MSComCtl2.DTPicker dtpDateOfBirth 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Top             =   5280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTrailingForeColor=   -2147483638
      Format          =   54329345
      CurrentDate     =   39552
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "***Please Note That All Non-Compulsory Fields Have Been Marked With An Asterisk"
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
      Left            =   600
      TabIndex        =   43
      Top             =   2550
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   5640
      Top             =   7560
      Width           =   5535
   End
   Begin VB.Label lblCompanyName 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
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
      Left            =   6360
      TabIndex        =   42
      Top             =   6285
      Width           =   1695
   End
   Begin VB.Label lblCompanyID 
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID"
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
      Left            =   6360
      TabIndex        =   41
      Top             =   5805
      Width           =   1575
   End
   Begin VB.Label lblPatientID 
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
      Left            =   1080
      TabIndex        =   40
      Top             =   3405
      Width           =   1335
   End
   Begin VB.Label lblCivilStatus 
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
      Left            =   6360
      TabIndex        =   39
      Top             =   4845
      Width           =   1575
   End
   Begin VB.Label lblAccountType 
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
      Left            =   6360
      TabIndex        =   38
      Top             =   5325
      Width           =   1575
   End
   Begin VB.Label lblPatientOccupation 
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
      Left            =   6360
      TabIndex        =   37
      Top             =   4365
      Width           =   1695
   End
   Begin VB.Label lblPhoneMob 
      BackStyle       =   0  'Transparent
      Caption         =   "* Phone No. (Mob)"
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
      Left            =   6360
      TabIndex        =   36
      Top             =   3885
      Width           =   1695
   End
   Begin VB.Label lblPhoneHome 
      BackStyle       =   0  'Transparent
      Caption         =   " Phone No. (Home)"
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
      Left            =   6360
      TabIndex        =   35
      Top             =   3405
      Width           =   1815
   End
   Begin VB.Label lblAddress 
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
      Left            =   1080
      TabIndex        =   34
      Top             =   6285
      Width           =   1335
   End
   Begin VB.Label lblNICNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "* NIC Number"
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
      Left            =   1080
      TabIndex        =   33
      Top             =   5805
      Width           =   1335
   End
   Begin VB.Label lblDateOfBirth 
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
      Left            =   1080
      TabIndex        =   32
      Top             =   5325
      Width           =   1335
   End
   Begin VB.Label lblGender 
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
      Left            =   1080
      TabIndex        =   31
      Top             =   4845
      Width           =   1335
   End
   Begin VB.Label lblSurname 
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
      Left            =   1080
      TabIndex        =   30
      Top             =   4365
      Width           =   1335
   End
   Begin VB.Label lblFirstName 
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
      Left            =   1080
      TabIndex        =   29
      Top             =   3885
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   600
      Y1              =   3000
      Y2              =   7320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   11160
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11160
      X2              =   11160
      Y1              =   3000
      Y2              =   7320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2760
      X2              =   11160
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   840
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblPersonalDetails 
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
      Left            =   960
      TabIndex        =   28
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   600
      Top             =   7560
      Width           =   4815
   End
End
Attribute VB_Name = "frmInpatientsMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Inpatients Maintenance Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 24/04/08
'Date Of Last Modification: 24/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Inpatient_Maintenance Table
'----------------------------------------------------------------------------

Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean

'The following Boolean Variable will determine if the Doctor's Date Of Birth
'entered by the user is valid
Dim dateFlag As Boolean

'The following variables will be used to autogenerate the Doctor ID
Dim iNumOfRecords As Integer    'This variable holds the number of records in the table
Dim strCode As String   'This variable will eventually hold the Doctor ID to be autogenerated


'The following variables will be used to autogenerate the Guardian ID to be
'displayed on the Guardians Maintenance form on form load
Dim iNumOfGuardians As Integer  'This variable holds the number of records in the table
Dim strDisplayID As String  'This variable will eventually hold the Guardian ID to be autogenerated



Private Sub cboAccountType_Click()   'This function will manipulate other controls according to the type of account
    
    'The following block of code will disable the Company related fields if the
    'Patient Account Type is "Individual"
    If cboAccountType.ListIndex = 0 Then
        lblCompanyID.Enabled = True
        txtCompanyID.Enabled = True
        lblCompanyName.Enabled = True
        txtCompanyName.Enabled = True
        cmdCompanySearchWizard.Enabled = True
    Else
        lblCompanyID.Enabled = False
        txtCompanyID.Enabled = False
        lblCompanyName.Enabled = False
        txtCompanyName.Enabled = False
        cmdCompanySearchWizard.Enabled = False
    End If
    
End Sub

Private Sub cmdAddNew_Click() 'This function adds a new recordset into the database

    enableAllFields     'Calling a Private Function To Enable All Fields
    clearAllFields      'Calling a Private Function To Clear All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    txtNICNumber.Text = "-" 'Since this textfield is not compulsory
    txtPhoneMob.Text = "-"  'Since this textfield is not compulsory
    txtPhoneHome.Text = "-" 'Since this textfield is not always compulsory
    
    
    'Enabling the Save Command Button & Close Command Button
    cmdSave.Enabled = True
    cmdClose.Enabled = True
    
    
    'Disabling the Company Related Records
    lblCompanyID.Enabled = False
    txtCompanyID.Enabled = False
    lblCompanyName.Enabled = False
    txtCompanyName.Enabled = False
    cmdCompanySearchWizard.Enabled = False

    
    Call Inpatients_Maintenance    'Calling the Inpatients_Maintenance Procedure to interact with the recordset
    
    'Generate Patient ID By Utilizing the Inpatients_Maintenance Table
    With rsInpatientMaintenance
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strCode = "INP0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Inpatients_Maintenance Table
            If iNumOfRecords < 10 Then
                strCode = "INP000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strCode = "INP00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strCode = "INP0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strCode = "INP" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    'The following line of code will enter the autogenerated Patient ID
    'into the Patient ID textfield
    txtPatientID.Text = strCode
    
End Sub


Private Sub cmdClose_Click()

    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdCompanySearchWizard_Click()
    
    frmCompanySearchWizard.Show
    
End Sub

Private Sub cmdDelete_Click()   'This function will delete a record from the database
    
    'Check for the record selection
    If txtPatientID.Text = "" Then
    
        MsgBox "Error! No Record Has Been Selected", vbCritical, "No Record Selected!"
    
    Else
    
        With rsInpatientMaintenance
        
            'Confirm the Delete procedure with the user
            If MsgBox("Are You Sure You Wish To Delete Patient ID " & txtPatientID.Text & "'s Record?", vbYesNo + vbQuestion, "Delete Record?") = vbYes Then
        
                .Delete 'Delete the record from the database
                
                'Display Success Message
                MsgBox "The Record Has Been Deleted Successfully!", vbInformation, "Successful Delete Procedure!"
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            Else
                
                'Display 'Delete Procedure Cancelled' Message
                MsgBox "The Delete Procedure Was Cancelled!", vbExclamation, "Delete Procedure Cancelled!"
                
                Form_Load   'Calling the Form_Load Procedure

                clearAllFields  'Calling a Private Function To Clear All Fields
        
            End If

            .Requery    'Requerying the Table
        
        End With
        
    End If

End Sub

Private Sub cmdLaunchInpatientSearch_Click() 'This function is fired when the Launch Inpatient Search Wizard Command Button is Clicked. It opens up the Inpatient Search Wizard
    
    
    enableAllFields     'Calling a Private Function To Enable All Fields
    enableAllButtons    'Calling a Private Function To Enable All Command Buttons
    
    cmdSave.Enabled = False     'Disabling the Save Command Button
    
    frmInpatientSearchWizard.Show      'Displays the Doctor Search Wizard
    
    'Disabling the "Step" Buttons
    cmdStep1.Enabled = False
    cmdStep3.Enabled = False
    
End Sub

Private Sub cmdSave_Click()     'This function will save all the user's data in the database
    
    
    'Checking if the Phone Number (Home) textfield and the Phone Number (Mob) textfield are empty
    If txtPhoneHome.Text = "-" And txtPhoneMob.Text = "-" Then
        txtPhoneHome.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtPhoneMob.BackColor = &H80000018 'Highlighting the textfield in a different colour
        MsgBox "Error! Both Phone Number Textfields Cannot Be Empty! At Least One Has To Be Provided!", vbCritical, "Error In Phone Numbers!"
        Exit Sub
    Else
        txtPhoneHome.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtPhoneMob.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
        
        
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        
        'Validation To Ensure That The NIC Number is 10 Characters In Length
        If txtNICNumber.Text <> "-" Then
            If Len(txtNICNumber.Text) <> 10 Then
                MsgBox "Error! The NIC Number Has To Consist Of 10 Characters!", vbCritical, "Error In NIC Number!"
                txtNICNumber.BackColor = &H80000018  'Highlighting the textfield in a different colour
                Exit Sub
            Else
                txtNICNumber.BackColor = &H80000004
            End If
        End If
        
        
        'Validation To Ensure That The Phone Numbers are not Greater than 15 Digits in Length
        If txtPhoneHome.Text <> "-" Then
            If Len(txtPhoneHome.Text) > 15 Then
                MsgBox "Error! The Phone No (Home) Textfield Cannot Consist Of More Than 15 Digits!", vbCritical, "Error In Phone No (Home)!"
                txtPhoneHome.BackColor = &H80000018
                Exit Sub
            Else
                txtPhoneHome.BackColor = &H80000004
            End If
        End If
        
        
        'Validation To Ensure That The Phone Numbers are not Greater than 15 Digits in Length
        If txtPhoneMob.Text <> "-" Then
            If Len(txtPhoneMob.Text) > 15 Then
                MsgBox "Error! The Phone No (Mob) Textfield Cannot Consist Of More Than 15 Digits!", vbCritical, "Error In Phone No (Mob)!"
                txtPhoneMob.BackColor = &H80000018
                Exit Sub
            Else
                txtPhoneMob.BackColor = &H80000004
            End If
        End If
        
        
        'Validation To Ensure That The Patient Occupation is not Greater than 30 Characters in Length
        If Len(txtPatientOccupation.Text) > 30 Then
            MsgBox "Error! The Patient Occupation Textfield Cannot Consist Of More Than 30 Characters", vbCritical, "Error In Patient Occupation!"
            txtPatientOccupation.BackColor = &H80000018
            Exit Sub
        Else
            txtPatientOccupation.BackColor = &H80000004
        End If
        
        
        With rsInpatientMaintenance
            
            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
                
                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtNICNumber.Text = "" Then
                    txtNICNumber.Text = "-"
                End If
                
                If txtPhoneMob.Text = "" Then
                    txtPhoneMob.Text = "-"
                End If
                
                If txtPhoneHome.Text = "" Then
                    txtPhoneHome.Text = "-"
                End If
                
                If txtCompanyID.Text = "" Then
                    txtCompanyID.Text = "-"
                End If
                
                If txtCompanyName.Text = "" Then
                    txtCompanyName.Text = "-"
                End If
                
                strPatientID = txtPatientID.Text
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtPatientID.Text
                .Fields(1) = txtFirstName.Text
                .Fields(2) = txtSurname.Text
                .Fields(3) = cboGender.Text
                .Fields(4) = dtpDateOfBirth.Value
                .Fields(5) = txtNICNumber.Text
                .Fields(6) = txtAddress.Text
                .Fields(7) = txtPhoneHome.Text
                .Fields(8) = txtPhoneMob.Text
                .Fields(9) = txtPatientOccupation.Text
                .Fields(10) = cboCivilStatus.Text
                .Fields(11) = cboAccountType.Text
                .Fields(12) = txtCompanyID.Text
                .Fields(13) = txtCompanyName.Text
            
                .Update
                
                'Display Success Message
                MsgBox "The Record Was Saved Successfully! You Will Now Be Taken To Step 2!", vbInformation, "Succesful Save Procedure!"
                
                
                loadGuardiansMaintenance    'Calling a public function to prepare the Guardians Maintenance form
                
                Unload Me
                
                frmGuardiansMaintenance.Show
            
            Else
            
                'Display 'No Modifications' Message
                MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
                .CancelUpdate   'Cancel the Save Procedure
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            End If
            
            .Requery    'Requerying the Table
            
        End With
        
    End If
        

End Sub

Private Function loadGuardiansMaintenance()
    
    frmGuardiansMaintenance.enableAllFields    'Calling a Private Function To Enable All Fields
    frmGuardiansMaintenance.clearAllFields      'Calling a Private Function To Clear All Fields
    frmGuardiansMaintenance.disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    frmGuardiansMaintenance.txtNICNumber.Text = "-" 'Since this textfield is not compulsory
    frmGuardiansMaintenance.txtPhoneMob.Text = "-"  'Since this textfield is not always compulsory
    frmGuardiansMaintenance.txtPhoneHome.Text = "-" 'Since this textfield is not always compulsory
    
    
    'Enabling the Save Command Button
    frmGuardiansMaintenance.cmdSave.Enabled = True
    
    'Disbaling the Search Frame
    frmGuardiansMaintenance.lblCriteria.Enabled = False
    frmGuardiansMaintenance.cboSearchType.Enabled = False
    frmGuardiansMaintenance.lblSearchText.Enabled = False
    frmGuardiansMaintenance.txtSearch.Enabled = False

    
    Call Guardians_Maintenance    'Calling the Guardians_Maintenance Procedure to interact with the recordset
    
    'Generate Guardian ID By Utilizing the Guardians_Maintenance Table
    With rsGuardiansMaintenance
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strDisplayID = "GRD0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfGuardians = .RecordCount
            iNumOfGuardians = iNumOfGuardians + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Guardians_Maintenance Table
            If iNumOfGuardians < 10 Then
                strDisplayID = "GRD000" & iNumOfGuardians
            ElseIf iNumOfGuardians < 100 Then
                strDisplayID = "GRD00" & iNumOfGuardians
            ElseIf iNumOfGuardians < 1000 Then
                strDisplayID = "GRD0" & iNumOfGuardians
            ElseIf iNumOfGuardians < 10000 Then
                strDisplayID = "GRD" & iNumOfGuardians
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    'The following line of code will enter the autogenerated Guardian ID
    'into the Guardian ID textfield
    frmGuardiansMaintenance.txtGuardianID.Text = strDisplayID
    
    frmGuardiansMaintenance.txtPatientID.Text = strPatientID    'Global Variable
    
    frmGuardiansMaintenance.dgrdGuardiansInfo.Enabled = False
    
End Function


Private Sub cmdStep2_Click()

    Call Guardians_Maintenance
    
    With rsGuardiansMaintenance
    
        .MoveFirst
        
         Do While .EOF = False
         
            If .Fields(1).Value = txtPatientID.Text Then
            
                'Entering the values in the particular record into the fields on the interface
                frmGuardiansMaintenance.txtGuardianID.Text = .Fields(0).Value
                frmGuardiansMaintenance.txtPatientID.Text = .Fields(1).Value
                frmGuardiansMaintenance.txtFirstName.Text = .Fields(2).Value
                frmGuardiansMaintenance.txtSurname.Text = .Fields(3).Value
                frmGuardiansMaintenance.cboGender.Text = .Fields(4).Value
                frmGuardiansMaintenance.txtNICNumber.Text = .Fields(5).Value
                frmGuardiansMaintenance.txtAddress.Text = .Fields(6).Value
                frmGuardiansMaintenance.txtPhoneHome.Text = .Fields(7).Value
                frmGuardiansMaintenance.txtPhoneMob.Text = .Fields(8).Value
                frmGuardiansMaintenance.txtOccupation.Text = .Fields(9).Value
                frmGuardiansMaintenance.txtRelationToPatient.Text = .Fields(10).Value
                Exit Do
                
            Else
            
                .MoveNext
            
            End If
            
        Loop
        
    End With
    
    
    'Enabling / Diabling the Navigation Buttons as necessary
    frmGuardiansMaintenance.cmdFirst.Enabled = False
    frmGuardiansMaintenance.cmdLast.Enabled = True
    frmGuardiansMaintenance.cmdPrevious.Enabled = False
    frmGuardiansMaintenance.cmdNext.Enabled = True

    'Enabling the Update Button
    frmGuardiansMaintenance.cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    frmGuardiansMaintenance.cmdStep1.Enabled = True
    frmGuardiansMaintenance.cmdStep3.Enabled = True
    
    frmGuardiansMaintenance.enableAllFields
    
    Unload Me
    
    frmGuardiansMaintenance.Show
    
    
    
End Sub



Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it


    'Checking if the Phone Number (Home) textfield and the Phone Number (Mob) textfield are empty
    If txtPhoneHome.Text = "-" And txtPhoneMob.Text = "-" Then
        txtPhoneHome.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtPhoneMob.BackColor = &H80000018 'Highlighting the textfield in a different colour
        MsgBox "Error! Both Phone Number Textfields Cannot Be Empty! At Least One Has To Be Provided!", vbCritical, "Error In Phone Numbers!"
        Exit Sub
    Else
        txtPhoneHome.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtPhoneMob.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
        
        
        
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        
        'Validation To Ensure That The NIC Number is 10 Characters In Length
        If txtNICNumber.Text <> "-" Then
            If Len(txtNICNumber.Text) <> 10 Then
                MsgBox "Error! The NIC Number Has To Consist Of 10 Characters!", vbCritical, "Error In NIC Number!"
                txtNICNumber.BackColor = &H80000018  'Highlighting the textfield in a different colour
                Exit Sub
            Else
                txtNICNumber.BackColor = &H80000004
            End If
        End If
        
        
        'Validation To Ensure That The Phone Numbers are not Greater than 15 Digits in Length
        If txtPhoneHome.Text <> "-" Then
            If Len(txtPhoneHome.Text) > 15 Then
                MsgBox "Error! The Phone No (Home) Textfield Cannot Consist Of More Than 15 Digits!", vbCritical, "Error In Phone No (Home)!"
                txtPhoneHome.BackColor = &H80000018
                Exit Sub
            Else
                txtPhoneHome.BackColor = &H80000004
            End If
        End If
        
        
        'Validation To Ensure That The Phone Numbers are not Greater than 15 Digits in Length
        If txtPhoneMob.Text <> "-" Then
            If Len(txtPhoneMob.Text) > 15 Then
                MsgBox "Error! The Phone No (Mob) Textfield Cannot Consist Of More Than 15 Digits!", vbCritical, "Error In Phone No (Mob)!"
                txtPhoneMob.BackColor = &H80000018
                Exit Sub
            Else
                txtPhoneMob.BackColor = &H80000004
            End If
        End If
        
        
        'Validation To Ensure That The Patient Occupation is not Greater than 30 Characters in Length
        If Len(txtPatientOccupation.Text) > 30 Then
            MsgBox "Error! The Patient Occupation Textfield Cannot Consist Of More Than 30 Characters", vbCritical, "Error In Patient Occupation!"
            txtPatientOccupation.BackColor = &H80000018
            Exit Sub
        Else
            txtPatientOccupation.BackColor = &H80000004
        End If
        
        
        With rsInpatientMaintenance
            
            'Making sure that the user wants to update the record
            If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
                
                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtNICNumber.Text = "" Then
                    txtNICNumber.Text = "-"
                End If
                
                If txtPhoneMob.Text = "" Then
                    txtPhoneMob.Text = "-"
                End If
                
                If txtPhoneHome.Text = "" Then
                    txtPhoneHome.Text = "-"
                End If
                
                If txtCompanyID.Text = "" Then
                    txtCompanyID.Text = "-"
                End If
                
                If txtCompanyName.Text = "" Then
                    txtCompanyName.Text = "-"
                End If
                
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtPatientID.Text
                .Fields(1) = txtFirstName.Text
                .Fields(2) = txtSurname.Text
                .Fields(3) = cboGender.Text
                .Fields(4) = dtpDateOfBirth.Value
                .Fields(5) = txtNICNumber.Text
                .Fields(6) = txtAddress.Text
                .Fields(7) = txtPhoneHome.Text
                .Fields(8) = txtPhoneMob.Text
                .Fields(9) = txtPatientOccupation.Text
                .Fields(10) = cboCivilStatus.Text
                .Fields(11) = cboAccountType.Text
                .Fields(12) = txtCompanyID.Text
                .Fields(13) = txtCompanyName.Text
            
                .Update
                
                'Display Success Message
                MsgBox "The Record Was Updated Successfully!", vbInformation, "Succesful Update Procedure"
                
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            Else
            
                'Display 'No Modifications' Message
                MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
                .CancelUpdate   'Cancel the Update Procedure
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            End If
            
            .Requery    'Requerying the Table
            
        End With
        
    End If
    
End Sub

Public Sub Form_Load()

    Call Connection  'Calling the Connection Procedure
        
    disableAllFields  'Calling a Private Function To Disable All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    'Enabling  the First Button and the Last Button
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Add New Button & the Close Button
    cmdAddNew.Enabled = True
    cmdClose.Enabled = True
    
    'Enabling the LaunchDoctorSearch Wizard Button
    cmdLaunchInpatientSearch.Enabled = True
    
    
End Sub

Private Function disableAllFields() 'This function will disable all fields on the interface

    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will disable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Then
        eachField.Enabled = False
    End If

    Next
    
    dtpDateOfBirth.Enabled = False  'Disabling the Date Of Birth Date Time Picker
    

End Function



Public Function enableAllFields() 'This function will enable all fields on the interface


    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will enable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Then
        eachField.Enabled = True
    End If

    Next
    
    dtpDateOfBirth.Enabled = True   'Enabling the Date Of Birth Date Time Picker

End Function


Private Function disableAllButtons() 'This function will disable all command buttons on the interface

    On Error Resume Next
    For Each eachButton In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will disable all Command Buttons
    If TypeOf eachButton Is CommandButton Then
        eachButton.Enabled = False
    End If

    Next

End Function



Private Function enableAllButtons() 'This function will enable all command buttons on the interface


    On Error Resume Next
    For Each eachButton In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will enable all Command Buttons
    If TypeOf eachButton Is CommandButton Then
        eachButton.Enabled = True
    End If

    Next
    
    'Disabling the Step 1 Button
    cmdStep1.Enabled = False

End Function


Public Function clearAllFields() 'This function will clear all fields on the interface


    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will clear all TextBoxes
    If TypeOf eachField Is TextBox Then
        eachField.Text = ""
    End If

    Next
    
    'The following lines will set the normal display values of the Gender
    'ComboBox, Doctor Category ComboBox and the Date Of Birth Date Time Picker
    cboGender.Text = "----------SELECT-----------"
    cboCivilStatus.Text = "----------SELECT-----------"
    cboAccountType.Text = "----------SELECT-----------"
    dtpDateOfBirth.Value = "4/14/2008"
    
End Function


Private Sub cmdFirst_Click()  'This function will Navigate to the First Record

    'Enabling / Diabling the Navigation Buttons as necessary
    cmdFirst.Enabled = False
    cmdLast.Enabled = True
    cmdPrevious.Enabled = False
    cmdNext.Enabled = True
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the "Step" Buttons
    cmdStep2.Enabled = True

    
    'Enabling the Company Search Wizard
    cmdCompanySearchWizard.Enabled = True
    
    Call Inpatients_Maintenance  'Calling the Inpatients_Maintenance Procedure to interact with the recordset
    
    With rsInpatientMaintenance
    
    
        .MoveFirst  'Moving to the first record
        
        'Entering the values in the particular record into the fields on the interface
        txtPatientID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtSurname.Text = .Fields(2).Value
        cboGender.Text = .Fields(3).Value
        dtpDateOfBirth.Value = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtPatientOccupation.Text = .Fields(9).Value
        cboCivilStatus.Text = .Fields(10).Value
        cboAccountType.Text = .Fields(11).Value
        txtCompanyID.Text = .Fields(12).Value
        txtCompanyName.Text = .Fields(13).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub


Private Sub cmdPrevious_Click() 'This function will Navigate to the Previous Record
    
    With rsInpatientMaintenance
    
        
        .MovePrevious   'Moving to the previous record
        
        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If
    
        'Entering the values in the particular record into the fields on the interface
        txtPatientID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtSurname.Text = .Fields(2).Value
        cboGender.Text = .Fields(3).Value
        dtpDateOfBirth.Value = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtPatientOccupation.Text = .Fields(9).Value
        cboCivilStatus.Text = .Fields(10).Value
        cboAccountType.Text = .Fields(11).Value
        txtCompanyID.Text = .Fields(12).Value
        txtCompanyName.Text = .Fields(13).Value
        
    End With
    
    cmdNext.Enabled = True  'Enabling the Next Button
    cmdLast.Enabled = True  'Enabling the Last Button
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the "Step" Buttons
    cmdStep2.Enabled = True

    
    'Enabling the Company Search Wizard
    cmdCompanySearchWizard.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub


Private Sub cmdNext_Click() 'This function will Navigate to the Next Record
    
    With rsInpatientMaintenance
    
        .MoveNext   'Moving to the Next Record
        
        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If
        
        'Entering the values in the particular record into the fields on the interface
        txtPatientID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtSurname.Text = .Fields(2).Value
        cboGender.Text = .Fields(3).Value
        dtpDateOfBirth.Value = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtPatientOccupation.Text = .Fields(9).Value
        cboCivilStatus.Text = .Fields(10).Value
        cboAccountType.Text = .Fields(11).Value
        txtCompanyID.Text = .Fields(12).Value
        txtCompanyName.Text = .Fields(13).Value
        
    End With
    
    cmdPrevious.Enabled = True  'Enabling the Previous Button
    cmdFirst.Enabled = True 'Enabling the First Button
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the "Step" Buttons
    cmdStep2.Enabled = True

    
    'Enabling the Company Search Wizard
    cmdCompanySearchWizard.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub


Private Sub cmdLast_Click() 'This function will Navigate to the Last Record
    
    'Enabling / Diabling the Navigation Buttons as necessary
    cmdLast.Enabled = False
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = False
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the "Step" Buttons
    cmdStep2.Enabled = True

    
    'Enabling the Company Search Wizard
    cmdCompanySearchWizard.Enabled = True
    
    Call Inpatients_Maintenance  'Calling the Inpatients_Maintenance Procedure to interact with the recordset
    
    With rsInpatientMaintenance
    
        .Requery
    
        .MoveLast   'Moving to the last record
        
        'Entering the values in the particular record into the fields on the interface
        txtPatientID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtSurname.Text = .Fields(2).Value
        cboGender.Text = .Fields(3).Value
        dtpDateOfBirth.Value = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtPatientOccupation.Text = .Fields(9).Value
        cboCivilStatus.Text = .Fields(10).Value
        cboAccountType.Text = .Fields(11).Value
        txtCompanyID.Text = .Fields(12).Value
        txtCompanyName.Text = .Fields(13).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub



Private Function textfieldsValidations() As Boolean  'This function will validate all fields
    
    Flag = True 'Setting the Flag variable to True
    dateFlag = True 'Setting the dateFlag variable to True
    
    'Checking if the First Name textfield is empty
    If txtFirstName.Text = "" Then
        txtFirstName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtFirstName.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Surname textfield is empty
    If txtSurname.Text = "" Then
        txtSurname.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtSurname.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the user has made a selection in the Gender ComboBox
    If cboGender.Text = "" Then
        cboGender.BackColor = &H80000018    'Highlighting the ComboBox in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboGender.BackColor = &H80000004    'Bringing the ComboBox BackColour back to normal
    End If
    
    'Checking if the Date Of Birth is valid
    If dtpDateOfBirth.Value = "4/14/2008" Then
        'Displaying an error message, asking the user to alter the date accordingly
        MsgBox "The Date You Have Provided Is Incorrect! Please Check Your Date!", vbCritical, "Incorrect Date"
        dateFlag = False    'Setting the dateFlag variable to False to indicate invalid data
    End If
    
    'Checking if the Address textfield is empty
    If txtAddress.Text = "" Then
        txtAddress.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtAddress.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Patient Occupation textfield is empty
    If txtPatientOccupation.Text = "" Then
        txtPatientOccupation.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtPatientOccupation.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the user has made a selection in the Civil Status  ComboBox
    If cboCivilStatus.Text = "" Then
        cboCivilStatus.BackColor = &H80000018    'Highlighting the ComboBox in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboCivilStatus.BackColor = &H80000004    'Bringing the ComboBox BackColour back to normal
    End If
    
    'Checking if the user has made a selection in the Account Type ComboBox
    If cboAccountType.Text = "" Then
        cboAccountType.BackColor = &H80000018    'Highlighting the ComboBox in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboAccountType.BackColor = &H80000004    'Bringing the ComboBox BackColour back to normal
    End If
    
    'If the user chooses 'Corporate' from the Account Type ComboBox
    If cboAccountType.ListIndex = 0 Then
        'Checking if the Channeling Charges textfield is empty
        If txtCompanyID.Text = "" Then
            txtCompanyID.BackColor = &H80000018 'Highlighting the textfield in a different colour
            txtCompanyName.BackColor = &H80000018 'Highlighting the textfield in a different colour
            Flag = False    'Setting the Flag variable to False to indicate invalid data
        Else
            txtCompanyID.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
            txtCompanyName.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        End If
    End If
    
    
    'Here, I am checking the state of the Flag variable and if it is False, I am displaying a
    'Message Box to instruct the user to enter data into all highlighted textfields.
    'The Save procedure will also be cancelled
    If Flag = False Then
        MsgBox "Error! Please Fill-in The Highlighted Textfields! They Are Compulsory!", vbCritical, "Please Fill Highlighted Textfields"
        textfieldsValidations = True    'Passing values to the Save procedure
    ElseIf dateFlag = False Then
        textfieldsValidations = True    'Passing values to the Save procedure
    Else
        textfieldsValidations = False   'Passing values to the Save procedure
    End If
    
End Function



Private Sub tmrErrMsg_Timer()

    Static i As Integer
    
    If i < 200000 Then     'Validation Msg Viewing Time Period
        picInvalidDataMsg.Visible = False
        picInvalidKeypressMsg.Visible = False
        picInvalidKeyMsg.Visible = False
        picInvalidTypingMsg.Visible = False
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If
    
End Sub


Private Sub txtPhoneHome_GotFocus() 'This procedure will ensure that the textfield is empty when the user types in it.
    
    If txtPhoneHome.Text = "-" Then
        txtPhoneHome.Text = ""
    End If
    
End Sub

Private Sub txtPhoneHome_LostFocus()    'This procedure will ensure that the textfield is not empty when the user is not typing in it.
    
    If txtPhoneHome.Text = "" Then
        txtPhoneHome.Text = "-"
    End If
    
End Sub


Private Sub txtPhoneMob_GotFocus()  'This procedure will ensure that the textfield is empty when the user types in it.
    
    If txtPhoneMob.Text = "-" Then
        txtPhoneMob.Text = ""
    End If
    
End Sub

Private Sub txtPhoneMob_LostFocus() 'This procedure will ensure that the textfield is not empty when the user is not typing in it.
    
    If txtPhoneMob.Text = "" Then
        txtPhoneMob.Text = "-"
    End If
    
End Sub

'This procedure will ensure that the textfield is empty when the user types in it.
Private Sub txtNICNumber_GotFocus()
    
    If txtNICNumber.Text = "-" Then
        txtNICNumber.Text = ""
    End If
    
End Sub

Private Sub txtNICNumber_LostFocus()    'This procedure will ensure that the textfield is not empty when the user is not typing in it.

    If txtNICNumber.Text = "" Then
        txtNICNumber.Text = "-"
    End If
    
End Sub

Private Sub txtPhoneHome_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 3360    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtPhoneMob_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 3840    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtNICNumber_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = Asc("X") Then
    ElseIf KeyAscii = Asc("x") Then
    ElseIf KeyAscii = Asc("V") Then
    ElseIf KeyAscii = Asc("v") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidTypingMsg.Top = 5760    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtFirstName_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets
    
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 3840    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtSurname_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets
    
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 4320    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtPatientOccupation_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets
    
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeyMsg.Top = 4320    'Validation Note View
        picInvalidKeyMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub
