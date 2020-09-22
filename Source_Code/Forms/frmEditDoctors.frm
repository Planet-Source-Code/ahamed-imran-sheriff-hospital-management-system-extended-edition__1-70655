VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDoctorsMaintenance 
   BackColor       =   &H80000004&
   Caption         =   "Doctor's Maintenance Module"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditDoctors.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInvalidKeyMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   9120
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   52
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.ComboBox cboAppointmentDuration 
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
      ItemData        =   "frmEditDoctors.frx":1EEB0
      Left            =   8400
      List            =   "frmEditDoctors.frx":1EEC6
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txtServiceCharges 
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
      MaxLength       =   4
      TabIndex        =   48
      Text            =   "-"
      Top             =   3960
      Width           =   2295
   End
   Begin VB.PictureBox picInvalidKeypressMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3360
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   46
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   120
      Top             =   4440
   End
   Begin VB.PictureBox picInvalidDataMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3360
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   44
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Digits Here! Only Alphabets Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdLaunchDocSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Launch Doctor-Search Wizard"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click here to launch the Search Wizard"
      Top             =   1800
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker dtpDateOfBirth 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   4920
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
      Format          =   56623105
      CurrentDate     =   39552
   End
   Begin VB.TextBox txtDoctorID 
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
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtLicenseNo 
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
      Left            =   2640
      TabIndex        =   10
      Text            =   "E"
      Top             =   7920
      Width           =   2295
   End
   Begin VB.TextBox txtMobPhone 
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
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   9
      Text            =   "-"
      Top             =   7440
      Width           =   2295
   End
   Begin VB.TextBox txtHomePhone 
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
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   8
      Top             =   6960
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
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5880
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
      ItemData        =   "frmEditDoctors.frx":1EEE2
      Left            =   2640
      List            =   "frmEditDoctors.frx":1EEEC
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4440
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
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "-"
      Top             =   5400
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
      Left            =   2640
      TabIndex        =   3
      Top             =   3960
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
      Left            =   2640
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdSetUpDocSchedule 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Set Up Doctor's Visiting Days Schedule"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Click here to set up a Doctor's Visiting Schedule"
      Top             =   5880
      Width           =   4095
   End
   Begin VB.CommandButton cmdSpecializationWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   10320
      TabIndex        =   12
      ToolTipText     =   "Click Here to select a Specialization"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txtDoctorSpecialization 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmEditDoctors.frx":1EEFE
      Height          =   855
      Left            =   8040
      Picture         =   "frmEditDoctors.frx":1F3E4
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmEditDoctors.frx":22128
      Height          =   855
      Left            =   6960
      Picture         =   "frmEditDoctors.frx":225A6
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      DisabledPicture =   "frmEditDoctors.frx":252EA
      Height          =   855
      Left            =   5880
      Picture         =   "frmEditDoctors.frx":256EC
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmEditDoctors.frx":28430
      Height          =   855
      Left            =   9120
      Picture         =   "frmEditDoctors.frx":288F9
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmEditDoctors.frx":2B63D
      Height          =   855
      Left            =   10200
      Picture         =   "frmEditDoctors.frx":2BAFC
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmEditDoctors.frx":2E840
      Height          =   750
      Left            =   7680
      Picture         =   "frmEditDoctors.frx":2EC55
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmEditDoctors.frx":30E11
      Height          =   750
      Left            =   6720
      Picture         =   "frmEditDoctors.frx":311ED
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmEditDoctors.frx":333A9
      Height          =   750
      Left            =   8640
      Picture         =   "frmEditDoctors.frx":3377F
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
      Width           =   890
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmEditDoctors.frx":3593B
      Height          =   750
      Left            =   9600
      Picture         =   "frmEditDoctors.frx":35D15
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6600
      Width           =   890
   End
   Begin VB.TextBox txtChannelingCharges 
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
      MaxLength       =   4
      TabIndex        =   14
      Text            =   "-"
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtReferringCharges 
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
      MaxLength       =   4
      TabIndex        =   15
      Text            =   "-"
      Top             =   5400
      Width           =   2295
   End
   Begin VB.ComboBox cboDoctorCategory 
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
      ItemData        =   "frmEditDoctors.frx":37ED1
      Left            =   8400
      List            =   "frmEditDoctors.frx":37EDE
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   15000
      TabIndex        =   26
      Top             =   12120
      Width           =   2295
   End
   Begin VB.Label lblAppointmentDuration 
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment Duration (Minutes)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   51
      Top             =   4965
      Width           =   1575
   End
   Begin VB.Label lblServiceCharges 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charges (Per Day)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   49
      Top             =   3960
      Width           =   1575
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
      Left            =   480
      TabIndex        =   43
      Top             =   2280
      Width           =   7935
   End
   Begin VB.Label lblDoctorSpecialization 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Specialization"
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
      Top             =   3045
      Width           =   1935
   End
   Begin VB.Label lblDoctorID 
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
      Left            =   960
      TabIndex        =   41
      Top             =   3045
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   5760
      Top             =   7680
      Width           =   5535
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   5760
      Top             =   6480
      Width           =   5535
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
      Left            =   840
      TabIndex        =   40
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   720
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2640
      X2              =   5400
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   5400
      X2              =   5400
      Y1              =   2760
      Y2              =   8760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   5400
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   480
      Y1              =   2760
      Y2              =   8760
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
      Left            =   960
      TabIndex        =   39
      Top             =   3525
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
      Left            =   960
      TabIndex        =   38
      Top             =   4005
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
      Left            =   960
      TabIndex        =   37
      Top             =   4485
      Width           =   1335
   End
   Begin VB.Label lblDateOFBirth 
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
      Left            =   960
      TabIndex        =   36
      Top             =   4965
      Width           =   1335
   End
   Begin VB.Label lblNICNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "*NIC Number"
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
      Left            =   960
      TabIndex        =   35
      Top             =   5445
      Width           =   1335
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
      Left            =   960
      TabIndex        =   34
      Top             =   5925
      Width           =   1335
   End
   Begin VB.Label lblHomePhone 
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
      Left            =   960
      TabIndex        =   33
      Top             =   7005
      Width           =   1695
   End
   Begin VB.Label lblMobPhone 
      BackStyle       =   0  'Transparent
      Caption         =   "*Phone No. (Mob)"
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
      Left            =   960
      TabIndex        =   32
      Top             =   7485
      Width           =   1695
   End
   Begin VB.Label lblLicenseNo 
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
      Left            =   960
      TabIndex        =   31
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label lblEmployeeDetails 
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
      Left            =   6240
      TabIndex        =   30
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   5760
      X2              =   6120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   11280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   11280
      X2              =   11280
      Y1              =   2760
      Y2              =   6360
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   5760
      X2              =   11280
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   5760
      X2              =   5760
      Y1              =   2760
      Y2              =   6360
   End
   Begin VB.Label lblChannelingCharges 
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
      Left            =   6360
      TabIndex        =   29
      Top             =   4485
      Width           =   1935
   End
   Begin VB.Label lblReferringCharges 
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
      Left            =   6360
      TabIndex        =   28
      Top             =   5445
      Width           =   1815
   End
   Begin VB.Label lblDoctorCategory 
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
      Left            =   6360
      TabIndex        =   27
      Top             =   3525
      Width           =   1575
   End
End
Attribute VB_Name = "frmDoctorsMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Doctors Maintenance Interface
'Programmer: Isham Sally
'Quality Assurance Engineer (Testing):  Imran Sheriff
'Start Date: 14/04/08
'Date Of Last Modification: 14/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Doctors_Maintenance Table
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



Private Sub cboDoctorCategory_Click()   'This function will manipulate other controls according to the type of doctor
    
    'The following block of code will disable the Referring Charges textfield
    'if the Doctor is a Permanent Doctor
    If cboDoctorCategory.ListIndex = 0 Then
        txtReferringCharges.Text = "-"
        txtReferringCharges.Enabled = False
        lblReferringCharges.Enabled = False
    Else
        txtReferringCharges.Enabled = True
        lblReferringCharges.Enabled = True
    End If
    
    
    'The following block of code will enable the Set Up Doctor Schedule
    'Command Button if the doctor is a Visiting Doctor
    If cboDoctorCategory.ListIndex = 1 Then
        cmdSetUpDocSchedule.Enabled = True
    Else
        cmdSetUpDocSchedule.Enabled = False
    End If
    
    
    'The following block of code will disable the Channeling Charges textfield
    'if the doctor is a Referring Doctor
    If cboDoctorCategory.ListIndex = 2 Then
        lblServiceCharges.Enabled = False
        txtServiceCharges.Enabled = False
        txtChannelingCharges.Text = "-"
        txtChannelingCharges.Enabled = False
        lblChannelingCharges.Enabled = False
        lblAppointmentDuration.Enabled = False
        cboAppointmentDuration.Enabled = False
    Else
        lblServiceCharges.Enabled = True
        txtServiceCharges.Enabled = True
        txtChannelingCharges.Enabled = True
        lblChannelingCharges.Enabled = True
        lblAppointmentDuration.Enabled = True
        cboAppointmentDuration.Enabled = True
    End If
    
    
End Sub

Private Sub cmdAddNew_Click() 'This function adds a new recordset into the database

    enableAllFields     'Calling a Private Function To Enable All Fields
    clearAllFields      'Calling a Private Function To Clear All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    txtLicenseNo.Text = "E" 'Since all doctor's License Numbers start with E
    txtNICNumber.Text = "-" 'Since this textfield is not compulsory
    txtMobPhone.Text = "-"  'Since this textfield is not compulsory
    txtServiceCharges.Text = "-"    'Since I do not want to store a null value in the database
    txtChannelingCharges.Text = "-" 'Since I do not want to store a null value in the database
    txtReferringCharges.Text = "-"  'Since I do not want to store a null value in the database
    
    'Enabling the Save Command Button & Close Command Button
    cmdSave.Enabled = True
    cmdClose.Enabled = True
    
    'Enabling the Specialization Wizard Button
    cmdSpecializationWizard.Enabled = True
    
    Call Doctors_Maintenance    'Calling the Doctors_Maintenance Procedure to interact with the recordset
    
    'Generate Doctor ID By Utilizing the Doctors_Maintenance Table
    With rsDoctorsMaintenance
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strCode = "DOC0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Doctors_Maintenance Table
            If iNumOfRecords < 10 Then
                strCode = "DOC000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strCode = "DOC00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strCode = "DOC0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strCode = "DOC" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    'The following line of code will enter the autogenerated Doctor ID
    'into the Doctor ID textfield
    txtDoctorID.Text = strCode
    
End Sub


Private Sub cmdClose_Click()

    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdDelete_Click()   'This function will delete a record from the database
    
    'Check for the record selection
    If txtDoctorID.Text = "" Then
    
        MsgBox "Error! No Record Has Been Selected", vbCritical, "No Record Selected!"
    
    Else
    
        With rsDoctorsMaintenance
        
            'Confirm the Delete procedure with the user
            If MsgBox("Are You Sure You Wish To Delete Doctor " & txtSurname.Text & "'s Record?", vbYesNo + vbQuestion, "Delete Record?") = vbYes Then
        
                .Delete 'Delete the record from the database
                
                'Display Success Message
                MsgBox "The Record Has Been Deleted Successfully!", vbInformation, "Successful Delete Procedure!"
                
                Call VisitTimes_Schedule
    
                With rsVisitTimesSchedule
    
                    .MoveFirst
    
                    While .EOF = False
        
                        If txtDoctorID.Text = .Fields(1).Value Then
                
                            .Delete
            
                        End If
            
                        .MoveNext
            
                    Wend
        
                    .Close
        
                End With
                
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

Private Sub cmdLaunchDocSearch_Click() 'This function is fired when the Launch Doctor-Search Wizard Command Button is Clicked. It opens up the Doctor Search Wizard
    
    
    enableAllFields     'Calling a Private Function To Enable All Fields
    enableAllButtons    'Calling a Private Function To Enable All Command Buttons
    
    cmdSave.Enabled = False     'Disabling the Save Command Button
    
    frmDoctorSearchWizard.Show      'Displays the Doctor Search Wizard
    
End Sub

Private Sub cmdSave_Click()     'This function will save all the user's data in the database
    
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
        
        
        'Validation To Ensure That The Channeling Charges Are Not Geater Than 3000
        If txtChannelingCharges.Text <> "-" Then
            If Val(txtChannelingCharges.Text) > 3000 Then
                MsgBox "Error! Channeling Charges Cannot Be Greater Than 3000!", vbCritical, "Error In Channeling Charges!"
                txtChannelingCharges.BackColor = &H80000018
                Exit Sub
            Else
                txtChannelingCharges.BackColor = &H80000004
            End If
        End If
        
        
        'Validation To Ensure That The Referring Charges Are Not Geater Than 1000
        If txtReferringCharges.Text <> "-" Then
            If Val(txtReferringCharges.Text) > 1000 Then
                MsgBox "Error! Referring Charges Cannot Be Greater Than 1000!", vbCritical, "Error In Referring Charges!"
                txtReferringCharges.BackColor = &H80000018
                Exit Sub
            Else
                txtReferringCharges.BackColor = &H80000004
            End If
        End If
        
        
        
        With rsDoctorsMaintenance
            
            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
                
                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtNICNumber.Text = "" Then
                    txtNICNumber.Text = "-"
                End If
                
                If txtMobPhone.Text = "" Then
                    txtMobPhone.Text = "-"
                End If
                
                If txtChannelingCharges.Text = "" Then
                    txtChannelingCharges.Text = "-"
                End If
                
                If txtReferringCharges.Text = "" Then
                    txtReferringCharges.Text = "-"
                End If
                
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtDoctorID.Text
                .Fields(1) = txtFirstName.Text
                .Fields(2) = txtSurname.Text
                .Fields(3) = cboGender.Text
                .Fields(4) = dtpDateOfBirth.Value
                .Fields(5) = txtNICNumber.Text
                .Fields(6) = txtAddress.Text
                .Fields(7) = txtHomePhone.Text
                .Fields(8) = txtMobPhone.Text
                .Fields(9) = txtLicenseNo.Text
                .Fields(10) = txtDoctorSpecialization.Text
                .Fields(11) = cboDoctorCategory.Text
                .Fields(12) = txtServiceCharges.Text
                .Fields(13) = txtChannelingCharges.Text
                .Fields(14) = cboAppointmentDuration.Text
                .Fields(15) = txtReferringCharges.Text
            
                .Update
                
                'Display Success Message
                MsgBox "The Record Was Saved Successfully!", vbInformation, "Succesful Save Procedure"
                
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
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

Private Sub cmdSetUpDocSchedule_Click()
    
    'Opens up the Wizard to set up the Doctor's Visiting Days Schedule.
    frmDoctorVisitingDaysWizard.Show
    
End Sub


Private Sub cmdSpecializationWizard_Click()
    frmDoctorSpecializationWizard.Show
End Sub

Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it

    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        'Here, i am ensuring that whilst editing, the user cannot enter
        'Referring Charges for a Permanent Doctor and cannot enter
        'Channeling Charges for a Referring Doctor
        If cboDoctorCategory.Text = "Permanent" Then
            txtReferringCharges.Text = "-"
        ElseIf cboDoctorCategory.Text = "Referring" Then
            txtChannelingCharges.Text = "-"
        End If
    
    
        'Validation To Ensure That The NIC Number is 10 Characters In Length
        If txtNICNumber.Text <> "-" Then
            If Len(txtNICNumber.Text) <> 10 Then
                MsgBox "Error! The NIC Number Has To Consist Of 10 Characters!", vbCritical, "Error In NIC Number!"
                Exit Sub
            End If
        End If
        
        
        'Validation To Ensure That The Phone Numbers are not Greater than 15 Digits in Length
        If Len(txtHomePhone.Text) > 15 Then
            MsgBox "Error! The Phone No (Home) Textfield Cannot Consist Of More Than 15 Digits!", vbCritical, "Error In Phone No (Home)!"
            Exit Sub
        End If
        
        
        'Validation To Ensure That The Phone Numbers are not Greater than 15 Digits in Length
        If txtMobPhone.Text <> "-" Then
            If Len(txtMobPhone.Text) > 15 Then
                MsgBox "Error! The Phone No (Mob) Textfield Cannot Consist Of More Than 15 Digits!", vbCritical, "Error In Phone No (Mob)!"
                Exit Sub
            End If
        End If
        
        
        'Validation To Ensure That The Channeling Charges Are Not Geater Than 3000
        If txtChannelingCharges.Text <> "-" Then
            If Val(txtChannelingCharges.Text) > 3000 Then
                MsgBox "Error! Channeling Charges Cannot Be Greater Than 3000!", vbCritical, "Error In Channeling Charges!"
                Exit Sub
            End If
        End If
        
        
        'Validation To Ensure That The Referring Charges Are Not Geater Than 1000
        If txtReferringCharges.Text <> "-" Then
            If Val(txtReferringCharges.Text) > 1000 Then
                MsgBox "Error! Referring Charges Cannot Be Greater Than 1000!", vbCritical, "Error In Referring Charges!"
                Exit Sub
            End If
        End If
        
        
        With rsDoctorsMaintenance
            
            'Making sure that the user wants to update the record
            If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
                
                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtNICNumber.Text = "" Then
                    txtNICNumber.Text = "-"
                End If
                
                If txtMobPhone.Text = "" Then
                    txtMobPhone.Text = "-"
                End If
                
                If txtChannelingCharges.Text = "" Then
                    txtChannelingCharges.Text = "-"
                End If
                
                If txtReferringCharges.Text = "" Then
                    txtReferringCharges.Text = "-"
                End If
                
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtDoctorID.Text
                .Fields(1) = txtFirstName.Text
                .Fields(2) = txtSurname.Text
                .Fields(3) = cboGender.Text
                .Fields(4) = dtpDateOfBirth.Value
                .Fields(5) = txtNICNumber.Text
                .Fields(6) = txtAddress.Text
                .Fields(7) = txtHomePhone.Text
                .Fields(8) = txtMobPhone.Text
                .Fields(9) = txtLicenseNo.Text
                .Fields(10) = txtDoctorSpecialization.Text
                .Fields(11) = cboDoctorCategory.Text
                .Fields(12) = txtServiceCharges.Text
                .Fields(13) = txtChannelingCharges.Text
                .Fields(14) = cboAppointmentDuration.Text
                .Fields(15) = txtReferringCharges.Text
            
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



Private Sub dtpDateOfBirth_CloseUp()
    
    dtpDateOfBirth.MaxDate = DateTime.Date

End Sub

Public Sub Form_Load()

    Call Connection  'Calling the Connection Procedure
    
    Call Doctors_Maintenance  'Calling the Doctors_Maintenance Procedure to interact with the recordset
    
    disableAllFields  'Calling a Private Function To Disable All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    'Enabling  the First Button and the Last Button
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Add New Button & the Close Button
    cmdAddNew.Enabled = True
    cmdClose.Enabled = True
    
    'Enabling the LaunchDoctorSearch Wizard Button
    cmdLaunchDocSearch.Enabled = True
    
    
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
    
    cmdSpecializationWizard.Enabled = False  'Disabling the Specialization Wizard Button

End Function



Private Function enableAllFields() 'This function will enable all fields on the interface


    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will enable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Then
        eachField.Enabled = True
    End If

    Next
    
    dtpDateOfBirth.Enabled = True   'Enabling the Date Of Birth Date Time Picker
    cmdSpecializationWizard.Enabled = True

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

End Function


Public Function clearAllFields() 'This function will clear all fields on the interface


    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will clear all TextBoxes
    If TypeOf eachField Is TextBox Then
        eachField.Text = ""
    End If

    Next
    
    'The following lines will set the normal display values of the
    'Date Of Birth Date Time Picker
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
    
    Call Doctors_Maintenance  'Calling the Doctors_Maintenance Procedure to interact with the recordset
    
    With rsDoctorsMaintenance
    
    
        .MoveFirst  'Moving to the first record
        
        'Entering the values in the particular record into the fields on the interface
        txtDoctorID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtSurname.Text = .Fields(2).Value
        cboGender.Text = .Fields(3).Value
        dtpDateOfBirth.Value = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtHomePhone.Text = .Fields(7).Value
        txtMobPhone.Text = .Fields(8).Value
        txtLicenseNo.Text = .Fields(9).Value
        txtDoctorSpecialization.Text = .Fields(10).Value
        cboDoctorCategory.Text = .Fields(11).Value
        txtServiceCharges.Text = .Fields(12).Value
        txtChannelingCharges.Text = .Fields(13).Value
        cboAppointmentDuration.Text = .Fields(14).Value
        txtReferringCharges.Text = .Fields(15).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
    'Here, I am enabling the SetUpDoctor'sVisitingDays Button only if the Doctor Type is Visiting
    If cboDoctorCategory.Text = "Visiting" Then
        cmdSetUpDocSchedule.Enabled = True
    Else
        cmdSetUpDocSchedule.Enabled = False
    End If
    
    disableIfReferringDoctor    'Calling a function to disable certain components if the doctor is a "Referring" doctor
    
    disableIfPermanentDoctor    'Calling a function to disable the Referring Charges textfield if the Doctor is a "Permanent" doctor
    
End Sub

Public Function disableIfReferringDoctor()


    'Here, I am disabling certain components if the doctor is a "Referring Doctor"
    If cboDoctorCategory.Text = "Referring" Then
    
        lblServiceCharges.Enabled = False
        txtServiceCharges.Enabled = False
        lblChannelingCharges.Enabled = False
        txtChannelingCharges.Enabled = False
        lblAppointmentDuration.Enabled = False
        cboAppointmentDuration.Enabled = False
        
    Else
        
        lblServiceCharges.Enabled = True
        txtServiceCharges.Enabled = True
        lblChannelingCharges.Enabled = True
        txtChannelingCharges.Enabled = True
        lblAppointmentDuration.Enabled = True
        cboAppointmentDuration.Enabled = True
        
    End If
    
    
End Function

Public Function disableIfPermanentDoctor()
    
    
    'Here, I am disabling the Referring Charges textfield if the doctor is a Permanent Doctor
    If cboDoctorCategory.Text = "Permanent" Then
        
        lblReferringCharges.Enabled = False
        txtReferringCharges.Enabled = False
        
    Else
    
        lblReferringCharges.Enabled = True
        txtReferringCharges.Enabled = True
        
    End If
    
    
End Function

Private Sub cmdPrevious_Click() 'This function will Navigate to the Previous Record
    
    With rsDoctorsMaintenance
    
        
        .MovePrevious   'Moving to the previous record
        
        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If
    
        'Entering the values in the particular record into the fields on the interface
        txtDoctorID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtSurname.Text = .Fields(2).Value
        cboGender.Text = .Fields(3).Value
        dtpDateOfBirth.Value = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtHomePhone.Text = .Fields(7).Value
        txtMobPhone.Text = .Fields(8).Value
        txtLicenseNo.Text = .Fields(9).Value
        txtDoctorSpecialization.Text = .Fields(10).Value
        cboDoctorCategory.Text = .Fields(11).Value
        txtServiceCharges.Text = .Fields(12).Value
        txtChannelingCharges.Text = .Fields(13).Value
        cboAppointmentDuration.Text = .Fields(14).Value
        txtReferringCharges.Text = .Fields(15).Value
        
    End With
    
    cmdNext.Enabled = True  'Enabling the Next Button
    cmdLast.Enabled = True  'Enabling the Last Button
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
    'Here, I am enabling the SetUpDoctor'sVisitingDays Button only if the Doctor Type is Visiting
    If cboDoctorCategory.Text = "Visiting" Then
        cmdSetUpDocSchedule.Enabled = True
    Else
        cmdSetUpDocSchedule.Enabled = False
    End If
    
    disableIfReferringDoctor    'Calling a function to disable certain components if the doctor is a "Referring" doctor
    
    disableIfPermanentDoctor    'Calling a function to disable the Referring Charges textfield if the Doctor is a "Permanent" doctor
    
End Sub


Private Sub cmdNext_Click() 'This function will Navigate to the Next Record
    
    With rsDoctorsMaintenance
    
        .MoveNext   'Moving to the Next Record
        
        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If
        
        'Entering the values in the particular record into the fields on the interface
        txtDoctorID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtSurname.Text = .Fields(2).Value
        cboGender.Text = .Fields(3).Value
        dtpDateOfBirth.Value = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtHomePhone.Text = .Fields(7).Value
        txtMobPhone.Text = .Fields(8).Value
        txtLicenseNo.Text = .Fields(9).Value
        txtDoctorSpecialization.Text = .Fields(10).Value
        cboDoctorCategory.Text = .Fields(11).Value
        txtServiceCharges.Text = .Fields(12).Value
        txtChannelingCharges.Text = .Fields(13).Value
        cboAppointmentDuration.Text = .Fields(14).Value
        txtReferringCharges.Text = .Fields(15).Value
        
    End With
    
    cmdPrevious.Enabled = True  'Enabling the Previous Button
    cmdFirst.Enabled = True 'Enabling the First Button
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
    'Here, I am enabling the SetUpDoctor'sVisitingDays Button only if the Doctor Type is Visiting
    If cboDoctorCategory.Text = "Visiting" Then
        cmdSetUpDocSchedule.Enabled = True
    Else
        cmdSetUpDocSchedule.Enabled = False
    End If
    
    disableIfReferringDoctor    'Calling a function to disable certain components if the doctor is a "Referring" doctor
    
    disableIfPermanentDoctor    'Calling a function to disable the Referring Charges textfield if the Doctor is a "Permanent" doctor
    
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
    
    Call Doctors_Maintenance  'Calling the Doctors_Maintenance Procedure to interact with the recordset
    
    With rsDoctorsMaintenance
    
        .Requery
    
        .MoveLast   'Moving to the last record
        
        'Entering the values in the particular record into the fields on the interface
        txtDoctorID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtSurname.Text = .Fields(2).Value
        cboGender.Text = .Fields(3).Value
        dtpDateOfBirth.Value = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtHomePhone.Text = .Fields(7).Value
        txtMobPhone.Text = .Fields(8).Value
        txtLicenseNo.Text = .Fields(9).Value
        txtDoctorSpecialization.Text = .Fields(10).Value
        cboDoctorCategory.Text = .Fields(11).Value
        txtServiceCharges.Text = .Fields(12).Value
        txtChannelingCharges.Text = .Fields(13).Value
        cboAppointmentDuration.Text = .Fields(14).Value
        txtReferringCharges.Text = .Fields(15).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
    'Here, I am enabling the SetUpDoctor'sVisitingDays Button only if the Doctor Type is Visiting
    If cboDoctorCategory.Text = "Visiting" Then
        cmdSetUpDocSchedule.Enabled = True
    Else
        cmdSetUpDocSchedule.Enabled = False
    End If
    
    disableIfReferringDoctor    'Calling a function to disable certain components if the doctor is a "Referring" doctor
    
    disableIfPermanentDoctor    'Calling a function to disable the Referring Charges textfield if the Doctor is a "Permanent" doctor
    
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
    If dtpDateOfBirth.Value = "4/14/2008" Or dtpDateOfBirth.Value > DateTime.Date Then
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
    
    'Checking if the Phone Number (Home) textfield is empty
    If txtHomePhone.Text = "" Then
        txtHomePhone.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtHomePhone.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the License Number textfield is empty
    If txtLicenseNo.Text = "E" Then
        txtLicenseNo.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtLicenseNo.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Doctor Specialization textfield is empty
    If txtDoctorSpecialization.Text = "" Then
        txtDoctorSpecialization.BackColor = &H80000018  'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtDoctorSpecialization.BackColor = &H80000004  'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the user has made a selection in the Doctor Category ComboBox
    If cboDoctorCategory.Text = "" Then
        cboDoctorCategory.BackColor = &H80000018    'Highlighting the ComboBox in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboDoctorCategory.BackColor = &H80000004    'Bringing the ComboBox BackColour back to normal
    End If
    
    'If the user chooses 'Permanent Doctor' from the Doctor Category ComboBox
    If cboDoctorCategory.ListIndex = 0 Then
    
        'Checking if the Channeling Charges textfield is empty
        If txtChannelingCharges.Text = "-" Then
            txtChannelingCharges.BackColor = &H80000018 'Highlighting the textfield in a different colour
            Flag = False    'Setting the Flag variable to False to indicate invalid data
        Else
            txtChannelingCharges.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        End If
        
        'Checking if the Service Charges textfield is empty
        If txtServiceCharges.Text = "-" Then
            txtServiceCharges.BackColor = &H80000018 'Highlighting the textfield in a different colour
            Flag = False
        Else
            txtServiceCharges.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        End If
        
        'Checking to ensure that the user has made a selection in the Appointment Duration combo box
        If cboAppointmentDuration.Text = "" Then
            cboAppointmentDuration.BackColor = &H80000018 'Highlighting the combobox in a different colour
            Flag = False
        Else
            cboAppointmentDuration.BackColor = &H80000004 'Bringing the combobox BackColour back to normal
        End If
        
    End If
    
    
    'If the user chooses 'Visiting Doctor' from te Doctor Category combo box
    If cboDoctorCategory.ListIndex = 1 Then
        If txtServiceCharges.Text = "-" Then
            txtServiceCharges.BackColor = &H80000018  'Highlighting the textfield in a different colour
            Flag = False
        Else
            txtServiceCharges.BackColor = &H80000004  'Bringing the textfield BackColour back to normal
        End If
    End If
    
    'If the user chooses 'Referring Doctor' from the Doctor Category ComboBox
    If cboDoctorCategory.ListIndex = 2 Then
        'Checking if the Referring Charges textfield is empty
        If txtReferringCharges.Text = "-" Then
            txtReferringCharges.BackColor = &H80000018  'Highlighting the textfield in a different colour
            Flag = False    'Setting the Flag variable to False to indicate invalid data
        Else
            txtReferringCharges.BackColor = &H80000004  'Bringing the textfield BackColour back to normal
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
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If
    
End Sub


Private Sub txtChannelingCharges_GotFocus()
    
    If txtChannelingCharges.Text = "-" Then
        txtChannelingCharges.Text = ""
    End If
    
End Sub


Private Sub txtServiceCharges_GotFocus()
    
    If txtServiceCharges.Text = "-" Then
        txtServiceCharges.Text = ""
    End If
    
End Sub


Private Sub txtChannelingCharges_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeyMsg.Top = 4440    'Validation Note View
        picInvalidKeyMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtServiceCharges_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeyMsg.Top = 3960    'Validation Note View
        picInvalidKeyMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtChannelingCharges_LostFocus()
    
    If txtChannelingCharges.Text = "" Then
        txtChannelingCharges.Text = "-"
    End If
    
End Sub


Private Sub txtServiceCharges_LostFocus()
    
    If txtServiceCharges.Text = "" Then
        txtServiceCharges.Text = "-"
    End If
    
End Sub


Private Sub txtFirstName_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets
    
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 3720    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtHomePhone_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 7200    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtMobPhone_GotFocus()
    
    If txtMobPhone.Text = "-" Then
        txtMobPhone.Text = ""
    End If
    
End Sub


Private Sub txtMobPhone_LostFocus()
    
    If txtMobPhone.Text = "" Then
        txtMobPhone.Text = "-"
    End If
    
End Sub


Private Sub txtMobPhone_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 7680    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub



Private Sub txtNICNumber_GotFocus() 'This procedure will ensure that the textfield is empty when the user types in it.
    
    If txtNICNumber.Text = "-" Then
        txtNICNumber.Text = ""
    End If
    
End Sub

Private Sub txtNICNumber_LostFocus()

    If txtNICNumber.Text = "" Then
        txtNICNumber.Text = "-"
    End If
    
End Sub



Private Sub txtReferringCharges_GotFocus()
    
    If txtReferringCharges.Text = "-" Then
        txtReferringCharges.Text = ""
    End If
    
End Sub

Private Sub txtReferringCharges_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeyMsg.Top = 5400    'Validation Note View
        picInvalidKeyMsg.Visible = True
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
        picInvalidKeypressMsg.Top = 5640    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub



Private Sub txtReferringCharges_LostFocus()
    
    If txtReferringCharges.Text = "" Then
        txtReferringCharges.Text = "-"
    End If
    
End Sub

Private Sub txtSurname_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets
    
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 4200    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtServiceCharges_Change()
    
    'Here, I am ensuring that the user cannot type 0 as the first digit
    If txtServiceCharges.Text = "0" Then
    
        MsgBox "Error! The Figure Cannot Begin With Zero!", vbCritical, "Cannot Begin Figure With 0!"
        txtServiceCharges.Text = ""
        Exit Sub
        
    End If
    
End Sub
