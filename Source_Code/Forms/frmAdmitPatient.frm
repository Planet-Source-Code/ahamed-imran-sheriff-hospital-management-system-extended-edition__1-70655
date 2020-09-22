VERSION 5.00
Begin VB.Form frmAdmitPatient 
   Caption         =   "Admit Patient"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAdmitPatient.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click here to launch the Search Wizard"
      Top             =   1950
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1950
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1950
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1950
      Width           =   855
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmAdmitPatient.frx":1F664
      Height          =   800
      Left            =   3840
      Picture         =   "frmAdmitPatient.frx":1FA3E
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   7845
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmAdmitPatient.frx":21BFA
      Height          =   800
      Left            =   2880
      Picture         =   "frmAdmitPatient.frx":21FD0
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7845
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmAdmitPatient.frx":2418C
      Height          =   800
      Left            =   960
      Picture         =   "frmAdmitPatient.frx":24568
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   7845
      Width           =   890
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmAdmitPatient.frx":26724
      Height          =   800
      Left            =   1920
      Picture         =   "frmAdmitPatient.frx":26B39
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7845
      Width           =   890
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmAdmitPatient.frx":28CF5
      Height          =   855
      Left            =   6600
      Picture         =   "frmAdmitPatient.frx":29173
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmAdmitPatient.frx":2BEB7
      Height          =   855
      Left            =   7680
      Picture         =   "frmAdmitPatient.frx":2C39D
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmAdmitPatient.frx":2F0E1
      Height          =   855
      Left            =   8760
      Picture         =   "frmAdmitPatient.frx":2F5A0
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox txtDepartmentName 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox txtRoomID 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdRoomIDWizardButton 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   24
      ToolTipText     =   "Click Here to select Customer"
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox txtWardNo 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtWardID 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdWardIDWizardButton 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   21
      ToolTipText     =   "Click Here to select Customer"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txtDepartmentID 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdDepartmentIDWizardButton 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   18
      ToolTipText     =   "Click Here to select Customer"
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtAssignedDoctorName 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtReferredDoctorID 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdReferredDoctorIDWizardButton 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      ToolTipText     =   "Click Here to select Customer"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox txtGuardianID 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtAssignedDoctorID 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdAssignedDoctorWizardButton 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   15
      ToolTipText     =   "Click Here to select Customer"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox txtPatientID 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtAdditionalNotes 
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
      Height          =   645
      Left            =   8160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox txtReferredDoctorName 
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
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox txtReasonForStatus 
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
      Height          =   645
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   5760
      Width           =   2295
   End
   Begin VB.ComboBox cboPatientStatus 
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
      ItemData        =   "frmAdmitPatient.frx":322E4
      Left            =   2880
      List            =   "frmAdmitPatient.frx":322F7
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtAdmissionID 
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
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtAdmissionTime 
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
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtAdmissionDate 
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
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   600
      Top             =   7680
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   5400
      Top             =   7680
      Width           =   5535
   End
   Begin VB.Label lblRoomID 
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
      Left            =   6120
      TabIndex        =   43
      Top             =   5805
      Width           =   1215
   End
   Begin VB.Label lblWardNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward No."
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
      Left            =   6120
      TabIndex        =   42
      Top             =   5325
      Width           =   1695
   End
   Begin VB.Label lblWardID 
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
      Left            =   6120
      TabIndex        =   41
      Top             =   4845
      Width           =   1215
   End
   Begin VB.Label lblDepartmentName 
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
      Left            =   6120
      TabIndex        =   40
      Top             =   4365
      Width           =   1695
   End
   Begin VB.Label lblDepartmentID 
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
      Left            =   6120
      TabIndex        =   39
      Top             =   3885
      Width           =   1695
   End
   Begin VB.Label lblGuardianID 
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
      Left            =   840
      TabIndex        =   38
      Top             =   3885
      Width           =   1335
   End
   Begin VB.Label lblAssignedDoctorName 
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned Doctor Name"
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
      Left            =   6120
      TabIndex        =   37
      Top             =   3405
      Width           =   2055
   End
   Begin VB.Label lblAdditionalNotes 
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
      Left            =   6120
      TabIndex        =   36
      Top             =   6285
      Width           =   1935
   End
   Begin VB.Label lblAssignedDoctorID 
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned Doctor ID"
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
      Left            =   6120
      TabIndex        =   35
      Top             =   2925
      Width           =   1695
   End
   Begin VB.Label lblReferredDoctorName 
      BackStyle       =   0  'Transparent
      Caption         =   "Referred Doctor Name"
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
      Left            =   840
      TabIndex        =   34
      Top             =   7125
      Width           =   1935
   End
   Begin VB.Label lblReferredDoctorID 
      BackStyle       =   0  'Transparent
      Caption         =   "Referred Doctor ID"
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
      Left            =   840
      TabIndex        =   33
      Top             =   6645
      Width           =   1695
   End
   Begin VB.Label lblReasonForStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason For Status"
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
      Left            =   840
      TabIndex        =   32
      Top             =   5805
      Width           =   1695
   End
   Begin VB.Label lblAdmissionTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Time"
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
      Left            =   840
      TabIndex        =   31
      Top             =   4845
      Width           =   1335
   End
   Begin VB.Label lblAdmissionDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Date"
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
      Left            =   840
      TabIndex        =   30
      Top             =   4365
      Width           =   1335
   End
   Begin VB.Label lblPatientStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Status"
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
      Left            =   840
      TabIndex        =   29
      Top             =   5325
      Width           =   1335
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
      Left            =   840
      TabIndex        =   28
      Top             =   3405
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   600
      Y1              =   2520
      Y2              =   7560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   10920
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   10920
      X2              =   10920
      Y1              =   2520
      Y2              =   7560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   3000
      X2              =   10920
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   840
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblFrameTitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Information"
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
      TabIndex        =   27
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label lblAdmissionID 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission ID"
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
      Left            =   840
      TabIndex        =   26
      Top             =   2925
      Width           =   1215
   End
End
Attribute VB_Name = "frmAdmitPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Inpatients Maintenance Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 24/04/08
'Date Of Last Modification: 24/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Inpatients_Admission Table
'------------------------------------------------------------------------------

Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean


Private Sub cmdAssignedDoctorWizardButton_Click()
    
    frmAssignedDocSelectionWizard.Show
    
End Sub

Private Sub cmdClose_Click()

    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub


Private Sub cmdDepartmentIDWizardButton_Click()
    
    frmDepartmentsSearchWizardAdmit.Show
    
End Sub

Private Sub cmdLaunchInpatientSearch_Click() 'This function is fired when the Launch Inpatient Search Wizard Command Button is Clicked. It opens up the Inpatient Search Wizard


    enableAllFields     'Calling a Private Function To Enable All Fields
    enableAllButtons    'Calling a Private Function To Enable All Command Buttons

    cmdSave.Enabled = False     'Disabling the Save Command Button

    frmInpatientSearchWizardAdmit.Show      'Displays the Inpatient Search Wizard


End Sub


Private Sub cmdReferredDoctorIDWizardButton_Click()
    
    frmReferringDoctorWizard.Show
    
End Sub

Private Sub cmdRoomIDWizardButton_Click()
    
    frmRoomsSearchWizard.Show
    
End Sub

Private Sub cmdSave_Click()     'This function will save all the user's data in the database


    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then


        With rsInpatientsAdmission

            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then

                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.

                If txtReferredDoctorID.Text = "" Then
                    txtReferredDoctorID.Text = "-"
                End If

                If txtReferredDoctorName.Text = "" Then
                    txtReferredDoctorName.Text = "-"
                End If

                If txtAdditionalNotes.Text = "" Then
                    txtAdditionalNotes.Text = "-"
                End If

                'Save the user-entered data into the recordset
                .Fields(0) = txtAdmissionID.Text
                .Fields(1) = txtPatientID.Text
                .Fields(2) = txtGuardianID.Text
                .Fields(3) = txtAdmissionDate.Text
                .Fields(4) = txtAdmissionTime.Text
                .Fields(5) = cboPatientStatus.Text
                .Fields(6) = txtReasonForStatus.Text
                .Fields(7) = txtReferredDoctorID.Text
                .Fields(8) = txtReferredDoctorName.Text
                .Fields(9) = txtAssignedDoctorID.Text
                .Fields(10) = txtAssignedDoctorName.Text
                .Fields(11) = txtDepartmentID.Text
                .Fields(12) = txtDepartmentName.Text
                .Fields(13) = txtWardID.Text
                .Fields(14) = txtWardNo.Text
                .Fields(15) = txtRoomID.Text
                .Fields(16) = txtAdditionalNotes.Text

                .Update

                'Display Success Message
                MsgBox "The Record Was Saved Successfully! The Inpatient Admission Process Is Over!", vbInformation, "Succesful Save Procedure!"

                Unload Me

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

Private Sub cmdStep1_Click()

    Call Inpatients_Maintenance

    With rsInpatientMaintenance

        .MoveFirst

        Do While .EOF = False

            If .Fields(0).Value = txtPatientID.Text Then

                'Entering the values in the particular record into the fields on the interface
                frmInpatientsMaintenance.txtPatientID.Text = .Fields(0).Value
                frmInpatientsMaintenance.txtFirstName.Text = .Fields(1).Value
                frmInpatientsMaintenance.txtSurname.Text = .Fields(2).Value
                frmInpatientsMaintenance.cboGender.Text = .Fields(3).Value
                frmInpatientsMaintenance.dtpDateOfBirth.Value = .Fields(4).Value
                frmInpatientsMaintenance.txtNICNumber.Text = .Fields(5).Value
                frmInpatientsMaintenance.txtAddress.Text = .Fields(6).Value
                frmInpatientsMaintenance.txtPhoneHome.Text = .Fields(7).Value
                frmInpatientsMaintenance.txtPhoneMob.Text = .Fields(8).Value
                frmInpatientsMaintenance.txtPatientOccupation.Text = .Fields(9).Value
                frmInpatientsMaintenance.cboCivilStatus.Text = .Fields(10).Value
                frmInpatientsMaintenance.cboAccountType.Text = .Fields(11).Value
                frmInpatientsMaintenance.txtCompanyID.Text = .Fields(12).Value
                frmInpatientsMaintenance.txtCompanyName.Text = .Fields(13).Value
                Exit Do
                
            Else

                .MoveNext
                
            End If

        Loop

    End With


    'Enabling / Diabling the Navigation Buttons as necessary
    frmInpatientsMaintenance.cmdFirst.Enabled = False
    frmInpatientsMaintenance.cmdLast.Enabled = True
    frmInpatientsMaintenance.cmdPrevious.Enabled = False
    frmInpatientsMaintenance.cmdNext.Enabled = True

    'Enabling the Update Button, Delete Button and The Close Button
    frmInpatientsMaintenance.cmdUpdate.Enabled = True
    frmInpatientsMaintenance.cmdDelete.Enabled = True
    frmInpatientsMaintenance.cmdClose.Enabled = True

    'Enabling the Wizard Buttons
    frmInpatientsMaintenance.cmdCompanySearchWizard.Enabled = True
    
    'Enabling the "Step" Buttons
    frmInpatientsMaintenance.cmdStep2.Enabled = True

    frmInpatientsMaintenance.enableAllFields

    Unload Me

    frmInpatientsMaintenance.Show



End Sub



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


    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then


        With rsInpatientsAdmission

            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then

                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.

                If txtReferredDoctorID.Text = "" Then
                    txtReferredDoctorID.Text = "-"
                End If

                If txtReferredDoctorName.Text = "" Then
                    txtReferredDoctorName.Text = "-"
                End If

                If txtAdditionalNotes.Text = "" Then
                    txtAdditionalNotes.Text = "-"
                End If

                'Save the user-entered data into the recordset
                .Fields(0) = txtAdmissionID.Text
                .Fields(1) = txtPatientID.Text
                .Fields(2) = txtGuardianID.Text
                .Fields(3) = txtAdmissionDate.Text
                .Fields(4) = txtAdmissionTime.Text
                .Fields(5) = cboPatientStatus.Text
                .Fields(6) = txtReasonForStatus.Text
                .Fields(7) = txtReferredDoctorID.Text
                .Fields(8) = txtReferredDoctorName.Text
                .Fields(9) = txtAssignedDoctorID.Text
                .Fields(10) = txtAssignedDoctorName.Text
                .Fields(11) = txtDepartmentID.Text
                .Fields(12) = txtDepartmentName.Text
                .Fields(13) = txtWardID.Text
                .Fields(14) = txtWardNo.Text
                .Fields(15) = txtRoomID.Text
                .Fields(16) = txtAdditionalNotes.Text

                .Update

                'Display Success Message
                MsgBox "The Record Was Updated Successfully!", vbInformation, "Succesful Save Procedure!"

                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Function To Clear All Fields

            Else

                'Display 'No Modifications' Message
                MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"

                .CancelUpdate   'Cancel the Save Procedure

                Form_Load   'Calling the Form_Load Procedure

                clearAllFields  'Calling a Function To Clear All Fields

            End If

            .Requery    'Requerying the Table

        End With

    End If

End Sub

Private Sub cmdWardIDWizardButton_Click()

    frmWardsSearchWizardAdmit.Show
    
End Sub

Public Sub Form_Load()

    Call Connection  'Calling the Connection Procedure
        
    disableAllFields  'Calling a Function To Disable All Fields
    disableAllButtons   'Calling a Function To Disable All Command Buttons
    
    'Enabling  the First Button and the Last Button
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Close Button
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
    

End Function



Public Function enableAllFields() 'This function will enable all fields on the interface


    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will enable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Then
        eachField.Enabled = True
    End If

    Next
    

End Function


Public Function disableAllButtons() 'This function will disable all command buttons on the interface

    On Error Resume Next
    For Each eachButton In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will disable all Command Buttons
    If TypeOf eachButton Is CommandButton Then
        eachButton.Enabled = False
    End If

    Next

End Function



Public Function enableAllButtons() 'This function will enable all command buttons on the interface


    On Error Resume Next
    For Each eachButton In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will enable all Command Buttons
    If TypeOf eachButton Is CommandButton Then
        eachButton.Enabled = True
    End If

    Next
    
    'Disabling the Step 1 Button
    'cmdStep1.Enabled = False

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
    cboPatientStatus.Text = "----------SELECT-----------"
    
End Function



Private Sub cmdFirst_Click()  'This function will Navigate to the First Record

    'Enabling / Diabling the Navigation Buttons as necessary
    cmdFirst.Enabled = False
    cmdLast.Enabled = True
    cmdPrevious.Enabled = False
    cmdNext.Enabled = True

    'Enabling the Update Button
    cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep2.Enabled = True
    
    
    'Enabling the Referring Doctor Search Wizard
    cmdReferredDoctorIDWizardButton.Enabled = True
    
    'Enabling the Assigned Doctor Search Wizard
    cmdAssignedDoctorWizardButton.Enabled = True
    
    'Enabling the Department ID Search Wizard
    cmdDepartmentIDWizardButton.Enabled = True
    
    'Enabling the Ward ID Wizard Button
    cmdWardIDWizardButton.Enabled = True
    
    'Enabling the Room ID Wizard Button
    cmdRoomIDWizardButton.Enabled = True

    Call Inpatients_Admission  'Calling the Inpatients_Admission Procedure to interact with the recordset

    With rsInpatientsAdmission


        .MoveFirst  'Moving to the first record

        'Entering the values in the particular record into the fields on the interface
        txtAdmissionID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtGuardianID.Text = .Fields(2).Value
        txtAdmissionDate.Text = .Fields(3).Value
        txtAdmissionTime.Text = .Fields(4).Value
        cboPatientStatus.Text = .Fields(5).Value
        txtReasonForStatus.Text = .Fields(6).Value
        txtReferredDoctorID.Text = .Fields(7).Value
        txtReferredDoctorName.Text = .Fields(8).Value
        txtAssignedDoctorID.Text = .Fields(9).Value
        txtAssignedDoctorName.Text = .Fields(10).Value
        txtDepartmentID.Text = .Fields(11).Value
        txtDepartmentName.Text = .Fields(12).Value
        txtWardID.Text = .Fields(13).Value
        txtWardNo.Text = .Fields(14).Value
        txtRoomID.Text = .Fields(15).Value
        txtAdditionalNotes.Text = .Fields(16).Value

    End With

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdPrevious_Click() 'This function will Navigate to the Previous Record

    With rsInpatientsAdmission


        .MovePrevious   'Moving to the previous record

        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If

        'Entering the values in the particular record into the fields on the interface
        txtAdmissionID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtGuardianID.Text = .Fields(2).Value
        txtAdmissionDate.Text = .Fields(3).Value
        txtAdmissionTime.Text = .Fields(4).Value
        cboPatientStatus.Text = .Fields(5).Value
        txtReasonForStatus.Text = .Fields(6).Value
        txtReferredDoctorID.Text = .Fields(7).Value
        txtReferredDoctorName.Text = .Fields(8).Value
        txtAssignedDoctorID.Text = .Fields(9).Value
        txtAssignedDoctorName.Text = .Fields(10).Value
        txtDepartmentID.Text = .Fields(11).Value
        txtDepartmentName.Text = .Fields(12).Value
        txtWardID.Text = .Fields(13).Value
        txtWardNo.Text = .Fields(14).Value
        txtRoomID.Text = .Fields(15).Value
        txtAdditionalNotes.Text = .Fields(16).Value

    End With

    cmdNext.Enabled = True  'Enabling the Next Button
    cmdLast.Enabled = True  'Enabling the Last Button

    'Enabling the Update Button
    cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep2.Enabled = True

    
    'Enabling the Referring Doctor Search Wizard
    cmdReferredDoctorIDWizardButton.Enabled = True
    
    'Enabling the Assigned Doctor Search Wizard
    cmdAssignedDoctorWizardButton.Enabled = True
    
    'Enabling the Department ID Search Wizard
    cmdDepartmentIDWizardButton.Enabled = True
    
    'Enabling the Ward ID Wizard Button
    cmdWardIDWizardButton.Enabled = True
    
    'Enabling the Room ID Wizard Button
    cmdRoomIDWizardButton.Enabled = True

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdNext_Click() 'This function will Navigate to the Next Record

    With rsInpatientsAdmission

        .MoveNext   'Moving to the Next Record

        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If

        'Entering the values in the particular record into the fields on the interface
        txtAdmissionID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtGuardianID.Text = .Fields(2).Value
        txtAdmissionDate.Text = .Fields(3).Value
        txtAdmissionTime.Text = .Fields(4).Value
        cboPatientStatus.Text = .Fields(5).Value
        txtReasonForStatus.Text = .Fields(6).Value
        txtReferredDoctorID.Text = .Fields(7).Value
        txtReferredDoctorName.Text = .Fields(8).Value
        txtAssignedDoctorID.Text = .Fields(9).Value
        txtAssignedDoctorName.Text = .Fields(10).Value
        txtDepartmentID.Text = .Fields(11).Value
        txtDepartmentName.Text = .Fields(12).Value
        txtWardID.Text = .Fields(13).Value
        txtWardNo.Text = .Fields(14).Value
        txtRoomID.Text = .Fields(15).Value
        txtAdditionalNotes.Text = .Fields(16).Value

    End With

    cmdPrevious.Enabled = True  'Enabling the Previous Button
    cmdFirst.Enabled = True 'Enabling the First Button

    'Enabling the Update Button
    cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep2.Enabled = True

    
    'Enabling the Referring Doctor Search Wizard
    cmdReferredDoctorIDWizardButton.Enabled = True
    
    'Enabling the Assigned Doctor Search Wizard
    cmdAssignedDoctorWizardButton.Enabled = True
    
    'Enabling the Department ID Search Wizard
    cmdDepartmentIDWizardButton.Enabled = True
    
    'Enabling the Ward ID Wizard Button
    cmdWardIDWizardButton.Enabled = True
    
    'Enabling the Room ID Wizard Button
    cmdRoomIDWizardButton.Enabled = True

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdLast_Click() 'This function will Navigate to the Last Record

    'Enabling / Disabling the Navigation Buttons as necessary
    cmdLast.Enabled = False
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = False

    'Enabling the Update Button
    cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep2.Enabled = True

    
    'Enabling the Referring Doctor Search Wizard
    cmdReferredDoctorIDWizardButton.Enabled = True
    
    'Enabling the Assigned Doctor Search Wizard
    cmdAssignedDoctorWizardButton.Enabled = True
    
    'Enabling the Department ID Search Wizard
    cmdDepartmentIDWizardButton.Enabled = True
    
    'Enabling the Ward ID Wizard Button
    cmdWardIDWizardButton.Enabled = True
    
    'Enabling the Room ID Wizard Button
    cmdRoomIDWizardButton.Enabled = True

    Call Inpatients_Admission  'Calling the Inpatients_Admission Procedure to interact with the recordset

    With rsInpatientsAdmission

        .Requery

        .MoveLast   'Moving to the last record

        'Entering the values in the particular record into the fields on the interface
        txtAdmissionID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtGuardianID.Text = .Fields(2).Value
        txtAdmissionDate.Text = .Fields(3).Value
        txtAdmissionTime.Text = .Fields(4).Value
        cboPatientStatus.Text = .Fields(5).Value
        txtReasonForStatus.Text = .Fields(6).Value
        txtReferredDoctorID.Text = .Fields(7).Value
        txtReferredDoctorName.Text = .Fields(8).Value
        txtAssignedDoctorID.Text = .Fields(9).Value
        txtAssignedDoctorName.Text = .Fields(10).Value
        txtDepartmentID.Text = .Fields(11).Value
        txtDepartmentName.Text = .Fields(12).Value
        txtWardID.Text = .Fields(13).Value
        txtWardNo.Text = .Fields(14).Value
        txtRoomID.Text = .Fields(15).Value
        txtAdditionalNotes.Text = .Fields(16).Value

    End With

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub



Private Function textfieldsValidations() As Boolean  'This function will validate all fields

    Flag = True 'Setting the Flag variable to True


    'Checking if the user has made a selection in the Patient Status ComboBox
    If cboPatientStatus.Text = "" Then
        cboPatientStatus.BackColor = &H80000018    'Highlighting the ComboBox in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboPatientStatus.BackColor = &H80000004    'Bringing the ComboBox BackColour back to normal
    End If

    'Checking if the Reason For Status textfield is empty
    If txtReasonForStatus.Text = "" Then
        txtReasonForStatus.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtReasonForStatus.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If

    'Checking if the Referred Doctor ID textfield is empty
    If txtReferredDoctorID.Text = "" Then
        txtReferredDoctorID.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtReferredDoctorName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtReferredDoctorID.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtReferredDoctorName.BackColor = &H80000004    'Bringing the textfield BackColour back to normal
    End If

    'Checking if the Assigned Doctor ID textfield is empty
    If txtAssignedDoctorID.Text = "" Then
        txtAssignedDoctorID.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtAssignedDoctorName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtAssignedDoctorID.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtAssignedDoctorName.BackColor = &H80000004    'Bringing the textfield BackColour back to normal
    End If

    'Checking if the Department ID textfield is empty
    If txtDepartmentID.Text = "" Then
        txtDepartmentID.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtDepartmentName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtDepartmentID.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtDepartmentName.BackColor = &H80000004    'Bringing the textfield BackColour back to normal
    End If

    'Checking if the Ward ID textfield is empty
    If txtWardID.Text = "" Then
        txtWardID.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtWardNo.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtWardID.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtWardNo.BackColor = &H80000004    'Bringing the textfield BackColour back to normal
    End If

    'Checking if the Room ID textfield is empty
    If txtRoomID.Text = "" Then
        txtRoomID.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtRoomID.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If

    'Here, I am checking the state of the Flag variable and if it is False, I am displaying a
    'Message Box to instruct the user to enter data into all highlighted textfields.
    'The Save procedure will also be cancelled
    If Flag = False Then
        MsgBox "Error! Please Fill-in The Highlighted Textfields! They Are Compulsory!", vbCritical, "Please Fill Highlighted Textfields"
        textfieldsValidations = True    'Passing values to the Save procedure
    Else
        textfieldsValidations = False   'Passing values to the Save procedure
    End If

End Function



Private Sub txtReasonForStatus_GotFocus() 'This procedure will ensure that the textfield is empty when the user types in it.

    If txtReasonForStatus.Text = "-" Then
        txtReasonForStatus.Text = ""
    End If

End Sub

Private Sub txtReasonForStatus_LostFocus()    'This procedure will ensure that the textfield is not empty when the user is not typing in it.

    If txtReasonForStatus.Text = "" Then
        txtReasonForStatus.Text = "-"
    End If

End Sub


Private Sub txtAdditionalNotes_GotFocus()  'This procedure will ensure that the textfield is empty when the user types in it.

    If txtAdditionalNotes.Text = "-" Then
        txtAdditionalNotes.Text = ""
    End If

End Sub

Private Sub txtAdditionalNotes_LostFocus() 'This procedure will ensure that the textfield is not empty when the user is not typing in it.

    If txtAdditionalNotes.Text = "" Then
        txtAdditionalNotes.Text = "-"
    End If

End Sub

