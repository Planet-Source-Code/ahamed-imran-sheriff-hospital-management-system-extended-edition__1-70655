VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Durdans Hospitals (Pvt) Ltd. - Hospital Management System"
   ClientHeight    =   10290
   ClientLeft      =   180
   ClientTop       =   750
   ClientWidth     =   14970
   Icon            =   "frmMDI.frx":0000
   Picture         =   "frmMDI.frx":6852
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpCommand     =   1
      HelpContext     =   1
      HelpFile        =   "HELP.hlp"
      HelpKey         =   "F1"
   End
   Begin VB.PictureBox picRightNavigation 
      Align           =   3  'Align Left
      Height          =   8895
      Left            =   0
      Picture         =   "frmMDI.frx":23FDD
      ScaleHeight     =   8835
      ScaleWidth      =   3405
      TabIndex        =   3
      Top             =   1020
      Width           =   3465
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Engine"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   5
         Left            =   1200
         TabIndex        =   19
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label lblUserAccount 
         BackStyle       =   0  'Transparent
         Caption         =   "User Account Panel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   720
         TabIndex        =   16
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label lblRecordExplorer 
         BackStyle       =   0  'Transparent
         Caption         =   "Record Explorer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   840
         TabIndex        =   15
         Top             =   200
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         X1              =   0
         X2              =   3360
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Label lblSolutionsProvider 
         BackStyle       =   0  'Transparent
         Caption         =   "Powered By : UDS Labs Inc.,"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   420
         Left            =   360
         TabIndex        =   14
         Top             =   8520
         Width           =   2775
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Log Off / Exit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   7
         Left            =   1200
         TabIndex        =   13
         Top             =   7500
         Width           =   2055
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   6
         Left            =   1200
         TabIndex        =   12
         Top             =   6600
         Width           =   2055
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Reports Quick Launch"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   4
         Left            =   1200
         TabIndex        =   11
         Top             =   4185
         Width           =   2055
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Payments"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   3
         Left            =   1200
         TabIndex        =   10
         Top             =   3435
         Width           =   1935
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Channeling Services"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   2
         Left            =   1200
         TabIndex        =   9
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Outpatients"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Inpatients"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   555
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   1000
         Width           =   2535
      End
      Begin VB.Image imgUserAccount 
         Height          =   3465
         Left            =   0
         Picture         =   "frmMDI.frx":2CA0E
         Top             =   5475
         Width           =   3405
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Record Explorer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   555
         Left            =   360
         TabIndex        =   6
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Image imgRecordExplorer 
         Height          =   5460
         Left            =   0
         Picture         =   "frmMDI.frx":3675E
         Top             =   0
         Width           =   3405
      End
   End
   Begin VB.PictureBox picTopNavigation 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      Picture         =   "frmMDI.frx":43DC8
      ScaleHeight     =   960
      ScaleMode       =   0  'User
      ScaleWidth      =   15115.58
      TabIndex        =   1
      Top             =   0
      Width           =   14970
      Begin VB.CommandButton Command5 
         Height          =   975
         Left            =   43523
         Picture         =   "frmMDI.frx":4CB37
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   435
         Index           =   9
         Left            =   14040
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image5 
         Height          =   765
         Left            =   13200
         Picture         =   "frmMDI.frx":4FE1B
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblShortcut 
         BackStyle       =   0  'Transparent
         Caption         =   "Log Off"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   435
         Index           =   8
         Left            =   12240
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image4 
         Height          =   735
         Left            =   11280
         Picture         =   "frmMDI.frx":5596E
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblDateTime 
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   540
         Width           =   4335
      End
      Begin VB.Label lblDesignation 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   280
         Width           =   2295
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   120
         Picture         =   "frmMDI.frx":5B368
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.StatusBar BottomStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9915
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
            Text            =   "Durdans Hospital Management System "
            TextSave        =   "Durdans Hospital Management System "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "AD000001"
            TextSave        =   "AD000001"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Administrator"
            TextSave        =   "Administrator"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuInpatients 
      Caption         =   "&Inpatients"
      Begin VB.Menu mnuPatientAdmission 
         Caption         =   "Patient Admission"
         Begin VB.Menu mnuInpatientsMaintenance 
            Caption         =   "Step 1 - In-patients Maintenance"
         End
         Begin VB.Menu mnuSeparator9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGuardiansMaintenance 
            Caption         =   "Step 2 - Guardians Maintenance"
         End
         Begin VB.Menu mnuSeparator10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRegisterAdmitPatient 
            Caption         =   "Step 3 - Register / Admit Patient"
         End
      End
      Begin VB.Menu mnuSeparator12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreatment 
         Caption         =   "Treatments"
         Begin VB.Menu mnuMedicalTreatmentsMaintenance 
            Caption         =   "Medical Treatments Maintenance"
            Begin VB.Menu mnuAddMedicalTreatments 
               Caption         =   "Add Medical Treatments"
            End
            Begin VB.Menu mnuEditMedicalTreatmentRecords 
               Caption         =   "Edit Medical Treatment Records"
            End
         End
         Begin VB.Menu mnuSeparator14 
            Caption         =   "-"
         End
         Begin VB.Menu mnuServiceTreatmentMaintenance 
            Caption         =   "Service Treatments Maintenance"
            Begin VB.Menu mnuAddServiceTreatments 
               Caption         =   "Add Service Treatments"
            End
            Begin VB.Menu mnuEditServiceTreatmentRecord 
               Caption         =   "Edit Service Treatment Records"
            End
         End
      End
      Begin VB.Menu mnuSeparator15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOveralPatientBill 
         Caption         =   "View Overall Patient Bill"
      End
      Begin VB.Menu mnuSeparator16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDischargePatient 
         Caption         =   "Discharge Patient"
      End
   End
   Begin VB.Menu mnuOutpatients 
      Caption         =   "&Outpatients"
      Begin VB.Menu mnuOutpatientsMaintenance 
         Caption         =   "Outpatients Maintenance"
      End
      Begin VB.Menu mnuSeparator17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreatments 
         Caption         =   "Treatments"
         Begin VB.Menu mnuMedicalTreatmentMaintenance 
            Caption         =   "Medical Treatments Maintenance"
            Begin VB.Menu mnuAddMedicalTreatment 
               Caption         =   "Add Medical Treatments"
            End
            Begin VB.Menu mnuEditMedicalTreatment 
               Caption         =   "Edit Medical Treatments"
            End
         End
         Begin VB.Menu mnuSeparator19 
            Caption         =   "-"
         End
         Begin VB.Menu mnuServiceTreatmentsMaintenance 
            Caption         =   "Service Treatments Maintenance"
            Begin VB.Menu mnuAddServiceTreatment 
               Caption         =   "Add Service Treatments"
            End
            Begin VB.Menu mnuEditServiceTreatmentRecords 
               Caption         =   "Edit Service Treatments Records"
            End
         End
      End
      Begin VB.Menu mnuSeparator20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOverallPatientBill 
         Caption         =   "View Overall Patient Bill"
      End
   End
   Begin VB.Menu mnuChanneling 
      Caption         =   "&Channeling"
      Begin VB.Menu mnuManageAppointments 
         Caption         =   "Manage Appointments"
      End
   End
   Begin VB.Menu mnuPayments 
      Caption         =   "&Payments"
      Begin VB.Menu mnuInpatientTransactions 
         Caption         =   "Inpatients"
         Begin VB.Menu mnuManagePatientBill 
            Caption         =   "Manage Patient Bill"
         End
         Begin VB.Menu mnuSeparator48 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchPayments 
            Caption         =   "Search Payments"
         End
      End
      Begin VB.Menu mnuSeparator21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOutpatientsBilling 
         Caption         =   "Outpatients"
         Begin VB.Menu mnuManagePatientsBill 
            Caption         =   "Manage Patient Bill"
         End
         Begin VB.Menu mnuSeparator49 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOSearchPayments 
            Caption         =   "Search Payments"
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuPatientReports 
         Caption         =   "Patient Reports"
         Begin VB.Menu mnuInpatientsMasterReport 
            Caption         =   "Inpatients Master Report"
         End
         Begin VB.Menu mnuOutpatientsMasterReport 
            Caption         =   "Outpatients Master Report"
         End
         Begin VB.Menu mnuChannelingPatientsMasterReport 
            Caption         =   "Channeling Patients Master Report"
         End
         Begin VB.Menu mnuSeparator22 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPatientAdmissionMasterReport 
            Caption         =   "Patient Admission Master Report"
         End
         Begin VB.Menu mnuPatientDischargeMasterReport 
            Caption         =   "Patient Discharge Master Report"
         End
         Begin VB.Menu mnuSeparator23 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDoctorsVisitsMasterReport 
            Caption         =   "Doctors Visits Master Report"
         End
         Begin VB.Menu mnuPatientsMedicalTreatmentsMasterReport 
            Caption         =   "Patients' Medical Treatments Master Report"
            Begin VB.Menu mnuInpatientMedicines 
               Caption         =   "Inpatients"
            End
            Begin VB.Menu mnuOutpatientsMedicines 
               Caption         =   "Outpatients"
            End
         End
         Begin VB.Menu mnuPatientsMedicalServicesMasterReport 
            Caption         =   "Patients' Medical Services Master Report"
            Begin VB.Menu mnuInpatientsServices 
               Caption         =   "Inpatients"
            End
            Begin VB.Menu mnuOutpatientsServices 
               Caption         =   "Outpatients"
            End
         End
      End
      Begin VB.Menu mnuSeparator24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDoctorReports 
         Caption         =   "Doctor Reports"
         Begin VB.Menu mnuAllDoctors 
            Caption         =   "Doctors Master Report"
         End
         Begin VB.Menu mnuDoctorsChannelingSchedule 
            Caption         =   "Doctors' Channeling Schedule Report"
         End
      End
      Begin VB.Menu mnuSeparator25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorporateReports 
         Caption         =   "Corporate Reports"
         Begin VB.Menu mnuAllCompaniesMasterReport 
            Caption         =   "All Companies Master Report"
         End
         Begin VB.Menu mnuIndividualCompanyPatientsReport 
            Caption         =   "Companies Patients Report By Division"
            Begin VB.Menu mnuIndividualCompanyInpatients 
               Caption         =   "Inpatients"
            End
            Begin VB.Menu mnuIndividualCompanyOutpatients 
               Caption         =   "Outpatients"
            End
         End
      End
      Begin VB.Menu mnuSeparator26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHospitalReports 
         Caption         =   "Hospital Reports"
         Begin VB.Menu mnuMedicinesMasterReport 
            Caption         =   "Medicines Master Report"
         End
         Begin VB.Menu mnuMedicalServicesMasterReport 
            Caption         =   "Medical Services Master Report"
         End
         Begin VB.Menu mnuSeparator27 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDepartmentsMasterReport 
            Caption         =   "Departments Master Report"
         End
         Begin VB.Menu mnuRoomsMasterReport 
            Caption         =   "Rooms Master Report"
         End
      End
      Begin VB.Menu mnuSeparator28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRevenueReports 
         Caption         =   "Revenue Reports"
         Begin VB.Menu mnuInpatientsInvoice 
            Caption         =   "Inpatient's Invoice"
         End
         Begin VB.Menu mnuSeparator100 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInpatientsRevenueReports 
            Caption         =   "Inpatients Revenue Reports"
         End
         Begin VB.Menu mnuSeparator102 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOutpatientsRevenueReports 
            Caption         =   "Outpatients Revenue Reports"
         End
         Begin VB.Menu mnuSeparator103 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDoctorEarningsReports 
            Caption         =   "Doctor's Earnings Reports"
         End
         Begin VB.Menu mnuSeparator108 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBillStatusReport 
            Caption         =   "Bill Status Report"
         End
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuDoctorMaintenance 
         Caption         =   "Doctor's Maintenance"
         Begin VB.Menu mnuDoctorsMaintenance 
            Caption         =   "Doctor's Maintenance"
         End
         Begin VB.Menu mnuSeparator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDoctorScheduleMaintenance 
            Caption         =   "Doctor's Schedule Maintenance"
         End
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicinesMaintenance 
         Caption         =   "Medicines Maintenance"
      End
      Begin VB.Menu mnuSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHopitalServicesMaintenance 
         Caption         =   "Hospital Services Maintenance"
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepartmentsMaintenance 
         Caption         =   "Departments Maintenance"
      End
      Begin VB.Menu mnuSeparator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWardsMaintenance 
         Caption         =   "Wards Maintenance"
      End
      Begin VB.Menu mnuSeparator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRoomsMaintenance 
         Caption         =   "Rooms Maintenance"
      End
      Begin VB.Menu mnuSeparator8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCorporateMaintenance 
         Caption         =   "Corporate Maintenance"
      End
   End
   Begin VB.Menu mnuUserAccount 
      Caption         =   "&User Account"
      Begin VB.Menu mnuManageUserAccounts 
         Caption         =   "Manage User Accounts"
      End
      Begin VB.Menu mnuSeparator98 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuSeparator97 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuSeparator101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSearchEngine 
         Caption         =   "Search Engine"
      End
      Begin VB.Menu mnuSeparator96 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMicrosoftMagnifier 
         Caption         =   "Microsoft Magnifier"
      End
      Begin VB.Menu mnuSeparator95 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMicrosoftNarrator 
         Caption         =   "Microsoft Narrator"
      End
      Begin VB.Menu mnuSeparator94 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemMediaPlayer 
         Caption         =   "System Media Player"
      End
      Begin VB.Menu mnuSeparator93 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalendar 
         Caption         =   "Calendar"
      End
      Begin VB.Menu mnuSeparator92 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemCalculator 
         Caption         =   "System Calculator"
      End
      Begin VB.Menu mnuSeparator91 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemNotepad 
         Caption         =   "System Notepad"
      End
      Begin VB.Menu mnuSeparator90 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackupDatabase 
         Caption         =   "Backup Database"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuSeparator88 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuSeparator87 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuSeparator86 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSeparator85 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCredits 
         Caption         =   "Credits"
      End
      Begin VB.Menu mnuSeparator84 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpFile 
         Caption         =   "Help File"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Menu Driven Interface (MDI)
'Programmer: Imran Sheriff & Isham Sally
'Quality Assurance Engineer (Testing): Imran Sheriff
'Start Date: 01/04/08
'Date Of Last Modification: 01/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: None
'---------------------------------------------------------------------

Option Explicit
Dim iExitReply As Integer 'This variable will hold the user's choice, once he has been asked whether he wants to exit or not
Dim iLogOutReply As Integer 'This variable will hold the user's choice, once he has been asked whether he wants to log out or not



Private Sub MDIForm_Load()
    
    'In the following lines of code, I am checking the user access level
    'and appropriately disabling certain restricted functions
    
    
    If accessLevel = "Administrator" Then
        lblDesignation.Caption = "Welcome, Administrator"
        frmQuickLaunch.lblDesignation.Caption = "Administrator"

        Call Enable_Controls    'Calling a User Defined Function In Order To Enable All Components
    
    
    ElseIf accessLevel = "Cashier" Then
        lblDesignation.Caption = "Welcome, Cashier"
        frmQuickLaunch.lblDesignation.Caption = "Cashier"
        
        Call Enable_Controls    'Calling a User Defined Function In Order To Enable All Components
        
        lblShortcut(0).Enabled = False
        lblShortcut(1).Enabled = False
        lblShortcut(2).Enabled = False
        lblShortcut(4).Enabled = False
        lblShortcut(5).Enabled = False
       

        mnuPatientAdmission.Enabled = False
        mnuTreatment.Enabled = False
        mnuDischargePatient.Enabled = False
        mnuOutpatientsMaintenance.Enabled = False
        mnuTreatments.Enabled = False
        mnuChanneling.Enabled = False
        mnuReports.Enabled = False
        mnuManageUserAccounts.Enabled = False
        mnuMaintenance.Enabled = False
        mnuSearchEngine.Enabled = False
        mnuBackupDatabase.Enabled = False
       
       
       
    ElseIf accessLevel = "Receptionist" Then
        
        lblDesignation.Caption = "Welcome, Receptionist"
        frmQuickLaunch.lblDesignation.Caption = "Receptionist"
        
        Call Enable_Controls    'Calling a User Defined Function In Order To Enable All Components
        
        lblShortcut(0).Enabled = False
        lblShortcut(1).Enabled = False
        lblShortcut(3).Enabled = False
        lblShortcut(4).Enabled = False
       
        mnuInpatients.Enabled = False
        mnuOutpatients.Enabled = False
        mnuPayments.Enabled = False
        mnuReports.Enabled = False
        mnuManageUserAccounts.Enabled = False
        mnuMaintenance.Enabled = False
        mnuBackupDatabase.Enabled = False

    
    ElseIf accessLevel = "Inpatient Staff" Then
        lblDesignation.Caption = "Welcome, Inpatient Staff"
        frmQuickLaunch.lblDesignation.Caption = "Inpatient Staff"
        
        Call Enable_Controls    'Calling a User Defined Function In Order To Enable All Components
        
        lblShortcut(1).Enabled = False
        lblShortcut(2).Enabled = False
        lblShortcut(3).Enabled = False
        lblShortcut(4).Enabled = False
        lblShortcut(5).Enabled = False
        
        mnuOutpatients.Enabled = False
        mnuChanneling.Enabled = False
        mnuPayments.Enabled = False
        mnuReports.Enabled = False
        mnuMaintenance.Enabled = False
        mnuManageUserAccounts.Enabled = False
        mnuSearchEngine.Enabled = False
        mnuBackupDatabase.Enabled = False
        
    ElseIf accessLevel = "Outpatient Staff" Then
        
        lblDesignation.Caption = "Welcome, Outpatient Staff"
        frmQuickLaunch.lblDesignation.Caption = "Outpatient Staff"
    
        Call Enable_Controls    'Calling a User Defined Function In Order To Enable All Components
        
        lblShortcut(0).Enabled = False
        lblShortcut(2).Enabled = False
        lblShortcut(3).Enabled = False
        lblShortcut(4).Enabled = False
        lblShortcut(5).Enabled = False
        
        mnuInpatients.Enabled = False
        mnuChanneling.Enabled = False
        mnuPayments.Enabled = False
        mnuReports.Enabled = False
        mnuMaintenance.Enabled = False
        mnuManageUserAccounts.Enabled = False
        mnuSearchEngine.Enabled = False
        mnuBackupDatabase.Enabled = False
    End If
    
    frmQuickLaunch.Show
    lblDateTime.Caption = "Today is " & FormatDateTime(Now, vbLongDate)
    
    
    BottomStatusBar.Panels(4).Text = userName
    BottomStatusBar.Panels(5).Text = accessLevel
    
    
End Sub

Private Function Enable_Controls()

    'This is a User Defined Function That Will Enable All The Components On The Screen.
    'Here, I am running a For loop to include all the controls on the interface and then
    'I enable them all, with one line of code

    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In Controls
        ctrl.Enabled = True
    Next
    
End Function



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'This event occurs when the user tries to quit the application by clicking the
    'standard red cross button, on the top left hand corner of the interface
    
    If UnloadMode = 0 Then
        iExitReply = MsgBox(userName & ", Are You Sure You Wish To Exit The Application?", vbYesNo + vbQuestion, "Exit Application?")
        If iExitReply = vbNo Then
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub lblShortcut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Here, I am creating Rollover Effects For Each Button On The Shortcut Panel

    Select Case (Index)
        Case 0: 'Manage User Accounts Button
            lblShortcut(0).ForeColor = &H800000
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = True
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
            
        Case 1: 'Manage Inpatients Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H800000
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = True
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
            
        Case 2: 'Manage Outpatients Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H800000
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = True
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
            
        Case 3: 'Channeling Services Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H800000
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = True
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
            
        Case 4: 'Reports Quick Launch Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H800000
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = True
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
            
        Case 5: 'Search Engine Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H800000
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = True
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
    
    
        Case 6: 'Change Password Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H800000
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = True
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
    
        Case 7: 'Log Off / Exit Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H800000
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = True
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
            
        Case 8: 'Log Off Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H800000
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = True
            lblShortcut(9).FontUnderline = False
            
        Case 9: 'Exit Button
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H800000
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = True
    End Select
    
End Sub


Private Sub imgRecordExplorer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'The Following Block Of Code Ensures That The Mouse Pointer
    'Returns To Normal When It Is Not Over A Button
    
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
    
End Sub


Private Sub imgUserAccount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'The Following Block Of Code Ensures That The Mouse Pointer
    'Returns To Normal When It Is Not Over A Button
    
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
    
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuAddMedicalTreatment_Click()
    frmAddMedicalTreatmentsOut.Show
End Sub

Private Sub mnuAddMedicalTreatments_Click()
    frmAddMedicalTreatmentsIn.Show
End Sub

Private Sub mnuAddServiceTreatment_Click()
    frmAddServiceTreatmentsOut.Show
End Sub

Private Sub mnuAddServiceTreatments_Click()
    frmAddServiceTreatmentsIn.Show
End Sub

Private Sub mnuBackupDatabase_Click()
    frmBackupDatabase.Show
End Sub

Private Sub mnuCalendar_Click()
    frmCalendar.Show
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCloseAll_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCorporateMaintenance_Click()
    frmCompaniesMaintenance.Show
End Sub

Private Sub mnuCredits_Click()
    frmCredits.Show
End Sub





Private Sub mnuEditMedicalTreatment_Click()
    frmMedicalTreatmentsOut.Show
End Sub

Private Sub mnuEditMedicalTreatmentRecords_Click()
    frmMedicalTreatmentsMaintenance.Show
End Sub

Private Sub mnuEditServiceTreatmentRecord_Click()
    frmServiceTreatmentsMaintenance.Show
End Sub

Private Sub mnuEditServiceTreatmentRecords_Click()
    frmServiceTreatmentsOut.Show
End Sub

Private Sub mnuHelpFile_Click()
    '---Opening the Help Guide File with the Common Dialog Object
    On Error GoTo e
        Const cdlHelpPartialKey = &H105     ' Calls the search engine in Windows Help
        
        CommonDialog.HelpCommand = cdlHelpPartialKey
        CommonDialog.Action = 6
    Exit Sub
e:
    MsgBox Err.Description
    'MsgBox "Error! The Help Guide does not exist!", vbCritical, "Help File Does Not Exist!"

End Sub

Private Sub mnuLogOff_Click()

    iLogOutReply = MsgBox(userName & ", Are You Sure You Wish To Log Out Of Your Account?", vbYesNo + vbQuestion, "Log Out?")
    If iLogOutReply = vbYes Then
        frmLogin.Show
        Unload Me
    End If
    
End Sub

Private Sub mnuMicrosoftMagnifier_Click()   'Opens Up The Magnifier Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\magnify.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run Microsoft Magnifier on your computer", vbError, "Error Loading Microsoft Magnifier!"
    Resume Next
End Sub

Private Sub mnuMicrosoftNarrator_Click()    'Opens Up The Narrator Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\narrator.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run Microsoft Narrator on your computer", vbError, "Error Loading Microsoft Narrator!"
    Resume Next
End Sub





Private Sub mnuRoomsMaintenance_Click()
    frmRoomsMaintenance.Show
End Sub

Private Sub mnuSearchEngine_Click()
    frmSearchEngine.Show
End Sub

Private Sub mnuSystemCalculator_Click() 'Opens Up The Calculator Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\calc.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run the Calculator Utility on your computer", vbError, "Error Loading Calculator!"
    Resume Next
End Sub

Private Sub mnuSystemMediaPlayer_Click()    'Opens Up The System Media Player Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\dvdplay.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run System Media Player on your computer", vbError, "Error Loading System Media Player!"
    Resume Next
End Sub

Private Sub mnuSystemNotepad_Click()    'Opens Up The Notepad Utility
    On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\notepad.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run the Notepad Utility on your computer", vbError, "Error Loading Notepad!"
    Resume Next
End Sub

Private Sub mnuTileHorizontally_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertically_Click()
    Me.Arrange vbTileVertical
End Sub


Private Sub picTopNavigation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'The Following Block Of Code Ensures That The Mouse Pointer
    'Returns To Normal When It Is Not Over A Button
    
            lblShortcut(0).ForeColor = &H80000006
            lblShortcut(1).ForeColor = &H80000006
            lblShortcut(2).ForeColor = &H80000006
            lblShortcut(3).ForeColor = &H80000006
            lblShortcut(4).ForeColor = &H80000006
            lblShortcut(5).ForeColor = &H80000006
            lblShortcut(6).ForeColor = &H80000006
            lblShortcut(7).ForeColor = &H80000006
            lblShortcut(8).ForeColor = &H80000006
            lblShortcut(9).ForeColor = &H80000006
            lblShortcut(0).FontUnderline = False
            lblShortcut(1).FontUnderline = False
            lblShortcut(2).FontUnderline = False
            lblShortcut(3).FontUnderline = False
            lblShortcut(4).FontUnderline = False
            lblShortcut(5).FontUnderline = False
            lblShortcut(6).FontUnderline = False
            lblShortcut(7).FontUnderline = False
            lblShortcut(8).FontUnderline = False
            lblShortcut(9).FontUnderline = False
    
End Sub

Private Sub lblShortcut_Click(Index As Integer)

    'The following block of code illustrates which interfaces are displayed on click of
    'each respective button
    
    Select Case (Index)
        Case 0: 'Manage Inpatients Button
            Call Inpatients_Maintenance  'Calling the Inpatients_Maintenance Procedure to interact with the recordset
            frmInpatientsMaintenance.Show
            
        Case 1: 'Manage Outpatients Button
            frmOutpatientsMaintenance.Show
            
        Case 2: 'Channeling Services Button
            frmChannelingAppointments.Show
            
        Case 3: 'Payments Button
            frmPaymentOptions.Show
            
        Case 4: 'Reports Quick Launch Button
            frmReportsQuickLaunch.Show
            
        Case 5: 'Search Engine Button
            frmSearchEngine.Show
            
        Case 6: 'Change Password Button
            frmChangePassword.Show
            
        Case 7: 'Log Off / Exit Button
            frmTurnOff.Show
            
        Case 8: 'Log Off Button
            iLogOutReply = MsgBox(userName & ", Are You Sure You Wish To Log Out Of Your Account?", vbYesNo + vbQuestion, "Log Out?")
            If iLogOutReply = vbYes Then
                frmLogin.Show
                Unload Me
            End If
            
        Case 9: 'Exit Button
            iExitReply = MsgBox(userName & ", Are You Sure You Wish To Quit The Application?", vbYesNo + vbQuestion, "Quit Application?")
            If iExitReply = vbYes Then
                End
            End If
    
    End Select
    
End Sub


Private Sub mnuChangePassword_Click()
    frmChangePassword.Show
End Sub

Private Sub mnuCompanyMaintenance_Click()
    frmCompaniesMaintenance.Show
End Sub

Private Sub mnuDepartmentsMaintenance_Click()
    frmDepartmentsMaintenance.Show
End Sub

Private Sub mnuDischargePatient_Click()
    frmDischargeDetailsMaintenance.Show
End Sub

Private Sub mnuDoctorScheduleMaintenance_Click()
    frmDoctorScheduleMaintenance.Show
End Sub

Private Sub mnuDoctorsMaintenance_Click()
    frmDoctorsMaintenance.Show
End Sub

Private Sub mnuDoctorsVisitMaintenance_Click()
    frmDoctorVisitsMaintenance.Show
End Sub

Private Sub mnuDoctorVisitMaintenance_Click()
    frmDoctorVisitsMaintenance.Show
End Sub

Private Sub mnuExit_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Quit The Application?", vbYesNo + vbQuestion, "Quit Application?") = vbYes Then
        End
    End If
    
End Sub

Private Sub mnuGuardiansMaintenance_Click()
    Call Guardians_Maintenance
    Set frmGuardiansMaintenance.dgrdGuardiansInfo.DataSource = rsGuardiansMaintenance
    frmGuardiansMaintenance.Show
End Sub

Private Sub mnuHopitalServicesMaintenance_Click()
    frmServicesMaintenance.Show
End Sub

Private Sub mnuInpatientsMaintenance_Click()
    Call Inpatients_Maintenance  'Calling the Inpatients_Maintenance Procedure to interact with the recordset
    frmInpatientsMaintenance.Show
End Sub

Private Sub mnuManageAppointments_Click()
    frmChannelingAppointments.Show
End Sub

Private Sub mnuManagePatientBill_Click()
    frmIPDOverallBilling.Show
End Sub

Private Sub mnuManagePatientsBill_Click()
    frmOPDOverallBilling.Show
End Sub

Private Sub mnuManageUserAccounts_Click()
    frmManageUserAccounts.Show
End Sub

Private Sub mnuMedicinesMaintenance_Click()
    frmMedicinesMaintenance.Show
End Sub

Private Sub mnuOSearchPayments_Click()
    frmSearchOutpatientPayments.Show
End Sub

Private Sub mnuOutpatientsMaintenance_Click()
    frmOutpatientsMaintenance.Show
End Sub

Private Sub mnuRegisterAdmitPatient_Click()
    Call Inpatients_Admission
    frmAdmitPatient.Show
End Sub

Private Sub mnuSearchPayments_Click()
    frmSearchPayments.Show
End Sub


Private Sub mnuViewOverallPatientBill_Click()
    frmOPDOverallBilling.cmdSave.Enabled = False
    frmOPDOverallBilling.cmdGoToPayments.Enabled = False
    frmOPDOverallBilling.Show
End Sub

Private Sub mnuViewOveralPatientBill_Click()
    frmIPDOverallBilling.cmdSave.Enabled = False
    frmIPDOverallBilling.cmdGoToPayments.Enabled = False
    frmIPDOverallBilling.Show
End Sub


Private Sub mnuWardsMaintenance_Click()
    frmWardsMaintenance.Show
End Sub

'----------------------------REPORTS MDI---------------------------------------------------

Private Sub mnuAllCompaniesMasterReport_Click()
    RptAllCompaniesMaster.Show
End Sub

Private Sub mnuAllDoctors_Click()
    RptDoctorsMaster.Show
End Sub

Private Sub mnuAllDoctorsSchedulesMasterReport_Click()
    RptAllDoctorsShedule.Show
End Sub

Private Sub mnuChannelingPatientsMasterReport_Click()
    frmReportChannelingMaster.Show
End Sub

Private Sub mnuDepartmentsMasterReport_Click()
    RptDepartmentMaster.Show
End Sub

Private Sub mnuDoctorsVisitsMasterReport_Click()
    RptVisitingDoctorsShedule.Show
End Sub

Private Sub mnuIndividualDoctorsScheduleReport_Click()
    RptAllDoctorsShedule.Show
End Sub

Private Sub mnuInpatientsInvoice_Click()
    frmReportInpatientInvoice.Show
End Sub

Private Sub mnuInpatientsMasterReport_Click()
    frmReportInpatientMaster.Show
End Sub

Private Sub mnuMedicalServicesMasterReport_Click()
    RptMedicalServicesMaster.Show
End Sub

Private Sub mnuMedicinesMasterReport_Click()
    RptMedicinesMaster.Show
End Sub

Private Sub mnuOutpatientsMasterReport_Click()
    frmReportOutpatientMaster.Show
End Sub

Private Sub mnuPatientAdmissionMasterReport_Click()
    frmReportPatientAdmission.Show
End Sub

Private Sub mnuPatientDischargeMasterReport_Click()
    frmReportPatientDischarge.Show
End Sub

Private Sub mnuRoomsMasterReport_Click()
    RptRoomsMaster.Show
End Sub

Private Sub mnuInpatientMedicines_Click()
    frmReportInpatientMedicalTreatment.Show
End Sub

Private Sub mnuOutpatientsMedicines_Click()
    frmReportOutpatientMedicalTreatments.Show
End Sub

Private Sub mnuInpatientsServices_Click()
    frmReportInpatientServiceTreatments.Show
End Sub

Private Sub mnuOutpatientsServices_Click()
    frmReportOutPatientPatientServiceTreatements.Show
End Sub

Private Sub mnuDoctorsChannelingSchedule_Click()
    RptAllDoctorsShedule.Show
End Sub
Private Sub mnuIndividualCompanyOutpatients_Click()
    RptIndividualCompanyOutpatients.Show
End Sub

Private Sub mnuIndividualCompanyInpatients_Click()
    RptIndividualCompanyInpatients.Show
End Sub

Private Sub mnuInpatientsRevenueReports_Click()
    frmReportInpatientRevenue.Show
End Sub

Private Sub mnuOutpatientsRevenueReports_Click()
    frmReportOutpatientRevenue.Show
End Sub

Private Sub mnuDoctorEarningsReports_Click()
    frmReportDoctorsEarnings.Show
End Sub

Private Sub mnuBillStatusReport_Click()
    frmReportAging.Show
End Sub
