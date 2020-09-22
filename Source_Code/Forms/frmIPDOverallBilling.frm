VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIPDOverallBilling 
   Caption         =   "Overall Billing Details"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmIPDOverallBilling.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Print"
      Enabled         =   0   'False
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8160
      Width           =   2535
   End
   Begin VB.TextBox txtOverallInBillID 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Save"
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Close"
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8160
      Width           =   2535
   End
   Begin VB.CommandButton cmdGoToPayments 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Go To Payments Form"
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7440
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid dgrdTotalServiceTreatments 
      Height          =   495
      Left            =   3120
      TabIndex        =   61
      Top             =   -240
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483629
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Inpatients Information Table"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAssignedDoctorID 
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
      TabIndex        =   6
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5640
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
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdInpatientSearchWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "Click Here to select an Inpatient"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtNettTotal 
      Alignment       =   2  'Center
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtVAT 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtHospitalCharges 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox txtDoctorsCharges 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtNoOfDays 
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
      Top             =   8040
      Width           =   2295
   End
   Begin VB.TextBox txtAccountType 
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
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox txtPatientID 
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
      TabIndex        =   2
      Top             =   2760
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
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3240
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
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtDepartmentID 
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
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtDepartmentName 
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
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtWardNo 
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
      TabIndex        =   9
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox txtRoomID 
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
      TabIndex        =   10
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox txtMedicalTreatmentCharges 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtRoomCharges 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtServiceTreatmentCharges 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   3240
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtpAdmissionDate 
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   7080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   56623105
      CurrentDate     =   39517
   End
   Begin MSComCtl2.DTPicker dtpTodaysDate 
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Top             =   7560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   56623105
      CurrentDate     =   39517
   End
   Begin MSDataGridLib.DataGrid dgrdTotalMedicalTreatments 
      Height          =   495
      Left            =   240
      TabIndex        =   60
      Top             =   -240
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483629
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Inpatients Information Table"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   7
      Left            =   10920
      TabIndex        =   59
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   6
      Left            =   10920
      TabIndex        =   58
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   5
      Left            =   10920
      TabIndex        =   57
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   4
      Left            =   10920
      TabIndex        =   56
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   3
      Left            =   10920
      TabIndex        =   55
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   2
      Left            =   10920
      TabIndex        =   54
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   1
      Left            =   10920
      TabIndex        =   53
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   0
      Left            =   10920
      TabIndex        =   52
      Top             =   2400
      Width           =   375
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
      Left            =   960
      TabIndex        =   51
      Top             =   4725
      Width           =   1695
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   50
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblNettTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "NETT TOTAL"
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
      Left            =   6000
      TabIndex        =   49
      Top             =   6405
      Width           =   1575
   End
   Begin VB.Label lblDiscount 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Left            =   6000
      TabIndex        =   48
      Top             =   5685
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   6000
      TabIndex        =   47
      Top             =   4755
      Width           =   1575
   End
   Begin VB.Label lblVAT 
      BackStyle       =   0  'Transparent
      Caption         =   "VAT (15%)"
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
      Left            =   6000
      TabIndex        =   46
      Top             =   5205
      Width           =   1575
   End
   Begin VB.Label lblHospitalCharges 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Charges (At 1000/= per day)"
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
      Left            =   6000
      TabIndex        =   45
      Top             =   4230
      Width           =   2295
   End
   Begin VB.Label lblDoctorsCharges 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Charges"
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
      Left            =   6000
      TabIndex        =   44
      Top             =   2325
      Width           =   2055
   End
   Begin VB.Label lblNoOfDays 
      BackStyle       =   0  'Transparent
      Caption         =   "Length Of Stay (Days)"
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
      TabIndex        =   43
      Top             =   8085
      Width           =   1935
   End
   Begin VB.Label lblTodayDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Today's Date"
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
      TabIndex        =   42
      Top             =   7605
      Width           =   1335
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
      Left            =   960
      TabIndex        =   41
      Top             =   4245
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
      Left            =   960
      TabIndex        =   40
      Top             =   2805
      Width           =   1335
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
      Left            =   960
      TabIndex        =   39
      Top             =   2325
      Width           =   1335
   End
   Begin VB.Label lblFrameTitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Details"
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
      TabIndex        =   38
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   840
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2880
      X2              =   5520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   5520
      X2              =   5520
      Y1              =   1920
      Y2              =   8520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   5520
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   600
      Y1              =   1920
      Y2              =   8520
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
      TabIndex        =   37
      Top             =   3285
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
      TabIndex        =   36
      Top             =   3765
      Width           =   1335
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
      Left            =   960
      TabIndex        =   35
      Top             =   5205
      Width           =   1335
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
      Left            =   960
      TabIndex        =   34
      Top             =   5685
      Width           =   1695
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
      Left            =   960
      TabIndex        =   33
      Top             =   6165
      Width           =   1695
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
      Left            =   960
      TabIndex        =   32
      Top             =   6645
      Width           =   1695
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
      Left            =   960
      TabIndex        =   31
      Top             =   7125
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Details"
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
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   5760
      X2              =   6120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   7800
      X2              =   11280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   11280
      X2              =   11280
      Y1              =   1920
      Y2              =   6960
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   5760
      X2              =   11280
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   5760
      X2              =   5760
      Y1              =   1920
      Y2              =   6960
   End
   Begin VB.Label lblMedicalTreatmentCharges 
      BackStyle       =   0  'Transparent
      Caption         =   "Medical Treatment Charges"
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
      Left            =   6000
      TabIndex        =   29
      Top             =   2805
      Width           =   2415
   End
   Begin VB.Label lblRoomCharges 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Charges"
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
      Left            =   6000
      TabIndex        =   28
      Top             =   3765
      Width           =   1575
   End
   Begin VB.Label lblServiceTreatmentCharges 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Treatment Charges"
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
      Left            =   6000
      TabIndex        =   27
      Top             =   3285
      Width           =   2415
   End
End
Attribute VB_Name = "frmIPDOverallBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Inpatient Overall Billing Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 11/05/08
'Date Of Last Modification: 11/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Inpatient_Payment_Details
'-----------------------------------------------------------------------------

Option Explicit

'The following variables will be used to autogenerate the Invoice ID
Dim iNumOfRecords As Integer    'This variable holds the number of records in the table
Dim strCode As String   'This variable will eventually hold the Invoice ID to be autogenerated
Dim iNumberOfRecords As Integer 'This variable will hold the number of records in the Inpatient_Payment_Details table
Dim strDisplay As String    'This variable will eventually display the OverallInBillID in the invisible textfield
Dim iTotalPayable As Double    'This variable will hold the value of the Total Payable
Dim billID As String    'This variable will hold the Billing ID of the patient



Private Sub cmdClose_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdGoToPayments_Click()

    'Ensuring that the user has selected a patient
    If txtPatientID.Text = "" Then
        MsgBox "You Cannot Go To The Payments Process Until You Select A Patient!", vbCritical, "Please Select A Patient!"
        Exit Sub
    End If
        
    On Error GoTo error_handler
               
           
    Call Inpatient_Billing    'Calling the Inpatient_Billing Procedure to interact with the recordset
    
    'Generate Invoice ID By Utilizing the Inpatient_Billing Table
    With rsInpatientBilling
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strCode = "IIN0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Inpatient_Billing Table
            If iNumOfRecords < 10 Then
                strCode = "IIN000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strCode = "IIN00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strCode = "IIN0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strCode = "IIN" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    
    iTotalPayable = Int(Val(txtNettTotal.Text))   'Storing the Nett Total in this variable
    
    'The following line of code will enter the autogenerated Invoice ID into the relevant textfield
    frmIPDCreateBill.txtInvoiceID.Text = strCode
    
    'Entering all relevant data onto the Payment Form
    frmIPDCreateBill.txtBillingDate.Text = DateTime.Date    'System Date
    frmIPDCreateBill.txtAdmissionID.Text = txtAdmissionID.Text
    frmIPDCreateBill.txtPatientID.Text = txtPatientID.Text
    frmIPDCreateBill.txtPatientName.Text = txtFirstName.Text & " " & txtSurname.Text
    frmIPDCreateBill.txtAccountType.Text = txtAccountType.Text
    frmIPDCreateBill.txtTotalCost.Text = txtTotal.Text
    frmIPDCreateBill.txtDiscount.Text = txtDiscount.Text
    frmIPDCreateBill.txtTotalPayable.Text = iTotalPayable
    
    
    'Here, I am calculating the Total Paid So Far
    Call TotalPaidSoFar
    Set frmIPDCreateBill.dgrdTotalPaidSoFar.DataSource = rsTotalPaidSoFar
    frmIPDCreateBill.txtTotalPaidSoFar.Text = frmIPDCreateBill.dgrdTotalPaidSoFar.Columns(0).Value
    

    'Here, I am calculating the balance owed by the patient
    frmIPDCreateBill.txtBalanceOwing.Text = Val(frmIPDCreateBill.txtTotalPayable.Text) - Val(frmIPDCreateBill.txtTotalPaidSoFar.Text)

    'Here, I am displaying the Bill Status
    If frmIPDCreateBill.txtTotalPaidSoFar.Text <> "0" Then
        If frmIPDCreateBill.txtBalanceOwing.Text = "0" Then
            frmIPDCreateBill.txtBillStatus.Text = "PAID"
        End If
    End If
        
    
    Unload Me   'Closing this form
    
    frmIPDCreateBill.Show   'Opening Up The Payments Form
    
    Exit Sub
    
error_handler:
    frmIPDCreateBill.txtTotalPaidSoFar.Text = "0"
    frmIPDCreateBill.txtBalanceOwing.Text = frmIPDCreateBill.txtTotalPayable.Text
    Unload Me

End Sub

Private Sub cmdInpatientSearchWizard_Click()
    
    frmInpatientSearchBilling.Show
    
End Sub



Private Sub cmdPrint_Click()
    
    On Error GoTo e
    DataEnvironment1.Commands("InpatientInvoice").Parameters(0) = billID
    DataEnvironment1.Commands("InpatientInvoice").Parameters(1) = ""
    RptInpatientInvoice.Show
    DataEnvironment1.rsInpatientInvoice.Close
        
    Unload Me
    Exit Sub
e:
    If Err.Number <> 3704 Then
        MsgBox Err.Description & "" & Err.Number, vbCritical
    End If

End Sub

Private Sub cmdSave_Click()

    'Ensuring that the user has selected a patient
    If txtAdmissionID.Text = "" Then
        MsgBox "Error! You Have Not Selected A Patient!", vbCritical, "Please Select A Patient!"
        Exit Sub
    End If


    Call Inpatient_Payment_Details    'Calling the Inpatient_Payment_Details Procedure to interact with the recordset

    'Generate OverallInBillID By Utilizing the Inpatient_Payment_Details Table
    With rsInpatientPaymentDetails

        If .RecordCount = 0 Then    'If there are no records in the table

            strDisplay = "IPD0001"

        Else

            'Calculating the number of records and storing in a variable
            iNumberOfRecords = .RecordCount
            iNumberOfRecords = iNumberOfRecords + 1   'incrementing the number by 1

            'The following block of code will generate the ID according
            'to the number of records in the Inpatient_Payment_Details Table
            If iNumberOfRecords < 10 Then
                strDisplay = "IPD000" & iNumberOfRecords
            ElseIf iNumberOfRecords < 100 Then
                strDisplay = "IPD00" & iNumberOfRecords
            ElseIf iNumberOfRecords < 1000 Then
                strDisplay = "IPD0" & iNumberOfRecords
            ElseIf iNumberOfRecords < 10000 Then
                strDisplay = "IPD" & iNumberOfRecords
            End If

        End If

        .Requery    'Requerying the Table

        .AddNew     'Adding a new recordset

    End With

    'The following line of code will enter the autogenerated OverallInBillID
    'into the invisible OverallInBillID textfield
    txtOverallInBillID.Text = strDisplay
    
    
    'Here, I am ensuring that the Discount textfield is not empty when I save
    If txtDiscount.Text = "" Then
        txtDiscount.Text = "-"
    End If

    
    'Now I am going to save the record in the database
    With rsInpatientPaymentDetails
    
        'Making sure that the user wants to save the record
        If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
        
            .Fields(0) = txtOverallInBillID.Text
             billID = txtOverallInBillID.Text     'Passing this value to a variable
            .Fields(1) = txtAdmissionID.Text
            .Fields(2) = txtPatientID.Text
            .Fields(3) = txtFirstName.Text
            .Fields(4) = txtSurname.Text
            .Fields(5) = txtAccountType.Text
            .Fields(6) = txtAssignedDoctorID.Text
            .Fields(7) = txtDepartmentID.Text
            .Fields(8) = txtDepartmentName.Text
            .Fields(9) = txtWardNo.Text
            .Fields(10) = txtRoomID.Text
            .Fields(11) = dtpAdmissionDate.Value
            .Fields(12) = dtpTodaysDate.Value
            .Fields(13) = txtNoOfDays.Text  'Length Of Stay
            .Fields(14) = txtDoctorsCharges.Text
            .Fields(15) = txtMedicalTreatmentCharges.Text
            .Fields(16) = txtServiceTreatmentCharges.Text
            .Fields(17) = txtRoomCharges.Text
            .Fields(18) = txtHospitalCharges.Text
            .Fields(19) = txtTotal.Text
            .Fields(20) = txtVAT.Text
            .Fields(21) = txtDiscount.Text
            .Fields(22) = txtNettTotal.Text
            
            .Update
            
            'Display Success Message
            MsgBox "The Record Was Saved Successfully!", vbInformation, "Succesful Save Procedure!"
                        
        Else
            
            'Display 'No Modifications' Message
            MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
            .CancelUpdate   'Cancel the Save Procedure
            
        End If

    End With
    
    cmdPrint.Enabled = True 'Enabling the Print button
    
End Sub

Private Sub Form_Load()
    
    'Displaying the system date in the Today's Date textfield
    dtpTodaysDate.Value = DateTime.Date
    
End Sub
