VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDoctorVisitingDaysWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor's Visiting Days Setup Wizard"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Close"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CheckBox chkMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   19
      Top             =   2430
      Width           =   255
   End
   Begin VB.CheckBox chkTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   18
      Top             =   3150
      Width           =   255
   End
   Begin VB.CheckBox chkWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   17
      Top             =   3870
      Width           =   255
   End
   Begin VB.CheckBox chkThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   16
      Top             =   4590
      Width           =   255
   End
   Begin VB.CheckBox chkFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   15
      Top             =   5310
      Width           =   255
   End
   Begin VB.CheckBox chkSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   14
      Top             =   6030
      Width           =   255
   End
   Begin VB.CheckBox chkSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   13
      Top             =   6750
      Width           =   255
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Save"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7680
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpTuesdayTimeIn 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   20
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpMondayTimeOut 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   21
      Top             =   2400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpMondayTimeIn 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   22
      Top             =   2400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpTuesdayTimeOut 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   23
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpWednesdayTimeIn 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   24
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpWednesdayTimeOut 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   25
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpThursdayTimeIn 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   26
      Top             =   4560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpThursdayTimeOut 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   27
      Top             =   4560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpFridayTimeIn 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   28
      Top             =   5280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpFridayTimeOut 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   29
      Top             =   5280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSaturdayTimeIn 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   30
      Top             =   6000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSaturdayTimeOut 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   31
      Top             =   6000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSundayTimeIn 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   32
      Top             =   6720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSundayTimeOut 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   33
      Top             =   6720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54919170
      CurrentDate     =   36494
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Visiting Doctor's Schedule Setup Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   46
      Top             =   240
      Width           =   4695
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   0
      Left            =   0
      Picture         =   "frmDoctorVisitingDays.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Schedule Setup Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   45
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblAvailableDays 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Days"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   44
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblMonday 
      BackStyle       =   0  'Transparent
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   43
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTimeIn 
      BackStyle       =   0  'Transparent
      Caption         =   "Time In"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   42
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblTimeOut 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   41
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3240
      X2              =   3240
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   6240
      X2              =   6240
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Label lblTuesday 
      BackStyle       =   0  'Transparent
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   40
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblWednesday 
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   39
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblThursday 
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   38
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblFriday 
      BackStyle       =   0  'Transparent
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   37
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblSaturday 
      BackStyle       =   0  'Transparent
      Caption         =   "Saturday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   36
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   35
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   9240
      X2              =   9240
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   240
      X2              =   240
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   240
      X2              =   9240
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line6 
      Index           =   1
      X1              =   240
      X2              =   9240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line7 
      Index           =   1
      X1              =   240
      X2              =   9240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblWizardFooter 
      BackStyle       =   0  'Transparent
      Caption         =   "Durdans Hospital Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   34
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Index           =   0
      Left            =   0
      Picture         =   "frmDoctorVisitingDays.frx":00A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   0
      Picture         =   "frmDoctorVisitingDays.frx":0140
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Schedule Setup Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Days"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "From (Time)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To (Time)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7320
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3360
      X2              =   3360
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   6360
      X2              =   6360
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Saturday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   9360
      X2              =   9360
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   360
      X2              =   9360
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line6 
      Index           =   0
      X1              =   360
      X2              =   9360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line7 
      Index           =   0
      X1              =   360
      X2              =   9360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Durdans Hospital Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   8520
      Width           =   3735
   End
End
Attribute VB_Name = "frmDoctorVisitingDaysWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    
    Call Connection 'Call Connection Procedure To Establish Connection With The Database
    
    Call VisitTimes_Schedule    'Calling the VisitTimes_Schedule Procedure To Interact With The Recordset
    
    With rsVisitTimesSchedule
    
        .MoveFirst  'Move To The First Record
    
        While .EOF = False  'Running Through All The Records
            
            'Here, I Am Checking The Doctor ID In Order To I Find A Match
            If .Fields(1).Value = frmDoctorsMaintenance.txtDoctorID.Text Then
                
                'In The Following Blocks Of If Else Conditions, I am enabling
                'the DateTime Pickers if the times have already been set.
                'Therefore for a new doctor, none of the DateTime Pickers
                'will be enabled.
                If .Fields(2).Value <> "12:00:00 AM" Then
                    chkMonday.Value = 1
                    dtpMondayTimeIn.Enabled = True
                    dtpMondayTimeOut.Enabled = True
                    dtpMondayTimeIn.Value = .Fields(2).Value
                    dtpMondayTimeOut.Value = .Fields(3).Value
                End If
                
                If .Fields(4).Value <> "12:00:00 AM" Then
                    chkTuesday.Value = 1
                    dtpTuesdayTimeIn.Enabled = True
                    dtpTuesdayTimeOut.Enabled = True
                    dtpTuesdayTimeIn.Value = .Fields(4).Value
                    dtpTuesdayTimeOut.Value = .Fields(5).Value
                End If
                    
                If .Fields(6).Value <> "12:00:00 AM" Then
                    chkWednesday.Value = 1
                    dtpWednesdayTimeIn.Enabled = True
                    dtpWednesdayTimeOut.Enabled = True
                    dtpWednesdayTimeIn.Value = .Fields(6).Value
                    dtpWednesdayTimeOut.Value = .Fields(7).Value
                End If
                
                If .Fields(8).Value <> "12:00:00 AM" Then
                    chkThursday.Value = 1
                    dtpThursdayTimeIn.Enabled = True
                    dtpThursdayTimeOut.Enabled = True
                    dtpThursdayTimeIn.Value = .Fields(8).Value
                    dtpThursdayTimeOut.Value = .Fields(9).Value
                End If
                
                If .Fields(10).Value <> "12:00:00 AM" Then
                    chkFriday.Value = 1
                    dtpFridayTimeIn.Enabled = True
                    dtpFridayTimeOut.Enabled = True
                    dtpFridayTimeIn.Value = .Fields(10).Value
                    dtpFridayTimeOut.Value = .Fields(11).Value
                End If
                
                If .Fields(12).Value <> "12:00:00 AM" Then
                    chkSaturday.Value = 1
                    dtpSaturdayTimeIn.Enabled = True
                    dtpSaturdayTimeOut.Enabled = True
                    dtpSaturdayTimeIn.Value = .Fields(12).Value
                    dtpSaturdayTimeOut.Value = .Fields(13).Value
                End If
                
                If .Fields(14).Value <> "12:00:00 AM" Then
                    chkSunday.Value = 1
                    dtpSundayTimeIn.Enabled = True
                    dtpSundayTimeOut.Enabled = True
                    dtpSundayTimeIn.Value = .Fields(14).Value
                    dtpSundayTimeOut.Value = .Fields(15).Value
                End If
            
            End If
                
            .MoveNext   'Moving to the next record
        
        Wend
        
    End With
                
End Sub



Private Sub chkMonday_Click()   'Enabling the DateTime Pickers

    If chkMonday.Value = 1 Then
        dtpMondayTimeIn.Enabled = True
        dtpMondayTimeOut.Enabled = True
    Else
        dtpMondayTimeIn.Enabled = False
        dtpMondayTimeOut.Enabled = False
    End If
    
End Sub



Private Sub chkTuesday_Click()  'Enabling the DateTime Pickers

    If chkTuesday.Value = 1 Then
        dtpTuesdayTimeIn.Enabled = True
        dtpTuesdayTimeOut.Enabled = True
    Else
        dtpTuesdayTimeIn.Enabled = False
        dtpTuesdayTimeOut.Enabled = False
    End If
    
End Sub

Private Sub chkWednesday_Click()    'Enabling the DateTime Pickers

    If chkWednesday.Value = 1 Then
        dtpWednesdayTimeIn.Enabled = True
        dtpWednesdayTimeOut.Enabled = True
    Else
        dtpWednesdayTimeIn.Enabled = False
        dtpWednesdayTimeOut.Enabled = False
    End If
    
End Sub

Private Sub chkThursday_Click() 'Enabling the DateTime Pickers
    
    If chkThursday.Value = 1 Then
        dtpThursdayTimeIn.Enabled = True
        dtpThursdayTimeOut.Enabled = True
    Else
        dtpThursdayTimeIn.Enabled = False
        dtpThursdayTimeOut.Enabled = False
    End If
    
End Sub

Private Sub chkFriday_Click()   'Enabling the DateTime Pickers
    
    If chkFriday.Value = 1 Then
        dtpFridayTimeIn.Enabled = True
        dtpFridayTimeOut.Enabled = True
    Else
        dtpFridayTimeIn.Enabled = False
        dtpFridayTimeOut.Enabled = False
    End If
    
End Sub

Private Sub chkSaturday_Click() 'Enabling the DateTime Pickers

    If chkSaturday.Value = 1 Then
        dtpSaturdayTimeIn.Enabled = True
        dtpSaturdayTimeOut.Enabled = True
    Else
        dtpSaturdayTimeIn.Enabled = False
        dtpSaturdayTimeOut.Enabled = False
    End If
    
End Sub

Private Sub chkSunday_Click()   'Enabling the DateTime Pickers
    
    If chkSunday.Value = 1 Then
        dtpSundayTimeIn.Enabled = True
        dtpSundayTimeOut.Enabled = True
    Else
        dtpSundayTimeIn.Enabled = False
        dtpSundayTimeOut.Enabled = False
    End If
    
        
End Sub


Private Sub cmdClose_Click()    'Closing the Wizard
    
    Unload Me
    
End Sub

Private Sub cmdSave_Click() 'If the Save Button is Clicked
    
    With rsVisitTimesSchedule
    
        .MoveFirst  'Moving to the first record
        
        While .EOF = False  'Running Through All The Records
            
            'Here, I Am Checking The Doctor ID In Order To I Find A Match.
            'When I Find A Match, I Will Be Setting The Values Of The DateTime
            'Pickers Accordingly.
            If .Fields(1).Value = frmDoctorsMaintenance.txtDoctorID.Text Then
            
                'Monday
                If chkMonday.Value = 0 Then
                    .Fields(2).Value = Null
                    .Fields(3).Value = Null
                Else
                    .Fields(2).Value = dtpMondayTimeIn.Hour & ":" & dtpMondayTimeIn.Minute & ":" & dtpMondayTimeIn.Second
                    .Fields(3).Value = dtpMondayTimeOut.Hour & ":" & dtpMondayTimeOut.Minute & ":" & dtpMondayTimeOut.Second
                End If
                
            
                'Tuesday
                If chkTuesday.Value = 0 Then
                    .Fields(4).Value = Null
                    .Fields(5).Value = Null
                Else
                    .Fields(4).Value = dtpTuesdayTimeIn.Hour & ":" & dtpTuesdayTimeIn.Minute & ":" & dtpTuesdayTimeIn.Second
                    .Fields(5).Value = dtpTuesdayTimeOut.Hour & ":" & dtpTuesdayTimeOut.Minute & ":" & dtpTuesdayTimeOut.Second
                End If
                
                
                'Wednesday
                If chkWednesday.Value = 0 Then
                    .Fields(6).Value = Null
                    .Fields(7).Value = Null
                Else
                    .Fields(6).Value = dtpWednesdayTimeIn.Hour & ":" & dtpWednesdayTimeIn.Minute & ":" & dtpWednesdayTimeIn.Second
                    .Fields(7).Value = dtpWednesdayTimeOut.Hour & ":" & dtpWednesdayTimeOut.Minute & ":" & dtpWednesdayTimeOut.Second
                End If
                
                
                'Thursday
                If chkThursday.Value = 0 Then
                    .Fields(8).Value = Null
                    .Fields(9).Value = Null
                Else
                    .Fields(8).Value = dtpThursdayTimeIn.Hour & ":" & dtpThursdayTimeIn.Minute & ":" & dtpThursdayTimeIn.Second
                    .Fields(9).Value = dtpThursdayTimeOut.Hour & ":" & dtpThursdayTimeOut.Minute & ":" & dtpThursdayTimeOut.Second
                End If
                
                
                'Friday
                If chkFriday.Value = 0 Then
                    .Fields(10).Value = Null
                    .Fields(11).Value = Null
                Else
                    .Fields(10).Value = dtpFridayTimeIn.Hour & ":" & dtpFridayTimeIn.Minute & ":" & dtpFridayTimeIn.Second
                    .Fields(11).Value = dtpFridayTimeOut.Hour & ":" & dtpFridayTimeOut.Minute & ":" & dtpFridayTimeOut.Second
                End If
                
                
                'Saturday
                If chkSaturday.Value = 0 Then
                    .Fields(12).Value = Null
                    .Fields(13).Value = Null
                Else
                    .Fields(12).Value = dtpSaturdayTimeIn.Hour & ":" & dtpSaturdayTimeIn.Minute & ":" & dtpSaturdayTimeIn.Second
                    .Fields(13).Value = dtpSaturdayTimeOut.Hour & ":" & dtpSaturdayTimeOut.Minute & ":" & dtpSaturdayTimeOut.Second
                End If
               
               
                'Sunday
                If chkSunday.Value = 0 Then
                    .Fields(14).Value = Null
                    .Fields(15).Value = Null
                Else
                    .Fields(14).Value = dtpSundayTimeIn.Hour & ":" & dtpSundayTimeIn.Minute & ":" & dtpSundayTimeIn.Second
                    .Fields(15).Value = dtpSundayTimeOut.Hour & ":" & dtpSundayTimeOut.Minute & ":" & dtpSundayTimeOut.Second
                End If
                
                .Update 'Updating the Recordset
                
                Unload Me   'Closing the Form
                
                Exit Sub
            
            End If
            
            .MoveNext   'Moving To The Next Record
        
        Wend
        
        
        
        'The Adding of a New Record Obviously Takes Place Only If There Is
        'No Matching Doctor ID To Be Found, Which Would Mean That The Doctor
        'Is New
        
        .AddNew 'Adding a New Record
        
        'Adding the Doctor ID Into The Relevant Field
        .Fields(1).Value = frmDoctorsMaintenance.txtDoctorID.Text
        
        
        'Monday
        If chkMonday.Value = 0 Then
            .Fields(2).Value = Null
            .Fields(3).Value = Null
        Else
            .Fields(2).Value = dtpMondayTimeIn.Hour & ":" & dtpMondayTimeIn.Minute & ":" & dtpMondayTimeIn.Second
            .Fields(3).Value = dtpMondayTimeOut.Hour & ":" & dtpMondayTimeOut.Minute & ":" & dtpMondayTimeOut.Second
        End If
                
                
        'Tuesday
        If chkTuesday.Value = 0 Then
            .Fields(4).Value = Null
            .Fields(5).Value = Null
        Else
            .Fields(4).Value = dtpTuesdayTimeIn.Hour & ":" & dtpTuesdayTimeIn.Minute & ":" & dtpTuesdayTimeIn.Second
            .Fields(5).Value = dtpTuesdayTimeOut.Hour & ":" & dtpTuesdayTimeOut.Minute & ":" & dtpTuesdayTimeOut.Second
        End If
                
                
        'Wednesday
        If chkWednesday.Value = 0 Then
            .Fields(6).Value = Null
            .Fields(7).Value = Null
        Else
            .Fields(6).Value = dtpWednesdayTimeIn.Hour & ":" & dtpWednesdayTimeIn.Minute & ":" & dtpWednesdayTimeIn.Second
            .Fields(7).Value = dtpWednesdayTimeOut.Hour & ":" & dtpWednesdayTimeOut.Minute & ":" & dtpWednesdayTimeOut.Second
        End If
                
                
        'Thursday
        If chkThursday.Value = 0 Then
            .Fields(8).Value = Null
            .Fields(9).Value = Null
        Else
            .Fields(8).Value = dtpThursdayTimeIn.Hour & ":" & dtpThursdayTimeIn.Minute & ":" & dtpThursdayTimeIn.Second
            .Fields(9).Value = dtpThursdayTimeOut.Hour & ":" & dtpThursdayTimeOut.Minute & ":" & dtpThursdayTimeOut.Second
        End If
                
                
        'Friday
        If chkFriday.Value = 0 Then
            .Fields(10).Value = Null
            .Fields(11).Value = Null
        Else
            .Fields(10).Value = dtpFridayTimeIn.Hour & ":" & dtpFridayTimeIn.Minute & ":" & dtpFridayTimeIn.Second
            .Fields(11).Value = dtpFridayTimeOut.Hour & ":" & dtpFridayTimeOut.Minute & ":" & dtpFridayTimeOut.Second
        End If
                
                
        'Saturday
        If chkSaturday.Value = 0 Then
            .Fields(12).Value = Null
            .Fields(13).Value = Null
        Else
            .Fields(12).Value = dtpSaturdayTimeIn.Hour & ":" & dtpSaturdayTimeIn.Minute & ":" & dtpSaturdayTimeIn.Second
            .Fields(13).Value = dtpSaturdayTimeOut.Hour & ":" & dtpSaturdayTimeOut.Minute & ":" & dtpSaturdayTimeOut.Second
        End If
               
               
        'Sunday
        If chkSunday.Value = 0 Then
            .Fields(14).Value = Null
            .Fields(15).Value = Null
        Else
            .Fields(14).Value = dtpSundayTimeIn.Hour & ":" & dtpSundayTimeIn.Minute & ":" & dtpSundayTimeIn.Second
            .Fields(15).Value = dtpSundayTimeOut.Hour & ":" & dtpSundayTimeOut.Minute & ":" & dtpSundayTimeOut.Second
        End If
                
        .Update 'Updating The Record
                
        Unload Me   'Closing the Form
        
        
    End With
    
    Unload Me
                
End Sub




