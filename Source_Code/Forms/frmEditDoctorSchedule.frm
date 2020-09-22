VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmDoctorScheduleMaintenance 
   Caption         =   "Maintain Doctor's Scedule Module"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditDoctorSchedule.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDoctorSearchWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Click Here to select a Doctor"
      Top             =   2640
      Width           =   375
   End
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetUpChannelingTimes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Set Up Doctor's Channeling Times"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2610
      Width           =   3615
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
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2610
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid dgrdVisitingDoctorsVisitSchedule 
      Height          =   4335
      Left            =   1680
      TabIndex        =   2
      Top             =   3120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   -2147483629
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
      Caption         =   "Visiting Doctors Visit Schedule"
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
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   360
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   4080
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblDoctorsScheduleInformation 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Schedule Information"
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
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   360
      Y1              =   2160
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   11520
      Y1              =   8040
      Y2              =   2160
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
      Left            =   1680
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "frmDoctorScheduleMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Doctor Schedule Maintenance
'Programmer: Deshan Subasinghe
'Quality Assurance Engineer (Testing):  Isham Sally
'Start Date: 25/04/08
'Date Of Last Modification: 25/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed:
'----------------------------------------------------------------------------

Private Sub cmdClose_Click()

    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdDoctorSearchWizard_Click()
    frmDoctorSearchWizardSchedule.Show
End Sub

Private Sub cmdSetUpChannelingTimes_Click()
    frmSetUpChannelingSchedule.Show
End Sub
