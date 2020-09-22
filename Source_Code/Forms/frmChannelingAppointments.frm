VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChannelingAppointments 
   Caption         =   "Maintain Channeling Appointments"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmChannelingAppointments.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   11820
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Click Here To Print This Record"
      Top             =   8280
      Width           =   1695
   End
   Begin VB.PictureBox picInvalidDataMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3120
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   36
      Top             =   6600
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Digits Here! Only Alphabets Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   37
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   480
      Top             =   960
   End
   Begin VB.PictureBox picInvalidKeyMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3120
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   34
      Top             =   7320
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   35
         Top             =   105
         Width           =   2175
      End
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Click Here To Save This Record"
      Top             =   8280
      Width           =   1695
   End
   Begin VB.TextBox txtTokenNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   7680
      Width           =   2175
   End
   Begin VB.TextBox txtContactNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      MaxLength       =   12
      TabIndex        =   22
      Top             =   7320
      Width           =   2175
   End
   Begin VB.TextBox txtLastName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      TabIndex        =   20
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox txtFirstName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      TabIndex        =   18
      Top             =   6600
      Width           =   2175
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Click Here To Close This Interface"
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheckChannelingDays 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Check Doctor's Channeling Days"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Click here to launch the Search Wizard"
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox txtAppointmentDuration 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtChosenDay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtChannelingCharges 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtSpecialization 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtDoctorID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click here to launch the Search Wizard"
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   13080
      Picture         =   "frmChannelingAppointments.frx":2069B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   14520
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpChosenDate 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   5520
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   20578305
      CurrentDate     =   39517
   End
   Begin MSComCtl2.DTPicker dtpAppointmentStartTime 
      Height          =   285
      Left            =   2280
      TabIndex        =   29
      Top             =   5880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   20578306
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpAppointmentEndTime 
      Height          =   285
      Left            =   2280
      TabIndex        =   30
      Top             =   6240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   20578306
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpStartTime 
      Height          =   285
      Left            =   2280
      TabIndex        =   31
      Top             =   4800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   20578306
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpEndTime 
      Height          =   285
      Left            =   2280
      TabIndex        =   32
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   20578306
      CurrentDate     =   36494
   End
   Begin MSDataGridLib.DataGrid dgrdChannelingInfo 
      Height          =   6075
      Left            =   4680
      TabIndex        =   33
      Top             =   1920
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10716
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
      Caption         =   "Appointments Schedule"
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
   Begin VB.Label lblAppointmentEndTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment End Time"
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
      Left            =   240
      TabIndex        =   27
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label lblAppointmentStartTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment Start Time"
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
      Left            =   240
      TabIndex        =   26
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblTokenNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Token No"
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
      Left            =   240
      TabIndex        =   25
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label lblContactNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
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
      Left            =   240
      TabIndex        =   23
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label lblLastName 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Last Name"
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
      Left            =   240
      TabIndex        =   21
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label lblFirstName 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient First Name"
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
      Left            =   240
      TabIndex        =   19
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label lblEndTime 
      BackStyle       =   0  'Transparent
      Caption         =   "End Time"
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
      Left            =   240
      TabIndex        =   16
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblStartTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
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
      Left            =   240
      TabIndex        =   15
      Top             =   4845
      Width           =   1815
   End
   Begin VB.Label lblAppointmentDuration 
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment Duration"
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
      Left            =   360
      TabIndex        =   13
      Top             =   3525
      Width           =   1935
   End
   Begin VB.Label lblPleaseChooseDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Choose Date"
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
      Left            =   240
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblChosenDay 
      BackStyle       =   0  'Transparent
      Caption         =   "Chosen Day"
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
      Left            =   240
      TabIndex        =   10
      Top             =   4485
      Width           =   1335
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
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3165
      Width           =   1935
   End
   Begin VB.Label lblSpecialization 
      BackStyle       =   0  'Transparent
      Caption         =   "Specialization"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2805
      Width           =   1695
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
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2445
      Width           =   855
   End
End
Attribute VB_Name = "frmChannelingAppointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name:   Change Password Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 14/05/08
'Date Of Last Modification: 14/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Channeling_Appointments Table
'--------------------------------------------------------------------------------


Dim iNumberOfRecords As Integer 'This variable will store the number of records belonging to the particular doctor
Dim datagridText As String  'This variable will hold the appointment end time of the last patient in the datagrid


Private Sub cmdCheckChannelingDays_Click()
    frmSelectADay.Show
End Sub

Private Sub cmdClose_Click()
    
    'Obtaining confirmation from the user
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdLaunchDocSearch_Click()
    frmDoctorSearchChanneling.Show
End Sub



Private Sub cmdPrint_Click()
    
    On Error GoTo e
    DataEnvironment1.Commands("ChannelingReceipt").Parameters(0) = txtFirstName
    DataEnvironment1.Commands("ChannelingReceipt").Parameters(1) = txtLastName
    RptChannelingReceipt.Show
    DataEnvironment1.rsChannelingReceipt.Close
        
    Unload Me
    Exit Sub
e:
    If Err.Number <> 3704 Then
        MsgBox Err.Description & "" & Err.Number, vbCritical
    End If

End Sub

Private Sub cmdSave_Click()
    
    
    Call All_Appointments
    With rsAllAppointments
    
        'Making sure that the user wants to save the record
        If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then

    
            .AddNew
            
            .Fields(0) = txtDoctorID.Text
            .Fields(1) = dtpChosenDate.Value
            .Fields(2) = txtTokenNo.Text
            .Fields(3) = txtSpecialization.Text
            .Fields(4) = txtChannelingCharges.Text
            .Fields(5) = txtAppointmentDuration.Text
            .Fields(6) = txtChosenDay.Text
            .Fields(7) = dtpStartTime.Value
            .Fields(8) = dtpEndTime.Value
            .Fields(9) = dtpAppointmentStartTime.Hour & ":" & dtpAppointmentStartTime.Minute & ":" & dtpAppointmentStartTime.Second
            .Fields(10) = dtpAppointmentEndTime.Hour & ":" & dtpAppointmentEndTime.Minute & ":" & dtpAppointmentEndTime.Second
            .Fields(11) = txtFirstName.Text
            .Fields(12) = txtLastName.Text
            .Fields(13) = txtContactNo.Text
            .Fields(14) = DateTime.Date
        
            .Update
            
            .Requery
            
            'Display Success Message
            MsgBox "The Record Was Saved Successfully!", vbInformation, "Succesful Save Procedure"
            
        End If
        
    End With
    
    Call Channeling_Appointments
    
    Set dgrdChannelingInfo.DataSource = rsChannelingAppointments
    
    cmdPrint.Enabled = True

    
End Sub

Private Sub dtpChosenDate_CloseUp()

    dtpChosenDate.MinDate = DateTime.Date
    
    Call Channeling_Appointments
    rsChannelingAppointments.Filter = "ChosenDate Like '" & dtpChosenDate.Value & "'"
    
    Set dgrdChannelingInfo.DataSource = rsChannelingAppointments
    
    iNumberOfRecords = rsChannelingAppointments.RecordCount
    
    If iNumberOfRecords = 0 Then
    
        dtpAppointmentStartTime.Value = dtpStartTime.Value
        dtpAppointmentEndTime.Value = dtpStartTime.Value
        dtpAppointmentEndTime.Minute = Val(dtpAppointmentEndTime.Minute) + Val(txtAppointmentDuration.Text)
        txtTokenNo.Text = "1"
        
        'Here, I am enabling the textfields where I will be entering the patient's information
        txtFirstName.Enabled = True
        txtLastName.Enabled = True
        txtContactNo.Enabled = True
        txtTokenNo.Enabled = True

        
    Else
    
    On Error GoTo error_handler
    
        'Here, I am enabling the textfields where I will be entering the patient's information
        txtFirstName.Enabled = True
        txtLastName.Enabled = True
        txtContactNo.Enabled = True
        txtTokenNo.Enabled = True

        txtTokenNo.Text = iNumberOfRecords + 1
        
        Dim gCol As MSDataGridLib.Column
        Set gCol = dgrdChannelingInfo.Columns("AppointmentEndTime")
        dtpAppointmentStartTime.Value = gCol.CellValue(dgrdChannelingInfo.RowBookmark(iNumberOfRecords - 1))
        
        If dtpAppointmentStartTime.Value = dtpEndTime.Value Then
            MsgBox "All Appointment Slots Have Been Booked! Please Choose Another Day!", vbCritical, "All Slots Booked!"
            Unload Me
            Exit Sub
            
        Else
        
            dtpAppointmentEndTime.Value = dtpAppointmentStartTime.Value
            dtpAppointmentEndTime.Minute = Val(dtpAppointmentEndTime.Minute) + Val(txtAppointmentDuration.Text)
        
        End If
    
    End If
    
    Exit Sub
    
error_handler:
    
    'If dtpAppointmentStartTime.Value = DateTime.Time Then
        
    dtpAppointmentEndTime.Minute = "00"
    dtpAppointmentEndTime.Hour = dtpAppointmentEndTime.Hour + 1
    
        
End Sub


Private Sub tmrErrMsg_Timer()
    
    Static i As Integer
    
    If i < 200000 Then     'Validation Msg Viewing Time Period
        picInvalidDataMsg.Visible = False
        picInvalidKeyMsg.Visible = False
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If
    
End Sub





Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeyMsg.Top = 7320    'Validation Note View
        picInvalidKeyMsg.Visible = True
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
        picInvalidDataMsg.Top = 6600    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub



Private Sub txtLastName_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets
    
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 6960    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub
