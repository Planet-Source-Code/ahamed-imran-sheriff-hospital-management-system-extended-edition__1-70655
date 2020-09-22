VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelectADay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Day"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optSunday 
      Enabled         =   0   'False
      Height          =   195
      Left            =   960
      TabIndex        =   42
      Top             =   6720
      Width           =   180
   End
   Begin VB.OptionButton optSaturday 
      Enabled         =   0   'False
      Height          =   195
      Left            =   960
      TabIndex        =   41
      Top             =   6000
      Width           =   180
   End
   Begin VB.OptionButton optFriday 
      Enabled         =   0   'False
      Height          =   195
      Left            =   960
      TabIndex        =   40
      Top             =   5280
      Width           =   180
   End
   Begin VB.OptionButton optThursday 
      Enabled         =   0   'False
      Height          =   195
      Left            =   960
      TabIndex        =   39
      Top             =   4560
      Width           =   180
   End
   Begin VB.OptionButton optWednesday 
      Enabled         =   0   'False
      Height          =   195
      Left            =   960
      TabIndex        =   38
      Top             =   3790
      Width           =   180
   End
   Begin VB.OptionButton optTuesday 
      Enabled         =   0   'False
      Height          =   195
      Left            =   960
      TabIndex        =   37
      Top             =   3120
      Width           =   180
   End
   Begin VB.OptionButton optMonday 
      Enabled         =   0   'False
      Height          =   195
      Left            =   960
      TabIndex        =   36
      Top             =   2450
      Width           =   180
   End
   Begin VB.CheckBox chkMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   2430
      Width           =   255
   End
   Begin VB.CheckBox chkTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   3150
      Width           =   255
   End
   Begin VB.CheckBox chkWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   7
      Top             =   3870
      Width           =   255
   End
   Begin VB.CheckBox chkThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   4590
      Width           =   255
   End
   Begin VB.CheckBox chkFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   5310
      Width           =   255
   End
   Begin VB.CheckBox chkSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   6030
      Width           =   255
   End
   Begin VB.CheckBox chkSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   6750
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&OK"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpTuesdayStartTime 
      Height          =   285
      Left            =   4200
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpMondayEndTime 
      Height          =   285
      Left            =   6840
      TabIndex        =   11
      Top             =   2400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpMondayStartTime 
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Top             =   2400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   39563.5416203704
   End
   Begin MSComCtl2.DTPicker dtpTuesdayEndTime 
      Height          =   285
      Left            =   6840
      TabIndex        =   13
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpWednesdayStartTime 
      Height          =   285
      Left            =   4200
      TabIndex        =   14
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpWednesdayEndTime 
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpThursdayStartTime 
      Height          =   285
      Left            =   4200
      TabIndex        =   16
      Top             =   4560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpThursdayEndTime 
      Height          =   285
      Left            =   6840
      TabIndex        =   17
      Top             =   4560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpFridayStartTime 
      Height          =   285
      Left            =   4200
      TabIndex        =   18
      Top             =   5280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpFridayEndTime 
      Height          =   285
      Left            =   6840
      TabIndex        =   19
      Top             =   5280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSaturdayStartTime 
      Height          =   285
      Left            =   4200
      TabIndex        =   20
      Top             =   6000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSaturdayEndTime 
      Height          =   285
      Left            =   6840
      TabIndex        =   21
      Top             =   6000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSundayStartTime 
      Height          =   285
      Left            =   4200
      TabIndex        =   22
      Top             =   6720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSundayEndTime 
      Height          =   285
      Left            =   6840
      TabIndex        =   23
      Top             =   6720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   70320130
      CurrentDate     =   36494
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select-A-Day Wizard"
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
      Left            =   3480
      TabIndex        =   43
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Day"
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
      Left            =   600
      TabIndex        =   35
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line8 
      X1              =   2040
      X2              =   2040
      Y1              =   1320
      Y2              =   7320
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
      Left            =   2400
      TabIndex        =   34
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
      Left            =   2760
      TabIndex        =   33
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblStartTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time"
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
      Left            =   4800
      TabIndex        =   32
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblEndTime 
      BackStyle       =   0  'Transparent
      Caption         =   "End Time"
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
      Left            =   7440
      TabIndex        =   31
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   3960
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   6600
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
      Left            =   2760
      TabIndex        =   30
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
      Left            =   2760
      TabIndex        =   29
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
      Left            =   2760
      TabIndex        =   28
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
      Left            =   2760
      TabIndex        =   27
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
      Left            =   2760
      TabIndex        =   26
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
      Left            =   2760
      TabIndex        =   25
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   9240
      X2              =   9240
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   240
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   9240
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   9240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line7 
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
      Left            =   2760
      TabIndex        =   24
      Top             =   8640
      Width           =   3735
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   0
      Picture         =   "frmSelectADay.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Index           =   0
      Left            =   0
      Picture         =   "frmSelectADay.frx":00A2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9810
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Channeling Times Setup Wizard"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Width           =   4935
   End
End
Attribute VB_Name = "frmSelectADay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    
    
    If frmChannelingAppointments.txtChosenDay.Text = "Monday" Then
        optMonday.Value = True
    ElseIf frmChannelingAppointments.txtChosenDay.Text = "Tuesday" Then
        optTuesday.Value = True
    ElseIf frmChannelingAppointments.txtChosenDay.Text = "Wednesday" Then
        optWednesday.Value = True
    ElseIf frmChannelingAppointments.txtChosenDay.Text = "Thursday" Then
        optThursday.Value = True
    ElseIf frmChannelingAppointments.txtChosenDay.Text = "Friday" Then
        optFriday.Value = True
    ElseIf frmChannelingAppointments.txtChosenDay.Text = "Saturday" Then
        optSaturday.Value = True
    ElseIf frmChannelingAppointments.txtChosenDay.Text = "Sunday" Then
        optSunday.Value = True
    End If
    
    
    Call Connection 'Call Connection Procedure To Establish Connection With The Database
    
    Call Doctor_Schedule    'Calling the Doctor_Schedule Procedure To Interact With The Recordset
    
    With rsDoctorSchedule
    
        .MoveFirst  'Move To The First Record
    
        While .EOF = False  'Running Through All The Records
            
            'Here, I Am Checking The Doctor ID In Order To I Find A Match
            If .Fields(1).Value = frmChannelingAppointments.txtDoctorID.Text Then
                
                'In The Following Blocks Of If Else Conditions, I am enabling
                'the DateTime Pickers if the times have already been set.
                'Therefore for a new doctor, none of the DateTime Pickers
                'will be enabled.
                If .Fields(2).Value <> "" Then
                    chkMonday.Value = 1
                    dtpMondayStartTime.Enabled = True
                    dtpMondayEndTime.Enabled = True
                    dtpMondayStartTime.Value = .Fields(2).Value
                    dtpMondayEndTime.Value = .Fields(3).Value
                End If
                
                If .Fields(4).Value <> "" Then
                    chkTuesday.Value = 1
                    dtpTuesdayStartTime.Enabled = True
                    dtpTuesdayEndTime.Enabled = True
                    dtpTuesdayStartTime.Value = .Fields(4).Value
                    dtpTuesdayEndTime.Value = .Fields(5).Value
                End If
                    
                If .Fields(6).Value <> "" Then
                    chkWednesday.Value = 1
                    dtpWednesdayStartTime.Enabled = True
                    dtpWednesdayEndTime.Enabled = True
                    dtpWednesdayStartTime.Value = .Fields(6).Value
                    dtpWednesdayEndTime.Value = .Fields(7).Value
                End If
                
                If .Fields(8).Value <> "" Then
                    chkThursday.Value = 1
                    dtpThursdayStartTime.Enabled = True
                    dtpThursdayEndTime.Enabled = True
                    dtpThursdayStartTime.Value = .Fields(8).Value
                    dtpThursdayEndTime.Value = .Fields(9).Value
                End If
                
                If .Fields(10).Value <> "" Then
                    chkFriday.Value = 1
                    dtpFridayStartTime.Enabled = True
                    dtpFridayEndTime.Enabled = True
                    dtpFridayStartTime.Value = .Fields(10).Value
                    dtpFridayEndTime.Value = .Fields(11).Value
                End If
                
                If .Fields(12).Value <> "" Then
                    chkSaturday.Value = 1
                    dtpSaturdayStartTime.Enabled = True
                    dtpSaturdayEndTime.Enabled = True
                    dtpSaturdayStartTime.Value = .Fields(12).Value
                    dtpSaturdayEndTime.Value = .Fields(13).Value
                End If
                
                If .Fields(14).Value <> "" Then
                    chkSunday.Value = 1
                    dtpSundayStartTime.Enabled = True
                    dtpSundayEndTime.Enabled = True
                    dtpSundayStartTime.Value = .Fields(14).Value
                    dtpSundayEndTime.Value = .Fields(15).Value
                End If
            
            End If
                
            .MoveNext   'Moving to the next record
        
        Wend
        
    End With
    
    If chkMonday.Value = 1 Then
        optMonday.Enabled = True
    End If
    
    If chkTuesday.Value = 1 Then
        optTuesday.Enabled = True
    End If
    
    If chkWednesday.Value = 1 Then
        optWednesday.Enabled = True
    End If
    
    If chkThursday.Value = 1 Then
        optThursday.Enabled = True
    End If
    
    If chkFriday.Value = 1 Then
        optFriday.Enabled = True
    End If
    
    If chkSaturday.Value = 1 Then
        optSaturday.Enabled = True
    End If
    
    If chkSunday.Value = 1 Then
        optSunday.Enabled = True
    End If
                
End Sub



Private Sub chkMonday_Click()   'Enabling the DateTime Pickers

    If chkMonday.Value = 1 Then
        dtpMondayStartTime.Enabled = True
        dtpMondayEndTime.Enabled = True
    Else
        dtpMondayStartTime.Enabled = False
        dtpMondayEndTime.Enabled = False
    End If
    
End Sub



Private Sub chkTuesday_Click()  'Enabling the DateTime Pickers

    If chkTuesday.Value = 1 Then
        dtpTuesdayStartTime.Enabled = True
        dtpTuesdayEndTime.Enabled = True
    Else
        dtpTuesdayStartTime.Enabled = False
        dtpTuesdayEndTime.Enabled = False
    End If
    
End Sub

Private Sub chkWednesday_Click()    'Enabling the DateTime Pickers

    If chkWednesday.Value = 1 Then
        dtpWednesdayStartTime.Enabled = True
        dtpWednesdayEndTime.Enabled = True
    Else
        dtpWednesdayStartTime.Enabled = False
        dtpWednesdayEndTime.Enabled = False
    End If
    
End Sub

Private Sub chkThursday_Click() 'Enabling the DateTime Pickers
    
    If chkThursday.Value = 1 Then
        dtpThursdayStartTime.Enabled = True
        dtpThursdayEndTime.Enabled = True
    Else
        dtpThursdayStartTime.Enabled = False
        dtpThursdayEndTime.Enabled = False
    End If
    
End Sub

Private Sub chkFriday_Click()   'Enabling the DateTime Pickers
    
    If chkFriday.Value = 1 Then
        dtpFridayStartTime.Enabled = True
        dtpFridayEndTime.Enabled = True
    Else
        dtpFridayStartTime.Enabled = False
        dtpFridayEndTime.Enabled = False
    End If
    
End Sub

Private Sub chkSaturday_Click() 'Enabling the DateTime Pickers

    If chkSaturday.Value = 1 Then
        dtpSaturdayStartTime.Enabled = True
        dtpSaturdayEndTime.Enabled = True
    Else
        dtpSaturdayStartTime.Enabled = False
        dtpSaturdayEndTime.Enabled = False
    End If
    
End Sub

Private Sub chkSunday_Click()   'Enabling the DateTime Pickers
    
    If chkSunday.Value = 1 Then
        dtpSundayStartTime.Enabled = True
        dtpSundayEndTime.Enabled = True
    Else
        dtpSundayStartTime.Enabled = False
        dtpSundayEndTime.Enabled = False
    End If
    
        
End Sub


Private Sub cmdOK_Click()
    
    Dim checkFlag As Boolean    'This variable wil help me to decide if any of the option buttons have been selected
    checkFlag = False
    Dim ctrl As Control 'This is a control variable which I will be using to run through all the option buttons
    On Error Resume Next
    For Each ctrl In Controls   'This is a For loop that will check if any of the option buttons have been selected
        If TypeOf ctrl Is OptionButton Then
            If ctrl.Value = True Then
                checkFlag = True
                Exit For
            End If
        End If
    Next
    
    If checkFlag = False Then
        'Display Error Message
        MsgBox "Error! You have Not Selected An Option Button!", vbCritical, "Error! No Selection!"
        Exit Sub
    End If
        
    'Enabling relevant components on the Channeling form
    frmChannelingAppointments.txtChosenDay.Enabled = True
    frmChannelingAppointments.dtpChosenDate.Enabled = True
        
    'Setting the relevant start times and end times on the channeling form
    If optMonday.Value = True Then
        frmChannelingAppointments.txtChosenDay.Text = "Monday"
        frmChannelingAppointments.dtpStartTime.Value = dtpMondayStartTime.Value
        frmChannelingAppointments.dtpEndTime.Value = dtpMondayEndTime.Value
    End If
    
    
    If optTuesday.Value = True Then
        frmChannelingAppointments.txtChosenDay.Text = "Tuesday"
        frmChannelingAppointments.dtpStartTime.Value = dtpTuesdayStartTime.Value
        frmChannelingAppointments.dtpEndTime.Value = dtpTuesdayEndTime.Value
    End If
    
    
    If optWednesday.Value = True Then
        frmChannelingAppointments.txtChosenDay.Text = "Wednesday"
        frmChannelingAppointments.dtpStartTime.Value = dtpWednesdayStartTime.Value
        frmChannelingAppointments.dtpEndTime.Value = dtpWednesdayEndTime.Value
    End If


    If optThursday.Value = True Then
        frmChannelingAppointments.txtChosenDay.Text = "Thursday"
        frmChannelingAppointments.dtpStartTime.Value = dtpThursdayStartTime.Value
        frmChannelingAppointments.dtpEndTime.Value = dtpThursdayEndTime.Value
    End If
    
    
    If optFriday.Value = True Then
        frmChannelingAppointments.txtChosenDay.Text = "Friday"
        frmChannelingAppointments.dtpStartTime.Value = dtpFridayStartTime.Value
        frmChannelingAppointments.dtpEndTime.Value = dtpFridayEndTime.Value
    End If


    If optSaturday.Value = True Then
        frmChannelingAppointments.txtChosenDay.Text = "Saturday"
        frmChannelingAppointments.dtpStartTime.Value = dtpSaturdayStartTime.Value
        frmChannelingAppointments.dtpEndTime.Value = dtpSaturdayEndTime.Value
    End If


    If optSunday.Value = True Then
        frmChannelingAppointments.txtChosenDay.Text = "Sunday"
        frmChannelingAppointments.dtpStartTime.Value = dtpSundayStartTime.Value
        frmChannelingAppointments.dtpEndTime.Value = dtpSundayEndTime.Value
    End If

    Unload Me

End Sub

Private Sub cmdClose_Click()    'Closing the Wizard
    
    Unload Me
    
End Sub

