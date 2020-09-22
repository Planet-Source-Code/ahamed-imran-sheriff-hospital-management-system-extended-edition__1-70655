VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSetUpChannelingSchedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor's Schedule Setup Wizard"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   9525
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7680
      Width           =   1695
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CheckBox chkSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   28
      Top             =   6750
      Width           =   255
   End
   Begin VB.CheckBox chkSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   24
      Top             =   6030
      Width           =   255
   End
   Begin VB.CheckBox chkFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   20
      Top             =   5310
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
   Begin VB.CheckBox chkWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   12
      Top             =   3870
      Width           =   255
   End
   Begin VB.CheckBox chkTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   3150
      Width           =   255
   End
   Begin VB.CheckBox chkMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   2430
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpTuesdayStartTime 
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpMondayEndTime 
      Height          =   285
      Left            =   6600
      TabIndex        =   7
      Top             =   2400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpMondayStartTime 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   2400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   39563.5416203704
   End
   Begin MSComCtl2.DTPicker dtpTuesdayEndTime 
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpWednesdayStartTime 
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpWednesdayEndTime 
      Height          =   285
      Left            =   6600
      TabIndex        =   15
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpThursdayStartTime 
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Top             =   4560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpThursdayEndTime 
      Height          =   285
      Left            =   6600
      TabIndex        =   19
      Top             =   4560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpFridayStartTime 
      Height          =   285
      Left            =   3600
      TabIndex        =   21
      Top             =   5280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpFridayEndTime 
      Height          =   285
      Left            =   6600
      TabIndex        =   23
      Top             =   5280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSaturdayStartTime 
      Height          =   285
      Left            =   3600
      TabIndex        =   25
      Top             =   6000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSaturdayEndTime 
      Height          =   285
      Left            =   6600
      TabIndex        =   27
      Top             =   6000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSundayStartTime 
      Height          =   285
      Left            =   3600
      TabIndex        =   29
      Top             =   6720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpSundayEndTime 
      Height          =   285
      Left            =   6600
      TabIndex        =   31
      Top             =   6720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   54394882
      CurrentDate     =   36494
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
      TabIndex        =   33
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Line Line7 
      X1              =   240
      X2              =   9240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   9240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   9240
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   240
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line3 
      X1              =   9240
      X2              =   9240
      Y1              =   1320
      Y2              =   7320
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
      Left            =   1440
      TabIndex        =   30
      Top             =   6720
      Width           =   1335
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
      Left            =   1440
      TabIndex        =   26
      Top             =   6000
      Width           =   1095
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
      Left            =   1440
      TabIndex        =   22
      Top             =   5280
      Width           =   1215
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
      Left            =   1440
      TabIndex        =   18
      Top             =   4560
      Width           =   1335
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
      Left            =   1440
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
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
      Left            =   1440
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   6240
      X2              =   6240
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   3240
      Y1              =   1320
      Y2              =   7320
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
      Left            =   7200
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
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
      Left            =   4200
      TabIndex        =   5
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
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
      Left            =   960
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
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
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   -120
      Picture         =   "frmSetUpSchedule.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Left            =   -120
      Picture         =   "frmSetUpSchedule.frx":00A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
End
Attribute VB_Name = "frmSetUpChannelingSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Setup Channeling Schedule Wizard
'Programmer: Deshan Subasinghe
'Quality Assurance Engineer (Testing): Imran Sheriff
'Start Date: 14/04/08
'Date Of Last Modification: 18/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Channeling_Appointments Table
'---------------------------------------------------------------------------


Private Sub Form_Load()
    
    
    Call Connection 'Call Connection Procedure To Establish Connection With The Database
    
    Call Doctor_Schedule    'Calling the Doctor_Schedule Procedure To Interact With The Recordset
    
    With rsDoctorSchedule
    
        .MoveFirst  'Move To The First Record
    
        While .EOF = False  'Running Through All The Records
            
            'Here, I Am Checking The Doctor ID In Order To I Find A Match
            If .Fields(1).Value = frmDoctorScheduleMaintenance.txtDoctorID.Text Then
                
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


Private Sub cmdClose_Click()    'Closing the Wizard
    
    Unload Me
    
End Sub

Private Sub cmdSave_Click() 'If the Save Button is Clicked
    
    With rsDoctorSchedule
    
        .MoveFirst  'Moving to the first record
        
        While .EOF = False  'Running Through All The Records
            
            'Here, I Am Checking The Doctor ID In Order To Find A Match.
            'When I Find A Match, I Will Be Setting The Values Of The DateTime
            'Pickers Accordingly.
            If .Fields(1).Value = frmDoctorScheduleMaintenance.txtDoctorID.Text Then
            
                'Monday
                If chkMonday.Value = 0 Then
                    .Fields(2).Value = Null
                    .Fields(3).Value = Null
                Else
                    .Fields(2).Value = dtpMondayStartTime.Hour & ":" & dtpMondayStartTime.Minute & ":" & dtpMondayStartTime.Second
                    .Fields(3).Value = dtpMondayEndTime.Hour & ":" & dtpMondayEndTime.Minute & ":" & dtpMondayEndTime.Second
                End If
                
            
                'Tuesday
                If chkTuesday.Value = 0 Then
                    .Fields(4).Value = Null
                    .Fields(5).Value = Null
                Else
                    .Fields(4).Value = dtpTuesdayStartTime.Hour & ":" & dtpTuesdayStartTime.Minute & ":" & dtpTuesdayStartTime.Second
                    .Fields(5).Value = dtpTuesdayEndTime.Hour & ":" & dtpTuesdayEndTime.Minute & ":" & dtpTuesdayEndTime.Second
                End If
                
                
                'Wednesday
                If chkWednesday.Value = 0 Then
                    .Fields(6).Value = Null
                    .Fields(7).Value = Null
                Else
                    .Fields(6).Value = dtpWednesdayStartTime.Hour & ":" & dtpWednesdayStartTime.Minute & ":" & dtpWednesdayStartTime.Second
                    .Fields(7).Value = dtpWednesdayEndTime.Hour & ":" & dtpWednesdayEndTime.Minute & ":" & dtpWednesdayEndTime.Second
                End If
                
                
                'Thursday
                If chkThursday.Value = 0 Then
                    .Fields(8).Value = Null
                    .Fields(9).Value = Null
                Else
                    .Fields(8).Value = dtpThursdayStartTime.Hour & ":" & dtpThursdayStartTime.Minute & ":" & dtpThursdayStartTime.Second
                    .Fields(9).Value = dtpThursdayEndTime.Hour & ":" & dtpThursdayEndTime.Minute & ":" & dtpThursdayEndTime.Second
                End If
                
                
                'Friday
                If chkFriday.Value = 0 Then
                    .Fields(10).Value = Null
                    .Fields(11).Value = Null
                Else
                    .Fields(10).Value = dtpFridayStartTime.Hour & ":" & dtpFridayStartTime.Minute & ":" & dtpFridayStartTime.Second
                    .Fields(11).Value = dtpFridayEndTime.Hour & ":" & dtpFridayEndTime.Minute & ":" & dtpFridayEndTime.Second
                End If
                
                
                'Saturday
                If chkSaturday.Value = 0 Then
                    .Fields(12).Value = Null
                    .Fields(13).Value = Null
                Else
                    .Fields(12).Value = dtpSaturdayStartTime.Hour & ":" & dtpSaturdayStartTime.Minute & ":" & dtpSaturdayStartTime.Second
                    .Fields(13).Value = dtpSaturdayEndTime.Hour & ":" & dtpSaturdayEndTime.Minute & ":" & dtpSaturdayEndTime.Second
                End If
               
               
                'Sunday
                If chkSunday.Value = 0 Then
                    .Fields(14).Value = Null
                    .Fields(15).Value = Null
                Else
                    .Fields(14).Value = dtpSundayStartTime.Hour & ":" & dtpSundayStartTime.Minute & ":" & dtpSundayStartTime.Second
                    .Fields(15).Value = dtpSundayEndTime.Hour & ":" & dtpSundayEndTime.Minute & ":" & dtpSundayEndTime.Second
                End If
                
                .Update 'Updating the Recordset
                
                Unload Me
                                
                Exit Sub
            
            End If
            
            .MoveNext   'Moving To The Next Record
        
        Wend
        
        
        
        'The Adding of a New Record Obviously Takes Place Only If There Is
        'No Matching Doctor ID To Be Found, Which Would Mean That The Doctor
        'Is New
        
        .AddNew 'Adding a New Record
        
        'Adding the Doctor ID Into The Relevant Field
        .Fields(1).Value = frmDoctorScheduleMaintenance.txtDoctorID.Text
        
        
        'Monday
        If chkMonday.Value = 0 Then
            .Fields(2).Value = Null
            .Fields(3).Value = Null
        Else
            .Fields(2).Value = dtpMondayStartTime.Hour & ":" & dtpMondayStartTime.Minute & ":" & dtpMondayStartTime.Second
            .Fields(3).Value = dtpMondayEndTime.Hour & ":" & dtpMondayEndTime.Minute & ":" & dtpMondayEndTime.Second
        End If
                
                
        'Tuesday
        If chkTuesday.Value = 0 Then
            .Fields(4).Value = Null
            .Fields(5).Value = Null
        Else
            .Fields(4).Value = dtpTuesdayStartTime.Hour & ":" & dtpTuesdayStartTime.Minute & ":" & dtpTuesdayStartTime.Second
            .Fields(5).Value = dtpTuesdayEndTime.Hour & ":" & dtpTuesdayEndTime.Minute & ":" & dtpTuesdayEndTime.Second
        End If
                
                
        'Wednesday
        If chkWednesday.Value = 0 Then
            .Fields(6).Value = Null
            .Fields(7).Value = Null
        Else
            .Fields(6).Value = dtpWednesdayStartTime.Hour & ":" & dtpWednesdayStartTime.Minute & ":" & dtpWednesdayStartTime.Second
            .Fields(7).Value = dtpWednesdayEndTime.Hour & ":" & dtpWednesdayEndTime.Minute & ":" & dtpWednesdayEndTime.Second
        End If
                
                
        'Thursday
        If chkThursday.Value = 0 Then
            .Fields(8).Value = Null
            .Fields(9).Value = Null
        Else
            .Fields(8).Value = dtpThursdayStartTime.Hour & ":" & dtpThursdayStartTime.Minute & ":" & dtpThursdayStartTime.Second
            .Fields(9).Value = dtpThursdayEndTime.Hour & ":" & dtpThursdayEndTime.Minute & ":" & dtpThursdayEndTime.Second
        End If
                
                
        'Friday
        If chkFriday.Value = 0 Then
            .Fields(10).Value = Null
            .Fields(11).Value = Null
        Else
            .Fields(10).Value = dtpFridayStartTime.Hour & ":" & dtpFridayStartTime.Minute & ":" & dtpFridayStartTime.Second
            .Fields(11).Value = dtpFridayEndTime.Hour & ":" & dtpFridayEndTime.Minute & ":" & dtpFridayEndTime.Second
        End If
                
                
        'Saturday
        If chkSaturday.Value = 0 Then
            .Fields(12).Value = Null
            .Fields(13).Value = Null
        Else
            .Fields(12).Value = dtpSaturdayStartTime.Hour & ":" & dtpSaturdayStartTime.Minute & ":" & dtpSaturdayStartTime.Second
            .Fields(13).Value = dtpSaturdayEndTime.Hour & ":" & dtpSaturdayEndTime.Minute & ":" & dtpSaturdayEndTime.Second
        End If
               
               
        'Sunday
        If chkSunday.Value = 0 Then
            .Fields(14).Value = Null
            .Fields(15).Value = Null
        Else
            .Fields(14).Value = dtpSundayStartTime.Hour & ":" & dtpSundayStartTime.Minute & ":" & dtpSundayStartTime.Second
            .Fields(15).Value = dtpSundayEndTime.Hour & ":" & dtpSundayEndTime.Minute & ":" & dtpSundayEndTime.Second
        End If
                
        .Update 'Updating The Record
                
        
    End With
    
    Unload Me   'Closing the form
                
End Sub




