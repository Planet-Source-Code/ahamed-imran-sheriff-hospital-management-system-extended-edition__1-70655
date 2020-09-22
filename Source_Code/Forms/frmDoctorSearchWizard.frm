VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmDoctorSearchWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Search Wizard"
   ClientHeight    =   8850
   ClientLeft      =   2925
   ClientTop       =   1470
   ClientWidth     =   8835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   8835
   Begin VB.ComboBox cboSearchType 
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
      ItemData        =   "frmDoctorSearchWizard.frx":0000
      Left            =   1800
      List            =   "frmDoctorSearchWizard.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1290
      Width           =   2295
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   5400
      TabIndex        =   14
      Top             =   1290
      Width           =   2295
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdApply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Apply"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dgrdDoctorsInfoTable 
      Height          =   4815
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8493
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
      Caption         =   "Doctor's Information Table"
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
   Begin VB.Label lblSearchText 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For :"
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
      Left            =   4200
      TabIndex        =   17
      Top             =   1335
      Width           =   1215
   End
   Begin VB.Shape shpSearchFrame 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   600
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria :"
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
      TabIndex        =   16
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Search Wizard"
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
      Left            =   2880
      TabIndex        =   13
      Top             =   240
      Width           =   2655
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   0
      Left            =   -360
      Picture         =   "frmDoctorSearchWizard.frx":0093
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
      Left            =   2520
      TabIndex        =   12
      Top             =   240
      Width           =   3975
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
      Index           =   1
      Left            =   2520
      TabIndex        =   11
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Index           =   0
      Left            =   -360
      Picture         =   "frmDoctorSearchWizard.frx":0135
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   -360
      Picture         =   "frmDoctorSearchWizard.frx":01D3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Appointments Maintenance Wizard"
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
      Left            =   2280
      TabIndex        =   10
      Top             =   240
      Width           =   4335
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   9000
      X2              =   9000
      Y1              =   2160
      Y2              =   7320
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   2160
      Y2              =   7320
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   0
      X2              =   9000
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line7 
      Index           =   0
      X1              =   0
      X2              =   9000
      Y1              =   2160
      Y2              =   2160
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
      Left            =   2760
      TabIndex        =   9
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "By : "
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
      Left            =   5040
      TabIndex        =   8
      Top             =   1470
      Width           =   615
   End
   Begin VB.Label lblSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Doctor : "
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
      Left            =   1440
      TabIndex        =   7
      Top             =   1455
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   720
      Top             =   1200
      Width           =   7455
   End
   Begin VB.Label Label13 
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
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2535
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Type"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   2535
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Channeling Days"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3255
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Time In : "
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
      Left            =   3600
      TabIndex        =   3
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out : "
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
      Left            =   6360
      TabIndex        =   2
      Top             =   3285
      Width           =   975
   End
End
Attribute VB_Name = "frmDoctorSearchWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'This variable will determine if the DataGrid has been clicked or not
Dim Flag As Boolean


Private Sub cmdClose_Click()    'This procedure will close the Wizard

    Unload Me   'Unloading the Wizard
    
    'Calling the clearAllFields procedure in the Doctors Maintenance Form
    frmDoctorsMaintenance.clearAllFields
    
    'Calling the Form_Load procedure in the Doctors Maintenance Form
    frmDoctorsMaintenance.Form_Load
    
End Sub

Private Sub dgrdDoctorsInfoTable_Click()    'This procedure is executed if the user clicks the DataGrid
    
    'Setting the Flag variable to True, to indicate that the user
    'has clicked the DataGrid
    Flag = True
    
End Sub


Private Sub Form_Load() 'Form Load Procedure

    Flag = False    'The Flag variable is being initialized to False
    
    Call Doctors_Maintenance    'Calling the Doctors_Maintenance Procedure to interact with the recordset
    
    Set dgrdDoctorsInfoTable.DataSource = rsDoctorsMaintenance  'Setting the DataSource of the DataGrid
    
End Sub


Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsDoctorsMaintenance
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[DoctorID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[FirstName] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
                    .Filter = "[Surname] Like '" & txtSearch.Text & "%" & "'"
                Case 3:
                    .Filter = "[Gender] Like '" & txtSearch.Text & "%" & "'"
                Case 4:
                    .Filter = "[NICNumber] Like '" & txtSearch.Text & "%" & "'"
                Case 5:
                    .Filter = "[LicenceNo] Like '" & txtSearch.Text & "%" & "'"
                Case 6:
                    .Filter = "[Specialization] Like '" & txtSearch.Text & "%" & "'"
                Case 7:
                    .Filter = "[DoctorCategory] Like '" & txtSearch.Text & "%" & "'"
            End Select
    
        End With
        
        Set dgrdDoctorsInfoTable.DataSource = rsDoctorsMaintenance  'Setting the DataSource of the DataGrid
    
    Else
    
        Form_Load   'Calling the Form_Load procedure
        
    End If
    
End Sub


Private Sub cmdApply_Click()    'This code is executed when the user clicks the Apply Button
    
    'Here, I am checkin to see if the user has chosen a record
    If Flag = True And rsDoctorsMaintenance.RecordCount > 0 Then
    
        With rsDoctorsMaintenance
        
            'Reset the textfields with the selected record
            frmDoctorsMaintenance.txtDoctorID.Text = .Fields(0).Value
            frmDoctorsMaintenance.txtFirstName.Text = .Fields(1).Value
            frmDoctorsMaintenance.txtSurname.Text = .Fields(2).Value
            frmDoctorsMaintenance.cboGender.Text = .Fields(3).Value
            frmDoctorsMaintenance.dtpDateOfBirth.Value = .Fields(4).Value
            frmDoctorsMaintenance.txtNICNumber.Text = .Fields(5).Value
            frmDoctorsMaintenance.txtAddress.Text = .Fields(6).Value
            frmDoctorsMaintenance.txtHomePhone.Text = .Fields(7).Value
            frmDoctorsMaintenance.txtMobPhone.Text = .Fields(8).Value
            frmDoctorsMaintenance.txtLicenseNo.Text = .Fields(9).Value
            frmDoctorsMaintenance.txtDoctorSpecialization.Text = .Fields(10).Value
            frmDoctorsMaintenance.cboDoctorCategory.Text = .Fields(11).Value
            frmDoctorsMaintenance.txtServiceCharges.Text = .Fields(12).Value
            frmDoctorsMaintenance.txtChannelingCharges.Text = .Fields(13).Value
            frmDoctorsMaintenance.cboAppointmentDuration.Text = .Fields(14).Value
            frmDoctorsMaintenance.txtReferringCharges.Text = .Fields(15).Value
            
            
            'Here, I am ensuring that the SetUpDoctor'sVisitingDays Button
            'will be enabled only if the Doctor is a Visiting Doctor
            If frmDoctorsMaintenance.cboDoctorCategory.Text = "Visiting" Then
                frmDoctorsMaintenance.cmdSetUpDocSchedule.Enabled = True
            Else
                frmDoctorsMaintenance.cmdSetUpDocSchedule.Enabled = False
            End If
            
            'Here, I am ensuring that certain components will be disabled if the doctor is a "Referring Doctor"
            If frmDoctorsMaintenance.cboDoctorCategory.Text = "Referring" Then
                frmDoctorsMaintenance.lblServiceCharges.Enabled = False
                frmDoctorsMaintenance.txtServiceCharges.Enabled = False
                frmDoctorsMaintenance.lblChannelingCharges.Enabled = False
                frmDoctorsMaintenance.txtChannelingCharges.Enabled = False
                frmDoctorsMaintenance.lblAppointmentDuration.Enabled = False
                frmDoctorsMaintenance.cboAppointmentDuration.Enabled = False
            Else
                frmDoctorsMaintenance.lblServiceCharges.Enabled = True
                frmDoctorsMaintenance.txtServiceCharges.Enabled = True
                frmDoctorsMaintenance.lblChannelingCharges.Enabled = True
                frmDoctorsMaintenance.txtChannelingCharges.Enabled = True
                frmDoctorsMaintenance.lblAppointmentDuration.Enabled = True
                frmDoctorsMaintenance.cboAppointmentDuration.Enabled = True
            End If
            
            
            'Here, I am ensuring that the Referring Charges textfield will be disabled if the Doctor is a "Permanent" Doctor
            If frmDoctorsMaintenance.cboDoctorCategory.Text = "Permanent" Then
                frmDoctorsMaintenance.lblReferringCharges.Enabled = False
                frmDoctorsMaintenance.txtReferringCharges.Enabled = False
            Else
                frmDoctorsMaintenance.lblReferringCharges.Enabled = True
                frmDoctorsMaintenance.txtReferringCharges.Enabled = True
            End If
        
            Unload Me   'Unload the Wizard
            
        End With
    
    Else    'Displaying an error message, asking the user to choose a record
        MsgBox "Please Select a Record First!", vbExclamation, "No Record Selected!"
        Exit Sub
    End If
    
End Sub
        
        
