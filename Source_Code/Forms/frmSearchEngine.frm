VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmSearchEngine 
   Caption         =   "Search Engine"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSearchEngine.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Click Here To Close This Interface"
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox txtSearchData 
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
      Left            =   7440
      TabIndex        =   3
      Top             =   2880
      Width           =   3855
   End
   Begin VB.ComboBox cboInfoTable 
      Height          =   315
      ItemData        =   "frmSearchEngine.frx":1C5A9
      Left            =   600
      List            =   "frmSearchEngine.frx":1C5BF
      TabIndex        =   1
      Text            =   "-------------------SELECT---------------------"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ComboBox cboInfoType 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSearchEngine.frx":1C65C
      Left            =   3960
      List            =   "frmSearchEngine.frx":1C65E
      TabIndex        =   0
      Text            =   "--------------------SELECT-------------------"
      Top             =   2880
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid dgrdInformation 
      Height          =   4095
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Line Line15 
      X1              =   10080
      X2              =   11400
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Text"
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
      Left            =   8880
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Line Line14 
      X1              =   7320
      X2              =   8640
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line13 
      X1              =   11400
      X2              =   11400
      Y1              =   3360
      Y2              =   2640
   End
   Begin VB.Line Line12 
      X1              =   7320
      X2              =   11400
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line11 
      X1              =   7320
      X2              =   7320
      Y1              =   2640
      Y2              =   3360
   End
   Begin VB.Line Line10 
      X1              =   6120
      X2              =   6840
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Criteria"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Line Line9 
      X1              =   3840
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line8 
      X1              =   6840
      X2              =   6840
      Y1              =   3360
      Y2              =   2640
   End
   Begin VB.Line Line7 
      X1              =   3840
      X2              =   6840
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line6 
      X1              =   3840
      X2              =   3840
      Y1              =   2640
      Y2              =   3360
   End
   Begin VB.Line Line5 
      X1              =   2760
      X2              =   3360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label lblPresentTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Information Table"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   960
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      X1              =   3360
      X2              =   3360
      Y1              =   3360
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   3360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   2640
      Y2              =   3360
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000001&
      FillColor       =   &H00FF0000&
      Height          =   4335
      Left            =   360
      Top             =   3480
      Width           =   11055
   End
End
Attribute VB_Name = "frmSearchEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Search Engine
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 03/05/08
'Date Of Last Modification: 03/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed:
'--------------------------------------------------------------------------------

Option Explicit

Private Sub cboInfoTable_Click()
    Dim strTitle As String
    
    cboInfoType.Clear
    
    Select Case (cboInfoTable.ListIndex)
        Case 0: 'Inpatient Information
        
            With cboInfoType
                .AddItem "Patient ID", 0
                .AddItem "First Name", 1
                .AddItem "Surname", 2
                .AddItem "NIC Number", 3
                .AddItem "Account Type", 4
                .AddItem "Company ID", 5
                .AddItem "Company Name", 6
                
            End With
            strTitle = cboInfoTable.List(0) & " Table"
            
            Set dgrdInformation.DataSource = rsInpatientMaintenance
            
            txtSearchData.Text = "" 'Clearing the textfield

            
        Case 1: 'Outpatient Information
        
            With cboInfoType
                .AddItem "Patient ID", 0
                .AddItem "First Name", 1
                .AddItem "Surname", 2
                .AddItem "NIC Number", 3
                .AddItem "Account Type", 4
                .AddItem "Company ID", 5
                .AddItem "Company Name", 6
            End With
            strTitle = cboInfoTable.List(1) & " Table"
            
            Set dgrdInformation.DataSource = rsOutpatientsMaintenance

            txtSearchData.Text = "" 'Clearing the textfield

            
        Case 2: 'Inpatient Payments
        
            With cboInfoType
                .AddItem "Patient ID", 0
                .AddItem "Patient Name", 1
            End With
            strTitle = cboInfoTable.List(2) & " Table"
            
            Set dgrdInformation.DataSource = rsInpatientBilling

            txtSearchData.Text = "" 'Clearing the textfield
        
        
        Case 3: 'Outpatient Payments
        
            With cboInfoType
                .AddItem "Patient ID", 0
                .AddItem "Patient Name", 1
            End With
            strTitle = cboInfoTable.List(3) & " Table"
            
            Set dgrdInformation.DataSource = rsOutpatientBilling
            
            txtSearchData.Text = "" 'Clearing the textfield
        
        
        Case 4: 'Doctors Maintenance
        
            With cboInfoType
                .AddItem "Doctor ID", 0
                .AddItem "First Name", 1
                .AddItem "Surname", 2
                .AddItem "Gender", 3
                .AddItem "NIC Number", 4
                .AddItem "License Number", 5
                .AddItem "Specialization", 6
            End With
            strTitle = cboInfoTable.List(4) & " Table"
            
            Set dgrdInformation.DataSource = rsDoctorsMaintenance

            txtSearchData.Text = "" 'Clearing the textfield
        
        Case 5: 'Doctors Channeling Schedule
        
            With cboInfoType
                .AddItem "Doctor ID", 0
            End With
            strTitle = cboInfoTable.List(5) & " Table"
            
            Set dgrdInformation.DataSource = rsDoctorSchedule

            txtSearchData.Text = "" 'Clearing the textfield
    
    End Select
    
    cboInfoType.Text = "--------------------SELECT-------------------"
    cboInfoType.Enabled = True
    dgrdInformation.Caption = strTitle
    
End Sub



Private Sub cmdClose_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

Call Connection
Call Inpatients_Maintenance
Call Outpatients_Maintenance
Call Inpatient_Billing
Call Outpatient_Billing
Call Doctors_Maintenance
Call Doctor_Schedule

End Sub



Private Sub txtSearchData_Change()

On Error GoTo backload

    If cboInfoType.ListIndex <> -1 Then
    
        If cboInfoTable.ListIndex = 0 Then  'Inpatients Information
                
            If Len(txtSearchData.Text) > 0 Then
         
                '-----Select the Type of search and filter the record
                Select Case (cboInfoType.ListIndex)
                    Case 0:
                        rsInpatientMaintenance.Filter = "PatientID Like '" & txtSearchData.Text & "%" & "'"
                    Case 1:
                        rsInpatientMaintenance.Filter = "FirstName Like '" & txtSearchData.Text & "%" & "'"
                    Case 2:
                        rsInpatientMaintenance.Filter = "Surname Like '" & txtSearchData.Text & "%" & "'"
                    Case 3:
                        rsInpatientMaintenance.Filter = "NICNumber Like '" & txtSearchData.Text & "%" & "'"
                    Case 4:
                        rsInpatientMaintenance.Filter = "AccountType Like '" & txtSearchData.Text & "%" & "'"
                    Case 5:
                        rsInpatientMaintenance.Filter = "CompanyID Like '" & txtSearchData.Text & "%" & "'"
                    Case 6:
                        rsInpatientMaintenance.Filter = "CompanyName Like '" & txtSearchData.Text & "%" & "'"
                End Select
                
            Else
                
                Call Inpatients_Maintenance
                
                Set dgrdInformation.DataSource = rsInpatientMaintenance
                
            End If
            
        ElseIf cboInfoTable.ListIndex = 1 Then 'Outpatients Information
        
            
            If Len(txtSearchData.Text) > 0 Then
            
                '-----Select the Type of search and filter the record
                Select Case (cboInfoType.ListIndex)
                    Case 0:
                        rsOutpatientsMaintenance.Filter = "PatientID Like '" & txtSearchData.Text & "%" & "'"
                    Case 1:
                        rsOutpatientsMaintenance.Filter = "FirstName Like '" & txtSearchData.Text & "%" & "'"
                    Case 2:
                        rsOutpatientsMaintenance.Filter = "Surname Like '" & txtSearchData.Text & "%" & "'"
                    Case 3:
                        rsOutpatientsMaintenance.Filter = "NICNumber Like '" & txtSearchData.Text & "%" & "'"
                    Case 4:
                        rsOutpatientsMaintenance.Filter = "AccountType Like '" & txtSearchData.Text & "%" & "'"
                    Case 5:
                        rsOutpatientsMaintenance.Filter = "CompanyID Like '" & txtSearchData.Text & "%" & "'"
                    Case 6:
                        rsOutpatientsMaintenance.Filter = "CompanyName Like '" & txtSearchData.Text & "%" & "'"
                End Select
                
            Else
                
                Call Outpatients_Maintenance
                    
                Set dgrdInformation.DataSource = rsOutpatientsMaintenance
                
            End If
            
          ElseIf cboInfoTable.ListIndex = 2 Then 'Inpatient Payments
          
            
            If Len(txtSearchData.Text) > 0 Then
                        
                '-----Select the Type of search and filter the record
                Select Case (cboInfoType.ListIndex)
                    Case 0:
                        rsInpatientBilling.Filter = "PatientID Like '" & txtSearchData.Text & "%" & "'"
                    Case 1:
                        rsInpatientBilling.Filter = "PatientName Like '" & txtSearchData.Text & "%" & "'"
                End Select
                
            Else
            
                Call Inpatient_Billing
                
                Set dgrdInformation.DataSource = rsInpatientBilling
                
            End If
            
          ElseIf cboInfoTable.ListIndex = 3 Then 'Outpatient Payments
          
            
            If Len(txtSearchData.Text) > 0 Then
                        
                '-----Select the Type of search and filter the record
                Select Case (cboInfoType.ListIndex)
                    Case 0:
                        rsOutpatientBilling.Filter = "PatientID Like '" & txtSearchData.Text & "%" & "'"
                    Case 1:
                        rsOutpatientBilling.Filter = "PatientName Like '" & txtSearchData.Text & "%" & "'"
                End Select
                
            Else
                
                Call Outpatient_Billing
                
                Set dgrdInformation.DataSource = rsOutpatientBilling
                
            End If
            
        ElseIf cboInfoTable.ListIndex = 4 Then 'Doctors Information
            
            
            If Len(txtSearchData.Text) > 0 Then
            
                '-----Select the Type of search and filter the record
                Select Case (cboInfoType.ListIndex)
                    Case 0:
                        rsDoctorsMaintenance.Filter = "DoctorID Like '" & txtSearchData.Text & "%" & "'"
                    Case 1:
                        rsDoctorsMaintenance.Filter = "FirstName Like '" & txtSearchData.Text & "%" & "'"
                    Case 2:
                        rsDoctorsMaintenance.Filter = "Surname Like '" & txtSearchData.Text & "%" & "'"
                    Case 3:
                        rsDoctorsMaintenance.Filter = "Gender Like '" & txtSearchData.Text & "%" & "'"
                    Case 4:
                        rsDoctorsMaintenance.Filter = "NICNumber Like '" & txtSearchData.Text & "%" & "'"
                    Case 5:
                        rsDoctorsMaintenance.Filter = "LicenseNo Like '" & txtSearchData.Text & "%" & "'"
                    Case 6:
                        rsDoctorsMaintenance.Filter = "Specialization Like '" & txtSearchData.Text & "%" & "'"
                End Select
                
            Else
                
                Call Doctors_Maintenance
                
                Set dgrdInformation.DataSource = rsDoctorsMaintenance
                
            End If
            
        ElseIf cboInfoTable.ListIndex = 5 Then 'Doctor's Channeling Schedule
                        
            If Len(txtSearchData.Text) > 0 Then

        
                '-----Select the Type of search and filter the record
                Select Case (cboInfoType.ListIndex)
                    Case 0:
                        rsDoctorSchedule.Filter = "DoctorID Like '" & txtSearchData.Text & "%" & "'"
                End Select
                
            Else
                
                Call Doctor_Schedule
                
                Set dgrdInformation.DataSource = rsDoctorSchedule
                
            End If
            
        End If
        
    Else
        
        If txtSearchData = "" Then
            MsgBox "Please Select A Search Criteria!", vbExclamation, "No Search Criteria!"
        End If
        
        txtSearchData = ""
        
    End If
    
Exit Sub

backload:
Call Form_Load

    
End Sub

