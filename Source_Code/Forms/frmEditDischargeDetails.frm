VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmDischargeDetailsMaintenance 
   Caption         =   "Discharge Details Maintenance"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditDischargeDetails.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSurname 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5400
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
      Left            =   7200
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmEditDischargeDetails.frx":1FA22
      Height          =   855
      Left            =   9960
      Picture         =   "frmEditDischargeDetails.frx":1FEE1
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmEditDischargeDetails.frx":22C25
      Height          =   855
      Left            =   8880
      Picture         =   "frmEditDischargeDetails.frx":230EE
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      DisabledPicture =   "frmEditDischargeDetails.frx":25E32
      Height          =   855
      Left            =   5640
      Picture         =   "frmEditDischargeDetails.frx":26234
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmEditDischargeDetails.frx":28F78
      Height          =   855
      Left            =   6720
      Picture         =   "frmEditDischargeDetails.frx":293F6
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmEditDischargeDetails.frx":2C13A
      Height          =   855
      Left            =   7800
      Picture         =   "frmEditDischargeDetails.frx":2C620
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdInpatientSearchWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Click Here to select Customer"
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox txtAdmissionID 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtDischargeID 
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
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtDischargeTime 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7320
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5880
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox txtDischargeDate 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox txtFirstName 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4920
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
      Height          =   765
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "frmEditDischargeDetails.frx":2F364
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmEditDischargeDetails.frx":2F368
      Height          =   750
      Left            =   9360
      Picture         =   "frmEditDischargeDetails.frx":2F742
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmEditDischargeDetails.frx":318FE
      Height          =   750
      Left            =   8400
      Picture         =   "frmEditDischargeDetails.frx":31CD4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmEditDischargeDetails.frx":33E90
      Height          =   750
      Left            =   6480
      Picture         =   "frmEditDischargeDetails.frx":3426C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmEditDischargeDetails.frx":36428
      Height          =   750
      Left            =   7440
      Picture         =   "frmEditDischargeDetails.frx":3683D
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   890
   End
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
      ItemData        =   "frmEditDischargeDetails.frx":389F9
      Left            =   3360
      List            =   "frmEditDischargeDetails.frx":38A0C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid dgrdDischargeInfo 
      Height          =   2535
      Left            =   5520
      TabIndex        =   13
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
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
      Caption         =   "Patient Discharge Information Table"
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
      Height          =   375
      Left            =   840
      TabIndex        =   35
      Top             =   5445
      Width           =   1335
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
      Left            =   2520
      TabIndex        =   34
      Top             =   2325
      Width           =   855
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
      Left            =   6000
      TabIndex        =   33
      Top             =   2325
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   5520
      Top             =   7320
      Width           =   5535
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
      TabIndex        =   32
      Top             =   4005
      Width           =   1575
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
      Height          =   375
      Left            =   840
      TabIndex        =   31
      Top             =   4485
      Width           =   1335
   End
   Begin VB.Label lblDischargeID 
      BackStyle       =   0  'Transparent
      Caption         =   "Discharge ID"
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
      Top             =   3525
      Width           =   1575
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
      Height          =   375
      Left            =   840
      TabIndex        =   29
      Top             =   4965
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
      Height          =   375
      Left            =   840
      TabIndex        =   28
      Top             =   5925
      Width           =   1335
   End
   Begin VB.Label lblDischargeTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Discharge Time"
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
      Left            =   840
      TabIndex        =   27
      Top             =   7365
      Width           =   1335
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
      Height          =   375
      Left            =   840
      TabIndex        =   26
      Top             =   6405
      Width           =   1335
   End
   Begin VB.Label lblDischargeDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Discharge Date"
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
      Left            =   840
      TabIndex        =   25
      Top             =   6885
      Width           =   1815
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
      Left            =   840
      TabIndex        =   24
      Top             =   7845
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   360
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   11520
      Y1              =   8760
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   360
      Y1              =   3000
      Y2              =   8760
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   720
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblDischargeInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Discharge Information"
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
      TabIndex        =   19
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   3960
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   6120
      Top             =   6120
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   735
      Left            =   1920
      Top             =   2040
      Width           =   8175
   End
End
Attribute VB_Name = "frmDischargeDetailsMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/
'--------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Discharge Details Maintenance Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 07/05/08
'Date Of Last Modification: 07/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Discharge_Details Table
'--------------------------------------------------------------------------------


Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons


'This variable will hold the Room ID of the patient being discharged
Dim strRoomIDStore As String

'This variable will help me to decide if the patient has settled the bill in full
Dim checkBillingFlag As Boolean

'The following variables will be used to autogenerate the Discharge ID to be
'displayed on the Discharge Details Maintenance form on form load
Dim iNumOfRecords As Integer  'This variable holds the number of records in the table
Dim strDisplay As String  'This variable will eventually hold the Discharge ID to be autogenerated



Private Sub cmdDelete_Click()   'This function will delete a record from the database
    
    'Check for the record selection
    If txtDischargeID.Text = "" Then
    
        MsgBox "Error! No Record Has Been Selected", vbCritical, "No Record Selected!"
    
    Else
    
        With rsDischargeMaintenance
        
            'Confirm the Delete procedure with the user
            If MsgBox("Are You Sure You Wish To Delete Discharge ID " & txtDischargeID.Text & "'s Record?", vbYesNo + vbQuestion, "Delete Record?") = vbYes Then
        
                .Delete 'Delete the record from the database
                
                'Display Success Message
                MsgBox "The Record Has Been Deleted Successfully!", vbInformation, "Successful Delete Procedure!"
                                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            Else
                
                'Display 'Delete Procedure Cancelled' Message
                MsgBox "The Delete Procedure Was Cancelled!", vbExclamation, "Delete Procedure Cancelled!"
                
                clearAllFields  'Calling a Private Function To Clear All Fields
        
            End If

            .Requery    'Requerying the Table
            
            Form_Load   'Calling the Form_Load Procedure
        
        End With
        
    End If

End Sub


Private Sub cmdInpatientSearchWizard_Click()
    
    frmInpatientSearchDischarge.Show
    
End Sub

Private Sub cmdSave_Click()     'This function will save all the user's data in the database
        
    checkBillingFlag = False
        
    If txtAdmissionID.Text = "" Then
        MsgBox "Error! You Have Not Selected A Patient!", vbInformation, "Error! No Selection!"
        Exit Sub
    End If
            
        
    
    
    With rsDischargeMaintenance
        
        'Making sure that the user wants to save the record
        If MsgBox("Are You Sure You Wish To Discharge This Patient?", vbYesNo + vbQuestion, "Discharge Patient?") = vbYes Then
            
            
            'Save the user-entered data into the recordset
            .Fields(0) = txtDischargeID.Text
            .Fields(1) = txtAdmissionID.Text
            .Fields(2) = txtPatientID.Text
            .Fields(3) = txtFirstName.Text
            .Fields(4) = txtSurname.Text
            .Fields(5) = txtAdmissionDate.Text
            .Fields(6) = txtAdmissionTime.Text
            .Fields(7) = txtDischargeDate.Text
            .Fields(8) = txtDischargeTime.Text
            .Fields(9) = txtAdditionalNotes.Text
            .Fields(10) = True
            
            
            .Update
            
            Call Inpatients_Admission
            
            With rsInpatientsAdmission
        
                .MoveFirst
            
                Do While .EOF = False
            
                    If txtAdmissionID.Text = .Fields(0).Value Then
                
                        strRoomIDStore = .Fields(15).Value
                        
                        Exit Do
                
                    Else
                
                        .MoveNext
                        
                    End If
            
                Loop
                
                .Close
            
            End With
            
            
            Call Rooms_Maintenance
            
            With rsRoomsMaintenance
            
                .MoveFirst
                
                Do While .EOF = False
                
                    If .Fields(0).Value = strRoomIDStore Then
                    
                        .Fields(8).Value = False
                        
                        .Update
                        
                        Exit Do
                        
                    Else
                    
                        .MoveNext
                        
                    End If
                    
                Loop
                
                
            End With
            
            'Display Success Message
            MsgBox "The Patient Has Been Discharged Successfully!", vbInformation, "Succesful Discharge Procedure!"
            
            Form_Load   'Calling the Form_Load Procedure
            
            clearAllFields  'Calling a Private Function To Clear All Fields
        
        Else
        
            'Display 'No Modifications' Message
            MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
            
            .CancelUpdate   'Cancel the Save Procedure
            
            Form_Load   'Calling the Form_Load Procedure
            
            clearAllFields  'Calling a Private Function To Clear All Fields
        
        End If
        
        .Requery    'Requerying the Table
        
    End With
        

End Sub



Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it
        
        
    If txtAdmissionID.Text = "" Then
        MsgBox "Error! You Have Not Selected A Patient!", vbInformation, "Error! No Selection!"
        Exit Sub
    End If
        
        
    With rsDischargeMaintenance
        
        'Making sure that the user wants to update the record
        If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
            
            
            'Save the user-entered data into the recordset
            .Fields(0) = txtDischargeID.Text
            .Fields(1) = txtAdmissionID.Text
            .Fields(2) = txtPatientID.Text
            .Fields(3) = txtFirstName.Text
            .Fields(4) = txtSurname.Text
            .Fields(5) = txtAdmissionDate.Text
            .Fields(6) = txtAdmissionTime.Text
            .Fields(7) = txtDischargeDate.Text
            .Fields(8) = txtDischargeTime.Text
            .Fields(9) = txtAdditionalNotes.Text
        
            .Update
            
            'Display Success Message
            MsgBox "The Record Was Updated Successfully!", vbInformation, "Succesful Update Procedure"
            
            
            Form_Load   'Calling the Form_Load Procedure
            
            clearAllFields  'Calling a Private Function To Clear All Fields
        
        Else
        
            'Display 'No Modifications' Message
            MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
            
            .CancelUpdate   'Cancel the Update Procedure
            
            Form_Load   'Calling the Form_Load Procedure
            
        
        End If
        
        .Requery    'Requerying the Table
        
    End With
        

    
End Sub

Private Sub dgrdDischargeInfo_Click()
    
    'Enabling the Update Button & Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True

    
    'Enabling the Navigation Buttons
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Wizard Buttons
    cmdInpatientSearchWizard.Enabled = True
    
    
    With rsDischargeMaintenance
    
        'Entering the values in the particular record into the fields on the interface
        txtDischargeID.Text = .Fields(0).Value
        txtAdmissionID.Text = .Fields(1).Value
        txtPatientID.Text = .Fields(2).Value
        txtFirstName.Text = .Fields(3).Value
        txtSurname.Text = .Fields(4).Value
        txtAdmissionDate.Text = .Fields(5).Value
        txtAdmissionTime.Text = .Fields(6).Value
        txtDischargeDate.Text = .Fields(7).Value
        txtDischargeTime.Text = .Fields(8).Value
        txtAdditionalNotes.Text = .Fields(9).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub


Private Sub cmdAddNew_Click() 'This function adds a new recordset into the database

    enableAllFields     'Calling a Private Function To Enable All Fields
    clearAllFields      'Calling a Private Function To Clear All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    
    'Disabling the Search Frame
    lblCriteria.Enabled = False
    cboSearchType.Enabled = False
    lblSearchText.Enabled = False
    txtSearch.Enabled = False
    
    
    'Disabling the DataGrid
    dgrdDischargeInfo.Enabled = False
    
    'Enabling the Save Command Button & Close Command Button
    cmdSave.Enabled = True
    cmdClose.Enabled = True

    
    'Enabling the Wizard Buttons
    cmdInpatientSearchWizard.Enabled = True

    
    Call Discharge_Maintenance    'Calling the Discharge_Maintenance Procedure to interact with the recordset
    
    'Generate Discharge ID By Utilizing the Discharge_Maintenance Table
    With rsDischargeMaintenance
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strDisplay = "DIS0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Discharge_Maintenance Table
            If iNumOfRecords < 10 Then
                strDisplay = "DIS000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strDisplay = "DIS00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strDisplay = "DIS0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strDisplay = "DIS" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    'The following line of code will enter the autogenerated Discharge ID
    'into the Discharge ID textfield
    txtDischargeID.Text = strDisplay
    
    txtDischargeDate.Text = DateTime.Date 'Setting the system date into this textfield.
    txtDischargeTime.Text = DateTime.Time   'Setting the system time into this textfield.
    
    txtAdditionalNotes.Text = "-"   'Setting the default value for this textfield
        
End Sub


Public Sub Form_Load()

    Call Connection  'Calling the Connection Procedure
    
    Call Discharge_Maintenance 'Calling the Discharge_Maintenance Procedure
    
    disableAllFields  'Calling a Private Function To Disable All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    'Enabling  the First Button and the Last Button
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Add New Button & Close Button
    cmdAddNew.Enabled = True
    cmdClose.Enabled = True
    
    'Enabling the Search Frame
    lblCriteria.Enabled = True
    cboSearchType.Enabled = True
    lblSearchText.Enabled = True
    txtSearch.Enabled = True

    dgrdDischargeInfo.Enabled = True
    
    Set dgrdDischargeInfo.DataSource = rsDischargeMaintenance   'Setting the DataSource of the DataGrid
    
    
End Sub

Public Function disableAllFields() 'This function will disable all fields on the interface

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
    
End Function


Public Function clearAllFields() 'This function will clear all fields on the interface


    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will clear all TextBoxes
    If TypeOf eachField Is TextBox Then
        eachField.Text = ""
    End If

    Next
    
End Function


Private Sub cmdFirst_Click()  'This function will Navigate to the First Record

    'Enabling / Diabling the Navigation Buttons as necessary
    cmdFirst.Enabled = False
    cmdLast.Enabled = True
    cmdPrevious.Enabled = False
    cmdNext.Enabled = True

    'Enabling the Update Button & Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the Wizard Buttons
    cmdInpatientSearchWizard.Enabled = True



    With rsDischargeMaintenance


        .MoveFirst  'Moving to the first record

        'Entering the values in the particular record into the fields on the interface
        txtDischargeID.Text = .Fields(0).Value
        txtAdmissionID.Text = .Fields(1).Value
        txtPatientID.Text = .Fields(2).Value
        txtFirstName.Text = .Fields(3).Value
        txtSurname.Text = .Fields(4).Value
        txtAdmissionDate.Text = .Fields(5).Value
        txtAdmissionTime.Text = .Fields(6).Value
        txtDischargeDate.Text = .Fields(7).Value
        txtDischargeTime.Text = .Fields(8).Value
        txtAdditionalNotes.Text = .Fields(9).Value

    End With

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdPrevious_Click() 'This function will Navigate to the Previous Record

    With rsDischargeMaintenance


        .MovePrevious   'Moving to the previous record

        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If

        'Entering the values in the particular record into the fields on the interface
        txtDischargeID.Text = .Fields(0).Value
        txtAdmissionID.Text = .Fields(1).Value
        txtPatientID.Text = .Fields(2).Value
        txtFirstName.Text = .Fields(3).Value
        txtSurname.Text = .Fields(4).Value
        txtAdmissionDate.Text = .Fields(5).Value
        txtAdmissionTime.Text = .Fields(6).Value
        txtDischargeDate.Text = .Fields(7).Value
        txtDischargeTime.Text = .Fields(8).Value
        txtAdditionalNotes.Text = .Fields(9).Value

    End With

    cmdNext.Enabled = True  'Enabling the Next Button
    cmdLast.Enabled = True  'Enabling the Last Button
    
    'Enabling the Wizard Buttons
    cmdInpatientSearchWizard.Enabled = True


    'Enabling the Update Button & Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdNext_Click() 'This function will Navigate to the Next Record

    With rsDischargeMaintenance

        .MoveNext   'Moving to the Next Record

        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If

        'Entering the values in the particular record into the fields on the interface
        txtDischargeID.Text = .Fields(0).Value
        txtAdmissionID.Text = .Fields(1).Value
        txtPatientID.Text = .Fields(2).Value
        txtFirstName.Text = .Fields(3).Value
        txtSurname.Text = .Fields(4).Value
        txtAdmissionDate.Text = .Fields(5).Value
        txtAdmissionTime.Text = .Fields(6).Value
        txtDischargeDate.Text = .Fields(7).Value
        txtDischargeTime.Text = .Fields(8).Value
        txtAdditionalNotes.Text = .Fields(9).Value

    End With

    cmdPrevious.Enabled = True  'Enabling the Previous Button
    cmdFirst.Enabled = True 'Enabling the First Button

    'Enabling the Update Button & Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the Wizard Buttons
    cmdInpatientSearchWizard.Enabled = True



    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdLast_Click() 'This function will Navigate to the Last Record

    'Enabling / Diabling the Navigation Buttons as necessary
    cmdLast.Enabled = False
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = False

    'Enabling the Update Button & the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True

    'Enabling the Wizard Buttons
    cmdInpatientSearchWizard.Enabled = True


    With rsDischargeMaintenance

        .Requery

        .MoveLast   'Moving to the last record

        'Entering the values in the particular record into the fields on the interface
        txtDischargeID.Text = .Fields(0).Value
        txtAdmissionID.Text = .Fields(1).Value
        txtPatientID.Text = .Fields(2).Value
        txtFirstName.Text = .Fields(3).Value
        txtSurname.Text = .Fields(4).Value
        txtAdmissionDate.Text = .Fields(5).Value
        txtAdmissionTime.Text = .Fields(6).Value
        txtDischargeDate.Text = .Fields(7).Value
        txtDischargeTime.Text = .Fields(8).Value
        txtAdditionalNotes.Text = .Fields(9).Value

    End With

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub
    


Private Sub txtSearch_Change()

    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsDischargeMaintenance
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[DischargeID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[AdmissionID] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
                    .Filter = "[PatientID] Like '" & txtSearch.Text & "%" & "'"
                Case 3:
                    .Filter = "[FirstName] Like '" & txtSearch.Text & "%" & "'"
                Case 4:
                    .Filter = "[Surname] Like '" & txtSearch.Text & "%" & "'"
            End Select
    
        End With
            
    Else
    
        clearAllFields  'Calling a Private Function To Clear All Fields
        
        disableAllFields  'Calling the disableAllFields procedure
        
        'Disabling the Update Button and the Delete Button
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False

        
        'Enable the Search Frame
        cboSearchType.Enabled = True
        txtSearch.Enabled = True
        
        Call Discharge_Maintenance
        
        Set dgrdDischargeInfo.DataSource = rsDischargeMaintenance
        
    End If
    
    
End Sub


Private Sub cmdClose_Click()

    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If

End Sub



