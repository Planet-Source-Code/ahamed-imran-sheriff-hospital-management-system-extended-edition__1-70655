VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMedicalTreatmentsOut 
   Caption         =   "Medical Treatments Maintenance"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmMedicalTreatmentsOut.frx":0000
   ScaleHeight     =   9045
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmMedicalTreatmentsOut.frx":20031
      Height          =   855
      Left            =   8400
      Picture         =   "frmMedicalTreatmentsOut.frx":204F0
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmMedicalTreatmentsOut.frx":23234
      Height          =   855
      Left            =   7320
      Picture         =   "frmMedicalTreatmentsOut.frx":2371A
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7440
      Width           =   975
   End
   Begin VB.PictureBox picInvalidTypingMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3840
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   32
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Can Type Only Whole Numeric Digits Here!"
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   105
         Width           =   2175
      End
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
      ItemData        =   "frmMedicalTreatmentsOut.frx":2645E
      Left            =   3360
      List            =   "frmMedicalTreatmentsOut.frx":2646E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmMedicalTreatmentsOut.frx":264B3
      Height          =   750
      Left            =   7440
      Picture         =   "frmMedicalTreatmentsOut.frx":268C8
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmMedicalTreatmentsOut.frx":28A84
      Height          =   750
      Left            =   6480
      Picture         =   "frmMedicalTreatmentsOut.frx":28E60
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmMedicalTreatmentsOut.frx":2B01C
      Height          =   750
      Left            =   8400
      Picture         =   "frmMedicalTreatmentsOut.frx":2B3F2
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmMedicalTreatmentsOut.frx":2D5AE
      Height          =   750
      Left            =   9360
      Picture         =   "frmMedicalTreatmentsOut.frx":2D988
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdPatientSearchWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      ToolTipText     =   "Click Here to select Patient"
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtPatientID 
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtTotal 
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
      Text            =   "0"
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox txtQty 
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
      MaxLength       =   3
      TabIndex        =   12
      Top             =   7200
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
      TabIndex        =   5
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox txtTreatmentID 
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
      TabIndex        =   2
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtUnitPrice 
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
      Left            =   2880
      TabIndex        =   11
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox txtMedicineName 
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
      TabIndex        =   9
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtDateOfIssue 
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
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdMedicineSearchWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Click Here to select Medicine"
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtMedicineID 
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
      Left            =   2880
      TabIndex        =   7
      Top             =   5280
      Width           =   1815
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
      Left            =   7320
      TabIndex        =   1
      Top             =   2280
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
      TabIndex        =   6
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   120
      Top             =   6240
   End
   Begin MSDataGridLib.DataGrid dgrdMedTreatmentInfo 
      Height          =   2535
      Left            =   5520
      TabIndex        =   14
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
      Caption         =   "Medical Treatments Information Table"
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   6840
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   6120
      Top             =   6120
      Width           =   4455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   4200
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblMedicalTreatments 
      BackStyle       =   0  'Transparent
      Caption         =   "Medical Treatments Information"
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
      TabIndex        =   31
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   720
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   360
      Y1              =   3000
      Y2              =   8640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   11520
      Y1              =   8640
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   360
      Y1              =   8640
      Y2              =   8640
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
      Height          =   375
      Left            =   840
      TabIndex        =   30
      Top             =   7725
      Width           =   1335
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Top             =   7245
      Width           =   735
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
      Left            =   840
      TabIndex        =   28
      Top             =   3885
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
      TabIndex        =   27
      Top             =   4365
      Width           =   1335
   End
   Begin VB.Label lblTreatmentID 
      BackStyle       =   0  'Transparent
      Caption         =   "Treatment ID"
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
      TabIndex        =   26
      Top             =   3405
      Width           =   1575
   End
   Begin VB.Label lblMedicineID 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine ID"
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
      Top             =   5325
      Width           =   1335
   End
   Begin VB.Label lblMedicineName 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine Name"
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
      Top             =   5805
      Width           =   1335
   End
   Begin VB.Label lblUnitPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
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
      TabIndex        =   23
      Top             =   6765
      Width           =   1335
   End
   Begin VB.Label lblDateOfIssue 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Issue"
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
      TabIndex        =   22
      Top             =   6285
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   735
      Left            =   2280
      Top             =   2040
      Width           =   7575
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
      TabIndex        =   21
      Top             =   2295
      Width           =   1215
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
      TabIndex        =   20
      Top             =   2295
      Width           =   855
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
      TabIndex        =   19
      Top             =   4845
      Width           =   1335
   End
End
Attribute VB_Name = "frmMedicalTreatmentsOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Medical Treatments Maintenance Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 26/04/08
'Date Of Last Modification: 26/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Medical_Treatments Table
'--------------------------------------------------------------------------------


Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean

'The following variables will be used to autogenerate the Treatment ID to be
'displayed on the Medical Treatments Maintenance form on form load
Dim iNumOfTreatments As Integer  'This variable holds the number of records in the table
Dim strDisplay As String  'This variable will eventually hold the Treatment ID to be autogenerated


Private Sub cmdMedicineSearchWizard_Click()
    
    frmMedsSearchMeds.Show
    
End Sub

Private Sub cmdPatientSearchWizard_Click()
    
    frmOutpatientsSearchMeds.Show
    
End Sub


Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it
        
        
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        
        
        With rsMedicalTreatmentsOut
            
            'Making sure that the user wants to update the record
            If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
                
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtTreatmentID.Text
                .Fields(1) = txtPatientID.Text
                .Fields(2) = txtFirstName.Text
                .Fields(3) = txtSurname.Text
                .Fields(4) = txtMedicineID.Text
                .Fields(5) = txtMedicineName.Text
                .Fields(6) = txtDateOfIssue.Text
                .Fields(7) = txtUnitPrice.Text
                .Fields(8) = txtQty.Text
                .Fields(9) = txtTotal.Text
            
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
        
    End If
    
End Sub

Private Sub dgrdGuardiansInfo_Click()
    
    'Enabling the Update Button & Delete Button
    cmdUpdate.Enabled = True

    
    'Enabling the Navigation Buttons
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    
    With rsMedicalTreatmentsOut
    
        'Entering the values in the particular record into the fields on the interface
        txtTreatmentID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        txtMedicineID.Text = .Fields(4).Value
        txtMedicineName.Text = .Fields(5).Value
        txtDateOfIssue.Text = .Fields(6).Value
        txtUnitPrice.Text = .Fields(7).Value
        txtQty.Text = .Fields(8).Value
        txtTotal.Text = .Fields(9).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub


Private Sub dgrdMedTreatmentInfo_Click()
    
    'Enabling the Update Button & the Delete Button
    cmdUpdate.Enabled = True
    
    'Enabling the Navigation Buttons
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Wizard Buttons
    cmdPatientSearchWizard.Enabled = True
    cmdMedicineSearchWizard.Enabled = True
    
    With rsMedicalTreatmentsOut
    
        'Entering the values in the particular record into the fields on the interface
        txtTreatmentID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        txtMedicineID.Text = .Fields(4).Value
        txtMedicineName.Text = .Fields(5).Value
        txtDateOfIssue.Text = .Fields(6).Value
        txtUnitPrice.Text = .Fields(7).Value
        txtQty.Text = .Fields(8).Value
        txtTotal.Text = .Fields(9).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub

Public Sub Form_Load()

    Call Connection  'Calling the Connection Procedure
    
    Call Medical_Treatments_Out 'Calling the Medical_Treatments Procedure
    
    disableAllFields  'Calling a Private Function To Disable All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    'Enabling  the First Button and the Last Button
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Add New Button & Close Button
    cmdClose.Enabled = True
    
    'Enabling the Search Frame
    lblCriteria.Enabled = True
    cboSearchType.Enabled = True
    lblSearchText.Enabled = True
    txtSearch.Enabled = True

    dgrdMedTreatmentInfo.Enabled = True
    
    Set dgrdMedTreatmentInfo.DataSource = rsMedicalTreatmentsOut   'Setting the DataSource of the DataGrid
    
    
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
    
    'Enabling the Wizard Buttons
    cmdPatientSearchWizard.Enabled = True
    cmdMedicineSearchWizard.Enabled = True


    With rsMedicalTreatmentsOut


        .MoveFirst  'Moving to the first record

        'Entering the values in the particular record into the fields on the interface
        txtTreatmentID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        txtMedicineID.Text = .Fields(4).Value
        txtMedicineName.Text = .Fields(5).Value
        txtDateOfIssue.Text = .Fields(6).Value
        txtUnitPrice.Text = .Fields(7).Value
        txtQty.Text = .Fields(8).Value
        txtTotal.Text = .Fields(9).Value

    End With

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdPrevious_Click() 'This function will Navigate to the Previous Record

    With rsMedicalTreatmentsOut


        .MovePrevious   'Moving to the previous record

        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If

        'Entering the values in the particular record into the fields on the interface
        txtTreatmentID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        txtMedicineID.Text = .Fields(4).Value
        txtMedicineName.Text = .Fields(5).Value
        txtDateOfIssue.Text = .Fields(6).Value
        txtUnitPrice.Text = .Fields(7).Value
        txtQty.Text = .Fields(8).Value
        txtTotal.Text = .Fields(9).Value

    End With

    cmdNext.Enabled = True  'Enabling the Next Button
    cmdLast.Enabled = True  'Enabling the Last Button
    
    'Enabling the Wizard Buttons
    cmdPatientSearchWizard.Enabled = True
    cmdMedicineSearchWizard.Enabled = True

    'Enabling the Update Button & Delete Button
    cmdUpdate.Enabled = True

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdNext_Click() 'This function will Navigate to the Next Record

    With rsMedicalTreatmentsOut

        .MoveNext   'Moving to the Next Record

        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If

        'Entering the values in the particular record into the fields on the interface
        txtTreatmentID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        txtMedicineID.Text = .Fields(4).Value
        txtMedicineName.Text = .Fields(5).Value
        txtDateOfIssue.Text = .Fields(6).Value
        txtUnitPrice.Text = .Fields(7).Value
        txtQty.Text = .Fields(8).Value
        txtTotal.Text = .Fields(9).Value

    End With

    cmdPrevious.Enabled = True  'Enabling the Previous Button
    cmdFirst.Enabled = True 'Enabling the First Button

    'Enabling the Update Button & Delete Button
    cmdUpdate.Enabled = True
    
    'Enabling the Wizard Buttons
    cmdPatientSearchWizard.Enabled = True
    cmdMedicineSearchWizard.Enabled = True


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

    'Enabling the Wizard Buttons
    cmdPatientSearchWizard.Enabled = True
    cmdMedicineSearchWizard.Enabled = True

    With rsMedicalTreatmentsOut

        .Requery

        .MoveLast   'Moving to the last record

        'Entering the values in the particular record into the fields on the interface
        txtTreatmentID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        txtMedicineID.Text = .Fields(4).Value
        txtMedicineName.Text = .Fields(5).Value
        txtDateOfIssue.Text = .Fields(6).Value
        txtUnitPrice.Text = .Fields(7).Value
        txtQty.Text = .Fields(8).Value
        txtTotal.Text = .Fields(9).Value

    End With

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub



Private Function textfieldsValidations() As Boolean  'This function will validate all fields
    
    Flag = True 'Setting the Flag variable to True

    
    'Checking if the Patient ID textfield is empty
    If txtPatientID.Text = "" Then
        txtPatientID.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtFirstName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtSurname.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtPatientID.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtFirstName.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtSurname.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    
    'Checking if the Medicine ID textfield is empty
    If txtMedicineID.Text = "" Then
        txtMedicineID.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtMedicineName.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtUnitPrice.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtMedicineID.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtMedicineName.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtUnitPrice.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    
    'Checking if the Total textfield has been filled in
    If txtTotal.Text = "0" Then
        txtQty.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtTotal.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtQty.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtTotal.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    
    'Here, I am checking the state of the Flag variable and if it is False, I am displaying a
    'Message Box to instruct the user to enter data into all highlighted textfields.
    'The Save procedure will also be cancelled
    If Flag = False Then
        MsgBox "Error! Please Fill-in The Highlighted Textfields! They Are Compulsory!", vbCritical, "Please Fill Highlighted Textfields"
        textfieldsValidations = True    'Passing values to the Save procedure
    Else
        textfieldsValidations = False   'Passing values to the Save procedure
    End If
    
End Function


Private Sub txtQty_Change()
    
    If txtQty.Text = "0" Then
    
        MsgBox "Error! The Figure Cannot Begin With Zero!", vbCritical, "Cannot Begin Figure With 0!"
        txtQty.Text = ""
        Exit Sub
        
    Else
    
        txtTotal.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
        
    End If
    
End Sub


Private Sub tmrErrMsg_Timer()

    Static i As Integer

    If i < 200000 Then     'Validation Msg Viewing Time Period
        picInvalidTypingMsg.Visible = False
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If

End Sub


Private Sub txtQty_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only digits

    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidTypingMsg.Top = 7200    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub



Private Sub txtSearch_Change()

    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsMedicalTreatmentsOut
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[TreatmentID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[PatientID] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
                    .Filter = "[FirstName] Like '" & txtSearch.Text & "%" & "'"
                Case 3:
                    .Filter = "[Surname] Like '" & txtSearch.Text & "%" & "'"
            End Select
    
        End With
            
    Else
    
        clearAllFields  'Calling a Private Function To Clear All Fields
        
        disableAllFields  'Calling the disableAllFields procedure
        
        'Disabling the Update Button and the Delete Button
        cmdUpdate.Enabled = False

        
        'Enable the Search Frame
        cboSearchType.Enabled = True
        txtSearch.Enabled = True
        
        Call Medical_Treatments_Out
        
        Set dgrdMedTreatmentInfo.DataSource = rsMedicalTreatmentsOut
        
    End If
    
    
End Sub


Private Sub cmdClose_Click()

    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If

End Sub


