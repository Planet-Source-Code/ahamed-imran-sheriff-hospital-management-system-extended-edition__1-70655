VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmAddMedicalTreatmentsOut 
   Caption         =   "Add Medical Treatments"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddMedicalTreatmentsOut.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInvalidTypingMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3360
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   26
      Top             =   6360
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Can Type Only Whole Numeric Digits Here!"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   105
         Width           =   2175
      End
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3240
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdPatientSearchWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Click Here to select a Patient"
      Top             =   2280
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdMedicineSearchWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Click Here to select a Medicine"
      Top             =   4440
      Width           =   375
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5400
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txtUnitPrice 
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox txtQty 
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
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   9
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox txtNettTotal 
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
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   0
      Top             =   4920
   End
   Begin VB.CommandButton cmdAdd 
      DisabledPicture =   "frmAddMedicalTreatmentsOut.frx":1E353
      Height          =   855
      Left            =   3720
      Picture         =   "frmAddMedicalTreatmentsOut.frx":1E755
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmAddMedicalTreatmentsOut.frx":21499
      Enabled         =   0   'False
      Height          =   855
      Left            =   4920
      Picture         =   "frmAddMedicalTreatmentsOut.frx":21962
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmAddMedicalTreatmentsOut.frx":246A6
      Height          =   855
      Left            =   6120
      Picture         =   "frmAddMedicalTreatmentsOut.frx":24B65
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7800
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgrdMedicalTreatmentsInfo 
      Height          =   4095
      Left            =   5760
      TabIndex        =   11
      Top             =   2280
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7223
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
      Left            =   720
      TabIndex        =   25
      Top             =   3285
      Width           =   1335
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
      Left            =   720
      TabIndex        =   24
      Top             =   2805
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
      Left            =   720
      TabIndex        =   23
      Top             =   2325
      Width           =   1575
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
      Left            =   720
      TabIndex        =   22
      Top             =   5445
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
      Left            =   720
      TabIndex        =   21
      Top             =   5880
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
      Left            =   720
      TabIndex        =   20
      Top             =   4965
      Width           =   1335
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
      Left            =   720
      TabIndex        =   19
      Top             =   4485
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
      Left            =   720
      TabIndex        =   18
      Top             =   6405
      Width           =   735
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
      Left            =   720
      TabIndex        =   17
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   1935
      Left            =   480
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   3375
      Left            =   480
      Top             =   4080
      Width           =   4815
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
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   6990
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   5535
      Left            =   5520
      Top             =   1920
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   5520
      X2              =   11160
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   3480
      Top             =   7680
      Width           =   3855
   End
End
Attribute VB_Name = "frmAddMedicalTreatmentsOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Add Medical Treatments Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 09/05/08
'Date Of Last Modification: 09/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Medical_Treatments_Out Table
'--------------------------------------------------------------------------------


Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean

'The following variables will be used to autogenerate the Treatment ID to be
'displayed on the Medical Treatments Maintenance form on form load
Dim iNumOfTreatments As Integer  'This variable holds the number of records in the table
Dim strDisplay As String  'This variable will eventually hold the Treatment ID to be autogenerated


Private Sub cmdAdd_Click()
    
    If textfieldsValidations = False Then
    
        If MsgBox("Are You Sure You Wish To Add This Record?", vbYesNo + vbQuestion, "Add This Record?") = vbYes Then
    
            'Enabling the DataGrid
            dgrdMedicalTreatmentsInfo.Enabled = True
    
            txtNettTotal.Text = Val(txtNettTotal.Text) + Val(txtTotal)
            
            Call Connection 'Calling the Connection function to set up a connection with the database
            
            Call Medical_Treatments_Out    'Calling the Medical_Treatments_Out Procedure to interact with the recordset
    
            'Generate Medical Treatment ID By Utilizing the Medical_Treatments_Out Table
            With rsMedicalTreatmentsOut
    
                If .RecordCount = 0 Then    'If there are no records in the table
                
                    strDisplay = "OMT0001"
            
                Else
                
                    'Calculating the number of records and storing in a variable
                    iNumOfTreatments = .RecordCount
                    iNumOfTreatments = iNumOfTreatments + 1   'incrementing the number by 1
                
                    'The following block of code will generate the ID according
                    'to the number of records in the Medical_Treatments_Out Table
                    If iNumOfTreatments < 10 Then
                        strDisplay = "OMT000" & iNumOfTreatments
                    ElseIf iNumOfTreatments < 100 Then
                        strDisplay = "OMT00" & iNumOfTreatments
                    ElseIf iNumOfTreatments < 1000 Then
                        strDisplay = "OMT0" & iNumOfTreatments
                    ElseIf iNumOfTreatments < 10000 Then
                        strDisplay = "OMT" & iNumOfTreatments
                    End If
                
                End If
            
                .Requery    'Requerying the Table
            
                .AddNew     'Adding a new recordset
        
            End With
                                       
            saveProcedure   'Calling a function which will save the record in the database
            
            Call OutpatientsMedicalTreatments    'Calling the Outpatients Medical Treatments Function
            
            Set dgrdMedicalTreatmentsInfo.DataSource = rsOutpatientsMedicalTreatments    'Setting the datasource for the datagrid
            
        Else
        
            'Display 'No Modifications' Message
            MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"

        End If
                
        'Checking if the user wants to add another record for the same patient
        If MsgBox("Do You Wish To Add Another Medication For This Patient?", vbYesNo + vbQuestion, "Add New Medication?") = vbYes Then
            
            'Clearing All Necessary Textfields
            txtMedicineID.Text = ""
            txtMedicineName.Text = ""
            txtUnitPrice.Text = ""
            txtQty.Text = ""
                
                
            txtDateOfIssue.Text = DateTime.Date 'Setting the default value for the DateOfIssue textfield
            
            txtTotal.Text = "0" 'Setting the default value for the Total Value textfield
        
            cmdClose.Enabled = False    'Disabling the Close button because I do not want the user to close the form henceforth
            
        Else
        
            Unload Me   'Closing the form
            
        End If
        
    End If

End Sub


Private Function saveProcedure()    'This procedure will save the record into the database.
    
        
    With rsMedicalTreatmentsOut
                    
                
        'Save the user-entered data into the recordset
        .Fields(0) = strDisplay
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
        MsgBox "The Record Was Added Successfully!", vbInformation, "Succesful Save Procedure!"
            
        .Requery    'Requerying the Table
                    
    End With
    
End Function


Private Sub cmdDelete_Click()
        
    With rsOutpatientsMedicalTreatments
    
        
        'Confirm the Delete procedure with the user
        If MsgBox("Are You Sure You Wish To Remove This Medication?", vbYesNo + vbQuestion, "Remove Medication?") = vbYes Then
                
            txtNettTotal.Text = Val(txtNettTotal.Text) - .Fields(9).Value
                
            .Delete 'Delete the record from the database
                
            'Display Success Message
            MsgBox "The Medication Has Been Removed Successfully!", vbInformation, "Successfully Removed Medication!"
                    
        Else
                
            'Display 'Medication Not Removed' Message
            MsgBox "The Medication Was Not Removed!", vbExclamation, "Medication Not Removed!"
                        
        End If

        .Requery    'Requerying the Table
        
        Set dgrdMedicalTreatmentsInfo.DataSource = rsOutpatientsMedicalTreatments  'Setting the Datasource for the Datagrid
        
        cmdDelete.Enabled = False   'Disabling the Remove Button at the end
        
        'In the following code, I will be enabling the user to close the form if there are no records
        If .RecordCount = 0 Then
        
            cmdClose.Enabled = True
            
        End If
        
    End With

End Sub

Private Sub cmdClose_Click()
    
    'Obtaining confirmation from the user
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdMedicineSearchWizard_Click()
    
    frmMedicinesWizardOut.Show
    
End Sub

Private Sub cmdPatientSearchWizard_Click()  'On click of the Inpatients Search Wizard Button
    
    frmOutpatientSearchMeds.Show
    
End Sub



Private Sub dgrdMedicalTreatmentsInfo_Click()
    
    'Here, I am enabling the Remove button only if the user has already added a record
    If txtNettTotal.Text <> "0" Then
        cmdDelete.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    
    txtDateOfIssue.Text = DateTime.Date 'Displaying the date in the DateOfIssue textfield.
    
End Sub

Private Sub txtQty_Change()
    
    If txtQty.Text = "0" Then
    
        MsgBox "Error! The Figure Cannot Begin With Zero!", vbCritical, "Cannot Begin Figure With 0!"
        txtQty.Text = ""
        Exit Sub
        
    Else
    
        txtTotal.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
        
    End If
    
End Sub


Private Sub txtQty_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only digits

    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidTypingMsg.Top = 6360    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
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


