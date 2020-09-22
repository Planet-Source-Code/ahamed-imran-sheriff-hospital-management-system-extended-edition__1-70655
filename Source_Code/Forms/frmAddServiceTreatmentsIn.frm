VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmAddServiceTreatmentsIn 
   Caption         =   "Add Service Treatments"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddServiceTreatmentsIn.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmAddServiceTreatmentsIn.frx":1DEB8
      Height          =   855
      Left            =   6240
      Picture         =   "frmAddServiceTreatmentsIn.frx":1E377
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7320
      Width           =   975
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3480
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3000
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdPatientSearchWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      ToolTipText     =   "Click Here to select a Patient"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox txtServiceID 
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
      TabIndex        =   4
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdServiceSearchWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "Click Here to select a Medicine"
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txtServiceName 
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
      TabIndex        =   6
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtServiceCharge 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtTreatmentDate 
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
      Top             =   6120
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdAdd 
      DisabledPicture =   "frmAddServiceTreatmentsIn.frx":210BB
      Height          =   855
      Left            =   3840
      Picture         =   "frmAddServiceTreatmentsIn.frx":214BD
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmAddServiceTreatmentsIn.frx":24201
      Enabled         =   0   'False
      Height          =   855
      Left            =   5040
      Picture         =   "frmAddServiceTreatmentsIn.frx":246CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgrdServiceTreatmentsInfo 
      Height          =   3255
      Left            =   5880
      TabIndex        =   9
      Top             =   2520
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5741
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
      Caption         =   "Service Treatments Information Table"
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
      TabIndex        =   21
      Top             =   3525
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
      Left            =   840
      TabIndex        =   20
      Top             =   3045
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
      Left            =   840
      TabIndex        =   19
      Top             =   2565
      Width           =   1575
   End
   Begin VB.Label lblServiceName 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Name"
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
      TabIndex        =   18
      Top             =   5205
      Width           =   1335
   End
   Begin VB.Label lblServiceID 
      BackStyle       =   0  'Transparent
      Caption         =   "Service ID"
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
      TabIndex        =   17
      Top             =   4725
      Width           =   1335
   End
   Begin VB.Label lblServiceCharge 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charge"
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
      TabIndex        =   16
      Top             =   5685
      Width           =   1455
   End
   Begin VB.Label lblTreatmentDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Treatment Date"
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
      TabIndex        =   15
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   1935
      Left            =   600
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   2415
      Left            =   600
      Top             =   4320
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
      Left            =   6720
      TabIndex        =   14
      Top             =   6270
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   4575
      Left            =   5640
      Top             =   2160
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   11280
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   3600
      Top             =   7200
      Width           =   3855
   End
End
Attribute VB_Name = "frmAddServiceTreatmentsIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Add Service Treatments Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing):  Isham Sally
'Start Date: 10/05/08
'Date Of Last Modification: 10/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Service_Treatments Table
'--------------------------------------------------------------------------------


Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean

'The following variables will be used to autogenerate the Treatment ID to be
'displayed on the Service Treatments Maintenance form on form load
Dim iNumOfTreatments As Integer  'This variable holds the number of records in the table
Dim strDisplay As String  'This variable will eventually hold the Treatment ID to be autogenerated


Private Sub cmdAdd_Click()
    
    If textfieldsValidations = False Then
    
        If MsgBox("Are You Sure You Wish To Add This Record?", vbYesNo + vbQuestion, "Add This Record?") = vbYes Then
    
            'Enabling the DataGrid
            dgrdServiceTreatmentsInfo.Enabled = True
    
            txtNettTotal.Text = Val(txtNettTotal.Text) + Val(txtServiceCharge.Text)
            
            Call Connection 'Calling the Connection function to set up a connection with the database
            
            Call Service_Treatments    'Calling the Service_Treatments Procedure to interact with the recordset
    
            'Generate Service Treatment ID By Utilizing the Service_Treatments Table
            With rsServiceTreatments
    
                If .RecordCount = 0 Then    'If there are no records in the table
                
                    strDisplay = "STR0001"
            
                Else
                
                    'Calculating the number of records and storing in a variable
                    iNumOfTreatments = .RecordCount
                    iNumOfTreatments = iNumOfTreatments + 1   'incrementing the number by 1
                
                    'The following block of code will generate the ID according
                    'to the number of records in the Service_Treatments Table
                    If iNumOfTreatments < 10 Then
                        strDisplay = "STR000" & iNumOfTreatments
                    ElseIf iNumOfTreatments < 100 Then
                        strDisplay = "STR00" & iNumOfTreatments
                    ElseIf iNumOfTreatments < 1000 Then
                        strDisplay = "STR0" & iNumOfTreatments
                    ElseIf iNumOfTreatments < 10000 Then
                        strDisplay = "STR" & iNumOfTreatments
                    End If
                
                End If
                            
                .Requery    'Requerying the Table
            
                .AddNew     'Adding a new recordset
        
            End With
                                       
            saveProcedure   'Calling a function which will save the record in the database
            
            Call InpatientsServiceTreatments    'Calling the InpatientsServiceTreatments Function
            
            Set dgrdServiceTreatmentsInfo.DataSource = rsInpatientsServiceTreatments    'Setting the datasource for the datagrid
            
        Else
        
            'Display 'No Modifications' Message
            MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"

        End If
                
        'Checking if the user wants to add another record for the same patient
        If MsgBox("Do You Wish To Add Another Service Treatment For This Patient?", vbYesNo + vbQuestion, "Add New Service Treatment?") = vbYes Then
            
            'Clearing All Necessary Textfields
            txtServiceID.Text = ""
            txtServiceName.Text = ""
            txtServiceCharge.Text = ""
                
                
            txtTreatmentDate.Text = DateTime.Date 'Setting the default value for the Treatment Date textfield
                    
            cmdClose.Enabled = False    'Disabling the Close button because I do not want the user to close the form henceforth
            
        Else
        
            Unload Me   'Closing the form
            
        End If
        
    End If

End Sub


Private Function saveProcedure()    'This procedure will save the record into the database.
    
        
    With rsServiceTreatments
                    
                
        'Save the user-entered data into the recordset
        .Fields(0) = strDisplay
        .Fields(1) = txtPatientID.Text
        .Fields(2) = txtFirstName.Text
        .Fields(3) = txtSurname.Text
        .Fields(4) = txtServiceID.Text
        .Fields(5) = txtServiceName.Text
        .Fields(6) = txtServiceCharge.Text
        .Fields(7) = txtTreatmentDate.Text
            
        .Update
                
        'Display Success Message
        MsgBox "The Record Was Added Successfully!", vbInformation, "Succesful Save Procedure!"
            
        .Requery    'Requerying the Table
                    
    End With
    
End Function


Private Sub cmdDelete_Click()
        
    With rsInpatientsServiceTreatments
    
        
        'Confirm the Delete procedure with the user
        If MsgBox("Are You Sure You Wish To Remove This Service?", vbYesNo + vbQuestion, "Remove Service?") = vbYes Then
                
            txtNettTotal.Text = Val(txtNettTotal.Text) - .Fields(6).Value
                
            .Delete 'Delete the record from the database
                
            'Display Success Message
            MsgBox "The Service Has Been Removed Successfully!", vbInformation, "Successfully Removed Service!"
                    
        Else
                
            'Display 'Service Not Removed' Message
            MsgBox "The Service Was Not Removed!", vbExclamation, "Service Not Removed!"
                        
        End If

        .Requery    'Requerying the Table
        
        Set dgrdServiceTreatmentsInfo.DataSource = rsInpatientsServiceTreatments  'Setting the Datasource for the Datagrid
        
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



Private Sub cmdPatientSearchWizard_Click()  'On click of the Inpatients Search Wizard Button
    
    frmInpatientsSearchAndFind.Show
    
End Sub



Private Sub cmdServiceSearchWizard_Click()
    
    frmServicesSearchAndFind.Show
    
End Sub

Private Sub dgrdServiceTreatmentsInfo_Click()
    
    'Here, I am enabling the Remove button only if the user has already added a record
    If txtNettTotal.Text <> "0" Then
        cmdDelete.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    
    txtTreatmentDate.Text = DateTime.Date 'Displaying the Date in the Treatment Date textfield.
    
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
    
    
    'Checking if the Service ID textfield is empty
    If txtServiceID.Text = "" Then
        txtServiceID.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtServiceName.BackColor = &H80000018   'Highlighting the textfield in a different colour
        txtServiceCharge.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtServiceID.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtServiceName.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        txtServiceCharge.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
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


