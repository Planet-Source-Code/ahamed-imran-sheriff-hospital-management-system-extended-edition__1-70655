VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmServicesMaintenance 
   Caption         =   "Hospital Servces Maintenance"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditHospitalServices.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInvalidDataMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3720
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   27
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Digits Here! Only Alphabets Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   28
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   120
      Top             =   3840
   End
   Begin VB.PictureBox picInvalidKeypressMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3720
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   105
         Width           =   2175
      End
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
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmEditHospitalServices.frx":21404
      Height          =   855
      Left            =   7560
      Picture         =   "frmEditHospitalServices.frx":218C3
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmEditHospitalServices.frx":24607
      Height          =   855
      Left            =   6480
      Picture         =   "frmEditHospitalServices.frx":24AD0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      DisabledPicture =   "frmEditHospitalServices.frx":27814
      Height          =   855
      Left            =   3240
      Picture         =   "frmEditHospitalServices.frx":27C16
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmEditHospitalServices.frx":2A95A
      Height          =   855
      Left            =   4320
      Picture         =   "frmEditHospitalServices.frx":2ADD8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmEditHospitalServices.frx":2DB1C
      Height          =   855
      Left            =   5400
      Picture         =   "frmEditHospitalServices.frx":2E002
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox txtServiceID 
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
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtAmount 
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
      TabIndex        =   4
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtAverageLengthOfStay 
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
      TabIndex        =   5
      Top             =   5400
      Width           =   2295
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
      Left            =   2880
      TabIndex        =   3
      Top             =   4200
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
      Height          =   1005
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
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
      ItemData        =   "frmEditHospitalServices.frx":30D46
      Left            =   3360
      List            =   "frmEditHospitalServices.frx":30D50
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmEditHospitalServices.frx":30D6E
      Height          =   750
      Left            =   7440
      Picture         =   "frmEditHospitalServices.frx":31183
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmEditHospitalServices.frx":3333F
      Height          =   750
      Left            =   6480
      Picture         =   "frmEditHospitalServices.frx":3371B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmEditHospitalServices.frx":358D7
      Height          =   750
      Left            =   8400
      Picture         =   "frmEditHospitalServices.frx":35CAD
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmEditHospitalServices.frx":37E69
      Height          =   750
      Left            =   9360
      Picture         =   "frmEditHospitalServices.frx":38243
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   890
   End
   Begin MSDataGridLib.DataGrid dgrdServicesInformation 
      Height          =   2535
      Left            =   5520
      TabIndex        =   7
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
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
      Caption         =   "Hospital Sevices Information Table"
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
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "***Please Note That All Non-Compulsory Fields Have Been Marked With An Asterisk"
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
      Left            =   360
      TabIndex        =   29
      Top             =   2800
      Width           =   7815
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria"
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
      Left            =   2640
      TabIndex        =   24
      Top             =   2190
      Width           =   615
   End
   Begin VB.Label lblSearchFor 
      BackStyle       =   0  'Transparent
      Caption         =   "Search For:"
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
      TabIndex        =   23
      Top             =   2190
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   3120
      Top             =   7680
      Width           =   5535
   End
   Begin VB.Label lblAmount 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount / Rate"
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
      TabIndex        =   22
      Top             =   4845
      Width           =   1455
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
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   3645
      Width           =   1575
   End
   Begin VB.Label lblAverageLengthOfStay 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Length Of Stay"
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
      Left            =   840
      TabIndex        =   20
      Top             =   5445
      Width           =   1935
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
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   4245
      Width           =   1575
   End
   Begin VB.Label lblAdditionalNotes 
      BackStyle       =   0  'Transparent
      Caption         =   "*Additional Notes"
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
      Top             =   6045
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   735
      Left            =   2280
      Top             =   1920
      Width           =   7575
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   6120
      Top             =   6360
      Width           =   4455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   3840
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblServiceInformationj 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Service Information"
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
      TabIndex        =   17
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   720
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   11520
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   360
      Y1              =   3240
      Y2              =   7440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   11520
      Y1              =   7440
      Y2              =   3240
   End
End
Attribute VB_Name = "frmServicesMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Hospital Services Maintenance Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 17/04/08
'Date Of Last Modification: 17/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Services_Maintenance Table
'-----------------------------------------------------------------------------

Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean


'The following variables will be used to autogenerate the Service ID
Dim iNumOfRecords As Integer    'This variable holds the number of records in the table
Dim strCode As String   'This variable will eventually hold the Service ID to be autogenerated


Private Sub cmdAddNew_Click()
    
    enableAllFields     'Calling a Private Function To Enable All Fields
    clearAllFields      'Calling a Private Function To Clear All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    
    'Enabling the Save Command Button & Close Command Button
    cmdSave.Enabled = True
    cmdClose.Enabled = True
    
    
    'Disabling the Search Frame
    lblCriteria.Enabled = False
    cboSearchType.Enabled = False
    lblSearchFor.Enabled = False
    txtSearch.Enabled = False
    
    'Disabling the DataGrid
    dgrdServicesInformation.Enabled = False
    
    
    Call Services_Maintenance    'Calling the Services_Maintenance Procedure to interact with the recordset
    
    'Generate Service ID By Utilizing the Services_Maintenance Table
    With rsServicesMaintenance
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strCode = "SER0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Services_Maintenance Table
            If iNumOfRecords < 10 Then
                strCode = "SER000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strCode = "SER00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strCode = "SER0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strCode = "SER" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
    'Disabling the Search Frame
    cboSearchType.Enabled = False
    txtSearch.Enabled = False
    
    'The following line of code will enter the autogenerated Service ID
    'into the Service ID textfield
    txtServiceID.Text = strCode
    
End Sub

Private Sub cmdClose_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it.
    
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        
        With rsServicesMaintenance
        
            'Making sure that the user wants to update the record
            If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
            
                'The following if else condition ensures that The Additional Notes
                'textfield will not be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtAdditionalNotes.Text = "" Then
                    txtAdditionalNotes.Text = "-"
                End If
                    
                    
                'Save the user-entered data into the recordset
                .Fields(0) = txtServiceID.Text
                .Fields(1) = txtServiceName.Text
                .Fields(2) = txtAmount.Text
                .Fields(3) = txtAverageLengthOfStay.Text
                .Fields(4) = txtAdditionalNotes.Text
                
                .Update
                    
                .Requery
            
            
                'Display Success Message
                MsgBox "The Record Was Updated Successfully!", vbInformation, "Succesful Update Procedure"
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            Else
            
                'Display 'No Modifications' Message
                MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
                .CancelUpdate   'Cancel the Update Procedure
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            End If
            
            .Requery    'Requerying the Table
            
        End With
        
    End If
    
End Sub

Private Sub dgrdServicesInformation_Click()
    
    'Enabling the Update Button & the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the Navigation Buttons
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    With rsServicesMaintenance
    
        'Entering the values in the particular record into the fields on the interface
        txtServiceID.Text = .Fields(0).Value
        txtServiceName.Text = .Fields(1).Value
        txtAmount.Text = .Fields(2).Value
        txtAverageLengthOfStay.Text = .Fields(3).Value
        txtAdditionalNotes.Text = .Fields(4).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub



Private Sub Form_Load()
    
    Call Connection  'Calling the Connection Procedure
    
    Call Services_Maintenance  'Calling the Services_Maintenance Procedure to interact with the recordset
    
    disableAllFields  'Calling a Private Function To Disable All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    'Enabling  the First Button and the Last Button
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Add New Button & the Close Button
    cmdAddNew.Enabled = True
    cmdClose.Enabled = True
    
    'Enabling the Search Frame
    lblCriteria.Enabled = True
    cboSearchType.Enabled = True
    lblSearchFor.Enabled = True
    txtSearch.Enabled = True
    
    
    'Enabling the DataGrid
    dgrdServicesInformation.Enabled = True
    
    Set dgrdServicesInformation.DataSource = rsServicesMaintenance  'Setting the DataSource of the DataGrid

End Sub

Private Function disableAllFields() 'This function will disable all fields on the interface

    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will disable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Then
        eachField.Enabled = False
    End If

    Next

End Function


Private Function disableAllButtons() 'This function will disable all command buttons on the interface

    On Error Resume Next
    For Each eachButton In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will disable all Command Buttons
    If TypeOf eachButton Is CommandButton Then
        eachButton.Enabled = False
    End If

    Next

End Function


Private Function enableAllFields() 'This function will enable all fields on the interface

    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will enable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Then
        eachField.Enabled = True
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
    
    
    'Clearing the Search Textfield to Enable All Records To Be
    'Displayed On The Grid
    txtSearch.Text = ""
    
    
    With rsServicesMaintenance
    
        .MoveFirst  'Moving to the first record
        
        'Entering the values in the particular record into the fields on the interface
        txtServiceID.Text = .Fields(0).Value
        txtServiceName.Text = .Fields(1).Value
        txtAmount.Text = .Fields(2).Value
        txtAverageLengthOfStay.Text = .Fields(3).Value
        txtAdditionalNotes.Text = .Fields(4).Value
    
    End With
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub


Private Sub cmdPrevious_Click() 'This function will Navigate to the Previous Record
    
    
    cmdNext.Enabled = True  'Enabling the Next Button
    cmdLast.Enabled = True  'Enabling the Last Button
    
    
    'Clearing the Search Textfield to Enable All Records To Be
    'Displayed On The Grid
    txtSearch.Text = ""
    
    
    With rsServicesMaintenance
    
        .MovePrevious   'Moving to the previous record
        
        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If
    
        'Entering the values in the particular record into the fields on the interface
        txtServiceID.Text = .Fields(0).Value
        txtServiceName.Text = .Fields(1).Value
        txtAmount.Text = .Fields(2).Value
        txtAverageLengthOfStay.Text = .Fields(3).Value
        txtAdditionalNotes.Text = .Fields(4).Value
        
    End With
    
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub


Private Sub cmdNext_Click() 'This function will Navigate to the Next Record
    
    
    cmdPrevious.Enabled = True  'Enabling the Previous Button
    cmdFirst.Enabled = True 'Enabling the First Button
    
    
    'Clearing the Search Textfield to Enable All Records To Be
    'Displayed On The Grid
    txtSearch.Text = ""
    
    
    With rsServicesMaintenance
    
        .MoveNext   'Moving to the Next Record
        
        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If
        
        'Entering the values in the particular record into the fields on the interface
        txtServiceID.Text = .Fields(0).Value
        txtServiceName.Text = .Fields(1).Value
        txtAmount.Text = .Fields(2).Value
        txtAverageLengthOfStay.Text = .Fields(3).Value
        txtAdditionalNotes.Text = .Fields(4).Value
        
    End With
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
    
End Sub


Private Sub cmdLast_Click() 'This function will Navigate to the Last Record
    
    'Enabling / Diabling the Navigation Buttons as necessary
    cmdLast.Enabled = False
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = False

    
    'Clearing the Search Textfield to Enable All Records To Be
    'Displayed On The Grid
    txtSearch.Text = ""
    
    
    With rsServicesMaintenance
    
        .MoveLast   'Moving to the last record
        
        'Entering the values in the particular record into the fields on the interface
        txtServiceID.Text = .Fields(0).Value
        txtServiceName.Text = .Fields(1).Value
        txtAmount.Text = .Fields(2).Value
        txtAverageLengthOfStay.Text = .Fields(3).Value
        txtAdditionalNotes.Text = .Fields(4).Value
        
    End With
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub




Private Sub tmrErrMsg_Timer()

    Static i As Integer

    If i < 200000 Then     'Validation Msg Viewing Time Period
        picInvalidKeypressMsg.Visible = False
        picInvalidDataMsg.Visible = False
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If

End Sub

Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsServicesMaintenance
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[ServiceID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[ServiceName] Like '" & txtSearch.Text & "%" & "'"
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
        
        Call Services_Maintenance
        
        Set dgrdServicesInformation.DataSource = rsServicesMaintenance
        
    End If
    
End Sub


Private Sub cmdSave_Click()     'This function will save all the user's data in the database
    
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        
        With rsServicesMaintenance
            
            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
                
                'The following if else condition ensures that The Additional Notes
                'textfield will not be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtAdditionalNotes.Text = "" Then
                    txtAdditionalNotes.Text = "-"
                End If
                
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtServiceID.Text
                .Fields(1) = txtServiceName.Text
                .Fields(2) = txtAmount.Text
                .Fields(3) = txtAverageLengthOfStay.Text
                .Fields(4) = txtAdditionalNotes.Text
                
                .Update
                
                .Requery    'Requerying the Table
                
                'Display Success Message
                MsgBox "The Record Was Saved Successfully!", vbInformation, "Succesful Save Procedure"
                
                
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
        
    End If
        

End Sub


Private Function textfieldsValidations() As Boolean  'This function will validate all fields
    
    Flag = True 'Setting the Flag variable to True
    
    'Checking if the Service Name textfield is empty
    If txtServiceName.Text = "" Then
        txtServiceName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtServiceName.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    
    'Checking if the Amount textfield is empty
    If txtAmount.Text = "" Then
        txtAmount.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtAmount.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Average Duration textfield is empty
    If txtAverageLengthOfStay.Text = "" Then
        txtAverageLengthOfStay.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtAverageLengthOfStay.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
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


Private Sub cmdDelete_Click()   'This function will delete a record from the database
    
    'Check for the record selection
    If txtServiceID.Text = "" Then
    
        MsgBox "Error! No Record Has Been Selected", vbCritical, "No Record Selected!"
    
    Else
    
        With rsServicesMaintenance
        
            'Confirm the Delete procedure with the user
            If MsgBox("Are You Sure You Wish To Delete Service ID " & txtServiceID.Text & "'s Record?", vbYesNo + vbQuestion, "Delete Record?") = vbYes Then
        
                .Delete 'Delete the record from the database
                
                'Display Success Message
                MsgBox "The Record Has Been Deleted Successfully!", vbInformation, "Successful Delete Procedure!"
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            Else
                
                'Display 'Delete Procedure Cancelled' Message
                MsgBox "The Delete Procedure Was Cancelled!", vbExclamation, "Delete Procedure Cancelled!"
                
                Form_Load   'Calling the Form_Load Procedure

                clearAllFields  'Calling a Private Function To Clear All Fields
        
            End If

            .Requery    'Requerying the Table
        
        End With
        
    End If

End Sub



Private Sub txtServiceName_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets
    
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 4200    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub



Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only Digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 4800    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub



