VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMedicinesMaintenance 
   Caption         =   "Medicines Maintenance Module"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditMedicines.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   240
      Top             =   4440
   End
   Begin VB.PictureBox picInvalidKeypressMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3480
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.ComboBox cboDosageForm 
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
      ItemData        =   "frmEditMedicines.frx":1FB21
      Left            =   2880
      List            =   "frmEditMedicines.frx":1FB37
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmEditMedicines.frx":1FB68
      Height          =   855
      Left            =   7560
      Picture         =   "frmEditMedicines.frx":20027
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmEditMedicines.frx":22D6B
      Height          =   855
      Left            =   6480
      Picture         =   "frmEditMedicines.frx":23234
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      DisabledPicture =   "frmEditMedicines.frx":25F78
      Height          =   855
      Left            =   3240
      Picture         =   "frmEditMedicines.frx":2637A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmEditMedicines.frx":290BE
      Height          =   855
      Left            =   4320
      Picture         =   "frmEditMedicines.frx":2953C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmEditMedicines.frx":2C280
      Height          =   855
      Left            =   5400
      Picture         =   "frmEditMedicines.frx":2C766
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7800
      Width           =   975
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
      Height          =   885
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   6480
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
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtUnitsInStock 
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
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox txtReOrderLevel 
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
      TabIndex        =   7
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox txtUnitPrice 
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
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox txtMedicineID 
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
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmEditMedicines.frx":2F4AA
      Height          =   750
      Left            =   9360
      Picture         =   "frmEditMedicines.frx":2F884
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmEditMedicines.frx":31A40
      Height          =   750
      Left            =   8400
      Picture         =   "frmEditMedicines.frx":31E16
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmEditMedicines.frx":33FD2
      Height          =   750
      Left            =   6480
      Picture         =   "frmEditMedicines.frx":343AE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmEditMedicines.frx":3656A
      Height          =   750
      Left            =   7440
      Picture         =   "frmEditMedicines.frx":3697F
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   890
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
      Top             =   2130
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
      ItemData        =   "frmEditMedicines.frx":38B3B
      Left            =   3360
      List            =   "frmEditMedicines.frx":38B48
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2130
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid dgrdMedicineInfo 
      Height          =   2655
      Left            =   5520
      TabIndex        =   9
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
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
      Caption         =   "Medicine Information Table"
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
      TabIndex        =   31
      Top             =   2800
      Width           =   7815
   End
   Begin VB.Label lblDosageForm 
      BackStyle       =   0  'Transparent
      Caption         =   "Dosage Form"
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
      Top             =   4605
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   3120
      Top             =   7680
      Width           =   5535
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
      TabIndex        =   27
      Top             =   6525
      Width           =   1695
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
      Height          =   255
      Left            =   840
      TabIndex        =   26
      Top             =   4125
      Width           =   1575
   End
   Begin VB.Label lblUnitsInStock 
      BackStyle       =   0  'Transparent
      Caption         =   "Units In Stock"
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
      TabIndex        =   25
      Top             =   5565
      Width           =   1575
   End
   Begin VB.Label lblReOrderLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-order Level"
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
      Top             =   6045
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
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   3645
      Width           =   1575
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
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   5085
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   11520
      Y1              =   7560
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   360
      Y1              =   3240
      Y2              =   7560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   11520
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   720
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblMedicineInformation 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine Information"
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
      TabIndex        =   21
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   3120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   6120
      Top             =   6360
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   735
      Left            =   2280
      Top             =   1900
      Width           =   7455
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
      TabIndex        =   20
      Top             =   2160
      Width           =   1695
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
      Left            =   2520
      TabIndex        =   19
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "frmMedicinesMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Medicines Maintenance Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 14/04/08
'Date Of Last Modification: 14/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Doctors_Maintenance Table
'----------------------------------------------------------------------------

Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean


'The following variables will be used to autogenerate the Medicine ID
Dim iNumOfRecords As Integer    'This variable holds the number of records in the table
Dim strCode As String   'This variable will eventually hold the Medicine ID to be autogenerated


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
    dgrdMedicineInfo.Enabled = False
    
    
    Call Medicines_Maintenance    'Calling the Medicines_Maintenance Procedure to interact with the recordset
    
    'Generate Medicine ID By Utilizing the Medicines_Maintenance Table
    With rsMedicinesMaintenance
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strCode = "MED0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Medicines_Maintenance Table
            If iNumOfRecords < 10 Then
                strCode = "MED000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strCode = "MED00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strCode = "MED0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strCode = "MED" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
    'Disabling the Search Frame
    cboSearchType.Enabled = False
    txtSearch.Enabled = False
    
    'The following line of code will enter the autogenerated Medicine ID
    'into the Medicine ID textfield
    txtMedicineID.Text = strCode
    
End Sub

Private Sub cmdClose_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it.
    
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        'Validation To Ensure That The Medicine Name is not Greater than 25 Characters in Length
        If Len(txtMedicineName.Text) > 25 Then
            MsgBox "Error! The Medicine Name Textfield Cannot Consist Of More Than 25 Characters!", vbCritical, "Error In Medicine Name!"
            Exit Sub
        End If
        
        With rsMedicinesMaintenance
        
            'Making sure that the user wants to update the record
            If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
            
                'The following if else condition ensures that The Additional Notes
                'textfield will not be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtAdditionalNotes.Text = "" Then
                    txtAdditionalNotes.Text = "-"
                End If
                    
                    
                'Save the user-entered data into the recordset
                .Fields(0) = txtMedicineID.Text
                .Fields(1) = txtMedicineName.Text
                .Fields(2) = cboDosageForm.Text
                .Fields(3) = txtUnitPrice.Text
                .Fields(4) = txtUnitsInStock.Text
                .Fields(5) = txtReOrderLevel.Text
                .Fields(6) = txtAdditionalNotes.Text
                
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

Private Sub dgrdMedicineInfo_Click()
    
    'Enabling the Update Button & the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the Navigation Buttons
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    
    With rsMedicinesMaintenance
    
        'Entering the values in the particular record into the fields on the interface
        txtMedicineID.Text = .Fields(0).Value
        txtMedicineName.Text = .Fields(1).Value
        cboDosageForm.Text = .Fields(2).Value
        txtUnitPrice.Text = .Fields(3).Value
        txtUnitsInStock.Text = .Fields(4).Value
        txtReOrderLevel.Text = .Fields(5).Value
        txtAdditionalNotes.Text = .Fields(6).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub

Private Sub Form_Load()
    
    Call Connection  'Calling the Connection Procedure
    
    Call Medicines_Maintenance  'Calling the Medicines_Maintenance Procedure to interact with the recordset
    
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
    dgrdMedicineInfo.Enabled = True
    
    Set dgrdMedicineInfo.DataSource = rsMedicinesMaintenance  'Setting the DataSource of the DataGrid

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
    
    'The following lines will set the normal display values of the Dosage Form ComboBox
    cboDosageForm.Text = ""
    
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
    
    
    With rsMedicinesMaintenance
    
        .MoveFirst  'Moving to the first record
        
        'Entering the values in the particular record into the fields on the interface
        txtMedicineID.Text = .Fields(0).Value
        txtMedicineName.Text = .Fields(1).Value
        cboDosageForm.Text = .Fields(2).Value
        txtUnitPrice.Text = .Fields(3).Value
        txtUnitsInStock.Text = .Fields(4).Value
        txtReOrderLevel.Text = .Fields(5).Value
        txtAdditionalNotes.Text = .Fields(6).Value
    
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
    
    
    With rsMedicinesMaintenance
    
        .MovePrevious   'Moving to the previous record
        
        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If
    
        'Entering the values in the particular record into the fields on the interface
        txtMedicineID.Text = .Fields(0).Value
        txtMedicineName.Text = .Fields(1).Value
        cboDosageForm.Text = .Fields(2).Value
        txtUnitPrice.Text = .Fields(3).Value
        txtUnitsInStock.Text = .Fields(4).Value
        txtReOrderLevel.Text = .Fields(5).Value
        txtAdditionalNotes.Text = .Fields(6).Value
        
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
    
    
    With rsMedicinesMaintenance
    
        .MoveNext   'Moving to the Next Record
        
        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If
        
        'Entering the values in the particular record into the fields on the interface
        txtMedicineID.Text = .Fields(0).Value
        txtMedicineName.Text = .Fields(1).Value
        cboDosageForm.Text = .Fields(2).Value
        txtUnitPrice.Text = .Fields(3).Value
        txtUnitsInStock.Text = .Fields(4).Value
        txtReOrderLevel.Text = .Fields(5).Value
        txtAdditionalNotes.Text = .Fields(6).Value
        
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
    
    With rsMedicinesMaintenance
    
        .MoveLast   'Moving to the last record
        
        'Entering the values in the particular record into the fields on the interface
        txtMedicineID.Text = .Fields(0).Value
        txtMedicineName.Text = .Fields(1).Value
        cboDosageForm.Text = .Fields(2).Value
        txtUnitPrice.Text = .Fields(3).Value
        txtUnitsInStock.Text = .Fields(4).Value
        txtReOrderLevel.Text = .Fields(5).Value
        txtAdditionalNotes.Text = .Fields(6).Value
        
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
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If
    
End Sub

Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsMedicinesMaintenance
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[MedicineID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[MedicineName] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
                    .Filter = "[DosageForm] Like '" & txtSearch.Text & "%" & "'"
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
        
        Call Medicines_Maintenance
        
        Set dgrdMedicineInfo.DataSource = rsMedicinesMaintenance
        
    End If
    
End Sub


Private Sub cmdSave_Click()     'This function will save all the user's data in the database
    
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        'Validation To Ensure That The Medicine Name is not Greater than 25 Characters in Length
        If Len(txtMedicineName.Text) > 25 Then
            MsgBox "Error! The Medicine Name Textfield Cannot Consist Of More Than 25 Characters!", vbCritical, "Error In Medicine Name!"
            Exit Sub
        End If
        
        
        With rsMedicinesMaintenance
            
            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
                
                'The following if else condition ensures that The Additional Notes
                'textfield will not be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtAdditionalNotes.Text = "" Then
                    txtAdditionalNotes.Text = "-"
                End If
                
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtMedicineID.Text
                .Fields(1) = txtMedicineName.Text
                .Fields(2) = cboDosageForm.Text
                .Fields(3) = txtUnitPrice.Text
                .Fields(4) = txtUnitsInStock.Text
                .Fields(5) = txtReOrderLevel.Text
                .Fields(6) = txtAdditionalNotes.Text
            
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
    
    'Checking if the Medicine Name textfield is empty
    If txtMedicineName.Text = "" Then
        txtMedicineName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtMedicineName.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the user has made a selection in the Dosage Form ComboBox
    If cboDosageForm.Text = "" Then
        cboDosageForm.BackColor = &H80000018    'Highlighting the ComboBox in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboDosageForm.BackColor = &H80000004    'Bringing the ComboBox BackColour back to normal
    End If
    
    'Checking if the Unit Price textfield is empty
    If txtUnitPrice.Text = "" Then
        txtUnitPrice.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtUnitPrice.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Units In Stock textfield is empty
    If txtUnitsInStock.Text = "" Then
        txtUnitsInStock.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtUnitsInStock.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Re-order Level textfield is empty
    If txtReOrderLevel.Text = "" Then
        txtReOrderLevel.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtReOrderLevel.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
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
    If txtMedicineID.Text = "" Then
    
        MsgBox "Error! No Record Has Been Selected", vbCritical, "No Record Selected!"
    
    Else
    
        With rsMedicinesMaintenance
        
            'Confirm the Delete procedure with the user
            If MsgBox("Are You Sure You Wish To Delete Medicine ID " & txtMedicineID.Text & "'s Record?", vbYesNo + vbQuestion, "Delete Record?") = vbYes Then
        
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


Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only Digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 5040    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub



Private Sub txtUnitsInStock_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only Digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 5520    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtReOrderLevel_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only Digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidKeypressMsg.Top = 6000    'Validation Note View
        picInvalidKeypressMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub
