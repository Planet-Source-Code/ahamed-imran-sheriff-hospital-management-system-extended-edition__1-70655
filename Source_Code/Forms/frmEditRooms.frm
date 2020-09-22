VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmRoomsMaintenance 
   Caption         =   "Rooms Maintenance Module"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditRooms.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRoomNumber 
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
      TabIndex        =   31
      Top             =   3840
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
      Left            =   7080
      TabIndex        =   28
      Top             =   2060
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmEditRooms.frx":1EB33
      Height          =   855
      Left            =   7440
      Picture         =   "frmEditRooms.frx":1EFF2
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmEditRooms.frx":21D36
      Height          =   855
      Left            =   6360
      Picture         =   "frmEditRooms.frx":221FF
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      DisabledPicture =   "frmEditRooms.frx":24F43
      Height          =   855
      Left            =   3120
      Picture         =   "frmEditRooms.frx":25345
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmEditRooms.frx":28089
      Height          =   855
      Left            =   4200
      Picture         =   "frmEditRooms.frx":28507
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmEditRooms.frx":2B24B
      Height          =   855
      Left            =   5280
      Picture         =   "frmEditRooms.frx":2B731
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox txtRoomID 
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
      TabIndex        =   15
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtDepartmentName 
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
      TabIndex        =   14
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtWardNumber 
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
      Top             =   5760
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
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox txtRoomCost 
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
      TabIndex        =   11
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdDepartmentSearchWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Click Here To Select A Department"
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox txtDepartmentID 
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
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdWardSearchWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Click Here To Select A Ward"
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtWardID 
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
      TabIndex        =   7
      Top             =   5280
      Width           =   1815
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
      ItemData        =   "frmEditRooms.frx":2E475
      Left            =   3360
      List            =   "frmEditRooms.frx":2E47F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2060
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmEditRooms.frx":2E499
      Height          =   750
      Left            =   7440
      Picture         =   "frmEditRooms.frx":2E8AE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmEditRooms.frx":30A6A
      Height          =   750
      Left            =   6480
      Picture         =   "frmEditRooms.frx":30E46
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmEditRooms.frx":33002
      Height          =   750
      Left            =   8400
      Picture         =   "frmEditRooms.frx":333D8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   890
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmEditRooms.frx":35594
      Height          =   750
      Left            =   9360
      Picture         =   "frmEditRooms.frx":3596E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   890
   End
   Begin MSDataGridLib.DataGrid dgrdRoomsInformation 
      Height          =   2655
      Left            =   5520
      TabIndex        =   5
      Top             =   3360
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
      Caption         =   "Rooms Information Table"
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
      TabIndex        =   33
      Top             =   2720
      Width           =   7815
   End
   Begin VB.Label lblRoomNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Number"
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
      Top             =   3885
      Width           =   1815
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
      Left            =   5880
      TabIndex        =   30
      Top             =   2100
      Width           =   1095
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
      TabIndex        =   29
      Top             =   2100
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   3000
      Top             =   7680
      Width           =   5535
   End
   Begin VB.Label lblDepartmentName 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Name"
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
      Width           =   1815
   End
   Begin VB.Label lblRoomID 
      BackStyle       =   0  'Transparent
      Caption         =   "Room ID"
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
      Top             =   3405
      Width           =   1575
   End
   Begin VB.Label lblWardNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward Number"
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
      Top             =   5805
      Width           =   1335
   End
   Begin VB.Label lblWardID 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward ID"
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
      Top             =   5325
      Width           =   1575
   End
   Begin VB.Label lblDepartmentID 
      BackStyle       =   0  'Transparent
      Caption         =   "Department ID"
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
      TabIndex        =   18
      Top             =   4365
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
      TabIndex        =   17
      Top             =   6765
      Width           =   1695
   End
   Begin VB.Label lblRoomCost 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Cost"
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
      Top             =   6285
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   735
      Left            =   2280
      Top             =   1850
      Width           =   7455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   6120
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   2880
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblRoomsInformation 
      BackStyle       =   0  'Transparent
      Caption         =   "Rooms Information"
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
      TabIndex        =   6
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   720
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   11520
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   360
      Y1              =   3120
      Y2              =   7560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   11520
      Y1              =   7560
      Y2              =   3120
   End
End
Attribute VB_Name = "frmRoomsMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Rooms Maintenance Interface
'Programmer: Isham Sally
'Quality Assurance Engineer (Testing): Imran Sheriff
'Start Date: 18/04/08
'Date Of Last Modification: 18/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Rooms_Maintenance Table, Departments_Maintenance Table, Wards_Maintenance Table
'-----------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean


'The following variables will be used to autogenerate the Ward ID
Dim iNumOfRecords As Integer    'This variable holds the number of records in the table
Dim strCode As String   'This variable will eventually hold the Room ID to be autogenerated


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
    dgrdRoomsInformation.Enabled = False
    
    
    Call Rooms_Maintenance    'Calling the Rooms_Maintenance Procedure to interact with the recordset
    
    'Generate Room ID By Utilizing the Rooms_Maintenance Table
    With rsRoomsMaintenance
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strCode = "ROM0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Rooms_Maintenance Table
            If iNumOfRecords < 10 Then
                strCode = "ROM000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strCode = "ROM00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strCode = "ROM0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strCode = "ROM" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
    'Disabling the Search Frame
    cboSearchType.Enabled = False
    txtSearch.Enabled = False
    
    'The following line of code will enter the autogenerated Room ID
    'into the Room ID textfield
    txtRoomID.Text = strCode
    
    'The following line of code will enter the Room Number
    'into the Room Number textfield
    txtRoomNumber.Text = "" & iNumOfRecords
    
End Sub

Private Sub cmdClose_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdDepartmentSearchWizard_Click()
    
    frmDepartmentSearchWizardRooms.Show
    
End Sub

Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it.
        
        
    With rsRoomsMaintenance
        
        'Making sure that the user wants to update the record
        If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
            
            'The following if else condition ensures that The Additional Notes
            'textfield will not be completely blank when saving in the database.
            'This has been done in order to avoid errors.
            If txtAdditionalNotes.Text = "" Then
                txtAdditionalNotes.Text = "-"
            End If
                    
                    
            'Save the user-entered data into the recordset
            .Fields(0) = txtRoomID.Text
            .Fields(1) = txtRoomNumber.Text
            .Fields(2) = txtDepartmentID.Text
            .Fields(3) = txtDepartmentName.Text
            .Fields(4) = txtWardID.Text
            .Fields(5) = txtWardNumber.Text
            .Fields(6) = txtRoomCost.Text
            .Fields(7) = txtAdditionalNotes.Text
                
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
    
End Sub

Private Sub dgrdRoomsInformation_Click()
    
    'Enabling the Update Button & the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the Navigation Buttons
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    With rsRoomsMaintenance
    
        'Entering the values in the particular record into the fields on the interface
        txtRoomID.Text = .Fields(0).Value
        txtRoomNumber.Text = .Fields(1).Value
        txtDepartmentID.Text = .Fields(2).Value
        txtDepartmentName.Text = .Fields(3).Value
        txtWardID.Text = .Fields(4).Value
        txtWardNumber.Text = .Fields(5).Value
        txtRoomCost.Text = .Fields(6).Value
        txtAdditionalNotes.Text = .Fields(7).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub



Private Sub cmdWardSearchWizard_Click()
    
    If txtDepartmentID.Text = "" Then
        MsgBox "You Have Not Selected A Department! Please Select A Department First!", vbExclamation, "No Department Selected!"
    Else
        frmWardSearchWizard.Show
    End If
    
End Sub

Private Sub Form_Load()
    
    Call Connection  'Calling the Connection Procedure
    
    Call Rooms_Maintenance  'Calling the Rooms_Maintenance Procedure to interact with the recordset
    
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
    dgrdRoomsInformation.Enabled = True
    
    Set dgrdRoomsInformation.DataSource = rsRoomsMaintenance  'Setting the DataSource of the DataGrid

End Sub

Private Function disableAllFields() 'This function will disable all fields on the interface

    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will disable all TextBoxes and ComboBoxes
    If TypeOf eachField Is TextBox Or TypeOf eachField Is ComboBox Then
        eachField.Enabled = False
    End If

    Next
    
    'Disabling the Department Search Wizard Button
    cmdDepartmentSearchWizard.Enabled = False
    
    'Disabling the Ward Search Wizard Button
    cmdWardSearchWizard.Enabled = False

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
    
    'Enabling the Department Search Wizard Button
    cmdDepartmentSearchWizard.Enabled = True
    
    'Enabling the Ward Search Wizard Button
    cmdWardSearchWizard.Enabled = True

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
    
    
    With rsRoomsMaintenance
    
        .MoveFirst  'Moving to the first record
        
        'Entering the values in the particular record into the fields on the interface
        txtRoomID.Text = .Fields(0).Value
        txtRoomNumber.Text = .Fields(1).Value
        txtDepartmentID.Text = .Fields(2).Value
        txtDepartmentName.Text = .Fields(3).Value
        txtWardID.Text = .Fields(4).Value
        txtWardNumber.Text = .Fields(5).Value
        txtRoomCost.Text = .Fields(6).Value
        txtAdditionalNotes.Text = .Fields(7).Value
    
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
    
    
    With rsRoomsMaintenance
    
        .MovePrevious   'Moving to the previous record
        
        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If
    
        'Entering the values in the particular record into the fields on the interface
        txtRoomID.Text = .Fields(0).Value
        txtRoomNumber.Text = .Fields(1).Value
        txtDepartmentID.Text = .Fields(2).Value
        txtDepartmentName.Text = .Fields(3).Value
        txtWardID.Text = .Fields(4).Value
        txtWardNumber.Text = .Fields(5).Value
        txtRoomCost.Text = .Fields(6).Value
        txtAdditionalNotes.Text = .Fields(7).Value
        
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
    
    
    With rsRoomsMaintenance
    
        .MoveNext   'Moving to the Next Record
        
        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If
        
        'Entering the values in the particular record into the fields on the interface
        txtRoomID.Text = .Fields(0).Value
        txtRoomNumber.Text = .Fields(1).Value
        txtDepartmentID.Text = .Fields(2).Value
        txtDepartmentName.Text = .Fields(3).Value
        txtWardID.Text = .Fields(4).Value
        txtWardNumber.Text = .Fields(5).Value
        txtRoomCost.Text = .Fields(6).Value
        txtAdditionalNotes.Text = .Fields(7).Value
        
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
    
    
    With rsRoomsMaintenance
    
        .MoveLast   'Moving to the last record
        
        'Entering the values in the particular record into the fields on the interface
        txtRoomID.Text = .Fields(0).Value
        txtRoomNumber.Text = .Fields(1).Value
        txtDepartmentID.Text = .Fields(2).Value
        txtDepartmentName.Text = .Fields(3).Value
        txtWardID.Text = .Fields(4).Value
        txtWardNumber.Text = .Fields(5).Value
        txtRoomCost.Text = .Fields(6).Value
        txtAdditionalNotes.Text = .Fields(7).Value
        
    End With
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub


Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsRoomsMaintenance
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[RoomID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[RoomNo] Like '" & txtSearch.Text & "%" & "'"
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
        
        Call Rooms_Maintenance
        
        Set dgrdRoomsInformation.DataSource = rsRoomsMaintenance 'Setting the Datasource for the DataGrid
        
    End If
    
End Sub


Private Sub cmdSave_Click()     'This function will save all the user's data in the database
        
        
    'Validation To Ensure That A Department Has Been Selected
    If txtDepartmentID.Text = "" Then
        MsgBox "Error! You Have To Select A Department!", vbCritical, "Error In Department ID!"
        txtDepartmentID.BackColor = &H80000018  'Highlighting the textfield in a different colour
        txtDepartmentName.BackColor = &H80000018  'Highlighting the textfield in a different colour
        Exit Sub
    Else
        txtDepartmentID.BackColor = &H80000004
        txtDepartmentName.BackColor = &H80000004
    End If

    
    'Validation To Ensure That A Ward Has Been Selected
    If txtWardID.Text = "" Then
        MsgBox "Error! You Have To Select A Ward!", vbCritical, "Error In Ward ID!"
        txtWardID.BackColor = &H80000018  'Highlighting the textfield in a different colour
        txtWardNumber.BackColor = &H80000018  'Highlighting the textfield in a different colour
        Exit Sub
    Else
        txtWardID.BackColor = &H80000004
        txtWardNumber.BackColor = &H80000004
    End If
    
    
    With rsRoomsMaintenance
            
        'Making sure that the user wants to save the record
        If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
                
            'The following if else condition ensures that The Additional Notes
            'textfield will not be completely blank when saving in the database.
            'This has been done in order to avoid errors.
            If txtAdditionalNotes.Text = "" Then
                txtAdditionalNotes.Text = "-"
            End If
                
                
            'Save the user-entered data into the recordset
            .Fields(0) = txtRoomID.Text
            .Fields(1) = txtRoomNumber.Text
            .Fields(2) = txtDepartmentID.Text
            .Fields(3) = txtDepartmentName.Text
            .Fields(4) = txtWardID.Text
            .Fields(5) = txtWardNumber.Text
            .Fields(6) = txtRoomCost.Text
            .Fields(7) = txtAdditionalNotes.Text
                
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
        

End Sub



Private Sub cmdDelete_Click()   'This function will delete a record from the database
    
    'Check for the record selection
    If txtRoomID.Text = "" Then
    
        MsgBox "Error! No Record Has Been Selected", vbCritical, "No Record Selected!"
    
    Else
    
        With rsRoomsMaintenance
        
            'Confirm the Delete procedure with the user
            If MsgBox("Are You Sure You Wish To Delete Room ID " & txtRoomID.Text & "'s Record?", vbYesNo + vbQuestion, "Delete Record?") = vbYes Then
        
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

