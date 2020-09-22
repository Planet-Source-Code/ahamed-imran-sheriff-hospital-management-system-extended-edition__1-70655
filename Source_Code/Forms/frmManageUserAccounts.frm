VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmManageUserAccounts 
   Caption         =   "Manage User Accounts"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmManageUserAccounts.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInvalidDataMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3240
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   33
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Digits Here! Only Alphabets Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   120
      Top             =   4800
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
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtUserID 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
   Begin VB.ComboBox cboDesignation 
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
      ItemData        =   "frmManageUserAccounts.frx":1DFDF
      Left            =   2400
      List            =   "frmManageUserAccounts.frx":1DFF2
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmManageUserAccounts.frx":1E03F
      Height          =   750
      Left            =   7680
      Picture         =   "frmManageUserAccounts.frx":1E454
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmManageUserAccounts.frx":20610
      Height          =   750
      Left            =   6720
      Picture         =   "frmManageUserAccounts.frx":209EC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmManageUserAccounts.frx":22BA8
      Height          =   750
      Left            =   8640
      Picture         =   "frmManageUserAccounts.frx":22F7E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmManageUserAccounts.frx":2513A
      Height          =   750
      Left            =   9600
      Picture         =   "frmManageUserAccounts.frx":25514
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   890
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmManageUserAccounts.frx":276D0
      Height          =   855
      Left            =   10080
      Picture         =   "frmManageUserAccounts.frx":27B8F
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmManageUserAccounts.frx":2A8D3
      Height          =   855
      Left            =   9000
      Picture         =   "frmManageUserAccounts.frx":2AD9C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      DisabledPicture =   "frmManageUserAccounts.frx":2DAE0
      Height          =   855
      Left            =   5760
      Picture         =   "frmManageUserAccounts.frx":2DEE2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmManageUserAccounts.frx":30C26
      Height          =   855
      Left            =   6840
      Picture         =   "frmManageUserAccounts.frx":310A4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmManageUserAccounts.frx":33DE8
      Height          =   855
      Left            =   7920
      Picture         =   "frmManageUserAccounts.frx":342CE
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7680
      Width           =   975
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
      ItemData        =   "frmManageUserAccounts.frx":37012
      Left            =   3240
      List            =   "frmManageUserAccounts.frx":37022
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2160
      Width           =   2415
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
      Left            =   2400
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox txtLastName 
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
      Left            =   2400
      TabIndex        =   4
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtConfirmation 
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   8040
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   7080
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid dgrdUserAccount 
      Height          =   2775
      Left            =   5520
      TabIndex        =   10
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4895
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
      Caption         =   "User Account Information Table"
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
      Left            =   480
      TabIndex        =   32
      Top             =   2880
      Width           =   7455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   735
      Left            =   2160
      Top             =   1920
      Width           =   7575
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
      TabIndex        =   31
      Top             =   2190
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
      TabIndex        =   30
      Top             =   2190
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   6360
      Top             =   6360
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   5520
      Top             =   7560
      Width           =   5775
   End
   Begin VB.Label lblConfirmation 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      TabIndex        =   29
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
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
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
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
      TabIndex        =   27
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000001&
      X1              =   3480
      X2              =   5160
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Account Information"
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
      TabIndex        =   26
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   720
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000001&
      X1              =   5160
      X2              =   5160
      Y1              =   6720
      Y2              =   8640
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   5160
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   480
      Y1              =   6720
      Y2              =   8640
   End
   Begin VB.Label lblDesignation 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation :"
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
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   5160
      X2              =   3480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff / User Information"
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
      Left            =   960
      TabIndex        =   24
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   840
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   5160
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   480
      Y1              =   3360
      Y2              =   6240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   5160
      X2              =   5160
      Y1              =   6240
      Y2              =   3360
   End
   Begin VB.Label lblUser_ID 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff/User ID :"
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
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblFirstName 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name :"
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
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblLastName 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name :"
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
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "* E-mail :"
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
      TabIndex        =   20
      Top             =   5160
      Width           =   1335
   End
End
Attribute VB_Name = "frmManageUserAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Manage User Accounts Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 22/04/08
'Date Of Last Modification: 22/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: UserAccount Table
'-----------------------------------------------------------------------------

Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean

'This variable will count the number of times the user keys in the "@" symbol.
Dim iNumOfSymbols As Integer

'The following variables will be used to autogenerate the User ID
Dim iNumOfRecords As Integer    'This variable holds the number of records in the table
Dim strCode As String   'This variable will eventually hold the User ID to be autogenerated

Dim rsSelectionOfFields As ADODB.Recordset    'This will limit the fields I show in the grid


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
    dgrdUserAccount.Enabled = False
    
    Call UserAccounts_Maintenance    'Calling the UserAccounts_Maintenance Procedure to interact with the recordset
    
    'Generate User ID By Utilizing the UserAccounts_Maintenance Table
    With rsUserAccount
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strCode = "EMP0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Doctors_Maintenance Table
            If iNumOfRecords < 10 Then
                strCode = "EMP000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strCode = "EMP00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strCode = "EMP0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strCode = "EMP" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    'The following line of code will enter the autogenerated User ID
    'into the User ID textfield & Username textfield.
    txtUserID.Text = strCode
    txtUsername.Text = strCode
    
End Sub



Private Sub cmdClose_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it.
    
    If txtPassword.Text <> txtConfirmation.Text Then
        MsgBox "Error! The Passwords You Provided Do Not Match", vbCritical, "Password Mismatch!"
        Exit Sub
    End If
    
    Flag = True 'Setting the Flag variable to True
    
    'Checking if the First Name textfield is empty
    If txtFirstName.Text = "" Then
        txtFirstName.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtFirstName.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Last Name textfield is empty
    If txtLastName.Text = "" Then
        txtLastName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtLastName.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the user has made a selection in the Designation ComboBox
    If cboDesignation.Text = "" Then
        cboDesignation.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboDesignation.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If

    
    
    'Here, I am checking the state of the Flag variable and if it is False, I am displaying a
    'Message Box to instruct the user to enter data into all highlighted textfields.
    'The Update procedure will also be cancelled
    If Flag = False Then
        MsgBox "Error! Please Fill-in The Highlighted Textfields! They Are Compulsory!", vbCritical, "Please Fill Highlighted Textfields"
        Exit Sub
    End If
    
    
    With rsSelectionOfFields
        
        'Making sure that the user wants to update the record
        If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
            
            'The following if else condition ensures that The Additional Notes
            'textfield will not be completely blank when saving in the database.
            'This has been done in order to avoid errors.
            If txtEmail.Text = "" Then
                txtEmail.Text = "-"
            End If
                    
                    
            'Save the user-entered data into the recordset
            .Fields(0) = txtUserID.Text
            .Fields(1) = txtFirstName.Text
            .Fields(2) = txtLastName.Text
            .Fields(3) = txtEmail.Text
            .Fields(4) = cboDesignation.Text
                
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

Private Sub dgrdUserAccount_Click()
    
    'Enabling the Update Button & the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    'Enabling the Navigation Buttons
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    With rsSelectionOfFields
    
        'Entering the values in the particular record into the fields on the interface
        txtUserID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtLastName.Text = .Fields(2).Value
        txtEmail.Text = .Fields(3).Value
        cboDesignation.Text = .Fields(4).Value
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtConfirmation.Text = ""
        
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub



Private Sub Form_Load()
    
    Call Connection  'Calling the Connection Procedure
    
    Call UserAccounts_Maintenance  'Calling the UserAccounts_Maintenance Procedure to interact with the recordset
    
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
    dgrdUserAccount.Enabled = True
    
    Call Selection_Of_Fields
    
    Set dgrdUserAccount.DataSource = rsSelectionOfFields  'Setting the DataSource of the DataGrid

End Sub

'This function will interact with the UserAccount table, whilst hiding
'sensitive information from the user
Private Function Selection_Of_Fields()
    
    Set rsSelectionOfFields = New ADODB.Recordset
    
    With rsSelectionOfFields
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select [UserID],[FirstName],[LastName],[EMail],[Designation] from UserAccount"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Function


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
    
    
    With rsSelectionOfFields
    
        .MoveFirst  'Moving to the first record
        
        'Entering the values in the particular record into the fields on the interface
        txtUserID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtLastName.Text = .Fields(2).Value
        txtEmail.Text = .Fields(3).Value
        cboDesignation.Text = .Fields(4).Value
        
    
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
    
    
    With rsSelectionOfFields
    
        .MovePrevious   'Moving to the previous record
        
        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If
    
        'Entering the values in the particular record into the fields on the interface
        txtUserID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtLastName.Text = .Fields(2).Value
        txtEmail.Text = .Fields(3).Value
        cboDesignation.Text = .Fields(4).Value
        
        
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
    
    
    With rsSelectionOfFields
    
        .MoveNext   'Moving to the Next Record
        
        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If
        
        'Entering the values in the particular record into the fields on the interface
        txtUserID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtLastName.Text = .Fields(2).Value
        txtEmail.Text = .Fields(3).Value
        cboDesignation.Text = .Fields(4).Value
        
        
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
    
    
    With rsSelectionOfFields
    
        .MoveLast   'Moving to the last record
        
        'Entering the values in the particular record into the fields on the interface
        txtUserID.Text = .Fields(0).Value
        txtFirstName.Text = .Fields(1).Value
        txtLastName.Text = .Fields(2).Value
        txtEmail.Text = .Fields(3).Value
        cboDesignation.Text = .Fields(4).Value
        
        
    End With
    
    'Enabling the Update Button and the Delete Button
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub




Private Sub tmrErrMsg_Timer()

    Static i As Integer

    If i < 200000 Then     'Validation Msg Viewing Time Period
        picInvalidDataMsg.Visible = False
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If

End Sub

Private Sub txtSearch_Change()  'This is executed when the user types in the Search textfield
    
    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsSelectionOfFields
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[UserID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[FirstName] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
                    .Filter = "[LastName] Like '" & txtSearch.Text & "%" & "'"
                Case 3:
                    .Filter = "[Designation] Like '" & txtSearch.Text & "%" & "'"
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
        
        
        Call Selection_Of_Fields
        
        Set dgrdUserAccount.DataSource = rsSelectionOfFields  'Setting the DataSource of the DataGrid
        

    End If
    
End Sub


Private Sub cmdSave_Click()     'This function will save all the user's data in the database
    
    'Checking if the Passwords Match
    If txtPassword.Text <> txtConfirmation.Text Then
        MsgBox "Error! The Passwords You Provided Do Not Match", vbCritical, "Password Mismatch!"
        Exit Sub
    End If
    
    'Checking the number of times the "@" symbol was pressed
    If iNumOfSymbols < 1 Then
        MsgBox "Error! The Email ID You Provided Does Not Contain The @ Symbol!", vbCritical, "No @ Symbol!"
        txtEmail.Text = ""  'Clearing the textfield
        iNumOfSymbols = 0   'Setting the value of the variable to 0
        Exit Sub
    ElseIf iNumOfSymbols > 1 Then
        MsgBox "Error! The Email ID You Provided Can Contain Only One @ Symbol!", vbCritical, "Too Many @ Symbols!"
        txtEmail.Text = ""  'Clearing the texfield
        iNumOfSymbols = 0   'Setting the value of the variable to 0
        Exit Sub
    End If
    
    
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        
        With rsUserAccount
            
            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
                
                'The following if else condition ensures that The Additional Notes
                'textfield will not be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtEmail.Text = "" Then
                    txtEmail.Text = "-"
                End If
                
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtUserID.Text
                .Fields(1) = txtFirstName.Text
                .Fields(2) = txtLastName.Text
                .Fields(3) = txtEmail.Text
                .Fields(4) = cboDesignation.Text
                .Fields(5) = txtUsername.Text
                .Fields(6) = txtPassword.Text
                
                .Update
                
                .Requery    'Requerying the Table
                
                'Display Success Message
                MsgBox "The Record Was Saved Successfully!", vbInformation, "Succesful Save Procedure"
                
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            Else
            
                'Display 'No Modifications' Message
                MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            End If
            
            '.Requery    'Requerying the Table
            
        End With
        
    End If
        

End Sub


Private Function textfieldsValidations() As Boolean  'This function will validate all fields
    
    Flag = True 'Setting the Flag variable to True
    
    'Checking if the First Name textfield is empty
    If txtFirstName.Text = "" Then
        txtFirstName.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtFirstName.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Last Name textfield is empty
    If txtLastName.Text = "" Then
        txtLastName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtLastName.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the User Name textfield is empty
    If txtUsername.Text = "" Then
        txtUsername.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtUsername.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Password textfield is empty
    If txtPassword.Text = "" Then
        txtPassword.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtPassword.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Confirm Password textfield is empty
    If txtConfirmation.Text = "" Then
        txtConfirmation.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtConfirmation.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the user has made a selection in the Designation ComboBox
    If cboDesignation.Text = "" Then
        cboDesignation.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboDesignation.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
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
    If txtUserID.Text = "" Then
    
        MsgBox "Error! No Record Has Been Selected", vbCritical, "No Record Selected!"
        
    Else
    
        With rsSelectionOfFields
        
            'Confirm the Delete procedure with the user
            If MsgBox("Are You Sure You Wish To Delete User ID " & txtUserID.Text & "'s Record?", vbYesNo + vbQuestion, "Delete Record?") = vbYes Then
        
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



Private Sub txtFirstName_KeyPress(KeyAscii As Integer)

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



Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only alphabets
    
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 4680    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    
    'Counting the number of times the "@" symbol was pressed
    
    If KeyAscii = Asc("@") Then
        iNumOfSymbols = iNumOfSymbols + 1
    End If
    
End Sub

