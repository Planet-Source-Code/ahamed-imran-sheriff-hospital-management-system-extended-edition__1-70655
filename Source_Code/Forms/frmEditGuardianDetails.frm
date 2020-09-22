VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmGuardiansMaintenance 
   Caption         =   "Guardians Maintenance Module"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditGuardianDetails.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   240
      Top             =   4560
   End
   Begin VB.PictureBox picInvalidDataMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3600
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   40
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Digits Here! Only Alphabets Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.PictureBox picInvalidTypingMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3600
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   38
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   39
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdClose 
      DisabledPicture =   "frmEditGuardianDetails.frx":1F95C
      Height          =   855
      Left            =   9000
      Picture         =   "frmEditGuardianDetails.frx":1FE1B
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7440
      Width           =   975
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
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3600
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
      Left            =   8880
      TabIndex        =   4
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdStep1 
      BackColor       =   &H80000013&
      Caption         =   "Step 1"
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
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2085
      Width           =   855
   End
   Begin VB.CommandButton cmdStep2 
      BackColor       =   &H80000013&
      Caption         =   "Step 2"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2085
      Width           =   855
   End
   Begin VB.CommandButton cmdStep3 
      BackColor       =   &H80000013&
      Caption         =   "Step 3"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2085
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      DisabledPicture =   "frmEditGuardianDetails.frx":22B5F
      Height          =   855
      Left            =   6840
      Picture         =   "frmEditGuardianDetails.frx":22FDD
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      DisabledPicture =   "frmEditGuardianDetails.frx":25D21
      Height          =   855
      Left            =   7920
      Picture         =   "frmEditGuardianDetails.frx":26207
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox txtRelationToPatient 
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
      MaxLength       =   15
      TabIndex        =   15
      Top             =   8400
      Width           =   2295
   End
   Begin VB.TextBox txtOccupation 
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
      MaxLength       =   30
      TabIndex        =   14
      Top             =   7920
      Width           =   2295
   End
   Begin VB.TextBox txtPhoneMob 
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
      MaxLength       =   15
      TabIndex        =   13
      Top             =   7440
      Width           =   2295
   End
   Begin VB.TextBox txtPhoneHome 
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
      MaxLength       =   15
      TabIndex        =   12
      Top             =   6960
      Width           =   2295
   End
   Begin VB.ComboBox cboGender 
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
      ItemData        =   "frmEditGuardianDetails.frx":28F4B
      Left            =   2880
      List            =   "frmEditGuardianDetails.frx":28F55
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox txtAddress 
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
      TabIndex        =   11
      Top             =   6000
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
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtNICNumber 
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
      MaxLength       =   10
      TabIndex        =   10
      Top             =   5520
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
      TabIndex        =   8
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtGuardianID 
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
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdLast 
      DisabledPicture =   "frmEditGuardianDetails.frx":28F67
      Height          =   750
      Left            =   9360
      Picture         =   "frmEditGuardianDetails.frx":29341
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      DisabledPicture =   "frmEditGuardianDetails.frx":2B4FD
      Height          =   750
      Left            =   8400
      Picture         =   "frmEditGuardianDetails.frx":2B8D3
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      DisabledPicture =   "frmEditGuardianDetails.frx":2DA8F
      Height          =   750
      Left            =   6480
      Picture         =   "frmEditGuardianDetails.frx":2DE6B
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdPrevious 
      DisabledPicture =   "frmEditGuardianDetails.frx":30027
      Height          =   750
      Left            =   7440
      Picture         =   "frmEditGuardianDetails.frx":3043C
      Style           =   1  'Graphical
      TabIndex        =   18
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
      ItemData        =   "frmEditGuardianDetails.frx":325F8
      Left            =   5160
      List            =   "frmEditGuardianDetails.frx":32605
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid dgrdGuardiansInfo 
      Height          =   2535
      Left            =   5520
      TabIndex        =   16
      Top             =   3120
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
      Caption         =   "Guardians Information Table"
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
      TabIndex        =   36
      Top             =   3645
      Width           =   1575
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
      Left            =   4320
      TabIndex        =   35
      Top             =   2175
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
      Left            =   7680
      TabIndex        =   34
      Top             =   2175
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   6000
      Top             =   7320
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   360
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Label lblRelationToPatient 
      BackStyle       =   0  'Transparent
      Caption         =   "Relation To Patient"
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
      TabIndex        =   33
      Top             =   8445
      Width           =   1815
   End
   Begin VB.Label lblOccupation 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
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
      Top             =   7965
      Width           =   1815
   End
   Begin VB.Label lblPhoneMob 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No. (Mob)"
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
      TabIndex        =   31
      Top             =   7485
      Width           =   1815
   End
   Begin VB.Label lblPhoneHome 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No. (Home)"
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
      Top             =   7005
      Width           =   1815
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Top             =   6045
      Width           =   1695
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
      Height          =   255
      Left            =   840
      TabIndex        =   28
      Top             =   4125
      Width           =   1575
   End
   Begin VB.Label lblGender 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Top             =   5085
      Width           =   1575
   End
   Begin VB.Label lblNICNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "NIC Number"
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
      Top             =   5565
      Width           =   1335
   End
   Begin VB.Label lblGuardianID 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian ID"
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
      Top             =   3165
      Width           =   1575
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
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   4605
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   11520
      Y1              =   8760
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   360
      Y1              =   2880
      Y2              =   8760
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   720
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblGuardianInformation 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian Information"
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
      TabIndex        =   23
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   3120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   6000
      Top             =   6120
      Width           =   4695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   735
      Left            =   3960
      Top             =   1920
      Width           =   7575
   End
End
Attribute VB_Name = "frmGuardiansMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Guardians Maintenance Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 26/04/08
'Date Of Last Modification: 26/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Guardians_Maintenance Table
'-------------------------------------------------------------------------------


Option Explicit

Dim eachField As Control  'Declaring a Control Variable for all Fields
Dim eachButton As Control 'Declaring a Control Variable fot all Command Buttons

'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean

'The following variables will be used to autogenerate the Admission ID to be
'displayed on the Admit Patient form on form load
Dim iNumOfInpatients As Integer  'This variable holds the number of records in the table
Dim strDisplayAdmissionID As String  'This variable will eventually hold the Admission ID to be autogenerated



Private Sub cmdSave_Click()     'This function will save all the user's data in the database
    
    
    'Checking if the Phone Number (Home) textfield and the Phone Number (Mob) textfield are empty
    If txtPhoneHome.Text = "-" And txtPhoneMob.Text = "-" Then
        txtPhoneHome.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtPhoneMob.BackColor = &H80000018 'Highlighting the textfield in a different colour
        MsgBox "Error! Both Phone Number Textfields Cannot Be Empty! At Least One Has To Be Provided!", vbCritical, "Error In Phone Numbers!"
        Exit Sub
    Else
        txtPhoneHome.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtPhoneMob.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
        
        
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        
        'Validation To Ensure That The NIC Number is 10 Characters In Length
        If txtNICNumber.Text <> "-" Then
            If Len(txtNICNumber.Text) <> 10 Then
                MsgBox "Error! The NIC Number Has To Consist Of 10 Characters!", vbCritical, "Error In NIC Number!"
                txtNICNumber.BackColor = &H80000018  'Highlighting the textfield in a different colour
                Exit Sub
            Else
                txtNICNumber.BackColor = &H80000004
            End If
        End If
        
        
        
        With rsGuardiansMaintenance
            
            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
                
                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtNICNumber.Text = "" Then
                    txtNICNumber.Text = "-"
                End If
                
                If txtPhoneMob.Text = "" Then
                    txtPhoneMob.Text = "-"
                End If
                
                If txtPhoneHome.Text = "" Then
                    txtPhoneHome.Text = "-"
                End If
                
                strGuardainID = txtGuardianID.Text
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtGuardianID.Text
                .Fields(1) = txtPatientID.Text
                .Fields(2) = txtFirstName.Text
                .Fields(3) = txtSurname.Text
                .Fields(4) = cboGender.Text
                .Fields(5) = txtNICNumber.Text
                .Fields(6) = txtAddress.Text
                .Fields(7) = txtPhoneHome.Text
                .Fields(8) = txtPhoneMob.Text
                .Fields(9) = txtOccupation.Text
                .Fields(10) = txtRelationToPatient.Text
                
            
                .Update
                
                'Display Success Message
                MsgBox "The Record Was Saved Successfully! You Will Now Be Taken To Step 3!", vbInformation, "Succesful Save Procedure!"
                
                
                loadPatientAdmission
                
                Unload Me
                
                frmAdmitPatient.Show
            
            Else
            
                'Display 'No Modifications' Message
                MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
                .CancelUpdate   'Cancel the Save Procedure
            
            End If
            
            .Requery    'Requerying the Table
            
        End With
        
    End If
        

End Sub

Private Function loadPatientAdmission()
    
    frmAdmitPatient.enableAllFields    'Calling a Private Function To Enable All Fields
    frmAdmitPatient.clearAllFields      'Calling a Private Function To Clear All Fields
    frmAdmitPatient.disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    frmAdmitPatient.txtReferredDoctorID.Text = "-" 'Since this textfield is not compulsory
    frmAdmitPatient.txtReferredDoctorName.Text = "-"  'Since this textfield is not always compulsory
    frmAdmitPatient.txtAdditionalNotes.Text = "-" 'Since this textfield is not always compulsory
    
    
    'Enabling the Save Command Button
    frmAdmitPatient.cmdSave.Enabled = True
    
    'Disabling the "Launch Inpatient Search Wizard" Button
    frmAdmitPatient.cmdLaunchInpatientSearch.Enabled = False
    
    'Enabling the Wizard Buttons
    frmAdmitPatient.cmdReferredDoctorIDWizardButton.Enabled = True
    frmAdmitPatient.cmdAssignedDoctorWizardButton.Enabled = True
    frmAdmitPatient.cmdDepartmentIDWizardButton.Enabled = True
    frmAdmitPatient.cmdWardIDWizardButton.Enabled = True
    frmAdmitPatient.cmdRoomIDWizardButton.Enabled = True
   

    
    Call Inpatients_Admission    'Calling the Inpatients_Admission Procedure to interact with the recordset
    
    'Generate Admission ID By Utilizing the Inpatients_Admission Table
    With rsInpatientsAdmission
    
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strDisplayAdmissionID = "ADM0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfInpatients = .RecordCount
            iNumOfInpatients = iNumOfInpatients + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Inpatients_Admission Table
            If iNumOfInpatients < 10 Then
                strDisplayAdmissionID = "ADM000" & iNumOfInpatients
            ElseIf iNumOfInpatients < 100 Then
                strDisplayAdmissionID = "ADM00" & iNumOfInpatients
            ElseIf iNumOfInpatients < 1000 Then
                strDisplayAdmissionID = "ADM0" & iNumOfInpatients
            ElseIf iNumOfInpatients < 10000 Then
                strDisplayAdmissionID = "ADM" & iNumOfInpatients
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    'The following line of code will enter the autogenerated Admission ID
    'into the Admission ID textfield
    frmAdmitPatient.txtAdmissionID.Text = strDisplayAdmissionID
    
    frmAdmitPatient.txtPatientID.Text = strPatientID    'Global Variable
    
    frmAdmitPatient.txtGuardianID = strGuardainID   'Global Variable
    
    frmAdmitPatient.txtAdmissionDate = DateTime.Date
    
    frmAdmitPatient.txtAdmissionTime = DateTime.Time
    
End Function

Private Sub cmdStep3_Click()

    Call Inpatients_Admission
    
    With rsInpatientsAdmission
    
        .MoveFirst
        
        Do While .EOF = False
            
            If .Fields(1).Value = txtPatientID.Text Then
            
                'Entering the values in the particular record into the fields on the interface
                frmAdmitPatient.txtAdmissionID = .Fields(0).Value
                frmAdmitPatient.txtPatientID.Text = .Fields(1).Value
                frmAdmitPatient.txtGuardianID.Text = .Fields(2).Value
                frmAdmitPatient.txtAdmissionDate.Text = .Fields(3).Value
                frmAdmitPatient.txtAdmissionTime.Text = .Fields(4).Value
                frmAdmitPatient.cboPatientStatus.Text = .Fields(5).Value
                frmAdmitPatient.txtReasonForStatus.Text = .Fields(6).Value
                frmAdmitPatient.txtReferredDoctorID.Text = .Fields(7).Value
                frmAdmitPatient.txtReferredDoctorName.Text = .Fields(8).Value
                frmAdmitPatient.txtAssignedDoctorID.Text = .Fields(9).Value
                frmAdmitPatient.txtAssignedDoctorName.Text = .Fields(10).Value
                frmAdmitPatient.txtDepartmentID.Text = .Fields(11).Value
                frmAdmitPatient.txtDepartmentName.Text = .Fields(12).Value
                frmAdmitPatient.txtWardID.Text = .Fields(13).Value
                frmAdmitPatient.txtWardNo.Text = .Fields(14).Value
                frmAdmitPatient.txtRoomID.Text = .Fields(15).Value
                frmAdmitPatient.txtAdditionalNotes.Text = .Fields(16).Value
                Exit Do
            
            Else
            
                .MoveNext
                
            End If
            
        Loop
        
    End With
    
    
    'Enabling / Diabling the Navigation Buttons as necessary
    frmAdmitPatient.cmdFirst.Enabled = False
    frmAdmitPatient.cmdLast.Enabled = True
    frmAdmitPatient.cmdPrevious.Enabled = False
    frmAdmitPatient.cmdNext.Enabled = True

    'Enabling the Update Button
    frmAdmitPatient.cmdUpdate.Enabled = True
    
    'Enabling the Wizard Buttons
    frmAdmitPatient.cmdReferredDoctorIDWizardButton.Enabled = True
    frmAdmitPatient.cmdAssignedDoctorWizardButton.Enabled = True
    frmAdmitPatient.cmdDepartmentIDWizardButton.Enabled = True
    frmAdmitPatient.cmdWardIDWizardButton.Enabled = True
    frmAdmitPatient.cmdRoomIDWizardButton.Enabled = True


    'Enabling the "Step" Buttons
    frmAdmitPatient.cmdStep1.Enabled = True
    frmAdmitPatient.cmdStep2.Enabled = True
    
    'frmAdmitPatient.enableAllFields
    frmAdmitPatient.enableAllFields
    
    Unload Me
    
    frmAdmitPatient.Show

End Sub

Private Sub cmdUpdate_Click()   'This function will update a record after the user has edited it


    'Checking if the Phone Number (Home) textfield and the Phone Number (Mob) textfield are empty
    If txtPhoneHome.Text = "-" And txtPhoneMob.Text = "-" Then
        txtPhoneHome.BackColor = &H80000018 'Highlighting the textfield in a different colour
        txtPhoneMob.BackColor = &H80000018 'Highlighting the textfield in a different colour
        MsgBox "Error! Both Phone Number Textfields Cannot Be Empty! At Least One Has To Be Provided!", vbCritical, "Error In Phone Numbers!"
        Exit Sub
    Else
        txtPhoneHome.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
        txtPhoneMob.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
        
        
        
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
        
        
        'Validation To Ensure That The NIC Number is 10 Characters In Length
        If txtNICNumber.Text <> "-" Then
            If Len(txtNICNumber.Text) <> 10 Then
                MsgBox "Error! The NIC Number Has To Consist Of 10 Characters!", vbCritical, "Error In NIC Number!"
                txtNICNumber.BackColor = &H80000018  'Highlighting the textfield in a different colour
                Exit Sub
            Else
                txtNICNumber.BackColor = &H80000004
            End If
        End If
        
        
        
        With rsGuardiansMaintenance
            
            'Making sure that the user wants to update the record
            If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
                
                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtNICNumber.Text = "" Then
                    txtNICNumber.Text = "-"
                End If
                
                If txtPhoneMob.Text = "" Then
                    txtPhoneMob.Text = "-"
                End If
                
                If txtPhoneHome.Text = "" Then
                    txtPhoneHome.Text = "-"
                End If
                
                
                
                'Save the user-entered data into the recordset
                .Fields(0) = txtGuardianID.Text
                .Fields(1) = txtPatientID.Text
                .Fields(2) = txtFirstName.Text
                .Fields(3) = txtSurname.Text
                .Fields(4) = cboGender.Text
                .Fields(5) = txtNICNumber.Text
                .Fields(6) = txtAddress.Text
                .Fields(7) = txtPhoneHome.Text
                .Fields(8) = txtPhoneMob.Text
                .Fields(9) = txtOccupation.Text
                .Fields(10) = txtRelationToPatient.Text
            
                .Update
                
                'Display Success Message
                MsgBox "The Record Was Updated Successfully!", vbInformation, "Succesful Update Procedure"
                
                
                Form_Load   'Calling the Form_Load Procedure
                
                clearAllFields  'Calling a Private Function To Clear All Fields
            
            Else
            
                'Display 'No Modifications' Message
                MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
                .CancelUpdate   'Cancel the Update Procedure
                
            
            End If
            
            .Requery    'Requerying the Table
            
        End With
        
    End If
    
End Sub

Private Sub dgrdGuardiansInfo_Click()
    
    'Enabling the Update Button
    cmdUpdate.Enabled = True

    
    'Enabling the Navigation Buttons
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep3.Enabled = True
    
    
    With rsGuardiansMaintenance
    
        'Entering the values in the particular record into the fields on the interface
        txtGuardianID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        cboGender.Text = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtOccupation.Text = .Fields(9).Value
        txtRelationToPatient.Text = .Fields(10).Value
        
    End With
    
    enableAllFields 'Calling a Private Function To Enable All Fields
    
End Sub

Private Sub Form_Load()

    Call Connection  'Calling the Connection Procedure
    
    disableAllFields  'Calling a Private Function To Disable All Fields
    disableAllButtons   'Calling a Private Function To Disable All Command Buttons
    
    'Enabling  the First Button and the Last Button
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    
    'Enabling the Close button
    cmdClose.Enabled = True
    
    'Enabling the Search Frame
    cboSearchType.Enabled = True
    txtSearch.Enabled = True

    dgrdGuardiansInfo.Enabled = True
    
    
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
    
    'Disabling the Step 2 Button
    cmdStep2.Enabled = False

End Function


Public Function clearAllFields() 'This function will clear all fields on the interface


    On Error Resume Next
    For Each eachField In Me.Controls  'Running a Loop through all the Controls

    'The following If Condition will clear all TextBoxes
    If TypeOf eachField Is TextBox Then
        eachField.Text = ""
    End If

    Next
    
    'The following lines will set the normal display values of the Gender
    'ComboBox.
    cboGender.Text = "----------SELECT-----------"
    
    
End Function


Private Sub cmdFirst_Click()  'This function will Navigate to the First Record

    'Enabling / Diabling the Navigation Buttons as necessary
    cmdFirst.Enabled = False
    cmdLast.Enabled = True
    cmdPrevious.Enabled = False
    cmdNext.Enabled = True

    'Enabling the Update Button
    cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep3.Enabled = True


    Call Guardians_Maintenance  'Calling the Guardians_Maintenance Procedure to interact with the recordset

    With rsGuardiansMaintenance


        .MoveFirst  'Moving to the first record

        'Entering the values in the particular record into the fields on the interface
        txtGuardianID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        cboGender.Text = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtOccupation.Text = .Fields(9).Value
        txtRelationToPatient.Text = .Fields(10).Value

    End With

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdPrevious_Click() 'This function will Navigate to the Previous Record

    With rsGuardiansMaintenance


        .MovePrevious   'Moving to the previous record

        'If the user reaches the first record, display a message box
        'to inform the user of this
        If .BOF Then
            MsgBox "This is the first record!", vbInformation, "First Record"
            .MoveFirst
        End If

        'Entering the values in the particular record into the fields on the interface
        txtGuardianID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        cboGender.Text = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtOccupation.Text = .Fields(9).Value
        txtRelationToPatient.Text = .Fields(10).Value

    End With

    cmdNext.Enabled = True  'Enabling the Next Button
    cmdLast.Enabled = True  'Enabling the Last Button

    'Enabling the Update Button
    cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep3.Enabled = True


    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdNext_Click() 'This function will Navigate to the Next Record

    With rsGuardiansMaintenance

        .MoveNext   'Moving to the Next Record

        'If the user reaches the last record, display a message box
        'to inform the user of this
        If .EOF Then
            MsgBox "This is the last record!", vbInformation, "Last Record"
            .MoveLast
        End If

        'Entering the values in the particular record into the fields on the interface
        txtGuardianID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        cboGender.Text = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtOccupation.Text = .Fields(9).Value
        txtRelationToPatient.Text = .Fields(10).Value

    End With

    cmdPrevious.Enabled = True  'Enabling the Previous Button
    cmdFirst.Enabled = True 'Enabling the First Button

    'Enabling the Update Button
    cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep3.Enabled = True



    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub


Private Sub cmdLast_Click() 'This function will Navigate to the Last Record

    'Enabling / Diabling the Navigation Buttons as necessary
    cmdLast.Enabled = False
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = False

    'Enabling the Update Button
    cmdUpdate.Enabled = True


    'Enabling the "Step" Buttons
    cmdStep1.Enabled = True
    cmdStep3.Enabled = True


    Call Guardians_Maintenance  'Calling the Guardians_Maintenance Procedure to interact with the recordset

    With rsGuardiansMaintenance

        .Requery

        .MoveLast   'Moving to the last record

        'Entering the values in the particular record into the fields on the interface
        txtGuardianID.Text = .Fields(0).Value
        txtPatientID.Text = .Fields(1).Value
        txtFirstName.Text = .Fields(2).Value
        txtSurname.Text = .Fields(3).Value
        cboGender.Text = .Fields(4).Value
        txtNICNumber.Text = .Fields(5).Value
        txtAddress.Text = .Fields(6).Value
        txtPhoneHome.Text = .Fields(7).Value
        txtPhoneMob.Text = .Fields(8).Value
        txtOccupation.Text = .Fields(9).Value
        txtRelationToPatient.Text = .Fields(10).Value

    End With

    enableAllFields 'Calling a Private Function To Enable All Fields

End Sub



Private Function textfieldsValidations() As Boolean  'This function will validate all fields
    
    Flag = True 'Setting the Flag variable to True

    
    'Checking if the First Name textfield is empty
    If txtFirstName.Text = "" Then
        txtFirstName.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtFirstName.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Surname textfield is empty
    If txtSurname.Text = "" Then
        txtSurname.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtSurname.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the user has made a selection in the Gender ComboBox
    If cboGender.Text = "" Then
        cboGender.BackColor = &H80000018    'Highlighting the ComboBox in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        cboGender.BackColor = &H80000004    'Bringing the ComboBox BackColour back to normal
    End If
    
    
    'Checking if the Address textfield is empty
    If txtAddress.Text = "" Then
        txtAddress.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtAddress.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Patient Occupation textfield is empty
    If txtOccupation.Text = "" Then
        txtOccupation.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtOccupation.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Relation To Patient textfield is empty
    If txtRelationToPatient.Text = "" Then
        txtRelationToPatient.BackColor = &H80000018 'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtRelationToPatient.BackColor = &H80000004 'Bringing the textfield BackColour back to normal
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



Private Sub tmrErrMsg_Timer()

    Static i As Integer

    If i < 200000 Then     'Validation Msg Viewing Time Period
        picInvalidDataMsg.Visible = False
        picInvalidTypingMsg.Visible = False
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If

End Sub


Private Sub txtPhoneHome_GotFocus() 'This procedure will ensure that the textfield is empty when the user types in it.
    
    If txtPhoneHome.Text = "-" Then
        txtPhoneHome.Text = ""
    End If
    
End Sub

Private Sub txtPhoneHome_LostFocus()    'This procedure will ensure that the textfield is not empty when the user is not typing in it.
    
    If txtPhoneHome.Text = "" Then
        txtPhoneHome.Text = "-"
    End If
    
End Sub


Private Sub txtPhoneMob_GotFocus()  'This procedure will ensure that the textfield is empty when the user types in it.
    
    If txtPhoneMob.Text = "-" Then
        txtPhoneMob.Text = ""
    End If
    
End Sub

Private Sub txtPhoneMob_LostFocus() 'This procedure will ensure that the textfield is not empty when the user is not typing in it.
    
    If txtPhoneMob.Text = "" Then
        txtPhoneMob.Text = "-"
    End If
    
End Sub

'This procedure will ensure that the textfield is empty when the user types in it.
Private Sub txtNICNumber_GotFocus()
    
    If txtNICNumber.Text = "-" Then
        txtNICNumber.Text = ""
    End If
    
End Sub

Private Sub txtNICNumber_LostFocus()    'This procedure will ensure that the textfield is not empty when the user is not typing in it.

    If txtNICNumber.Text = "" Then
        txtNICNumber.Text = "-"
    End If
    
End Sub


Private Sub txtNICNumber_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = Asc("X") Then
    ElseIf KeyAscii = Asc("x") Then
    ElseIf KeyAscii = Asc("V") Then
    ElseIf KeyAscii = Asc("v") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidTypingMsg.Top = 5520    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub
    
    
Private Sub txtPhoneHome_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only digits

    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidTypingMsg.Top = 6960    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub


Private Sub txtPhoneMob_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only digits

    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidTypingMsg.Top = 7440    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub


Private Sub txtFirstName_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets

    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 4080    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub


Private Sub txtSurname_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets

    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 4560    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub


Private Sub txtOccupation_KeyPress(KeyAscii As Integer)

    'Keypress Validation to allow only alphabets

    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidDataMsg.Top = 7920    'Validation Note View
        picInvalidDataMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If

End Sub

Private Sub txtSearch_Change()

    If Len(txtSearch.Text) > 0 Then 'Checking if the user has typed in the textfield
    
        With rsGuardiansMaintenance
        
            'Filter the Records As The User Types, According to the Criteria
            Select Case (cboSearchType.ListIndex)
                Case 0:
                    .Filter = "[GuardianID] Like '" & txtSearch.Text & "%" & "'"
                Case 1:
                    .Filter = "[FirstName] Like '" & txtSearch.Text & "%" & "'"
                Case 2:
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
        
        Call Guardians_Maintenance
        
        Set dgrdGuardiansInfo.DataSource = rsGuardiansMaintenance
        
    End If
    
    
End Sub

Private Sub cmdStep1_Click()

    Call Inpatients_Maintenance
    
    With rsInpatientMaintenance
    
        .MoveFirst
        
        Do While .EOF = False
            
            If .Fields(0).Value = txtPatientID.Text Then
                
                'Entering the values in the particular record into the fields on the interface
                frmInpatientsMaintenance.txtPatientID.Text = .Fields(0).Value
                frmInpatientsMaintenance.txtFirstName.Text = .Fields(1).Value
                frmInpatientsMaintenance.txtSurname.Text = .Fields(2).Value
                frmInpatientsMaintenance.cboGender.Text = .Fields(3).Value
                frmInpatientsMaintenance.dtpDateOfBirth.Value = .Fields(4).Value
                frmInpatientsMaintenance.txtNICNumber.Text = .Fields(5).Value
                frmInpatientsMaintenance.txtAddress.Text = .Fields(6).Value
                frmInpatientsMaintenance.txtPhoneHome.Text = .Fields(7).Value
                frmInpatientsMaintenance.txtPhoneMob.Text = .Fields(8).Value
                frmInpatientsMaintenance.txtPatientOccupation.Text = .Fields(9).Value
                frmInpatientsMaintenance.cboCivilStatus.Text = .Fields(10).Value
                frmInpatientsMaintenance.cboAccountType.Text = .Fields(11).Value
                frmInpatientsMaintenance.txtCompanyID.Text = .Fields(12).Value
                frmInpatientsMaintenance.txtCompanyName.Text = .Fields(13).Value
                Exit Do
                
            Else
            
                .MoveNext
                
            End If
            
        Loop
        
    End With
    
    
    'Enabling / Diabling the Navigation Buttons as necessary
    frmInpatientsMaintenance.cmdFirst.Enabled = False
    frmInpatientsMaintenance.cmdLast.Enabled = True
    frmInpatientsMaintenance.cmdPrevious.Enabled = False
    frmInpatientsMaintenance.cmdNext.Enabled = True

    'Enabling the Update Button and the Delete Button
    frmInpatientsMaintenance.cmdUpdate.Enabled = True
    frmInpatientsMaintenance.cmdDelete.Enabled = True
    
    'Enabling the Wizard Buttons
    frmInpatientsMaintenance.cmdCompanySearchWizard.Enabled = True
    
    'Enabling the "Step" Buttons
    frmInpatientsMaintenance.cmdStep2.Enabled = True
    
    frmInpatientsMaintenance.enableAllFields
    
    Unload Me
    
    frmInpatientsMaintenance.Show
    
    
    
End Sub

Private Sub cmdClose_Click()

    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If

End Sub
