VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   11520
   ClientLeft      =   -120
   ClientTop       =   -135
   ClientWidth     =   15360
   ForeColor       =   &H80000018&
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Height          =   496
      Left            =   9960
      Picture         =   "frmLogin.frx":450D3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   511
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   10425
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6780
      Width           =   2445
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   10425
      TabIndex        =   0
      Top             =   6270
      Width           =   2445
   End
   Begin VB.CommandButton cmdGo 
      DisabledPicture =   "frmLogin.frx":45D17
      Enabled         =   0   'False
      Height          =   496
      Left            =   13080
      Picture         =   "frmLogin.frx":45FB7
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click To Access System"
      Top             =   6720
      Width           =   511
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Off System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   10560
      TabIndex        =   6
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   9000
      TabIndex        =   5
      Top             =   6795
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   9000
      TabIndex        =   4
      Top             =   6285
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Login Interface
'Programmer: Isham Sally
'Quality Assurance Engineer (Testing): Imran Sheriff
'Start Date: 01/04/08
'Date Of Last Modification: 01/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: User_Account Table
'---------------------------------------------------------------------

Option Explicit
Dim rsLogin As ADODB.Recordset 'Creating a Recordset Variable
Dim iLoginFailure As Integer    'This variable will count the number of times the user's login is unsuccessful.


Private Sub Form_Initialize()
    
    Call Connection 'Calling the Connection Procedure.
    
    'Creating a New Recordset To Be Used For Login Purposes Only
    Set rsLogin = New ADODB.Recordset
    
    With rsLogin
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .CursorLocation = adUseClient
    End With

    iLoginFailure = 1    ' When a login attempt is unsuccessful, I decrement this variable's value.
    
End Sub


Private Sub txtUsername_Change()

    'This block of code will enable or disable the "Go" button accordingly
    
    If txtUsername.Text = "" Then
        cmdGO.Enabled = False
    Else
        cmdGO.Enabled = True
    End If
    
    
End Sub

Private Sub txtUserName_keypress(KeyAscii As Integer)
    
    'This block of code prevents the user from using "Copy-Paste" (Ctrl+C, Ctrl+V) functions.
    
    If KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then   'This is for using the Enter key
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
    
End Sub

Private Sub txtPassword_Keypress(KeyAscii As Integer)

    'This block of code prevents the user from using "Copy-Paste" (Ctrl+C, Ctrl+V) functions.

    If KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Then
        KeyAscii = 0
        ElseIf KeyAscii = 13 Then   'This is for using the Enter key
        KeyAscii = 0
        SendKeys "{Tab}"
       Call cmdGO_Click
    End If
    
    
End Sub

Private Sub cmdGO_Click()

    If iLoginFailure <= 3 Then  'Checking If The User Is Still Allowed To Login
    
       'Selecting the Related Login Record from the User_Account Table.
        rsLogin.Open "select * from UserAccount where Username='" & txtUsername.Text & "'", conn
        
        With rsLogin
            
            
            If .RecordCount = 0 Then    'This Means That There Is No Matching Record
                
                iLoginFailure = iLoginFailure + 1   'Decrementing The Value Of i On Each Unsuccessful Login Attempt
                MsgBox "Sorry! Invalid User Name! Please Try Again!", vbCritical, "Invalid Login!"
                txtUsername.BackColor = &H80000018  'Highlighting The Textbox With The Error
                txtPassword.BackColor = &H80000005  'Highlighting The Textbox With The Error
                txtUsername.Text = ""
                txtUsername.SetFocus
                
            End If
        
            If .RecordCount <> 0 Then   'This Means That There Is A Matching Record
                If txtPassword.Text = .Fields(6).Value Then 'Checking Password
                                        
                    If .Fields(4) = "Administrator" Then    'Checking Designation
                    
                        'Passing Necessary Values To Global Variables
            
                        userName = .Fields(1).Value & " " & .Fields(2).Value
                        accessLevel = "Administrator"
                        userID = .Fields(0).Value
                        frmMDI.Show
                        Unload Me
                        
                    ElseIf .Fields(4) = "Cashier" Then    'Checking Designation
                    
                        'Passing Necessary Values To Global Variables
                        userName = .Fields(1).Value & " " & .Fields(2).Value
                        accessLevel = "Cashier"
                        userID = .Fields(0).Value
                        Unload Me
                        frmMDI.Show
                        
                    ElseIf .Fields(4) = "Inpatient Staff" Then    'Checking Designation
                    
                        'Passing Necessary Values To Global Variables
                        userName = .Fields(1).Value & " " & .Fields(2).Value
                        accessLevel = "Inpatient Staff"
                        userID = .Fields(0).Value
                        frmMDI.Show
                        Unload Me
                        
                    ElseIf .Fields(4) = "Outpatient Staff" Then    'Checking Designation
                    
                        'Passing Necessary Values To Global Variables
                        userName = .Fields(1).Value & " " & .Fields(2).Value
                        accessLevel = "Outpatient Staff"
                        userID = .Fields(0).Value
                        frmMDI.Show
                        Unload Me
                        
                    ElseIf .Fields(4) = "Receptionist" Then    'Checking Designation
                    
                        'Passing Necessary Values To Global Variables
                        userName = .Fields(1).Value & " " & .Fields(2).Value
                        accessLevel = "Receptionist"
                        userID = .Fields(0).Value
                        frmMDI.Show
                        Unload Me
                        
                    End If
                    
                Else
                
                    'Error Mesage For Invalid Password
                    iLoginFailure = iLoginFailure + 1   'Decrementing The Value Of i On Each Unsuccessful Login Attempt
                    MsgBox "Sorry! Invalid Password! Please Try Again!", vbCritical, "Invalid Login!"
                    txtPassword.BackColor = &H80000018  'Highlighting The Textbox With The Error
                    txtUsername.BackColor = &H80000005  'Highlighting The Textbox With The Error
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                    
                End If
                
            End If
            
            .Close  'Closing Recordset
                        
        End With
        
        Else
            'Error Message If User's Login Attempt Is Unsuccesful On Three
            'Consecutive Occasions
            MsgBox "Sorry! You Have To Login Within Three Tries! Unloading...", vbCritical, "Login Failure"
        End
        
    End If
    
End Sub


Private Sub cmdExit_Click()

    'This block of code will be executed if the user decides to quit the application
    'from the Login page
    
    Dim ans As Variant
    ans = MsgBox("Are You Sure You Wish To Quit The Application?", vbYesNo + vbQuestion, "Quit Application?")
    
    If ans = vbYes Then
        End
    End If
    
End Sub

