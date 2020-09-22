VERSION 5.00
Begin VB.Form frmChangePassword 
   Caption         =   "Change Password Module"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmChangePassword.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtOldPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtConfirmPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox txtNewPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Height          =   830
      Left            =   8400
      Picture         =   "frmChangePassword.frx":1E728
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   990
   End
   Begin VB.CommandButton cmdClose 
      Height          =   830
      Left            =   8400
      Picture         =   "frmChangePassword.frx":20D9C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   990
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   9600
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   9600
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   9600
      X2              =   9600
      Y1              =   3360
      Y2              =   6480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   8160
      Y1              =   3360
      Y2              =   6480
   End
   Begin VB.Label lblOldPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password :"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblConfirmPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password :"
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
      Left            =   2880
      TabIndex        =   9
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblNewPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password :"
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
      Left            =   2880
      TabIndex        =   8
      Top             =   5160
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
      Left            =   2880
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   2520
      X2              =   2520
      Y1              =   3360
      Y2              =   6480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   2520
      X2              =   7920
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   7920
      Y1              =   6480
      Y2              =   3360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2520
      X2              =   2760
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblUserAccountInformation 
      BackStyle       =   0  'Transparent
      Caption         =   "User Account Information "
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
      Left            =   2880
      TabIndex        =   6
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   5640
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name:   Change Password Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 23/04/08
'Date Of Last Modification: 23/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: UserAccount Table
'----------------------------------------------------------------------------



Private Sub Form_Load() 'On Form Load
    Call Connection 'Establishing connectivity with the database
    Call User_Account   'Calling the User_Account procedure to interact with the recordset
    txtUsername.Text = userID   'Automatically including the User ID in the Username textfield
End Sub


Private Sub cmdSave_Click() 'When the Save Button is clicked
    
    With rsUserAccount
    
        .MoveFirst  'Moving to the first record
        
        While .EOF = False    'Running through all the records in the database
            
            If .Fields(0).Value = txtUsername.Text Then 'Checking for the right Employee ID
                
                If .Fields(6).Value <> txtOldPassword Then   'Checking if the Old Password typed by the user is correct
                    MsgBox "Error! The Old Password You Provided Was Incorrect! Please Check Your Password!", vbCritical, "Password Mismatch!"
                    txtOldPassword.Text = ""    'Clearing the Old Password textfield
                    Exit Sub
                End If
                    
                If txtNewPassword.Text <> txtConfirmPassword.Text Then  'Checking if the new passwords match
                    MsgBox "Error! The New Passwords You Provided Do Not Match! Please Check Your Passwords!", vbCritical, "Password Mismatch!"
                    txtNewPassword.Text = ""    'Clearing the New Password textfield
                    txtConfirmPassword.Text = ""    'Clearing the Confirm Password textfield
                    Exit Sub
                End If
                        
                        
                If txtNewPassword.Text = txtConfirmPassword.Text Then   'Checking if the passwords match
                    'Making sure that the user wants to save the record
                    If MsgBox("Are You Sure You Wish To Change Your Password?", vbYesNo + vbQuestion, "Change Password?") = vbYes Then
                        .Fields(6).Value = txtNewPassword.Text
                        .Update 'Updating the recordset
                        'Display Success Message
                        MsgBox "Your Password Has Been Changed Successfully!", vbInformation, "Password Changed Succesfully!"
                        .MoveLast
                    Else
                        'Display 'No Modifications' Message
                        MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                        .MoveLast
                    End If
                End If
            
            Else
                
                .MoveNext   'Moving to the next record
            
            End If
            
        Wend
        
        Unload Me
        
    End With
    
End Sub


Private Sub cmdClose_Click()    'On click of the Close Button
    
    'Obtaining confirmation from the user
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub
