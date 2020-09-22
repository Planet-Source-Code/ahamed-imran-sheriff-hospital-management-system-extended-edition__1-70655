VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackupDatabase 
   Caption         =   "Backup Database"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmBackupDatabase.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.DirListBox Dir 
      Appearance      =   0  'Flat
      Height          =   1665
      Left            =   2040
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.DriveListBox Drive 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4320
      Width           =   3135
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   5760
      Width           =   3135
   End
   Begin VB.CommandButton cmdClose 
      Height          =   855
      Left            =   8760
      Picture         =   "frmBackupDatabase.frx":1E873
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdBackup 
      Height          =   855
      Left            =   8760
      Picture         =   "frmBackupDatabase.frx":215B7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.Timer timPrgBar 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10200
      Top             =   6840
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   135
      Left            =   1680
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   8280
      X2              =   4200
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lbl_fra_BackUp 
      BackStyle       =   0  'Transparent
      Caption         =   "Back Up Information"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   1920
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   8280
      X2              =   8280
      Y1              =   6360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   8280
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   1680
      Y1              =   3360
      Y2              =   6360
   End
   Begin VB.Label lblBackUpPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Back Up Path :"
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
      Left            =   4440
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      Caption         =   "Back Up Filename :"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   8520
      X2              =   8520
      Y1              =   3360
      Y2              =   6360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   9960
      X2              =   9960
      Y1              =   3360
      Y2              =   6360
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   8520
      X2              =   9960
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   8520
      X2              =   9960
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "frmBackupDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Backup Database Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing):  Isham Sally
'Start Date: 23/04/08
'Date Of Last Modification: 23/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed:
'--------------------------------------------------------------------------------

Option Explicit
Dim FileSystemObject As Object
Dim strfilename As String


Private Sub cmdBackup_Click()
On Error GoTo e
    'set the copying functionality
    strfilename = "" + txtPath.Text + "\" + txtFile.Text + ".mdb"
    'Set the object contractions
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    'Copy the file according the path settings
    FileSystemObject.copyfile App.Path & "\sdp.mdb", strfilename
    PrgBar.Visible = True
    timPrgBar.Enabled = True
Exit Sub
e:
MsgBox "Invalid Path Setting, Please Try Again", vbCritical, "Invalid Path Setting!"
End Sub


Private Sub Dir_Click()
    txtPath.Text = "" & Dir.Path
End Sub

Private Sub Drive_Change()
    Dim d, fs As Object
    
    'Set the constrctions to created objectes
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.getdrive(fs.getdrivename(Drive.Drive))
    
    'Set the contents of the selected drives
    If d.isready Then
        Dir.Path = Drive.Drive
        Dir.SetFocus
    Else
        MsgBox "The Drive Is Not Ready!", vbExclamation, "Drive Not Ready!"
    End If
End Sub

Private Sub Form_Load()
    'Display Today 's date
    txtFile.Text = FormatDateTime(Now, vbLongDate)
End Sub


Private Sub timPrgBar_Timer()
    Static iCnt As Integer
    'Run the timer and check the condition
    If iCnt <= 100 Then
        PrgBar.Value = iCnt
        iCnt = iCnt + 1
    Else
        MsgBox "The Backup Procedure Has Been Successfully Completed!", vbInformation, "Successful Backup Procedure!"
        Drive.SetFocus
        PrgBar.Visible = False
        timPrgBar.Enabled = False
    End If
End Sub



Private Sub cmdClose_Click()    'On click of the Close Button
    
    'Obtaining confirmation from the user
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub
