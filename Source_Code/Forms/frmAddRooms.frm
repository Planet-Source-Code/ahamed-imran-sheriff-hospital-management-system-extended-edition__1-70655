VERSION 5.00
Begin VB.Form frmAddRooms 
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddRooms.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   11655
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboRoomType 
      Height          =   315
      ItemData        =   "frmAddRooms.frx":2A96D
      Left            =   4680
      List            =   "frmAddRooms.frx":2A97A
      TabIndex        =   1
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8400
      Picture         =   "frmAddRooms.frx":2A99F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Height          =   855
      Left            =   8400
      Picture         =   "frmAddRooms.frx":2D6E3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtRoomID 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4680
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
   End
   Begin VB.TextBox txtCapacity 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4680
      TabIndex        =   2
      Top             =   5400
      Width           =   2895
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4680
      TabIndex        =   3
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   9600
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   9600
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   9600
      X2              =   9600
      Y1              =   3480
      Y2              =   6840
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   8160
      X2              =   8160
      Y1              =   3480
      Y2              =   6840
   End
   Begin VB.Label lblROOM_Type 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Type"
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
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblCapacity 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Capacity"
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
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblROOM_ID 
      BackStyle       =   0  'Transparent
      Caption         =   "Room ID :"
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
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   2520
      X2              =   2520
      Y1              =   3480
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   2520
      X2              =   7920
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   7920
      Y1              =   6840
      Y2              =   3480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2520
      X2              =   2760
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lbl_fra_Room 
      BackStyle       =   0  'Transparent
      Caption         =   "New Room Information"
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
      TabIndex        =   7
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   5280
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lblROOM_Description 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Description"
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
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
End
Attribute VB_Name = "frmAddRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------
'Form Name: Add Roooms
'Software Architect : Ahamed Imran Sheriff (CB002260)
'Junior Programmer : Nimesh Wijemanne(CB002362)
'Date Completed: 10/01/08
'Beta Version
'--------------------------------------------------------

Option Explicit
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim x As Integer


Private Sub Form_Load()
    Dim connstring As String
    Set con = New ADODB.Connection
    connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase.mdb"
    con.Open connstring
    Set rs = New ADODB.Recordset
    rs.Open "RoomInfo", con, adOpenDynamic, adLockOptimistic
End Sub


Private Sub cmdSave_Click()
On Error GoTo err_handler:
    x = MsgBox("Do you wish to add a new Room Record?", vbYesNo + vbQuestion, "Add Record?")
    If x = vbYes Then
        If Not (rs.RecordCount = 0) Then
            rs.MoveLast
        End If
        rs.AddNew
        rs.Fields(0) = txtRoomID.Text
        rs.Fields(1) = cboRoomType.Text
        rs.Fields(2) = txtCapacity.Text
        rs.Fields(4) = txtDescription.Text
        rs.Update
        txtRoomID.Text = ""
        cboRoomType.Text = ""
        txtCapacity.Text = ""
        txtDescription.Text = ""
        MsgBox "The record has been updated succesfully!", vbOKOnly + vbInformation, "Updated!"
    Else
        txtRoomID.Text = ""
        cboRoomType.Text = ""
        txtCapacity.Text = ""
        txtDescription.Text = ""
        MsgBox "No modifications have taken place!", vbOKOnly + vbInformation, "No Modifications!"
    End If
    Exit Sub
err_handler:
    MsgBox Err.Description, vbCritical
End Sub


Private Sub cmdClose_Click()
    Dim ans As Variant
    ans = MsgBox("Do you wish to close this module?", vbYesNo + vbQuestion, "Close Module?")
    If ans = vbYes Then
        Unload Me
    End If
End Sub


Private Sub txtRoomID_KeyPress(KeyAscii As Integer)

    'Character Uppercase Validation for the Module Code
    If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = 32 Then
    ElseIf KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

