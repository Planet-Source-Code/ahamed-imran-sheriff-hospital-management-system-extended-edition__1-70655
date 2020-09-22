VERSION 5.00
Begin VB.Form frmMakeBookings 
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmRoomBookings.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   11655
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8160
      Picture         =   "frmRoomBookings.frx":29784
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Width           =   975
   End
   Begin VB.ComboBox cboID 
      Height          =   315
      ItemData        =   "frmRoomBookings.frx":2C4C8
      Left            =   4440
      List            =   "frmRoomBookings.frx":2C4CA
      TabIndex        =   0
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ComboBox cboRoom 
      Height          =   315
      ItemData        =   "frmRoomBookings.frx":2C4CC
      Left            =   4440
      List            =   "frmRoomBookings.frx":2C4D9
      TabIndex        =   3
      Top             =   5040
      Width           =   2895
   End
   Begin VB.ComboBox cboModule 
      Height          =   315
      ItemData        =   "frmRoomBookings.frx":2C4FE
      Left            =   4440
      List            =   "frmRoomBookings.frx":2C500
      TabIndex        =   2
      Top             =   4320
      Width           =   2895
   End
   Begin VB.ComboBox cboBatch 
      Height          =   315
      ItemData        =   "frmRoomBookings.frx":2C502
      Left            =   4440
      List            =   "frmRoomBookings.frx":2C504
      TabIndex        =   1
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4440
      TabIndex        =   5
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      Height          =   855
      Left            =   8160
      Picture         =   "frmRoomBookings.frx":2C506
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8160
      Picture         =   "frmRoomBookings.frx":2F24A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox cboAvailable 
      Height          =   315
      ItemData        =   "frmRoomBookings.frx":31F8E
      Left            =   4440
      List            =   "frmRoomBookings.frx":31F90
      TabIndex        =   4
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose ID : "
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
      TabIndex        =   14
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Room Type"
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
      TabIndex        =   13
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Module : "
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
      TabIndex        =   12
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Batch : "
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
      TabIndex        =   11
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   2640
      TabIndex        =   10
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label lblROOM_Description 
      BackStyle       =   0  'Transparent
      Caption         =   "Rooms Available : "
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
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   7680
      X2              =   4440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lbl_fra_Room 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Bookings"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2280
      X2              =   2520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   7680
      X2              =   7680
      Y1              =   7800
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   2280
      X2              =   7680
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   2280
      X2              =   2280
      Y1              =   2280
      Y2              =   7800
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   7920
      Y1              =   2280
      Y2              =   7800
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   9360
      X2              =   9360
      Y1              =   2280
      Y2              =   7800
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   9360
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   7920
      X2              =   9360
      Y1              =   7800
      Y2              =   7800
   End
End
Attribute VB_Name = "frmMakeBookings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------
'Form Name: Make Bookings
'Software Architect : Ahamed Imran Sheriff (CB002260)
'Junior Programmer : Nimesh Wijemanne(CB002362)
'Date Completed: 10/01/08
'Beta Version
'--------------------------------------------------------

Option Explicit
Dim ireply As Integer
Dim icnt As Integer
Dim icnt1 As Integer
Dim icnt2 As Integer
Dim icnt3 As Integer
Dim bChange As Boolean
Dim strBatCode As String
Dim strModCode As String
Dim strRoomCode As String
Dim strUsrCode As String
Dim strType As String
Dim strHold As String
Dim strRoom As String




Private Sub cboID_Click()

    cboBatch.Clear
    
    With rsLectBatch
        .MoveFirst
        While (.EOF = False)
            
            If (cboID.Text = .Fields(0)) Then
                cboBatch.AddItem .Fields(2)
                cboBatch.AddItem .Fields(3)
                cboBatch.AddItem .Fields(4)
                cboBatch.AddItem .Fields(5)
                .MoveLast
            End If
            .MoveNext
        
        Wend
        rsLectBatch.Requery
    End With
    
        cboModule.Clear
    
    With rsLectMd
        .MoveFirst
        While (.EOF = False)
            
            If (cboID.Text = .Fields(0)) Then
                cboModule.AddItem .Fields(2)
                cboModule.AddItem .Fields(3)
                cboModule.AddItem .Fields(4)
                cboModule.AddItem .Fields(5)
                .MoveLast
            End If
            .MoveNext
        
        Wend
        rsLectMd.Requery
    End With
    
End Sub

Private Sub cboRoom_Click()
    strType = cboRoom.Text
    cboAvailable.Clear
    
    rsRooms.MoveFirst
    While rsRooms.EOF = False
        If rsRooms.Fields(3) = False And rsRooms.Fields(1) = strType Then
            cboAvailable.AddItem rsRooms.Fields(0)
        End If
        rsRooms.MoveNext
    Wend
    rsRooms.Requery
    
    'rsRooms.Close
End Sub

Private Sub cmdClear_Click()
    cboBatch.Text = ""
    cboModule.Text = ""
    cboRoom.Text = ""
    cboAvailable.Text = ""
    txtDescription.Text = ""
End Sub

Private Sub cmdClose_Click()
    Dim ans As Variant
    ans = MsgBox("Do you wish to close this module?", vbYesNo + vbQuestion, "Close Module?")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdSave_Click()
    
    If cboAvailable.Text = "" Then
        MsgBox "Sorry! There are no rooms available at the moment!", vbOKOnly + vbExclamation, "No Rooms!"
        Exit Sub
    End If
    
    'On Error GoTo e
        ireply = MsgBox("Do You Wish to Make a Booking?", vbYesNo + vbQuestion)
        With rsResRoom
            'filling values to the blank record added by the addnew
            'statement
            .AddNew
            .Fields(1) = cboAvailable
            .Fields(2) = cboID
            .Fields(3) = cboModule.Text
            .Fields(4) = cboBatch.Text
            strRoom = cboAvailable
            
            If ireply = vbYes Then
                .Update       ' makes the changes permanent
                Call chkRoom
                MsgBox "Your Room Booking Has Been Made Successfully!", vbInformation
            Else
                .CancelUpdate ' cancels the changes made including
                              ' the addnew operation
            End If
        End With
    rsResRoom.Requery   ' execute the query again on the DB and get new
                          ' records after change to the database
    Exit Sub
'e:
   'MsgBox Err.Description, vbCritical

End Sub

Private Function chkRoom()  'Userdefine function
    
    With rsRooms
        .MoveFirst
        While .EOF = False
            If strRoom = .Fields(0) Then
                .Fields(3) = True
                .Update
                txtDescription = .Fields(4)
            End If
            .MoveNext
        Wend
    End With
    
End Function


Private Sub Form_Load()
    
    Connection          'calling conection sub procedure
    
    
    Call Lecturer_Batch_Set
    Call Lecturer_Module_Set
    Call Room_Resv
    Call Batch_Info
    Call Module_Info
    Call Room_Info

    icnt = 0
    While (rsLectBatch.EOF = False)
        strUsrCode = rsLectBatch.Fields(0)
        
        cboID.AddItem strUsrCode, icnt
        icnt = icnt + 1
        rsLectBatch.MoveNext
    Wend
    rsLectBatch.Requery
    rsLectBatch.MoveFirst
    
    With rsLectBatch
    While (.EOF = False)
    
        strBatCode = rsLectBatch.Fields(2)
        
        If (cboID.Text = strUsrCode) Then
                cboBatch.AddItem strBatCode, icnt1
                icnt2 = icnt2 + 1
                .MoveLast
        End If
        .MoveNext
    
    Wend
    rsBatch.Requery
    End With
End Sub




