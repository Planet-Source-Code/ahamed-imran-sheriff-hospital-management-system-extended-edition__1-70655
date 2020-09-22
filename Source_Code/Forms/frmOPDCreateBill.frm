VERSION 5.00
Begin VB.Form frmOPDCreateBill 
   Caption         =   "Create Patient Bill"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmOPDCreateBill.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Click Here To Save This Record"
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Click Here To Close This Interface"
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Click Here To Print This Record"
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "GO"
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
      Left            =   10440
      TabIndex        =   18
      Top             =   6120
      Width           =   495
   End
   Begin VB.PictureBox picInvalidTypingMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8640
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   50
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   51
         Top             =   105
         Width           =   2175
      End
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   8280
      Top             =   1320
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO"
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
      Left            =   10320
      TabIndex        =   9
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtPatientName 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.OptionButton optCreditCard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9360
      TabIndex        =   13
      Top             =   4080
      Width           =   255
   End
   Begin VB.OptionButton optCash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   4080
      Width           =   255
   End
   Begin VB.OptionButton optCheque 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000003&
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
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
      Left            =   7680
      TabIndex        =   19
      Text            =   "0"
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox txtTotalPayable 
      Alignment       =   1  'Right Justify
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox txtTotalCost 
      Alignment       =   1  'Right Justify
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox txtChequeNo 
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
      Left            =   7680
      TabIndex        =   14
      Text            =   "-"
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtCardNo 
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
      Left            =   7680
      TabIndex        =   15
      Text            =   "-"
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtBankName 
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
      Left            =   7680
      TabIndex        =   16
      Text            =   "-"
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtTotalRecieved 
      Alignment       =   1  'Right Justify
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
      Left            =   7680
      TabIndex        =   17
      Text            =   "0"
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox txtAmountPaid 
      Alignment       =   1  'Right Justify
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
      Left            =   7560
      TabIndex        =   8
      Text            =   "0"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtBillStatus 
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
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   7560
      TabIndex        =   10
      Text            =   "UNPAID"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtInvoiceID 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtBillingDate 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtAccountType 
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4800
      Width           =   2295
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   3
      Left            =   10080
      TabIndex        =   49
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   2
      Left            =   10080
      TabIndex        =   48
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   1
      Left            =   9960
      TabIndex        =   47
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   0
      Left            =   4920
      TabIndex        =   46
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label lblRs 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Index           =   5
      Left            =   4920
      TabIndex        =   45
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label lblPatientName 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name"
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
      TabIndex        =   44
      Top             =   4245
      Width           =   1335
   End
   Begin VB.Label lblCreditCard 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card"
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
      Left            =   9600
      TabIndex        =   43
      Top             =   4170
      Width           =   1095
   End
   Begin VB.Label lblCash 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash"
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
      Left            =   8160
      TabIndex        =   42
      Top             =   4170
      Width           =   495
   End
   Begin VB.Label lblCheque 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   6480
      TabIndex        =   41
      Top             =   4170
      Width           =   855
   End
   Begin VB.Label lblBalance 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   40
      Top             =   6645
      Width           =   1335
   End
   Begin VB.Label lblTotalPayable 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payable"
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
      TabIndex        =   39
      Top             =   6645
      Width           =   1335
   End
   Begin VB.Label lblAccountType 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type"
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
      TabIndex        =   38
      Top             =   4845
      Width           =   1335
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
      TabIndex        =   37
      Top             =   3645
      Width           =   1335
   End
   Begin VB.Label lblFrameTitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Details"
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
      TabIndex        =   36
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   720
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2400
      X2              =   5400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   5400
      X2              =   5400
      Y1              =   2040
      Y2              =   7080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   5400
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   480
      Y1              =   2040
      Y2              =   7080
   End
   Begin VB.Label lblTotalCost 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost"
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
      TabIndex        =   35
      Top             =   5445
      Width           =   1335
   End
   Begin VB.Label lblDiscount 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      TabIndex        =   34
      Top             =   6045
      Width           =   1335
   End
   Begin VB.Label lblChequeNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No."
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
      Height          =   255
      Left            =   6120
      TabIndex        =   33
      Top             =   4725
      Width           =   1695
   End
   Begin VB.Label lblCardNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Card No."
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
      Height          =   255
      Left            =   6120
      TabIndex        =   32
      Top             =   5205
      Width           =   1695
   End
   Begin VB.Label lblBankName 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   31
      Top             =   5685
      Width           =   1695
   End
   Begin VB.Label lblTotalRecieved 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Recieved"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   30
      Top             =   6165
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Info"
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
      Left            =   6120
      TabIndex        =   29
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   6000
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   7560
      X2              =   11160
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   11160
      X2              =   11160
      Y1              =   2040
      Y2              =   3600
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   11160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   5640
      Y1              =   2040
      Y2              =   3600
   End
   Begin VB.Label lblAmountPaid 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid"
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
      Left            =   6120
      TabIndex        =   28
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label lblBillStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Status"
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
      Left            =   6120
      TabIndex        =   27
      Top             =   3045
      Width           =   2055
   End
   Begin VB.Label lblInvoiceID 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice ID"
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
      Top             =   2445
      Width           =   1335
   End
   Begin VB.Label lblBillingDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Date"
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
      Top             =   3045
      Width           =   1335
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Plan"
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
      Left            =   6120
      TabIndex        =   23
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   6000
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000001&
      X1              =   7680
      X2              =   11160
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000001&
      X1              =   11160
      X2              =   11160
      Y1              =   3840
      Y2              =   7080
   End
   Begin VB.Line Line18 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   11160
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line19 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   5640
      Y1              =   3840
      Y2              =   7080
   End
End
Attribute VB_Name = "frmOPDCreateBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Outpatients Create Bill Interface
'Programmer: Isham Sally
'Quality Assurance Engineer (Testing): Imran Sheriff
'Start Date: 11/05/08
'Date Of Last Modification: 11/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Outpatient_Billing Table
'----------------------------------------------------------------------------



'The Following Boolean Variable is being used to determine
'if the data the user enters is valid or not
Dim Flag As Boolean

'This variable will hold the invoice ID when saving for the purpose of report generation.
Dim invoiceid As String


Private Sub cmdClose_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub


Private Sub cmdGO_Click()
            
    If Val(txtAmountPaid.Text) <> Val(txtTotalPayable.Text) Then
        MsgBox "Error! The Amount Paid Has To Be Equal To The Total Payable!", vbCritical, "Amount Paid Has To Equal Total Payable!"
        txtBillStatus.Text = "UNPAID"
        txtAmountPaid.Text = ""
        Exit Sub
    Else
        txtBillStatus.Text = "PAID"
    End If
    
    cmdGO.Enabled = False
    txtAmountPaid.Enabled = False
    
End Sub


Private Sub cmdOK_Click()
    
    'Here, I am checking to ensure that the user does not type in a value that is lower than the Amount Paid
    If Val(txtTotalRecieved.Text) < Val(txtAmountPaid.Text) Then
        MsgBox "Error! The Total Received Cannot Be Less Than The Amount Paid!", vbCritical, "Figure Is Too Low!"
        txtTotalRecieved.Text = "0"
        Exit Sub
    End If
    
    'Calculating the balance to be given to the user
    txtBalance.Text = Val(txtTotalRecieved.Text) - Val(txtAmountPaid.Text)

    cmdOK.Enabled = False
    txtTotalRecieved.Enabled = False
    txtBalance.Enabled = False
    
End Sub

Private Sub cmdPrint_Click()
    
    On Error GoTo e
    DataEnvironment1.Commands("OutpatientReceipt").Parameters(0) = invoiceid
    RptOutpatientReceipt.Show
    DataEnvironment1.rsOutpatientReceipt.Close
        
    Unload Me
    Exit Sub
e:
    If Err.Number <> 3704 Then
        MsgBox Err.Description & "" & Err.Number, vbCritical
    End If

End Sub

Private Sub cmdSave_Click()
    'Checking if the user has filled in the Payment Plan frame
    If optCheque.Value <> True And optCash.Value <> True And optCreditCard.Value <> True Then
        MsgBox "Error! You Have Not Chosen A Payment Type!", vbCritical, "Please Choose Payment Type!"
        Exit Sub
    End If
        
        
    'Checking the return value of the function that validates the user's data
    If textfieldsValidations = False Then
    
        
        With rsOutpatientBilling
            
            'Making sure that the user wants to save the record
            If MsgBox("Are You Sure You Wish To Record This Payment?", vbYesNo + vbQuestion, "Record Payment?") = vbYes Then
                
                'The following block of if else conditions ensure that no
                'textfield will be completely blank when saving in the database.
                'This has been done in order to avoid errors.
                If txtChequeNo.Text = "" Then
                    txtChequeNo.Text = "-"
                End If
                
                If txtCardNo.Text = "" Then
                    txtCardNo.Text = "-"
                End If
                
                If txtBankName.Text = "" Then
                    txtBankName.Text = "-"
                End If
                
                If txtDiscount.Text = "" Then
                    txtDiscount.Text = "-"
                End If
                                
                'Save the user-entered data into the recordset
                .Fields(0) = txtInvoiceID.Text
                invoiceid = txtInvoiceID.Text
                .Fields(1) = txtBillingDate.Text
                .Fields(2) = txtPatientID.Text
                .Fields(3) = txtPatientName.Text
                .Fields(4) = txtAccountType.Text
                .Fields(5) = txtTotalCost.Text
                .Fields(6) = txtDiscount.Text
                .Fields(7) = txtTotalPayable.Text
                .Fields(8) = txtAmountPaid.Text
                .Fields(9) = txtBillStatus.Text
                
                If optCheque.Value = True Then
                    .Fields(10).Value = "Cheque"
                ElseIf optCash.Value = True Then
                    .Fields(10).Value = "Cash"
                ElseIf optCreditCard.Value = True Then
                    .Fields(10).Value = "CreditCard"
                End If
                
                .Fields(11) = txtChequeNo.Text
                .Fields(12) = txtCardNo.Text
                .Fields(13) = txtBankName.Text
                .Fields(14) = txtTotalRecieved.Text
                .Fields(15) = txtBalance.Text
            
                .Update
                
                'Display Success Message
                MsgBox "The Payment Was Recorded Successfully!", vbInformation, "Payment Recorded Successfully!"
                
                            
            Else
            
                'Display 'No Modifications' Message
                MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
                .CancelUpdate   'Cancel the Save Procedure
            
            End If
            
            .Requery    'Requerying the Table
                 
            
        End With
        
    End If
    
    cmdPrint.Enabled = True 'Enabling the Print button
    
End Sub


Private Sub optCash_Click()

    'Here, I am checking to see if the Payment Info frame has been filled
    If txtAmountPaid.Text = "0" Then
        MsgBox "Error! Please Fill-in The Payment Info Frame First!", vbCritical, "Payment Info Missing!"
        Exit Sub
    End If
    
    disablePaymentPlan  'Calling a function to diable all fields in the Payment Plan frame
    
    lblTotalRecieved.Enabled = True
    txtTotalRecieved.Enabled = True
    lblBalance.Enabled = True
    txtBalance.Enabled = True
    cmdOK.Enabled = True
    
End Sub

Private Sub optCheque_Click()

    'Here, I am checking to see if the Payment Info frame has been filled
    If txtAmountPaid.Text = "0" Then
        MsgBox "Error! Please Fill-in The Payment Info Frame First!", vbCritical, "Payment Info Missing!"
        Exit Sub
    End If

    disablePaymentPlan  'Calling a function to diable all fields in the Payment Plan frame
    
    lblChequeNo.Enabled = True
    txtChequeNo.Enabled = True
    lblBankName.Enabled = True
    txtBankName.Enabled = True

End Sub

Private Sub optCreditCard_Click()
    
    'Here, I am checking to see if the Payment Info frame has been filled
    If txtAmountPaid.Text = "0" Then
        MsgBox "Error! Please Fill-in The Payment Info Frame First!", vbCritical, "Payment Info Missing!"
        Exit Sub
    End If

    disablePaymentPlan  'Calling a function to diable all fields in the Payment Plan frame
    
    lblCardNo.Enabled = True
    txtCardNo.Enabled = True
    lblBankName.Enabled = True
    txtBankName.Enabled = True
    
End Sub

Public Function disablePaymentPlan()    'This function will disable all fields in the Payment Plan frame
    
    lblChequeNo.Enabled = False
    txtChequeNo.Enabled = False
    lblCardNo.Enabled = False
    txtCardNo.Enabled = False
    lblBankName.Enabled = False
    txtBankName.Enabled = False
    lblTotalRecieved.Enabled = False
    txtTotalRecieved.Enabled = False
    lblBalance.Enabled = False
    txtBalance.Enabled = False
    
End Function


Private Sub txtAmountPaid_GotFocus()
    
    txtAmountPaid.Text = ""
    
End Sub

Private Sub txtAmountPaid_LostFocus()
    
    If txtAmountPaid.Text = "" Then
        txtAmountPaid.Text = "0"
    End If
    
End Sub


Private Sub tmrErrMsg_Timer()

    Static i As Integer
    
    If i < 200000 Then     'Validation Msg Viewing Time Period
        picInvalidTypingMsg.Visible = False
        tmrErrMsg.Enabled = False
    Else
        i = i + 1
    End If
    
End Sub


Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidTypingMsg.Top = 2400    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtTotalRecieved_GotFocus()
    
    txtTotalRecieved.Text = ""
    
End Sub

Private Sub txtTotalRecieved_LostFocus()
    
    If txtTotalRecieved.Text = "" Then
        txtTotalRecieved.Text = "0"
    End If
    
End Sub

Private Sub txtTotalRecieved_KeyPress(KeyAscii As Integer)
    
    'Keypress Validation to allow only digits
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        picInvalidTypingMsg.Top = 6120    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub


Public Function textfieldsValidations()
    
    Flag = True 'Setting the Flag variable to True
    
    'Checking if the Amount Paid textfield is empty
    If txtAmountPaid.Text = "0" Then
        txtAmountPaid.BackColor = &H80000018   'Highlighting the textfield in a different colour
        Flag = False    'Setting the Flag variable to False to indicate invalid data
    Else
        txtAmountPaid.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
    End If
    
    'Checking if the Cheque No textfield is empty
    If txtChequeNo.Enabled = True Then
        If txtChequeNo.Text = "-" Then
            txtChequeNo.BackColor = &H80000018   'Highlighting the textfield in a different colour
            Flag = False    'Setting the Flag variable to False to indicate invalid data
        Else
            txtChequeNo.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        End If
    End If
    
    'Checking if the Card No textfield is empty
    If txtCardNo.Enabled = True Then
        If txtCardNo.Text = "-" Then
            txtCardNo.BackColor = &H80000018   'Highlighting the textfield in a different colour
            Flag = False    'Setting the Flag variable to False to indicate invalid data
        Else
            txtCardNo.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        End If
    End If
    
    'Checking if the Bank Name textfield is empty
    If txtBankName.Enabled = True Then
        If txtBankName.Text = "-" Then
            txtBankName.BackColor = &H80000018   'Highlighting the textfield in a different colour
            Flag = False    'Setting the Flag variable to False to indicate invalid data
        Else
            txtBankName.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        End If
    End If
    
    'Checking if the Total Recieved textfield is empty
    If txtTotalRecieved.Enabled = True Then
        If txtTotalRecieved.Text = "0" Then
            txtTotalRecieved.BackColor = &H80000018   'Highlighting the textfield in a different colour
            Flag = False    'Setting the Flag variable to False to indicate invalid data
        Else
            txtTotalRecieved.BackColor = &H80000004   'Bringing the textfield BackColour back to normal
        End If
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


