VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmIPDCreateBill 
   Caption         =   "Create Inpatient Bill"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmIPDCreateBill.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Click Here To Close This Interface"
      Top             =   8040
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Click Here To Close This Interface"
      Top             =   8040
      Width           =   1695
   End
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Click Here To Save This Record"
      Top             =   8040
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
      Left            =   10320
      TabIndex        =   21
      Top             =   6600
      Width           =   495
   End
   Begin VB.Timer tmrErrMsg 
      Interval        =   1000
      Left            =   9120
      Top             =   1560
   End
   Begin VB.PictureBox picInvalidTypingMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   9240
      ScaleHeight     =   825
      ScaleWidth      =   2385
      TabIndex        =   59
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sorry! You Cannot Type Alphabets Here! Only Digits Are Allowed!"
         Height          =   615
         Left            =   120
         TabIndex        =   60
         Top             =   105
         Width           =   2175
      End
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
      Left            =   10800
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.OptionButton optCheque 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   4560
      Width           =   255
   End
   Begin VB.OptionButton optCash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   4560
      Width           =   255
   End
   Begin VB.OptionButton optCreditCard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9360
      TabIndex        =   16
      Top             =   4560
      Width           =   255
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4200
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtAdmissionID 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3240
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4680
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2760
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2280
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "UNPAID"
      Top             =   3120
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
      Left            =   8520
      TabIndex        =   11
      Text            =   "0"
      Top             =   2640
      Width           =   1815
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
      Left            =   8040
      TabIndex        =   20
      Text            =   "0"
      Top             =   6600
      Width           =   1815
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
      Left            =   8040
      TabIndex        =   19
      Text            =   "-"
      Top             =   6120
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
      Left            =   8040
      TabIndex        =   18
      Text            =   "-"
      Top             =   5640
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
      Left            =   8040
      TabIndex        =   17
      Text            =   "-"
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtBalanceOwing 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   7080
      Width           =   1815
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   5160
      Width           =   1815
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5640
      Width           =   1815
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox txtTotalPaidSoFar 
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
      ForeColor       =   &H80000006&
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   6600
      Width           =   1815
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
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0"
      Top             =   7080
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dgrdTotalPaidSoFar 
      Height          =   1095
      Left            =   6480
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1931
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
      Caption         =   "Total Paid So Far"
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
      Index           =   6
      Left            =   9960
      TabIndex        =   58
      Top             =   7200
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
      Left            =   9960
      TabIndex        =   57
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
      Index           =   3
      Left            =   10440
      TabIndex        =   56
      Top             =   2720
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
      Left            =   4680
      TabIndex        =   54
      Top             =   7200
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
      Left            =   4680
      TabIndex        =   53
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
      Index           =   0
      Left            =   4680
      TabIndex        =   52
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
      Index           =   4
      Left            =   4680
      TabIndex        =   51
      Top             =   5280
      Width           =   375
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
      Height          =   255
      Left            =   6480
      TabIndex        =   50
      Top             =   4650
      Width           =   855
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
      TabIndex        =   49
      Top             =   4650
      Width           =   495
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
      TabIndex        =   48
      Top             =   4650
      Width           =   1095
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
      TabIndex        =   47
      Top             =   4245
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
      TabIndex        =   46
      Top             =   3765
      Width           =   1335
   End
   Begin VB.Line Line19 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   5640
      Y1              =   4320
      Y2              =   7560
   End
   Begin VB.Line Line18 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   11400
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000001&
      X1              =   11400
      X2              =   11400
      Y1              =   4320
      Y2              =   7560
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000001&
      X1              =   7680
      X2              =   11400
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   6000
      Y1              =   4320
      Y2              =   4320
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
      TabIndex        =   45
      Top             =   4200
      Width           =   2535
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
      Left            =   4680
      TabIndex        =   44
      Top             =   5715
      Width           =   375
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
      TabIndex        =   43
      Top             =   2805
      Width           =   1335
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
      TabIndex        =   42
      Top             =   2325
      Width           =   1335
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
      TabIndex        =   41
      Top             =   3165
      Width           =   2055
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
      TabIndex        =   40
      Top             =   2685
      Width           =   1815
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   5640
      Y1              =   2040
      Y2              =   3960
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   11400
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   11400
      X2              =   11400
      Y1              =   2040
      Y2              =   3960
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   7560
      X2              =   11400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   6000
      Y1              =   2040
      Y2              =   2040
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
      TabIndex        =   39
      Top             =   1920
      Width           =   2535
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
      TabIndex        =   38
      Top             =   6645
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
      TabIndex        =   37
      Top             =   6165
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
      TabIndex        =   36
      Top             =   5685
      Width           =   1695
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
      TabIndex        =   35
      Top             =   5205
      Width           =   1695
   End
   Begin VB.Label lblBalanceOwing 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Owing"
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
      Top             =   7125
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
      TabIndex        =   33
      Top             =   5685
      Width           =   1335
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
      TabIndex        =   32
      Top             =   5205
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   480
      Y1              =   2040
      Y2              =   7560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   5400
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   5400
      X2              =   5400
      Y1              =   2040
      Y2              =   7560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2400
      X2              =   5400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   720
      Y1              =   2040
      Y2              =   2040
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
      TabIndex        =   31
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblAdmissionID 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission ID"
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
      Top             =   3285
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
      TabIndex        =   29
      Top             =   4725
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
      TabIndex        =   28
      Top             =   6165
      Width           =   1335
   End
   Begin VB.Label lblTotalPaidSoFar 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Paid So Far"
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
      Top             =   6645
      Width           =   1815
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
      TabIndex        =   26
      Top             =   7125
      Width           =   1335
   End
End
Attribute VB_Name = "frmIPDCreateBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Inpatients Create Bill Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing): Isham Sally
'Start Date: 11/05/08
'Date Of Last Modification: 11/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed: Inpatient_Billing Table
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
    
    'Here, I am checking to ensure that the user does not type in a value that is higher than the Balance Owing
    If Val(txtAmountPaid.Text) > Val(txtBalanceOwing.Text) Then
        MsgBox "Error! The Amount Paid Cannot Be Greater Than The Balance Owing!", vbCritical, "Figure Is Too High!"
        txtAmountPaid.Text = "0"
        Exit Sub
    End If
    
    'Incrementing the value in the Total Paid So far textfield
    txtTotalPaidSoFar.Text = Val(txtTotalPaidSoFar.Text) + Val(txtAmountPaid.Text)
    
    'Decrementing the value in the Balance Owing textfield
    txtBalanceOwing.Text = Val(txtBalanceOwing.Text) - Val(txtAmountPaid.Text)
    
    If txtBalanceOwing.Text = "0" Then
        txtBillStatus.Text = "PAID"
    Else
        txtBillStatus.Text = "UNPAID"
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
    'For Reports
    On Error GoTo e
    DataEnvironment1.Commands("InpatientReceipt").Parameters(0) = invoiceid
    RptInpatientReceipt.Show
    DataEnvironment1.rsInpatientReceipt.Close
        
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
    
        
        With rsInpatientBilling
            
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
                'invoiceid = ""  'Clearing the string
                invoiceid = txtInvoiceID.Text
                .Fields(1) = txtBillingDate.Text
                .Fields(2) = txtAdmissionID.Text
                .Fields(3) = txtPatientID.Text
                
                .Fields(4) = txtPatientName.Text
                .Fields(5) = txtAccountType.Text
                .Fields(6) = txtTotalCost.Text
                .Fields(7) = txtDiscount.Text
                .Fields(8) = txtTotalPayable.Text
                .Fields(9) = txtTotalPaidSoFar.Text
                .Fields(10) = txtBalanceOwing.Text
                .Fields(11) = txtAmountPaid.Text
                .Fields(12) = txtBillStatus.Text
                
                If optCheque.Value = True Then
                    .Fields(13).Value = "Cheque"
                ElseIf optCash.Value = True Then
                    .Fields(13).Value = "Cash"
                ElseIf optCreditCard.Value = True Then
                    .Fields(13).Value = "Credit Card"
                End If
                
                .Fields(14) = txtChequeNo.Text
                .Fields(15) = txtCardNo.Text
                .Fields(16) = txtBankName.Text
                .Fields(17) = txtTotalRecieved.Text
                .Fields(18) = txtBalance.Text
            
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
        picInvalidTypingMsg.Top = 2640    'Validation Note View
        picInvalidTypingMsg.Visible = True
        tmrErrMsg.Enabled = True
        KeyAscii = 0
    End If
    
End Sub








Private Sub txtBankName_GotFocus()
    
    txtBankName.Text = ""
    
End Sub

Private Sub txtBankName_LostFocus()
    
    If txtBankName.Text = "" Then
        txtBankName.Text = "-"
    End If
    
End Sub

Private Sub txtCardNo_GotFocus()
    
    txtCardNo.Text = ""
    
End Sub

Private Sub txtCardNo_LostFocus()
    
    If txtCardNo.Text = "" Then
        txtCardNo.Text = "-"
    End If
    
End Sub

Private Sub txtChequeNo_GotFocus()
    
    txtChequeNo.Text = ""
    
End Sub

Private Sub txtChequeNo_LostFocus()
    
    If txtChequeNo.Text = "" Then
        txtChequeNo.Text = "-"
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
        picInvalidTypingMsg.Top = 6600    'Validation Note View
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
