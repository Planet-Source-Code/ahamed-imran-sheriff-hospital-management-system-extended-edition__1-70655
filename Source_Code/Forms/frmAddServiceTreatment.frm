VERSION 5.00
Begin VB.Form frmAddServiceTreatment 
   Caption         =   "Add Service Treatments Module"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmAddServiceTreatment.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   4680
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      ToolTipText     =   "Click Here to select Customer"
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddServiceTreatment.frx":1EBC5
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddServiceTreatment.frx":21909
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddServiceTreatment.frx":2464D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   8640
      Picture         =   "frmAddServiceTreatment.frx":27391
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   4680
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   0
      ToolTipText     =   "Click Here to select Customer"
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Notes"
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
      Left            =   2640
      TabIndex        =   22
      Top             =   6405
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Treatment Date"
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
      Left            =   2640
      TabIndex        =   20
      Top             =   5445
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charge"
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
      Left            =   2640
      TabIndex        =   19
      Top             =   5925
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Name"
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
      Left            =   2640
      TabIndex        =   18
      Top             =   4965
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service ID"
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
      Left            =   2640
      TabIndex        =   17
      Top             =   4485
      Width           =   1335
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   8040
      X2              =   5400
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Treatment Information"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   2040
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   8040
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1680
      X2              =   1680
      Y1              =   2520
      Y2              =   7920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   8040
      X2              =   8040
      Y1              =   7920
      Y2              =   2520
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Treatment ID"
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
      TabIndex        =   15
      Top             =   3045
      Width           =   1575
   End
   Begin VB.Label Label16 
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
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   4005
      Width           =   1335
   End
   Begin VB.Label Label12 
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
      Left            =   2640
      TabIndex        =   13
      Top             =   3525
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   5415
      Left            =   8400
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frmaddServiceTreatment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
