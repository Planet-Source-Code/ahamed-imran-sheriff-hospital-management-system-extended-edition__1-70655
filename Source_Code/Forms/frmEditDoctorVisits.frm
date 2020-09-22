VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDoctorVisitsMaintenance 
   Caption         =   "Doctor Visits Maintenance Module"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmEditDoctorVisits.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUpdate 
      Height          =   855
      Left            =   7800
      Picture         =   "frmEditDoctorVisits.frx":1FF2A
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Height          =   855
      Left            =   6720
      Picture         =   "frmEditDoctorVisits.frx":22C6E
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Height          =   855
      Left            =   5640
      Picture         =   "frmEditDoctorVisits.frx":259B2
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   855
      Left            =   8880
      Picture         =   "frmEditDoctorVisits.frx":286F6
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   9960
      Picture         =   "frmEditDoctorVisits.frx":2B43A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   7320
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   19
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   18
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      Height          =   285
      Left            =   2640
      TabIndex        =   17
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdCusWizard 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      ToolTipText     =   "Click Here to select Customer"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   2640
      TabIndex        =   15
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   12
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      ToolTipText     =   "Click Here to select Customer"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox cboType 
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
      ItemData        =   "frmEditDoctorVisits.frx":2E17E
      Left            =   7320
      List            =   "frmEditDoctorVisits.frx":2E194
      TabIndex        =   5
      Text            =   "Visit ID"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   750
      Left            =   7440
      Picture         =   "frmEditDoctorVisits.frx":2E1E0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   750
      Left            =   6480
      Picture         =   "frmEditDoctorVisits.frx":3039C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdNext 
      Height          =   750
      Left            =   8400
      Picture         =   "frmEditDoctorVisits.frx":32558
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   890
   End
   Begin VB.CommandButton cmdLast 
      Height          =   750
      Left            =   9360
      Picture         =   "frmEditDoctorVisits.frx":34714
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   890
   End
   Begin MSDataGridLib.DataGrid DataGrid_Module 
      Height          =   2535
      Left            =   5520
      TabIndex        =   6
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
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
      Caption         =   "Visits Information Table"
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      Height          =   1095
      Left            =   5520
      Top             =   7320
      Width           =   5535
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
      Left            =   840
      TabIndex        =   29
      Top             =   7365
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit Date"
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
      Top             =   4005
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit Time"
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
      TabIndex        =   27
      Top             =   4485
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit ID"
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
      Top             =   3525
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor ID"
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
      TabIndex        =   25
      Top             =   4965
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
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
      TabIndex        =   24
      Top             =   5445
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   840
      TabIndex        =   23
      Top             =   6405
      Width           =   1335
   End
   Begin VB.Label Label5 
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
      Left            =   840
      TabIndex        =   22
      Top             =   6885
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Charges"
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
      TabIndex        =   21
      Top             =   5925
      Width           =   1815
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "By"
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
      Left            =   6840
      TabIndex        =   9
      Top             =   2330
      Width           =   615
   End
   Begin VB.Label lblSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Doctor Visits :"
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
      Left            =   2400
      TabIndex        =   8
      Top             =   2325
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000001&
      Height          =   735
      Left            =   2280
      Top             =   2040
      Width           =   7575
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000001&
      Height          =   975
      Left            =   6120
      Top             =   6120
      Width           =   4455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   2760
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lbl_fra_Staff 
      BackStyle       =   0  'Transparent
      Caption         =   "Visits Information"
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
      TabIndex        =   7
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   720
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   360
      X2              =   360
      Y1              =   3000
      Y2              =   8640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   11520
      Y1              =   8640
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   11520
      X2              =   360
      Y1              =   8640
      Y2              =   8640
   End
End
Attribute VB_Name = "frmDoctorVisitsMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
