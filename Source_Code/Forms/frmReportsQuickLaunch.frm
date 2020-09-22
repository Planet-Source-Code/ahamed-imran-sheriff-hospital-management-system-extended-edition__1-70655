VERSION 5.00
Begin VB.Form frmReportsQuickLaunch 
   Caption         =   "Reports Quick Launch"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmReportsQuickLaunch.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8280
      Width           =   2175
   End
   Begin VB.CommandButton cmdOutpatientsServiceTreatmentsReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Outpatients Service Treatments Report"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   3975
   End
   Begin VB.CommandButton cmdInpatientsServiceTreatmentsReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Inpatients Service Treatments Report"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   3975
   End
   Begin VB.CommandButton cmdOutpatientsMedicalTreatmentsReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Outpatients Medical Treatments Report"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   3975
   End
   Begin VB.CommandButton cmdInpatientsMedicalTreatmentsReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Inpatients Medical Treatments Report"
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   3975
   End
   Begin VB.CommandButton cmdBillStatusReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Bill Status Report"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   3975
   End
   Begin VB.CommandButton cmdOutpatientsRevenueReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Outpatients Revenue Report"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   3975
   End
   Begin VB.CommandButton cmdInpatientsRevenueReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Inpatients Revenue Report"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   3975
   End
   Begin VB.CommandButton cmdInpatientsInvoice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Inpatients Invoice"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Patient Discharge Master Report"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   3975
   End
   Begin VB.CommandButton cmdPatientAdmissionMasterReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Patient Admission Master Report"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Channeling Patients Master Report"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton cmdOutpatientsMasterReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Outpatients Master Report"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   3975
   End
   Begin VB.CommandButton cmdInpatientsMasterReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Inpatients Master Report"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000001&
      X1              =   11040
      X2              =   11040
      Y1              =   5160
      Y2              =   7920
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000001&
      X1              =   6120
      X2              =   6120
      Y1              =   5160
      Y2              =   7920
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000001&
      X1              =   6120
      X2              =   11040
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000001&
      X1              =   8520
      X2              =   11040
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000001&
      X1              =   6120
      X2              =   6480
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue Reports"
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
      Left            =   6600
      TabIndex        =   15
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   5520
      X2              =   5520
      Y1              =   5160
      Y2              =   7920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   600
      Y1              =   5160
      Y2              =   7920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   5520
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   3240
      X2              =   5520
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   600
      X2              =   960
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Treatments Reports"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   8400
      X2              =   8400
      Y1              =   2040
      Y2              =   4800
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   3480
      X2              =   3480
      Y1              =   2040
      Y2              =   4800
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   3480
      X2              =   8400
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   5640
      X2              =   8400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   3480
      X2              =   3840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Reports"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "frmReportsQuickLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBillStatusReport_Click()
    frmReportAging.Show
End Sub

Private Sub cmdClose_Click()
    
    'Obtaining confirmation from the user
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdInpatientsInvoice_Click()
    frmReportInpatientInvoice.Show
End Sub

Private Sub cmdInpatientsMasterReport_Click()
    frmReportInpatientMaster.Show
End Sub

Private Sub cmdInpatientsMedicalTreatmentsReport_Click()
    frmReportInpatientMedicalTreatment.Show
End Sub

Private Sub cmdInpatientsRevenueReport_Click()
    frmReportInpatientRevenue.Show
End Sub

Private Sub cmdInpatientsServiceTreatmentsReport_Click()
    frmReportInpatientServiceTreatments.Show
End Sub

Private Sub cmdOutpatientsMasterReport_Click()
    frmReportOutpatientMaster.Show
End Sub

Private Sub cmdOutpatientsMedicalTreatmentsReport_Click()
    frmReportOutpatientMedicalTreatments.Show
End Sub

Private Sub cmdOutpatientsRevenueReport_Click()
    frmReportOutpatientRevenue.Show
End Sub

Private Sub cmdOutpatientsServiceTreatmentsReport_Click()
    frmReportOutPatientPatientServiceTreatements.Show
End Sub

Private Sub cmdPatientAdmissionMasterReport_Click()
    frmReportPatientAdmission.Show
End Sub

Private Sub Command1_Click()
    frmReportChannelingMaster.Show
End Sub

Private Sub Command2_Click()
    frmReportPatientDischarge.Show
End Sub
