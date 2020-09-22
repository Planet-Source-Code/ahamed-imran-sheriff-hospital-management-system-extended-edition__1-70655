VERSION 5.00
Begin VB.Form frmDoctorSpecializationWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Specialization Selection Wizard"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstSpecializations 
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
      Height          =   6075
      ItemData        =   "frmDoctorSpecializationWizard.frx":0000
      Left            =   1200
      List            =   "frmDoctorSpecializationWizard.frx":0061
      TabIndex        =   5
      Top             =   1200
      Width           =   6255
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Cancel"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&OK"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label lblWizardHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Specialization Selection Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblWizardFooter 
      BackStyle       =   0  'Transparent
      Caption         =   "Durdans Hospital Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   0
      Left            =   0
      Picture         =   "frmDoctorSpecializationWizard.frx":023B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Schedule Setup Wizard"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image imgbg2 
      Height          =   8865
      Index           =   0
      Left            =   0
      Picture         =   "frmDoctorSpecializationWizard.frx":02DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Image imgCenter 
      Height          =   840
      Index           =   2
      Left            =   0
      Picture         =   "frmDoctorSpecializationWizard.frx":037B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9810
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   9360
      X2              =   9360
      Y1              =   2160
      Y2              =   7320
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   2160
      Y2              =   7320
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   360
      X2              =   9360
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line7 
      Index           =   0
      X1              =   360
      X2              =   9360
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H80000006&
      Height          =   735
      Left            =   1080
      Top             =   1200
      Width           =   7455
   End
End
Attribute VB_Name = "frmDoctorSpecializationWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    frmDoctorsMaintenance.txtDoctorSpecialization.Text = lstSpecializations.Text
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    cmdOK.Enabled = False   'Disabling the OK Button
    
End Sub


Private Sub lstSpecializations_Click()

    cmdOK.Enabled = True
    
End Sub
