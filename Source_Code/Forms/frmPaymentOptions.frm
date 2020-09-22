VERSION 5.00
Begin VB.Form frmPaymentOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Select Patient Type"
   ClientHeight    =   3450
   ClientLeft      =   5100
   ClientTop       =   4065
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Close"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton cmdOutpatient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Outpatient"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton cmdInpatient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Inpatient"
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblSelectOption 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Patient Type"
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
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   195
      Width           =   2655
   End
   Begin VB.Image imgCenter 
      Height          =   600
      Index           =   0
      Left            =   0
      Picture         =   "frmPaymentOptions.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5370
   End
   Begin VB.Image imgbg2 
      Height          =   3105
      Index           =   0
      Left            =   0
      Picture         =   "frmPaymentOptions.frx":00A2
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5370
   End
   Begin VB.Label Label12 
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
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   8520
      Width           =   3735
   End
   Begin VB.Label Label12 
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
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   8520
      Width           =   3735
   End
End
Attribute VB_Name = "frmPaymentOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdInpatient_Click()
    frmIPDOverallBilling.Show
    Unload Me
End Sub


Private Sub cmdOutpatient_Click()
    frmOPDOverallBilling.Show
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
