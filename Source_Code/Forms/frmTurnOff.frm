VERSION 5.00
Begin VB.Form frmTurnOff 
   BorderStyle     =   0  'None
   Caption         =   "Turn Off System?"
   ClientHeight    =   3030
   ClientLeft      =   4455
   ClientTop       =   3585
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTurnOff.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdTurnOff 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Height          =   778
      Left            =   2640
      MaskColor       =   &H8000000D&
      Picture         =   "frmTurnOff.frx":A6DA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   800
   End
   Begin VB.CommandButton cmdLogOff 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Height          =   778
      Left            =   1200
      MaskColor       =   &H8000000D&
      Picture         =   "frmTurnOff.frx":C372
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   800
   End
   Begin VB.Label lblTurnOff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Off"
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
      Left            =   2310
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Log Off / Turn Off"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   200
      Width           =   3015
   End
   Begin VB.Label lblLogOff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Log Off"
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
      Left            =   880
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "frmTurnOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLogOff_Click()
    frmLogin.Show
    Unload Me
    Unload frmMDI
End Sub

Private Sub cmdTurnoff_Click()
    End
End Sub

