VERSION 5.00
Begin VB.Form frmReportChannelingMaster 
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Please Choose The Required Channeling Day"
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cbobillstat 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Please Choose The Channeling Day"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmReportChannelingMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cbobillstat.AddItem ("Sunday")
    cbobillstat.AddItem ("Monday")
    cbobillstat.AddItem ("Tuesday")
    cbobillstat.AddItem ("Wednesday")
    cbobillstat.AddItem ("Thursday")
    cbobillstat.AddItem ("Friday")
    cbobillstat.AddItem ("Saturday")
End Sub

Private Sub cmdOK_Click()
    On Error GoTo e
        
        DataEnvironment1.Commands("ChannelingMaster").Parameters(0) = cbobillstat
        With RptChannelingMaster

            .Sections("Section2").Controls("lblbillstat").Caption = cbobillstat
            .Show
        End With
        DataEnvironment1.rsChannelingMaster.Close
        Unload Me
    Exit Sub
e:
    If Err.Number <> 3704 Then
        MsgBox Err.Description, vbCritical
    End If
End Sub


