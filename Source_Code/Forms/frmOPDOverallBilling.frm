VERSION 5.00
Begin VB.Form frmOPDOverallBilling 
   Caption         =   "Overall Patient Billing Details Module"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmOPDOverallBilling.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   11835
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox txtOverallOutBillID 
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
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdGoToPayments 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Go To Payments Form"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   2535
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   2535
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtAssignedDoctorID 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   7320
      Width           =   1815
   End
   Begin VB.TextBox txtPatientID 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdOutpatientSearchWizard 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      ToolTipText     =   "Click Here to select an Outpatient"
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtSurname 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtFirstName 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2880
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtDoctorsCharges 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox txtHospitalCharges 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox txtVAT 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox txtNettTotal 
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
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7800
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
      Index           =   4
      Left            =   7080
      TabIndex        =   33
      Top             =   7920
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
      Left            =   7080
      TabIndex        =   32
      Top             =   6960
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
      Left            =   7080
      TabIndex        =   31
      Top             =   6480
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
      Left            =   7080
      TabIndex        =   30
      Top             =   6000
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
      Left            =   7080
      TabIndex        =   29
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label lblAssignedDoctorID 
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned Doctor ID"
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
      Left            =   2280
      TabIndex        =   28
      Top             =   4965
      Width           =   2055
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
      Left            =   2400
      TabIndex        =   27
      Top             =   2445
      Width           =   1095
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   1920
      Y1              =   4680
      Y2              =   8280
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   7800
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   7800
      X2              =   7800
      Y1              =   4680
      Y2              =   8280
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   3840
      X2              =   7800
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   2160
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lblFrameName 
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
      Left            =   2280
      TabIndex        =   26
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label lblSurname 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
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
      TabIndex        =   25
      Top             =   3405
      Width           =   1335
   End
   Begin VB.Label lblFirstName 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      TabIndex        =   24
      Top             =   2925
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   1920
      Y1              =   2160
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   7800
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   7800
      X2              =   7800
      Y1              =   2160
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   4200
      X2              =   7800
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   1920
      X2              =   2280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblFrame 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Details"
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
      Left            =   2400
      TabIndex        =   23
      Top             =   2040
      Width           =   2535
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
      Left            =   2400
      TabIndex        =   22
      Top             =   3765
      Width           =   1335
   End
   Begin VB.Label lblDoctorsCharges 
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
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      Top             =   5445
      Width           =   2055
   End
   Begin VB.Label lblHospitalCharges 
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Charges"
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
      Left            =   2280
      TabIndex        =   20
      Top             =   5925
      Width           =   1575
   End
   Begin VB.Label lblVAT 
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
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
      Left            =   2280
      TabIndex        =   19
      Top             =   6885
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   2280
      TabIndex        =   18
      Top             =   6405
      Width           =   1575
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
      Left            =   2280
      TabIndex        =   17
      Top             =   7365
      Width           =   1575
   End
   Begin VB.Label lblNettTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "NETT TOTAL"
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
      Left            =   2280
      TabIndex        =   16
      Top             =   7845
      Width           =   1575
   End
   Begin VB.Label Label24 
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
      Left            =   6600
      TabIndex        =   15
      Top             =   7365
      Width           =   375
   End
End
Attribute VB_Name = "frmOPDOverallBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Outpatient Overall Billing Interface
'Programmer: Isham Sally
'Quality Assurance Engineer (Testing): Imran Sheriff
'Start Date: 11/05/08
'Date Of Last Modification: 11/05/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed:
'-----------------------------------------------------------------------------

Option Explicit

'The following variables will be used to autogenerate the Invoice ID
Dim iNumOfRecords As Integer    'This variable holds the number of records in the table
Dim strCode As String   'This variable will eventually hold the Invoice ID to be autogenerated
Dim iNumberOfRecords As Integer 'This variable will hold the number of records in the Inpatient_Payment_Details table
Dim strDisplay As String    'This variable will eventually display the OverallInBillID in the invisible textfield
Dim iTotalPayable As Double    'This variable will hold the value of the Total Payable
Dim billID As String    'This variable will hold the Billing ID for the purpose of report generation.



Private Sub cmdClose_Click()
    
    If MsgBox(userName & ", Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdGoToPayments_Click()
    
    
    'Ensuring that the user has selected a patient
    If txtPatientID.Text = "" Then
        MsgBox "You Cannot Go To The Payments Process Until You Select A Patient!", vbCritical, "Please Select A Patient!"
        Exit Sub
    End If
        
    Call Outpatient_Billing
    With rsOutpatientBilling
        .MoveFirst
        Do While .EOF = False
            If .Fields(2).Value = txtPatientID.Text Then
                MsgBox "This Patient Has Already Settled The Bill!", vbCritical, "Bill Already Settled!"
                Exit Sub
            Else
                .MoveNext
            End If
        Loop
    End With
    
           
    Call Outpatient_Billing    'Calling the Outpatient_Billing Procedure to interact with the recordset
    
    'Generate Invoice ID By Utilizing the Outpatient_Billing Table
    With rsOutpatientBilling
        
        
        If .RecordCount = 0 Then    'If there are no records in the table
            
            strCode = "OIN0001"
        
        Else
            
            'Calculating the number of records and storing in a variable
            iNumOfRecords = .RecordCount
            iNumOfRecords = iNumOfRecords + 1   'incrementing the number by 1
            
            'The following block of code will generate the ID according
            'to the number of records in the Outpatient_Billing Table
            If iNumOfRecords < 10 Then
                strCode = "OIN000" & iNumOfRecords
            ElseIf iNumOfRecords < 100 Then
                strCode = "OIN00" & iNumOfRecords
            ElseIf iNumOfRecords < 1000 Then
                strCode = "OIN0" & iNumOfRecords
            ElseIf iNumOfRecords < 10000 Then
                strCode = "OIN" & iNumOfRecords
            End If
            
        End If
        
        .Requery    'Requerying the Table
        
        .AddNew     'Adding a new recordset
        
    End With
    
    
    iTotalPayable = Int(Val(txtNettTotal.Text))   'Storing the Nett Total in this variable
    
    'The following line of code will enter the autogenerated Invoice ID into the relevant textfield
    frmOPDCreateBill.txtInvoiceID.Text = strCode
    
    'Entering all relevant data onto the Payment Form
    frmOPDCreateBill.txtBillingDate.Text = DateTime.Date    'System Date
    frmOPDCreateBill.txtPatientID.Text = txtPatientID.Text
    frmOPDCreateBill.txtPatientName.Text = txtFirstName.Text & " " & txtSurname.Text
    frmOPDCreateBill.txtAccountType.Text = txtAccountType.Text
    frmOPDCreateBill.txtTotalCost.Text = txtTotal.Text
    frmOPDCreateBill.txtDiscount.Text = txtDiscount.Text
    frmOPDCreateBill.txtTotalPayable.Text = iTotalPayable
    

    
        
    
    Unload Me   'Closing this form
    
    frmOPDCreateBill.Show   'Opening Up The Payments Form
    
    Exit Sub
    
    
End Sub

Private Sub cmdOutpatientSearchWizard_Click()
    
    frmOutpatientSearchBilling.Show
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo e
        DataEnvironment1.Commands("OutpatientInvoice").Parameters(0) = billID   'Passing the value of the variable as a parameter.
        RptOutpatientInvoice.Show
        DataEnvironment1.rsOutpatientInvoice.Close
        
        Unload Me
        Exit Sub
e:
        If Err.Number <> 3704 Then
            MsgBox Err.Description & "" & Err.Number, vbCritical
        End If

End Sub

Private Sub cmdSave_Click()

    'Ensuring that the user has selected a patient
    If txtPatientID.Text = "" Then
        MsgBox "Error! You Have Not Selected A Patient!", vbCritical, "Please Select A Patient!"
        Exit Sub
    End If
    
    Call Outpatient_Payment_Details    'Calling the Outpatient_Payment_Details Procedure to interact with the recordset

    'Generate OverallOutBillID By Utilizing the Outpatient_Payment_Details Table
    With rsOutpatientPaymentDetails

        If .RecordCount = 0 Then    'If there are no records in the table

            strDisplay = "OPD0001"

        Else

            'Calculating the number of records and storing in a variable
            iNumberOfRecords = .RecordCount
            iNumberOfRecords = iNumberOfRecords + 1   'incrementing the number by 1

            'The following block of code will generate the ID according
            'to the number of records in the Outpatient_Payment_Details Table
            If iNumberOfRecords < 10 Then
                strDisplay = "OPD000" & iNumberOfRecords
            ElseIf iNumberOfRecords < 100 Then
                strDisplay = "OPD00" & iNumberOfRecords
            ElseIf iNumberOfRecords < 1000 Then
                strDisplay = "OPD0" & iNumberOfRecords
            ElseIf iNumberOfRecords < 10000 Then
                strDisplay = "OPD" & iNumberOfRecords
            End If

        End If

        .Requery    'Requerying the Table

        .AddNew     'Adding a new recordset

    End With

    'The following line of code will enter the autogenerated OverallOutBillID
    'into the invisible OverallOutBillID textfield
    txtOverallOutBillID.Text = strDisplay
    
    'Here, I am ensuring that the Discount textfield is not empty when I save
    If txtDiscount.Text = "" Then
        txtDiscount.Text = "-"
    End If
    
    'Now I am going to save the record in the database
    With rsOutpatientPaymentDetails
    
        'Making sure that the user wants to save the record
        If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
        
            .Fields(0) = txtOverallOutBillID.Text
            billID = txtOverallOutBillID.Text   'Storing values in a variable
            .Fields(1) = txtPatientID.Text
            .Fields(2) = txtFirstName.Text
            .Fields(3) = txtSurname.Text
            .Fields(4) = txtAccountType.Text
            .Fields(5) = txtAssignedDoctorID.Text
            .Fields(6) = txtDoctorsCharges.Text
            .Fields(7) = txtHospitalCharges.Text
            .Fields(8) = txtTotal.Text
            .Fields(9) = txtVAT.Text
            .Fields(10) = txtDiscount.Text
            .Fields(11) = txtNettTotal.Text
            
            .Update
            
            'Display Success Message
            MsgBox "The Record Was Saved Successfully!", vbInformation, "Succesful Save Procedure!"
                        
        Else
            
            'Display 'No Modifications' Message
            MsgBox "No Modifications Have Taken Place!", vbInformation, "No Modifications!"
                
            .CancelUpdate   'Cancel the Save Procedure
            
        End If

    End With
    
    cmdPrint.Enabled = True 'Enabling the Print button

End Sub
