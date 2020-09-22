Attribute VB_Name = "mdlConnection"

'Global ADO Object Variables
Option Explicit
Public conn As ADODB.Connection
Public rsUserAccount As ADODB.Recordset
Public rsDoctorsMaintenance As ADODB.Recordset
Public rsMedicinesMaintenance As ADODB.Recordset
Public rsServicesMaintenance As ADODB.Recordset
Public rsDepartmentsMaintenance As ADODB.Recordset
Public rsCompaniesMaintenance As ADODB.Recordset
Public rsWardsMaintenance As ADODB.Recordset
Public rsWardsSelection As ADODB.Recordset
Public rsRoomsMaintenance As ADODB.Recordset
Public rsVisitTimesSchedule As ADODB.Recordset
Public rsRelevantDoctorSchedule As ADODB.Recordset
Public rsDoctorSchedule As ADODB.Recordset
Public rsInpatientMaintenance As ADODB.Recordset
Public rsGuardiansMaintenance As ADODB.Recordset
Public rsInpatientsAdmission As ADODB.Recordset
Public rsReferringDoctor As ADODB.Recordset
Public rsAssignedDoctor As ADODB.Recordset
Public rsRelevantWardsSelection As ADODB.Recordset
Public rsRoomsSelection As ADODB.Recordset
Public rsOutpatientsMaintenance As ADODB.Recordset
Public rsMedicalTreatments As ADODB.Recordset
Public rsServiceTreatments As ADODB.Recordset
Public rsMedicalTreatmentsOut As ADODB.Recordset
Public rsServiceTreatmentsOut As ADODB.Recordset
Public rsDischargeMaintenance As ADODB.Recordset
Public rsInpatientsMedicalTreatments As ADODB.Recordset
Public rsOutpatientsMedicalTreatments As ADODB.Recordset
Public rsInpatientsServiceTreatments As ADODB.Recordset
Public rsOutpatientsServiceTreatments As ADODB.Recordset
Public rsTotalMedicalTreatments As ADODB.Recordset
Public rsTotalServiceTreatments As ADODB.Recordset
Public rsInpatientBilling As ADODB.Recordset
Public rsTotalPaidSoFar As ADODB.Recordset
Public rsOutpatientBilling As ADODB.Recordset
Public rsInpatientPaymentDetails As ADODB.Recordset
Public rsOutpatientPaymentDetails As ADODB.Recordset
Public rsChannelingAppointments As ADODB.Recordset
Public rsAllAppointments As ADODB.Recordset


Public Sub Connection()
    
    'The Purpose of this function is to open a connection to link the database
    
    Set conn = New ADODB.Connection
    
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & App.Path & "\sdp.mdb;Persist Security Info=False"
    
    conn.Open
    
End Sub


Public Sub User_Account()

    'The Purpose of this function is to open the recordset "User_Account"
    
    Set rsUserAccount = New ADODB.Recordset
    
    With rsUserAccount
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from UserAccount"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub

Public Sub Doctors_Maintenance()

    'The Purpose of this function is to manage the recordset "Doctors_Maintenance"
    
    Set rsDoctorsMaintenance = New ADODB.Recordset
    
    With rsDoctorsMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Doctors_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Medicines_Maintenance()

    'The Purpose of this function is to manage the recordset "Medicines_Maintenance"
    
    Set rsMedicinesMaintenance = New ADODB.Recordset
    
    With rsMedicinesMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Medicines_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Services_Maintenance()

    'The Purpose of this function is to manage the recordset "Services_Maintenance"
    
    Set rsServicesMaintenance = New ADODB.Recordset
    
    With rsServicesMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Services_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Departments_Maintenance()

    'The Purpose of this function is to manage the recordset "Departments_Maintenance"
    
    Set rsDepartmentsMaintenance = New ADODB.Recordset
    
    With rsDepartmentsMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Departments_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Companies_Maintenance()

    'The Purpose of this function is to manage the recordset "Companies_Maintenance"
    
    Set rsCompaniesMaintenance = New ADODB.Recordset
    
    With rsCompaniesMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Companies_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub



Public Sub Wards_Maintenance()

    'The Purpose of this function is to manage the recordset "Wards_Maintenance"
    
    Set rsWardsMaintenance = New ADODB.Recordset
    
    With rsWardsMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Wards_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub



Public Sub RelevantWard_Selection()

    'The Purpose of this function is to select all the wards that belong to a particular department.
    
    Set rsWardsSelection = New ADODB.Recordset
    
    With rsWardsSelection
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Wards_Maintenance where [DepartmentID] = '" & frmRoomsMaintenance.txtDepartmentID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub



Public Sub Relevant_Ward_Selection()

    'The Purpose of this function is to select all the wards that belong to a particular department.
    
    Set rsRelevantWardsSelection = New ADODB.Recordset
    
    With rsRelevantWardsSelection
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Wards_Maintenance where [DepartmentID] = '" & frmAdmitPatient.txtDepartmentID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Rooms_Maintenance()

    'The Purpose of this function is to manage the recordset "Rooms_Maintenance"
    
    Set rsRoomsMaintenance = New ADODB.Recordset
    
    With rsRoomsMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Rooms_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub



Public Sub Rooms_Selection()

    'The Purpose of this function is to choose a room from the recordset "Rooms_Maintenance"
    
    Set rsRoomsSelection = New ADODB.Recordset
    
    With rsRoomsSelection
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Rooms_Maintenance where [WardID] = '" & frmAdmitPatient.txtWardID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub
Public Sub VisitTimes_Schedule()

    'The Purpose of this function is to manage the recordset "VisitingDoctors_VisitTimes"
    
    Set rsVisitTimesSchedule = New ADODB.Recordset
    
    With rsVisitTimesSchedule
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from VisitingDoctors_VisitTimes"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Relevant_Doctor_Schedule()

    'The Purpose of this function is to manage the recordset "VisitingDoctors_VisitTimes"
    'This function will select a particular doctor only
    
    Set rsRelevantDoctorSchedule = New ADODB.Recordset
    
    With rsRelevantDoctorSchedule
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from VisitingDoctors_VisitTimes where [DoctorID] = '" & frmDoctorScheduleMaintenance.txtDoctorID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub Doctor_Schedule()

    'The Purpose of this function is to manage the recordset "Doctors_Schedules"
    
    Set rsDoctorSchedule = New ADODB.Recordset
    
    With rsDoctorSchedule
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Doctors_Schedules"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub UserAccounts_Maintenance()

    'The Purpose of this function is to manage the recordset "UserAccount"

    Set rsUserAccount = New ADODB.Recordset

    With rsUserAccount
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from UserAccount"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub Inpatients_Maintenance()

    'The Purpose of this function is to manage the recordset "Inpatient_Maintenance"

    Set rsInpatientMaintenance = New ADODB.Recordset

    With rsInpatientMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Inpatients_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub



Public Sub Guardians_Maintenance()

    'The Purpose of this function is to manage the recordset "Guardians_Maintenance"

    Set rsGuardiansMaintenance = New ADODB.Recordset

    With rsGuardiansMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Guardians_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub



Public Sub Inpatients_Admission()

    'The Purpose of this function is to manage the recordset "Inpatients_Admission"

    Set rsInpatientsAdmission = New ADODB.Recordset

    With rsInpatientsAdmission
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Inpatients_Admission"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub ReferringDoctor_Selection()

    'The Purpose of this function is to select a Referring Doctor
    
    Set rsReferringDoctor = New ADODB.Recordset
    
    With rsReferringDoctor
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Doctors_Maintenance where [DoctorCategory] = 'Referring'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub AssignedDoctor_Selection()

    'The Purpose of this function is to select a Referring Doctor
    
    Set rsAssignedDoctor = New ADODB.Recordset
    
    With rsAssignedDoctor
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Doctors_Maintenance where [DoctorCategory] <> 'Referring'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Outpatients_Maintenance()

    'The Purpose of this function is to manage the recordset "Outpatients_Maintenance"

    Set rsOutpatientsMaintenance = New ADODB.Recordset

    With rsOutpatientsMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Outpatients_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub Medical_Treatments()

    'The Purpose of this function is to manage the recordset "Medical_Treatments"

    Set rsMedicalTreatments = New ADODB.Recordset

    With rsMedicalTreatments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Medical_Treatments"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub



Public Sub Service_Treatments()

    'The Purpose of this function is to manage the recordset "Service_Treatments"

    Set rsServiceTreatments = New ADODB.Recordset

    With rsServiceTreatments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Service_Treatments"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub Medical_Treatments_Out()

    'The Purpose of this function is to manage the recordset "Medical_Treatments_Out"

    Set rsMedicalTreatmentsOut = New ADODB.Recordset

    With rsMedicalTreatmentsOut
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Medical_Treatments_Out"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub Service_Treatments_Out()

    'The Purpose of this function is to manage the recordset "Service_Treatments_Out"

    Set rsServiceTreatmentsOut = New ADODB.Recordset

    With rsServiceTreatmentsOut
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Service_Treatments_Out"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub



Public Sub Discharge_Maintenance()

    'The Purpose of this function is to manage the recordset "Discharge_Maintenance"

    Set rsDischargeMaintenance = New ADODB.Recordset

    With rsDischargeMaintenance
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Discharge_Maintenance"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub InpatientsMedicalTreatments()

    'The Purpose of this function is to select an Inpatient's Medical Records
    
    Set rsInpatientsMedicalTreatments = New ADODB.Recordset
    
    With rsInpatientsMedicalTreatments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Medical_Treatments where [PatientID] = '" & frmAddMedicalTreatmentsIn.txtPatientID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub OutpatientsMedicalTreatments()

    'The Purpose of this function is to select an Outpatient's Medical Records
    
    Set rsOutpatientsMedicalTreatments = New ADODB.Recordset
    
    With rsOutpatientsMedicalTreatments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Medical_Treatments_Out where [PatientID] = '" & frmAddMedicalTreatmentsOut.txtPatientID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub InpatientsServiceTreatments()

    'The Purpose of this function is to select an Inpatient's Service Records
    
    Set rsInpatientsServiceTreatments = New ADODB.Recordset
    
    With rsInpatientsServiceTreatments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Service_Treatments where [PatientID] = '" & frmAddServiceTreatmentsIn.txtPatientID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub OutpatientsServiceTreatments()

    'The Purpose of this function is to select an Outpatient's Service Records
    
    Set rsOutpatientsServiceTreatments = New ADODB.Recordset
    
    With rsOutpatientsServiceTreatments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Service_Treatments_Out where [PatientID] = '" & frmAddServiceTreatmentsOut.txtPatientID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub TotalMedicalTreatments()

    'The Purpose of this function is to get the total charges of all medical treatments
    
    Set rsTotalMedicalTreatments = New ADODB.Recordset
    
    With rsTotalMedicalTreatments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select SUM(UnitPrice) as 'Sum' from Medical_Treatments where [PatientID] = '" & frmIPDOverallBilling.txtPatientID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub TotalServiceTreatments()

    'The Purpose of this function is to get the total charges of all medical treatments
    
    Set rsTotalServiceTreatments = New ADODB.Recordset
    
    With rsTotalServiceTreatments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select SUM(ServiceCharge) as 'Sum' from Service_Treatments where [PatientID] = '" & frmIPDOverallBilling.txtPatientID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Inpatient_Billing()

    'The Purpose of this function is to manage the recordset "Inpatient_Billing"

    Set rsInpatientBilling = New ADODB.Recordset

    With rsInpatientBilling
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Inpatient_Billing"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub TotalPaidSoFar()

    'The Purpose of this function is to get the total amount paid by the patient todate
    
    Set rsTotalPaidSoFar = New ADODB.Recordset
    
    With rsTotalPaidSoFar
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select SUM(AmountPaid) as 'Sum' from Inpatient_Billing where [PatientID] = '" & frmIPDOverallBilling.txtPatientID.Text & "'"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub


Public Sub Outpatient_Billing()

    'The Purpose of this function is to manage the recordset "Outpatient_Billing"

    Set rsOutpatientBilling = New ADODB.Recordset

    With rsOutpatientBilling
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Outpatient_Billing"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub Inpatient_Payment_Details()

    'The Purpose of this function is to manage the recordset "Inpatient_Payment_Details"

    Set rsInpatientPaymentDetails = New ADODB.Recordset

    With rsInpatientPaymentDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Inpatient_Payment_Details"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub Outpatient_Payment_Details()

    'The Purpose of this function is to manage the recordset "Outpatient_Payment_Details"

    Set rsOutpatientPaymentDetails = New ADODB.Recordset

    With rsOutpatientPaymentDetails
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Outpatient_Payment_Details"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub Channeling_Appointments()

    'The Purpose of this function is to manage the recordset "Channeling_Appointments" with a certain query to get selected information

    Set rsChannelingAppointments = New ADODB.Recordset

    With rsChannelingAppointments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "SELECT DoctorID, TokenNo, AppointmentStartTime, AppointmentEndTime,  FirstName, LastName, ContactNo, ChosenDate FROM Channeling_Appointments WHERE (DoctorID = '" & frmChannelingAppointments.txtDoctorID.Text & "')"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub


Public Sub All_Appointments()

    'The Purpose of this function is to manage the recordset "Channeling_Appointments"

    Set rsAllAppointments = New ADODB.Recordset

    With rsAllAppointments
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .Source = "Select * from Channeling_Appointments"
        .CursorLocation = adUseClient
        .Open
    End With
    
End Sub
