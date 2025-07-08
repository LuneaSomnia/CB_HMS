VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Out Patient System"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10770
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnuReception 
      Caption         =   "Reception"
      Begin VB.Menu MnuPatientDetails 
         Caption         =   "Patient Details"
      End
      Begin VB.Menu MnuBillingCompanies 
         Caption         =   "Billing Companies"
      End
   End
   Begin VB.Menu MnuTreatment 
      Caption         =   "Treatment"
      Index           =   1
      Begin VB.Menu MnuObservation 
         Caption         =   "Nurse Observation"
      End
      Begin VB.Menu mnuWaitingRoom 
         Caption         =   "Waiting Room"
      End
      Begin VB.Menu MnuDiagnosis 
         Caption         =   "Pending Diagnosis"
      End
   End
   Begin VB.Menu MnuAttendance 
      Caption         =   "Attendance"
      Begin VB.Menu MnuPharmacy 
         Caption         =   "Pharmacy"
      End
      Begin VB.Menu MnuLaboratory 
         Caption         =   "Laboratory"
      End
      Begin VB.Menu MnuAdmission 
         Caption         =   "Admission"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MnuReception_Click()

End Sub
