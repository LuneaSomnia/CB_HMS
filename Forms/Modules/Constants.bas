Attribute VB_Name = "Constants"
Global GlbCardNumber As String
Global GlbVisitDate As String
Global GlbCurrentForm As Integer
Global GblServiceProviderName
Global GblServiceProviderID
Global GlbCalledFromPatients As Boolean
Global GlbCalledFromAppointments As Boolean
Global GlbTestDestinationLocation As String
Global GlbTestSourceLocation As String
Global ConnectionPassword As String
Global GlbCurrentUser As String
Public Enum SendModule
    EnumConsultation = 1
    EnumObservation = 2
    EnumDoctors = 3
    EnumPharmacy = 4
    EnumCashier = 5
    EnumLab = 6
    EnumSecurity = 7
    EnumPatientHistory = 8
    'EnumAdmission = 8
    EnumReports = 9
    EnumDashBoard = 10
    EnumWard = 11
    EnumStockInventory = 12
End Enum
