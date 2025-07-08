VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nairobi Cardiovascular Clinic Management System"
   ClientHeight    =   9045
   ClientLeft      =   4410
   ClientTop       =   2070
   ClientWidth     =   15300
   Icon            =   "MDIForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm.frx":0442
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10560
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":7F1C
            Key             =   "write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":836E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":87C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":8C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":9064
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":94B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":9908
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":9D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":A074
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":A4C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":A918
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":AD6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":B1BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":B60E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":BA60
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm.frx":BEB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   1588
      ButtonWidth     =   1879
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Patients"
            Key             =   "write"
            Object.Tag             =   "1"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Observation"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Doctors"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pharmacy"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cashier"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Laboratory"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Security"
            Description     =   "Security Module"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DashBoard"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Wards"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            Object.Tag             =   "2"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            ImageIndex      =   11
         EndProperty
      EndProperty
      MouseIcon       =   "MDIForm.frx":C304
   End
   Begin MSComctlLib.StatusBar SSBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8670
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7840
            MinWidth        =   7840
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   15677
            MinWidth        =   15677
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIForm.frx":C756
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuReception 
      Caption         =   "Reception"
      Begin VB.Menu MnuPatientDetails 
         Caption         =   "Patient Details"
      End
      Begin VB.Menu MnuBillingCompanies 
         Caption         =   "Billing Companies"
      End
      Begin VB.Menu mnuConsultationFee 
         Caption         =   "Consultation Fee"
      End
   End
   Begin VB.Menu MnuTreatment 
      Caption         =   "Pre Treatment"
      Index           =   1
      Begin VB.Menu MnuObservation 
         Caption         =   "Nurse Observation"
      End
      Begin VB.Menu mnuWaitingRoom 
         Caption         =   "Waiting Room"
      End
      Begin VB.Menu MnuHistory 
         Caption         =   "Patient History"
      End
      Begin VB.Menu MnuSearch 
         Caption         =   "Search Engine"
      End
      Begin VB.Menu MnuDiagnosis 
         Caption         =   "Pending Diagnosis"
      End
   End
   Begin VB.Menu MnuPostTreatment 
      Caption         =   "Post Treatment"
      Begin VB.Menu MnuPharmacy 
         Caption         =   "Pharmacy"
      End
      Begin VB.Menu MnuLaboratory 
         Caption         =   "Laboratory"
      End
      Begin VB.Menu MnuAdmission 
         Caption         =   "Admission"
      End
      Begin VB.Menu MnuReferrals 
         Caption         =   "Referrals"
      End
      Begin VB.Menu MnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScanSave 
         Caption         =   "SCAN and SAVE"
      End
   End
   Begin VB.Menu mnuSms 
      Caption         =   "SMS Module"
      Begin VB.Menu mnuListMaintenance 
         Caption         =   "Sms Contact List Maintenance"
      End
      Begin VB.Menu mnuBulkSmS 
         Caption         =   "Bulk SMS"
      End
   End
   Begin VB.Menu MnuCashier 
      Caption         =   "Cashier"
      Begin VB.Menu MnuPayments 
         Caption         =   "Payments"
      End
      Begin VB.Menu MnuCreditPayments 
         Caption         =   "Credit Payments"
      End
   End
   Begin VB.Menu MnuReports 
      Caption         =   "Reports"
      Begin VB.Menu MnuListOfReports 
         Caption         =   "General Reports"
      End
      Begin VB.Menu MnuFinancial 
         Caption         =   "Financial Reports"
      End
      Begin VB.Menu MnuStockUpdate 
         Caption         =   "Stock Update Audit"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuMaintenace 
      Caption         =   "Maintenance"
      Begin VB.Menu MnuAilmentsCategory 
         Caption         =   "Ailments Category Definition"
         Index           =   4
      End
      Begin VB.Menu MnuAilments 
         Caption         =   "Ailments List Maintenance"
      End
      Begin VB.Menu MNULAINI 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCategory 
         Caption         =   "Medicine Category Definition"
      End
      Begin VB.Menu MnuNewMedicine 
         Caption         =   "New Medicine Maintenance"
      End
      Begin VB.Menu Mnuline6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuStock 
         Caption         =   "Medicine Stock Maintenance"
      End
      Begin VB.Menu MnuNightShift 
         Caption         =   "Night Shift Stock Maintenance"
      End
      Begin VB.Menu MnuStockTake 
         Caption         =   "Stock Take Reconciliation"
      End
      Begin VB.Menu MnuLine2 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu MnuSecurity 
         Caption         =   "Security Definition"
      End
      Begin VB.Menu MnuChangePass 
         Caption         =   "Change Own Password"
      End
      Begin VB.Menu MnuParameters 
         Caption         =   "System Global Parameters"
      End
      Begin VB.Menu MnuLine8 
         Caption         =   "-"
      End
      Begin VB.Menu MnuUnlockRecord 
         Caption         =   "Unlock Queue Record"
      End
      Begin VB.Menu Mnuline7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAudit 
         Caption         =   "Audit Trail"
      End
   End
   Begin VB.Menu MnuEODMenu 
      Caption         =   "End Of Day"
      Begin VB.Menu MnuEOD 
         Caption         =   "End of Day Process"
      End
      Begin VB.Menu mnuline3 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuShiftChange 
         Caption         =   "Day/Night Shift Change"
      End
   End
   Begin VB.Menu MnuTerminate 
      Caption         =   "Terminate Session"
      Begin VB.Menu MnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "Close Application"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub jsButtonmenubar_ItemClick(ByVal ItemBar As JSBtnBar16.JSGroupItem)

    Select Case ItemBar.Caption
        Case "Patient Details"
            FrmPatients.Show
        Case "Billing Companies"
            FrmBillingCompanies.Show
        Case "Consultation Fees"
            FrmConsultation.Show
        Case "Nurse Observation"
            FrmObservation.Show
        Case "Doctors Waiting Room"
            FrmWaitingRoom.Show
        Case "Patient History"
            FrmHistory.Show
        Case "Pharmacy"
            FrmPharmacy.Show
        Case "Laboratory"
            FrmLaboratoryWaiting.Show
        Case "Referral"
            FrmReferrals.Show
        Case "Admission"
            FrmAdmission.Show
    End Select
    
End Sub

Private Sub MDIForm_Load()
    GlbCurrentForm = EnumConsultation
    SSBar.Panels(1).Text = Format(GlbSysDate, "DD MMMM YYYY") & "           " & "Current User = " & UCase(GlbCurrentUser)
    SSBar.Panels(4).Text = Format(GlbSysDate, "DD MMMM YYYY")
    If Format(GlbSysDate, "ddmmmyyyy") <> Format(Date, "ddmmmyyyy") Then
        Beep
        FrmDateDiff.Show 1
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    MnuClose_Click
End Sub

Private Sub mnuAbout_Click()
    FrmAbout.Show 1
End Sub

Private Sub MnuAdmission_Click()
    FrmWard.Show
End Sub

Private Sub MnuAilments_Click()
    If VerifyAccess(GlbCurrentUser, "Doctors") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmAilments.Show
End Sub

Private Sub MnuAilmentsCategory_Click(Index As Integer)
    If VerifyAccess(GlbCurrentUser, "Doctors") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmAilmentCategory.Show
End Sub

Private Sub MnuAudit_Click()
    If VerifyAccess(GlbCurrentUser, "System Administrator") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmAudit.Show 1
End Sub

Private Sub MnuBillingCompanies_Click()
    If VerifyAccess(GlbCurrentUser, "System Administrator") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmBillingCompanies.Show
End Sub

Private Sub MnuCategory_Click()
    FrmCategory.Show
End Sub

Private Sub MnuChangePass_Click()
    FrmConfirmPassword.Show
End Sub

Private Sub MnuClose_Click()
    
    Dim Resp
    Resp = MsgBox("This will terminate your Session. Are you sure you want to Terminate?", vbQuestion + vbYesNo)
        If Resp = vbYes Then
        DoEvents
            On Error Resume Next
            Kill App.Path & "\" & App.EXEName & ".exe"
            Kill App.Path & "\Siloam Medical System.exe"
            End
        Else
            Exit Sub
        End If
End Sub

Private Sub mnuConsultationFee_Click()
    If VerifyAccess(GlbCurrentUser, "System Administrator") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmConsultation.Show
End Sub

Private Sub MnuCreditPayments_Click()
    If VerifyAccess(GlbCurrentUser, "Cashier") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmCreditPayment.Show
End Sub

Private Sub MnuDiagnosis_Click()
    If VerifyAccess(GlbCurrentUser, "Doctors") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmTreatment.Show
End Sub

Private Sub MnuEOD_Click()
    If VerifyAccess(GlbCurrentUser, "System Administrator") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmEOD.Show
End Sub

Private Sub MnuHistory_Click()
    If VerifyAccess(GlbCurrentUser, "Doctors") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmHistory.Show
End Sub

Private Sub MnuLaboratory_Click()
    If VerifyAccess(GlbCurrentUser, "Laboratory") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmLaboratoryWaiting.Show
End Sub

Private Sub MnuListOfReports_Click()
    If VerifyAccess(GlbCurrentUser, "Reports") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmReports.Show
End Sub

Private Sub MnuLogOff_Click()
    Unload MDIMain
    frmLogin.Show
End Sub

Private Sub MnuNewMedicine_Click()
    If VerifyAccess(GlbCurrentUser, "Pharmacy") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmNewMedicine.Show
End Sub

Private Sub MnuNightShift_Click()
    If VerifyAccess(GlbCurrentUser, "Pharmacy") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmNightShiftStock.Show
End Sub

Private Sub MnuObservation_Click()
    If VerifyAccess(GlbCurrentUser, "Observation") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmObservation.Show
End Sub

Private Sub MnuParameters_Click()
    If VerifyAccess(GlbCurrentUser, "System Administrator") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmParmeters.Show
End Sub

Private Sub MnuPatientDetails_Click()
    If VerifyAccess(GlbCurrentUser, "Patients") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmPatients.Show
End Sub

Private Sub MnuPayments_Click()
    If VerifyAccess(GlbCurrentUser, "Cashier") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmCashier.Show
End Sub

Private Sub MnuPharmacy_Click()
    If VerifyAccess(GlbCurrentUser, "Pharmacy") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmPharmacy.Show
End Sub

Private Sub MnuReferrals_Click()
    FrmReferrals.Show
End Sub

Private Sub mnuScanSave_Click()
    FrmConverter.Show
End Sub

Private Sub MnuSearch_Click()
    FrmSearchEngine.Show
End Sub

Private Sub MnuSecurity_Click()
    If VerifyAccess(GlbCurrentUser, "System Administrator") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmSecurity.Show
End Sub

Private Sub mnuShiftChange_Click()
    FrmShifChange.Show
End Sub

Private Sub MnuStock_Click()
    If VerifyAccess(GlbCurrentUser, "Pharmacy") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmProducts.Show
End Sub

Private Sub MnuStockTake_Click()
    FrmStockReconciliation.Show
End Sub

Private Sub MnuStockUpdate_Click()
    If VerifyAccess(GlbCurrentUser, "System Administrator") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmStockAudit.Show
End Sub

Private Sub MnuUnlockRecord_Click()
    FrmUnlock.Show 1
End Sub

Private Sub mnuWaitingRoom_Click()
    If VerifyAccess(GlbCurrentUser, "Doctors") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
    FrmWaitingRoom.Show
End Sub

Private Sub Toolbar_ButtonClick(ByVal ToolButton As MSComctlLib.Button)
On Error GoTo Errorhandler
Dim DUMMYFORM As Form  'I dont think I use this. delete if you dare
    If GlbCurrentForm = 0 Then Exit Sub
    Select Case ToolButton.Index
        Case 1
            If VerifyAccess(GlbCurrentUser, "Patients") = False Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, EnumConsultation)
        Case 2
            If VerifyAccess(GlbCurrentUser, "Observation") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, EnumObservation)
        Case 3
            If VerifyAccess(GlbCurrentUser, "Doctors") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, EnumDoctors)
        Case 4
            If VerifyAccess(GlbCurrentUser, "Pharmacy") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, EnumPharmacy)
        Case 5
            If VerifyAccess(GlbCurrentUser, "Cashier") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, EnumCashier)
        Case 6
            If VerifyAccess(GlbCurrentUser, "Laboratory") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, EnumLab)
        Case 7
            If VerifyAccess(GlbCurrentUser, "System Administrator") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, 7)
        Case 8
            'If VerifyAccess(GlbCurrentUser, "System Administrator") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, 10)
        Case 9
            If VerifyAccess(GlbCurrentUser, "Patients") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, 11)
        Case 10
            If VerifyAccess(GlbCurrentUser, "Reports") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            DMY = SwitchScreen(GlbCurrentForm, 9)
        Case 11
            FrmSearchEngine.Show
        Case 12
            MnuClose_Click
    End Select
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
End Sub
