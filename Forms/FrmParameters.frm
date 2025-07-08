VERSION 5.00
Begin VB.Form FrmParmeters 
   Caption         =   "System Parameter Settings"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   Icon            =   "FrmParameters.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Day or Night Shift"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   21
      Top             =   4560
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmParameters.frx":0442
         Left            =   3240
         List            =   "FrmParameters.frx":044C
         TabIndex        =   23
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Current Shift"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nurse Room  && Doctors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   17
      Top             =   1920
      Width           =   6015
      Begin VB.CheckBox ChkNurseDoc 
         Caption         =   "Doctors Peform Nurses Functions.  (Nurse Room is not Present.)"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   11415
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   9000
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton CmdAccept 
         Caption         =   "Accept Changes"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Process Flow Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   6015
      Begin VB.CheckBox ChkReception 
         Caption         =   "To Reception"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox ChkLaboratory 
         Caption         =   "To Lab"
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox ChkPharmacy 
         Caption         =   "To Pharmacy"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox ChkDoctor 
         Caption         =   "To Doctor"
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox ChkCashier 
         Caption         =   "To Cashier"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox ChkObservation 
         Caption         =   "To Observation"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox CboScreen 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Parameters 
      Caption         =   "System Global Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox ChkLabScan 
         Caption         =   "Allow Importation of Scanned Results"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1320
         Value           =   2  'Grayed
         Width           =   3255
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Disable Laboratory Module"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   4695
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Disable Pharmacy Module"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   4695
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Disable Inpatient Module"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CheckBox ChkExcludeLab 
         Caption         =   "Exclude Laboratory Posting By LabTech"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4575
      End
      Begin VB.CheckBox ChkExcludePharmacy 
         Caption         =   "Exclude Drug Disbursment but Print Invoice (No Pharmacy)"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
   End
End
Attribute VB_Name = "FrmParmeters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset

Private Sub CboScreen_Click()
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM PROCESSFLOW WHERE SCREENID = '" & Mid(CboScreen, 1, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        If RsRecords.EOF = False Then
            With RsRecords
                ChkReception.Value = !CONSULTATION
                ChkObservation.Value = !OBSERVATION
                ChkDoctor.Value = !DOCTORS
                ChkCashier.Value = !CASHIER
                ChkPharmacy.Value = !PHARMACY
                ChkLaboratory.Value = !LAB
            End With
        End If
End Sub

Private Sub ChkExcludeLab_change()
    If ChkExcludeLab.Value = True Then ChkLabScan.Value = True
End Sub

Private Sub ChkExcludeLab_Click()
    If ChkExcludeLab.Value = 1 Then
        ChkLabScan.Value = 1
    Else
        ChkLabScan.Value = 0
    End If
End Sub

Private Sub CmdAccept_Click()
Dim Resp
    Resp = MsgBox("Are you sure you wish to commit changes to Settings?", vbQuestion + vbYesNo)
    If Resp = vbYes Then
    If CboScreen = "" Then GoTo NEXT1 'IN FUTURE, ONLY THE UPDATE WILL BE SKIPPED AND OTHER SETTINGS WILL BE SAVED.
    Conn.Execute "UPDATE PROCESSFLOW SET CONSULTATION = '" & ChkReception.Value & "', OBSERVATION = '" & ChkObservation & "',DOCTORS = '" & ChkDoctor & "',CASHIER= '" & ChkCashier & "',PHARMACY= '" & ChkPharmacy & "',LAB= '" & ChkLaboratory & "' WHERE SCREENID = '" & Mid(CboScreen, 1, 1) & "'"
    MsgBox "Settings Saved Successfully", vbInformation
    End If
NEXT1:
    'NURSE & DOCTOR COMBINATION SETTINGS
    Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = '" & ChkNurseDoc.Value & "' WHERE ITEMNAME = 'NurseDoctorRolesCombined'"
    
    'INHOUSE PHARMACY NOT PRESENT. ONLY PRINTING OF PRESCRIPTION
    Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = '" & ChkExcludePharmacy.Value & "' WHERE ITEMNAME = 'ExcludePharmacy'"
    
    'INHOUSE LABORATORY NOT PRESENT. SCANNING OF OUTSOURCE LABTESTS ALLOWED
    Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = '" & ChkExcludeLab.Value & "' WHERE ITEMNAME = 'ExcludeLaboratory'"
    
    LoadSettings
End Sub
Private Sub LoadSettings()
    ChkNurseDoc.Value = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'NurseDoctorRolesCombined'")
    ChkExcludePharmacy.Value = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'ExcludePharmacy'")
    ChkExcludeLab.Value = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'ExcludeLaboratory'")
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM PROCESSFLOW", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboScreen.AddItem RsRecords!SCREENID & " - " & RsRecords!SCREENNAME
            RsRecords.MoveNext
        Wend
    RsRecords.Close
    LoadSettings
    centerform Me
End Sub
