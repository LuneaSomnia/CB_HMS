VERSION 5.00
Begin VB.Form FrmABI 
   Caption         =   "Ankle - Brachial Index (ABI) Examination"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14745
   Icon            =   "FrmABI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   14745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Height          =   9495
      Left            =   12840
      TabIndex        =   72
      Top             =   480
      Width           =   1815
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   73
         Top             =   8880
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Interpreting the ABI"
      Height          =   2175
      Left            =   6480
      TabIndex        =   61
      Top             =   7800
      Width           =   6255
      Begin VB.Label Label41 
         Caption         =   "0.00 - 0.40 Servere P.A.D"
         Height          =   255
         Left            =   1800
         TabIndex        =   67
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label40 
         Caption         =   "0.41 - 0.90 Mild to Moderate P.A.D"
         Height          =   255
         Left            =   1800
         TabIndex        =   66
         Top             =   1500
         Width           =   2535
      End
      Begin VB.Label Label39 
         Caption         =   "0.91 - 0.99 Borderline P.A.D"
         Height          =   255
         Left            =   1800
         TabIndex        =   65
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label38 
         Caption         =   "1.0 - 1.29 Normal"
         Height          =   255
         Left            =   1800
         TabIndex        =   64
         Top             =   900
         Width           =   1935
      End
      Begin VB.Label Label37 
         Caption         =   "> 1.30 Noncompressible"
         Height          =   255
         Left            =   1800
         TabIndex        =   63
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label36 
         Caption         =   "ACC/AHA Guidelines for management of Patients with P.A.D (12/0)"
         Height          =   255
         Left            =   720
         TabIndex        =   62
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Notes/ RX"
      Height          =   2175
      Left            =   120
      TabIndex        =   60
      Top             =   7800
      Width           =   6255
      Begin VB.TextBox TxtNotes 
         Appearance      =   0  'Flat
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "ABI Results Right"
      Height          =   735
      Left            =   6480
      TabIndex        =   53
      Top             =   6960
      Width           =   6255
      Begin VB.CheckBox ChkRightNormal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   4920
         TabIndex        =   59
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox ChkRightInconclusive 
         Caption         =   "Inconclusive"
         Height          =   255
         Left            =   2640
         TabIndex        =   58
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox ChkRightAbnormal 
         Caption         =   "Abnormal"
         Height          =   255
         Left            =   600
         TabIndex        =   57
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ABI Results Left"
      Height          =   735
      Left            =   120
      TabIndex        =   52
      Top             =   6960
      Width           =   6255
      Begin VB.CheckBox ChkLeftNormal 
         Caption         =   "Normal"
         Height          =   255
         Left            =   4800
         TabIndex        =   56
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox ChkLeftInconclusive 
         Caption         =   "Inconclusive"
         Height          =   255
         Left            =   2520
         TabIndex        =   55
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox ChkLeftAbnormal 
         Caption         =   "Abnormal"
         Height          =   195
         Left            =   480
         TabIndex        =   54
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   120
      TabIndex        =   36
      Top             =   4920
      Width           =   12615
      Begin VB.TextBox TxtRightABI 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   50
         Top             =   1485
         Width           =   1695
      End
      Begin VB.TextBox TxtHigherRightArm 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   48
         Top             =   1005
         Width           =   1695
      End
      Begin VB.TextBox TxtHigherRightAnkle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   46
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtLeftABI 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   43
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox TxtHigherLeftArm 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   41
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxtHigherLeftAnkle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   39
         Top             =   680
         Width           =   1695
      End
      Begin VB.Label Label44 
         Caption         =   "Calculationg the Right Side :"
         Height          =   255
         Left            =   6840
         TabIndex        =   71
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label35 
         Caption         =   "Left ABI"
         Height          =   255
         Left            =   11520
         TabIndex        =   51
         Top             =   1545
         Width           =   855
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Ankle / Arm = "
         Height          =   255
         Left            =   6840
         TabIndex        =   49
         Top             =   1485
         Width           =   2775
      End
      Begin VB.Label Label33 
         Caption         =   "Higher arm pressure (of either arm)"
         Height          =   255
         Left            =   6840
         TabIndex        =   47
         Top             =   1050
         Width           =   2775
      End
      Begin VB.Label Label32 
         Caption         =   "Higher left ankle pressure (DP or PT)"
         Height          =   255
         Left            =   6840
         TabIndex        =   45
         Top             =   645
         Width           =   2775
      End
      Begin VB.Label Label31 
         Caption         =   "Right ABI"
         Height          =   255
         Left            =   4920
         TabIndex        =   44
         Top             =   1630
         Width           =   855
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Ankle / Arm = "
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label29 
         Caption         =   "Higher arm pressure (of either arm)"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1125
         Width           =   2775
      End
      Begin VB.Label Label28 
         Caption         =   "Higher right ankle pressure (DP or PT)"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label27 
         Caption         =   "Calculating Left Side :"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pressure"
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   12615
      Begin VB.TextBox TxtRightAnkleSystolicPT 
         Height          =   285
         Left            =   9120
         TabIndex        =   32
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox TxtRightAnkleSystolicDP 
         Height          =   285
         Left            =   9120
         TabIndex        =   29
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox TxtRightArmSystolic 
         Height          =   285
         Left            =   9120
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtleftAnklePT 
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox TxtleftAnkleDP 
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox TxtLeftArmSystolic 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label43 
         Caption         =   "Right Arm"
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
         Left            =   7320
         TabIndex        =   70
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label42 
         Caption         =   "Left Arm"
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
         Left            =   240
         TabIndex        =   69
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Right Ankle :"
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
         Left            =   7680
         TabIndex        =   35
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Left Ankle :"
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
         Left            =   720
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "mmHg (PT)"
         Height          =   255
         Left            =   10800
         TabIndex        =   33
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Systolic Pressure : "
         Height          =   255
         Left            =   7680
         TabIndex        =   31
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "mmHg (DP)"
         Height          =   255
         Left            =   10800
         TabIndex        =   30
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Systolic Pressure : "
         Height          =   255
         Left            =   7680
         TabIndex        =   28
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "mmHg"
         Height          =   255
         Left            =   10800
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Systolic Pressure : "
         Height          =   255
         Left            =   7680
         TabIndex        =   25
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "mmHg (PT)"
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Systolic Pressure : "
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "mmHg (DP)"
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Systolic Pressure : "
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "mmHg"
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Systolic Pressure : "
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "RIGHT PRESSURES"
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
         Left            =   9000
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "LEFT PRESSURES"
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
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   12615
      Begin VB.TextBox TxtRested 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Use higher ankle pressure. Calculate ratio of each ankle to brachaial pressure using formula below."
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         TabIndex        =   12
         Top             =   1320
         Width           =   6015
      End
      Begin VB.Label Label9 
         Caption         =   "7.  User higher ankle pressure (DP or PT) for each ankle."
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
         Left            =   6240
         TabIndex        =   11
         Top             =   960
         Width           =   6135
      End
      Begin VB.Label Label8 
         Caption         =   "6.  Repeat proceedure on left arm followed by left ankle with Doppler."
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
         Left            =   6240
         TabIndex        =   10
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label7 
         Caption         =   "5.  Measure systolic reading on right anlke with Doppler"
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
         Left            =   6240
         TabIndex        =   9
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label6 
         Caption         =   "4.  Measure Systolic reading in right arm with Doppler."
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
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "3.  Apply Ultrasound gel."
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
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "2.  Place Standard blood pressure cuffs around patient's ankles and arms"
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   "Time Patient Rested"
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
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "1.  Have Patient rest Supine for 10 Minutes."
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.TextBox TxtPatientsName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Patient's Name :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   1695
   End
End
Attribute VB_Name = "FrmABI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Private Sub CmdSave_Click()
    
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open " SELECT * FROM ABI_EXAMINATION WHERE CARDNUMBER = '" & lvOptionalCardNo & "'", Conn, adOpenStatic, adLockOptimistic
            With RsRecords
                If .EOF = True Then
                    .AddNew
                Else
                
                End If
                    !CARDNUMBER = lvOptionalCardNo
                    !TIMERESTED = TxtRested
                    !RIGHTARMSYSTOLIC = TxtRightArmSystolic
                    !RIGHTANKLESYSTOLICDP = TxtRightAnkleSystolicDP
                    !RIGHTANKLESYSTOLICPT = TxtRightAnkleSystolicPT
                    !LEFTARMSYSTOLIC = TxtLeftArmSystolic
                    !LEFTANKLESYSTOLICDP = TxtleftAnkleDP
                    !LEFTANKLESYSTOLICPT = TxtleftAnklePT
                    !HIGHERRIGHTANKLE = TxtHigherRightAnkle
                    !HIGHERLEFTANKLE = TxtHigherLeftAnkle
                    !HIGHERARMRIGHT = TxtHigherRightArm
                    !HIGHERARMLEFT = TxtHigherLeftArm
                    !RIGHTABNORMAL = ChkRightAbnormal.Value
                    !rightInconclusive = ChkRightInconclusive.Value
                    !rightnormal = ChkRightNormal.Value
                    !leftabnormal = ChkLeftAbnormal.Value
                    !leftinconclusive = ChkLeftInconclusive.Value
                    !leftnormal = ChkLeftNormal
                    !NotesRX = TxtNotes
                
               .Update
            End With
End Sub
Private Sub RetrieveLastABI(ByRef CardNo As String)
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open " SELECT * FROM ABI_EXAMINATION WHERE CARDNUMBER = '" & lvOptionalCardNo & "'", Conn, adOpenStatic, adLockOptimistic
            With RsRecords
                If .EOF = False Then
                    '!CARDNUMBER = lvOptionalCardNo
                     TxtRested = !TIMERESTED
                    TxtRightArmSystolic = !RIGHTARMSYSTOLIC
                    TxtRightAnkleSystolicDP = !RIGHTANKLESYSTOLICDP
                    TxtRightAnkleSystolicPT = !RIGHTANKLESYSTOLICPT
                    TxtLeftArmSystolic = !LEFTARMSYSTOLIC
                    TxtleftAnkleDP = !LEFTANKLESYSTOLICDP
                    TxtleftAnklePT = !LEFTANKLESYSTOLICPT
                    TxtHigherRightAnkle = !HIGHERRIGHTANKLE
                    TxtHigherLeftAnkle = !HIGHERLEFTANKLE
                    TxtHigherRightArm = !HIGHERARMRIGHT
                    TxtHigherLeftArm = !HIGHERARMLEFT
                    ChkRightAbnormal.Value = !RIGHTABNORMAL
                    ChkRightInconclusive = !rightInconclusive
                    ChkRightNormal = !rightnormal
                    ChkLeftAbnormal = !leftabnormal
                    ChkLeftInconclusive = !leftinconclusive
                    ChkLeftNormal = !leftnormal
                    TxtNotes = !NotesRX
                End If
            End With
End Sub
Private Sub Form_Load()
    centerform Me
        If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT FIRSTNAME,SECONDNAME,SURNAME FROM PATIENT_DETAILS WHERE CARDNUMBER = '" & lvOptionalCardNo & "'", Conn, adOpenStatic, adLockOptimistic
            With RsRecords
                If .EOF = False Then
                    TxtPatientsName = !FirstName & "  " & !SECONDNAME & "  " & !SURNAME
                End If
            End With
        RsRecords.Close
        
    RetrieveLastABI (lvOptionalCardNo)
End Sub

Private Sub TxtHigherLeftArm_Change()
On Error GoTo ErrorHandler
    If TxtHigherLeftAnkle <> "" Then
        If TxtHigherLeftArm = "" Then Exit Sub
         TxtLeftABI = (TxtHigherLeftAnkle / TxtHigherLeftArm)
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
    TxtLeftABI = ""
End Sub

Private Sub TxtHigherRightArm_Change()
    If TxtHigherRightAnkle <> "" Then
        TxtRightABI = (TxtHigherRightAnkle / TxtHigherRightArm)
    End If
End Sub
