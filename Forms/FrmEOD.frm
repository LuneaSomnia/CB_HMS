VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEOD 
   Caption         =   "&H0000FFFF&"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "FrmEOD.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4680
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkReset 
      Caption         =   "Reset System Dates (In the Event EOD Process has been run more than once or otherwise.)"
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
      TabIndex        =   15
      Top             =   1320
      Width           =   8295
   End
   Begin VB.Frame Frame3 
      Caption         =   "RESET SYSTEM DATE"
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   8775
      Begin VB.CommandButton CmdAcceptNewDate 
         Caption         =   "Accept New System Date"
         Height          =   495
         Left            =   6360
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTCurrentDate 
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49741825
         CurrentDate     =   40998
      End
      Begin MSComCtl2.DTPicker DTPreviousDate 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49741825
         CurrentDate     =   40998
      End
      Begin VB.Label Label4 
         Caption         =   "Current System Date:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Previous System Date"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Timer TmrEOD 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4440
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar ProgressEOD 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   8775
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit End Of Day Process"
         Height          =   495
         Left            =   6360
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton CmdStartEOD 
         Caption         =   "Start End Of Day Process"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SYSTEM DATE"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.TextBox TxtPreviousEODate 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   2535
      End
      Begin VB.TextBox TxtCurrentEODate 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Previous System Date:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Current System Date:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmEOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrPreviousDate As Date
Dim StrCurrentDate As Date
Dim RsDates As New ADODB.Recordset

Private Sub ChkReset_Click()
    If ChkReset.Value = 1 Then
        Me.Height = 5235
    Else
        Me.Height = 3645
    End If
End Sub

Private Sub CmdAcceptNewDate_Click()
Dim Resp As Variant
    Resp = MsgBox("Are you sure you wish to Reset System Date?", vbQuestions + vbYesNo + vbQuestion)
    If Resp = vbYes Then
        Conn.Execute "UPDATE DUAL SET PREVIOUSDATE = '" & Format(DTPreviousDate, "DD MMMM YYYY") & "', CURRENTDATE = '" & Format(DTCurrentDate, "DD MMM YYYY") & "'"
        
        'LOAD SYSTEM GLOBAL DATE
        RsDates.Open "SELECT CurrentDate FROM DUAL", Conn, adOpenStatic, adLockOptimistic
            With RsDates
                If .EOF = False Then
                   GlbSysDate = Format(!CurrentDate, "dd/MMM/YYYY")
                End If
            End With
        RsDates.Close
        
        MDIMain.SSBar.Panels(1).Text = Format(GlbSysDate, "DD MMMM YYYY") & "           " & "Current User = " & UCase(GlbCurrentUser)
        MDIMain.SSBar.Panels(4).Text = Format(GlbSysDate, "DD MMMM YYYY")
    Else
        Exit Sub
    End If
    'LOAD DATES FROM DATABASE
    RsDates.Open "SELECT * FROM DUAL", Conn, adOpenStatic, adLockOptimistic
        With RsDates
            If .EOF = False Then
                TxtPreviousEODate = Format(!PREVIOUSDATE, "DDDD, MMMM - DD - YYYY")
                TxtCurrentEODate = Format(!CurrentDate, "DDDD, MMMM - DD - YYYY")
            End If
        End With
    RsDates.Close
        MsgBox "New Date Updated Succesfully", vbInformation
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdStartEOD_Click()
On Error GoTo ERRORHANDLER
Resp = MsgBox("END OF DAY PROCESS IS RUN ONLY ONCE A DAY (IN THE MORNING). ARE YOU SURE YOU WISH TO RUN EOD PROCESS?", vbQuestion + vbYesNo)
    If Resp = vbYes Then
        ProgressEOD.Value = 0
        TmrEOD.Enabled = True
        CmdStartEOD.Enabled = True
        
    'CHANGE THE SHIFT BACK TO DAY SHIFT.
    Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = 0 WHERE ITEMNAME = 'DayShift_0_NightShift_1'"
    Else
        MsgBox "End of Day Process Has Been Aborted", vbInformation
    End If
Exit Sub
ERRORHANDLER:
    MsgBox Err.Description
   ' Resume
End Sub


Private Sub Form_Load()
    Me.Height = 3645
    centerform Me
    
    'LOAD DATES FROM DATABASE
    RsDates.Open "SELECT * FROM DUAL", Conn, adOpenStatic, adLockOptimistic
        With RsDates
            If .EOF = False Then
                TxtPreviousEODate = !PREVIOUSDATE
                TxtCurrentEODate = !CurrentDate
            End If
        End With
    RsDates.Close
End Sub

Private Sub TmrEOD_Timer()
    ProgressEOD.Value = ProgressEOD.Value + 1
    If ProgressEOD.Value = 20 Then
        ProgressEOD.Value = 100
        TmrEOD.Enabled = False
        StrPreviousDate = TxtCurrentEODate
        TxtPreviousEODate = Format(TxtCurrentEODate, "DDDD, MMMM - DD - YYYY")
        TxtCurrentEODate = CDate(TxtCurrentEODate) + 1
        StrCurrentDate = TxtCurrentEODate
        TxtCurrentEODate = Format(TxtCurrentEODate, "DDDD, MMMM - DD - YYYY")
        
        Conn.Execute "UPDATE DUAL SET PREVIOUSDATE = '" & Format(StrPreviousDate, "DD MMMM YYYY") & "', CURRENTDATE = '" & Format(StrCurrentDate, "DD MMM YYYY") & "'"
             '******88
            ''''''' Conn.Execute "UPDATE DUAL SET PREVIOUSDATE = '" & Format(DTPreviousDate, "DD MMMM YYYY") & "', CURRENTDATE = '" & Format(DTCurrentDate, "DD MMM YYYY") & "'"
             
             'LOAD SYSTEM GLOBAL DATE
             RsDates.Open "SELECT CurrentDate FROM DUAL", Conn, adOpenStatic, adLockOptimistic
                 With RsDates
                     If .EOF = False Then
                        GlbSysDate = Format(!CurrentDate, "dd/MMM/YYYY")
                     End If
                 End With
             RsDates.Close
             
             MDIMain.SSBar.Panels(1).Text = Format(GlbSysDate, "DD MMMM YYYY") & "           " & "Current User = " & UCase(GlbCurrentUser)
             MDIMain.SSBar.Panels(4).Text = Format(GlbSysDate, "DD MMMM YYYY")
             
             '********
        MsgBox "End of Day Process Completed Succesfully", vbInformation
    End If
End Sub
