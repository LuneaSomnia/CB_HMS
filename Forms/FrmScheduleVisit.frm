VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmScheduleVisit 
   Caption         =   "Schedule Re-Visit"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   Icon            =   "FrmScheduleVisit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtCount 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Appointments for Selected Date"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   7935
      Begin VB.ComboBox CboDoctor 
         Height          =   315
         Left            =   5760
         TabIndex        =   15
         Top             =   3480
         Width           =   2055
      End
      Begin VSFlex6DAOCtl.vsFlexGrid GridAppointments 
         Height          =   3015
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5318
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483624
         ForeColor       =   16711680
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   65535
         ForeColorSel    =   16711680
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483624
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   4
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   -1  'True
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Appointments By Doctor"
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
         Left            =   3360
         TabIndex        =   16
         Top             =   3480
         Width           =   2295
      End
   End
   Begin VB.CommandButton CmdSaveSchedule 
      Caption         =   "Save Schedule"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Schedule Details"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin MSComCtl2.DTPicker DTTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   49348610
         CurrentDate     =   41024
      End
      Begin MSComCtl2.DTPicker DTAppointmentDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49348609
         CurrentDate     =   40848
      End
      Begin VB.TextBox TxtCardNumber 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox TxtNames 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Time"
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
         Left            =   3360
         TabIndex        =   14
         Top             =   1480
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Next Visit"
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
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Card Number"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Number of Appointments"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label LblSelectedDate 
      Alignment       =   2  'Center
      Caption         =   "Date Selected"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   7935
   End
End
Attribute VB_Name = "FrmScheduleVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset
Dim RsRecords As New ADODB.Recordset

Private Sub CboDoctor_Change()
    If CboDoctor = "ALL DOCTORS" Then
        FillAppointments DTAppointmentDate
    Else
        FillAppointments DTAppointmentDate, CboDoctor
    End If
End Sub

Private Sub CboDoctor_Click()
    If CboDoctor = "ALL DOCTORS" Then
        FillAppointments DTAppointmentDate
    Else
        FillAppointments DTAppointmentDate, CboDoctor
    End If
End Sub

Private Sub CmdSaveSchedule_Click()

    'CHECK TO CONFIRM THAT DATE IS NOT EARLIER THAN CURRENT DATE
    If DTAppointmentDate < GlbSysDate Then MsgBox "Appointment Date can NOT be earlier than Current System Date.", vbMsgBoxRight + vbSystemModal + vbExclamation, "Schedule Not Set": Exit Sub
    'CHECK IF APPOINTMENT WITH SAME DATE EXISTS
    If FindRecord("APPOINTMENTS", "APPDATE", "CARDNUMBER = '" & TxtCardNumber & "' AND APPDATE = '" & DTAppointmentDate & "'") <> "" Then
        MsgBox "Patient With Card Number " & TxtCardNumber & " Already has an appointment on " & Format(DTAppointmentDate, "dd mmmm yyyy"), vbExclamation, "Scheduling Cancelled"
        Exit Sub
    End If
    
    'PICK RELEVANT FIELDS AND SAVE IN APPOINTMENTS
    Dim RsSchedule As New ADODB.Recordset
    If RsSchedule.State = 1 Then Set RsSchedule = Nothing
    RsSchedule.Open "SELECT * FROM PATIENT_DETAILS WHERE CARDNUMBER = '" & TxtCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
        If RsSchedule.EOF = False Then
            With RsSchedule
            Conn.Execute "INSERT INTO APPOINTMENTS(CARDNUMBER,FIRSTNAME,SECONDNAME,SURNAME,TELEPHONE,APPDATE,APPTIME,DOCTOR,STATUS)" & _
                         "VALUES('" & TxtCardNumber.Text & "','" & !FirstName & "','" & !SECONDNAME & "','" & !SURNAME & "','" & !TELEPHONE & "','" & DTAppointmentDate & "','" & DTTime & "','" & GlbCurrentUser & "','0')"
            End With
        End If
    MsgBox "Visit For  " & Format(DTAppointmentDate, "DD MMMm YYYY") & "  Has been saved Successfully", vbInformation
    CmdSaveSchedule.Enabled = False
End Sub

Private Sub DtVisitDate_Change()
    LblDate = Format(DtVisitDate, "dd mmmm yyyy")
End Sub

Private Sub DTAppointmentDate_Change()
    LblSelectedDate.Caption = Format(DTAppointmentDate, "DD MMMM YYYY")
    FillAppointments DTAppointmentDate
End Sub
Private Sub FillAppointments(ByVal AppDate As Date, Optional Doc As String)
    On Error GoTo ERRORHANDLER
    Dim Rcount As Integer
    If AppDate = 0 Then Exit Sub
    GridAppointments.Clear
    GridAppointments.Rows = 1
    GridAppointments.Cols = 3
    GridAppointments.ColAlignment(1) = flexAlignCenterCenter
    GridAppointments.ColWidth(1) = 3105
    GridAppointments.ColWidth(2) = 3990
    GridAppointments.FormatString = "FIRST NAME |   SECOND NAME  |   SURNAME     |  PHONE NUMBER |APPOINTMENT DATE| TIME |  DOCTOR TO BE SEEN  | ID | CARD NO"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        If Doc = "" Then
            RsGrid.Open "SELECT * FROM APPOINTMENTS WHERE APPDATE = '" & Format(AppDate, "DD MMM YYYY") & "' ORDER BY APPDATE DESC", Conn, adOpenStatic, adLockOptimistic
        Else
            RsGrid.Open "SELECT * FROM APPOINTMENTS WHERE APPDATE = '" & Format(AppDate, "DD MMM YYYY") & "' AND DOCTOR = '" & Right(CboDoctor, Len(CboDoctor) - 4) & "' ORDER BY APPDATE DESC", Conn, adOpenStatic, adLockOptimistic
        End If
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridAppointments.AddItem !FirstName & vbTab & !SECONDNAME & vbTab & !SURNAME & vbTab & !TELEPHONE & vbTab & !AppDate & vbTab & Right(!APPTIME, 11) & vbTab & !DOCTOR & vbTab & !AppointmentID & vbTab & !CARDNUMBER
                        .MoveNext
                        Rcount = Rcount + 1
                    Wend
                End With
            End If
        RsGrid.Close
        TxtCount = Rcount
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Description
   ' Resume
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
    TxtCardNumber = StrDocCardNo
    GetNames StrDocCardNo
    centerform Me
    FillAppointments 0
    
    'POPULATE DOCTORS COMBO
    If RsRecords.State = 1 Then Set RsRecords = Nothing
            CboDoctor.AddItem "ALL DOCTORS"
    RsRecords.Open "SELECT DISTINCT DOCTOR FROM APPOINTMENTS WHERE APPDATE >= '" & GlbSysDate & "'", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboDoctor.AddItem "DR. " & RsRecords!DOCTOR
            RsRecords.MoveNext
        Wend
    RsRecords.Close
    
    DTAppointmentDate.Value = GlbSysDate
    
End Sub
Private Sub GetNames(ByVal CardNo)
    If CardNo = "" Then Exit Sub
    Dim RsNames As New ADODB.Recordset
    RsNames.Open "Select Firstname, SecondName, Surname  from Patient_details where cardnumber = '" & CardNo & "'", Conn, adOpenStatic, adLockOptimistic
    TxtNames = RsNames!FirstName + " " + RsNames!SECONDNAME + " " + RsNames!SURNAME
    RsNames.Close
End Sub

