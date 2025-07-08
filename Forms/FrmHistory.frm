VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmHistory 
   Caption         =   "Patient History"
   ClientHeight    =   8325
   ClientLeft      =   -1470
   ClientTop       =   555
   ClientWidth     =   12105
   Icon            =   "FrmHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   10200
      TabIndex        =   22
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   10200
      TabIndex        =   21
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   8520
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search For Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   8295
      Begin VB.CheckBox ChkSecondName 
         Caption         =   "Second Name"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox ChkFirstName 
         Caption         =   "First Name"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtSearchSurname 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton OptSurname 
         Caption         =   "Surname"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox TxtSearchSecondName 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox TxtPhoneNumber 
         Height          =   285
         Left            =   6000
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton OptPhone 
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   4560
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton OptIDNumber 
         Caption         =   "ID Number"
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton OptCardNumber 
         Caption         =   "Card Number"
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtSearchIDNumber 
         Height          =   285
         Left            =   6000
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TxtSearchFirstName 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox TxtSearchCard 
         Height          =   285
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Patient History"
      Height          =   5055
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   11895
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   4695
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   8281
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
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   11895
      Begin VB.TextBox TxtAge 
         BackColor       =   &H80000000&
         Height          =   405
         Left            =   10680
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtSecondName 
         Height          =   285
         Left            =   6720
         TabIndex        =   15
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox TxtCardNumber 
         Height          =   285
         Left            =   6720
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox TxtBillingCompany 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtFirstname 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   "AGE"
         Height          =   255
         Left            =   11040
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Card Number"
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Billing Company"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Second Name"
         Height          =   255
         Left            =   5520
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset
Private Sub FillHistory()
   ' On Error GoTo ErrorHandler
   KARI = Date
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 7
    Grid.ColAlignment(1) = flexAlignCenterCenter
    'Grid.ColDataType(7) = flexDTBoolean
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |   VISIT DATE "
        If RsGrid.State = adStateOpen Then RsGrid.Close
         'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE < '" & Format(KARI, "DDMMMYYYY") & "' ORDER BY VISITDATE DESC", Conn, adOpenDynamic, adLockOptimistic
         RsGrid.Open "SELECT top 200 * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER ORDER BY VISITDATE DESC", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & !FirstName & " " & !SECONDNAME & " " & !SURNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER & vbTab & !VisitDate
                        .MoveNext
                    Wend
                End With
            End If
    Exit Sub
'ErrorHandler:
  '  MsgBox Err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdRefresh_Click()
    FillHistory
End Sub

Private Sub CmdSearch_Click()
    'If TxtSearchCard <> "" Then
        Dim RsSearch As New ADODB.Recordset
        Dim SQLStatement As String
        KARI = GlbSysDate
        Grid.Clear
        Grid.Rows = 1
        Grid.Cols = 7
        Grid.ColAlignment(1) = flexAlignCenterCenter
        'Grid.ColDataType(7) = flexDTBoolean
        Grid.ColWidth(1) = 3105
        Grid.ColWidth(2) = 3990
        Grid.FormatString = "CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |   VISIT DATE "
        If RsSearch.State = adStateOpen Then RsSearch.Close
        Me.MousePointer = vbHourglass
            If ChkFirstName.Value = 1 And ChkSecondName.Value = 1 Then
                Grid.FormatString = "CARD NUMBER| PATIENT FULL NAME       |  BILLING COMPANY     |ID NUMBER |   VISIT DATE "
                SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS WHERE FIRSTNAME = '" & TxtSearchFirstName & "' AND SECONDNAME = '" & TxtSearchSecondName & "'"
            ElseIf ChkFirstName.Value = 1 Then
                Grid.FormatString = "CARD NUMBER| PATIENT FULL NAME       |  BILLING COMPANY     |ID NUMBER |   VISIT DATE "
                SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS WHERE FIRSTNAME = '" & TxtSearchFirstName & "'"
            ElseIf ChkSecondName.Value = 1 Then
                Grid.FormatString = "CARD NUMBER| PATIENT FULL NAME       |  BILLING COMPANY     |ID NUMBER |   VISIT DATE "
                SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS WHERE SECONDNAME = '" & TxtSearchSecondName & "'"
            ElseIf OptSurname.Value = True Then
                Grid.FormatString = "CARD NUMBER| PATIENT FULL NAME       |  BILLING COMPANY     |ID NUMBER |   VISIT DATE "
                SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS WHERE SURNAME = '" & TxtSearchSurname & "'"
            ElseIf OptIDNumber = True Then
                Grid.FormatString = "CARD NUMBER| PATIENT FULL NAME       |  BILLING COMPANY     |ID NUMBER |   VISIT DATE "
                SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS WHERE IDNUMBER = '" & TxtSearchIDNumber & "'"
            ElseIf OptPhone.Value = True Then
                Grid.FormatString = "CARD NUMBER| PATIENT FULL NAME       |  BILLING COMPANY     |ID NUMBER |   VISIT DATE "
                SQLStatement = "SELECT  top 500 * FROM PATIENT_DETAILS WHERE TELEPHONE = '" & TxtPhoneNumber & "'"
            ElseIf OptCardNumber = True Then
                Grid.FormatString = "CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |   VISIT DATE "
                'INNER JOIN STATEMENT WITH TABLE COMPLAINS DISQUALIFIES NEW RECORDS FROM BEING SEEN ON THE SEARCH.
                'SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER WHERE PATIENT_DETAILS.CARDNUMBER = '" & TxtSearchCard & "'  ORDER BY VISITDATE DESC"
                SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS  WHERE PATIENT_DETAILS.CARDNUMBER = '" & TxtSearchCard & "'"
            End If
                If SQLStatement = "" Then Exit Sub
                RsSearch.Open SQLStatement, Conn, adOpenDynamic, adLockOptimistic
                    If RsSearch.BOF = False And RsSearch.EOF = False Then
                        With RsSearch
                            While Not .EOF
                                If OptCardNumber.Value = True Then
                                    'Grid.AddItem !CARDNUMBER & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FIRSTNAME & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER & vbTab & !VisitDate
                                    Grid.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                                Else
                                    Grid.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                                End If
                                .MoveNext
                            Wend
                        End With
                    End If
        Me.MousePointer = vbNormal
    'End If
End Sub
Private Sub SearchFilter_FirstName(ByVal FirstName As String)
        Dim RsSearch As New ADODB.Recordset
        Dim SQLStatement As String
        KARI = GlbSysDate
        Grid.Clear
        Grid.Rows = 1
        Grid.Cols = 7
        Grid.ColAlignment(1) = flexAlignCenterCenter
        'Grid.ColDataType(7) = flexDTBoolean
        Grid.ColWidth(1) = 3105
        Grid.ColWidth(2) = 3990
        Grid.FormatString = "CARD NUMBER|   PATIENTS FULL NAME                      . |   BILLING COMPANY     |ID NUMBER |   VISIT DATE "
        If RsSearch.State = adStateOpen Then RsSearch.Close
        Me.MousePointer = vbHourglass
        WAMBOH = TxtSearchFirstName + "%"
                SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS  WHERE PATIENT_DETAILS.FIRSTNAME LIKE '" & WAMBOH & "'"
                If SQLStatement = "" Then Exit Sub
                
                RsSearch.Open SQLStatement, Conn, adOpenStatic, adLockOptimistic
                    If RsSearch.BOF = False And RsSearch.EOF = False Then
                        With RsSearch
                            While Not .EOF
                                    Grid.AddItem !CardNumber & vbTab & !FirstName & " " & !SECONDNAME & " " & !SURNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                                .MoveNext
                            Wend
                        End With
                    End If
        Me.MousePointer = vbNormal
End Sub

Private Sub SearchFilter_SecondName(ByVal FirstName As String)
        Dim RsSearch As New ADODB.Recordset
        Dim SQLStatement As String
        KARI = GlbSysDate
        Grid.Clear
        Grid.Rows = 1
        Grid.Cols = 7
        Grid.ColAlignment(1) = flexAlignCenterCenter
        'Grid.ColDataType(7) = flexDTBoolean
        Grid.ColWidth(1) = 3105
        Grid.ColWidth(2) = 3990
        Grid.FormatString = "CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |   VISIT DATE "
        If RsSearch.State = adStateOpen Then RsSearch.Close
        Me.MousePointer = vbHourglass
        WAMBOH = TxtSearchSecondName + "%"
                SQLStatement = "SELECT top 500 * FROM PATIENT_DETAILS  WHERE PATIENT_DETAILS.FIRSTNAME = '" & TxtSearchFirstName & "' AND SECONDNAME LIKE '" & WAMBOH & "'"
                If SQLStatement = "" Then Exit Sub
                
                RsSearch.Open SQLStatement, Conn, adOpenStatic, adLockOptimistic
                    If RsSearch.BOF = False And RsSearch.EOF = False Then
                        With RsSearch
                            While Not .EOF
                                    Grid.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                                .MoveNext
                            Wend
                        End With
                    End If
        Me.MousePointer = vbNormal
End Sub
Private Sub Form_Load()
    centerform Me
    FillHistory
End Sub

Private Sub Grid_DblClick()
On Error GoTo ErrorHandler
    If GlbCalledFromPatients = True Then
        FrmPatients.TxtCardNumber = Grid.TextMatrix(Grid.Row, 0)
        Unload Me
    ElseIf GlbCalledFromAppointments = True Then
        FrmPatients.TxtBookCardNumber = Grid.TextMatrix(Grid.Row, 0)
        Unload Me
    Else
    'ASSIGN VALUES TO GLOBAL VARIABLES.
    If Grid.Row = 0 Then Exit Sub
    StrDocCardNo = Grid.TextMatrix(Grid.Row, 0)
    StrDocVisitNumber = FindRecord("COMPLAINS", "VISITNUMBER", "CARDNUMBER = '" & StrDocCardNo & "' ORDER BY VISITNUMBER DESC")    ' Grid.TextMatrix(Grid.Row, 1)
    StrDocVisitDate = Grid.TextMatrix(Grid.Row, 5)
    BlnHISTORY = True
    
    MsgBox StrDocCardNo + " " + StrDocVisitNumber
    
    lvFullNames = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
    lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SECONDNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
    lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SURNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
    AuditTrail GlbCurrentUser, EnumPatientHistory, GlbSysDate, Time, "Loaded Patient History For CardNumber - " + "" & StrDocCardNo & "" + " - " + "" & lvFullNames & ""
    FrmTreatment.Show
    End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub


Private Sub TxtSearchFirstName_Change()
    SearchFilter_FirstName TxtSearchFirstName
End Sub

Private Sub TxtSearchSecondName_Change()
    SearchFilter_SecondName TxtSearchSecondName
End Sub


