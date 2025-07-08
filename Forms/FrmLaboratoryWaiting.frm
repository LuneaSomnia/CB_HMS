VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmLaboratoryWaiting 
   Caption         =   "Laboratory Waiting List"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "FrmLaboratoryWaiting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Refresh"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   11055
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh List"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "C&lose"
         Height          =   615
         Left            =   8400
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laboratory Test Requests"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   9340
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
   End
End
Attribute VB_Name = "FrmLaboratoryWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset

Private Sub POPULATELABORATORY()
   ' On Error GoTo ErrorHandler
   KARI = GlbSysDate
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 3
    Grid.ColAlignment(1) = flexAlignCenterCenter
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 99999
    Grid.FormatString = "CARD NUMBER|   PATIENTS FULL NAME  | REQUESTING DOCTOR | VISIT NUMBER | LAB TEST REQUEST DETAILS "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT PATIENT_DETAILS.*,COMPLAINS.LABREQUEST,COMPLAINS.DOCTOR,COMPLAINS.VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND COMPLAINS.TOLABORATORY = '1'", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !CARDNUMBER & vbTab & !SURNAME & " " & !FIRSTNAME & " " & !SECONDNAME & vbTab & "DR. " & !DOCTOR & vbTab & !VISITNUMBER & vbTab & !LABREQUEST
                        .MoveNext
                    Wend
                End With
            End If
        RsGrid.Close
    Exit Sub
'ErrorHandler:
  '  MsgBox Err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdRefresh_Click()
    POPULATELABORATORY
End Sub

Private Sub Form_Load()
    POPULATELABORATORY
    centerform Me
End Sub



Private Sub Grid_DblClick()
    If Grid.TextMatrix(Grid.Row, 0) = "CARD NUMBER" Then GoTo LoadBlank
    StrDocCardNo = Grid.TextMatrix(Grid.Row, 0)
    StrDocVisitNumber = Grid.TextMatrix(Grid.Row, 3)
    'StrDocVisitDate = Grid.TextMatrix(Grid.Row, 9)
    StrDocVisitDate = FindRecord("COMPLAINS", "VISITDATE", "VISITNUMBER = '" & StrDocVisitNumber & "'")
    StrLabAmount = FindRecord("PRESCRIPTION", "CASHAMOUNT", "CARDNUMBER='" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "' AND DESCRIPTION = 'Lab Test'")
LoadBlank:
    FrmLaboratory.Show
    Unload Me
End Sub

