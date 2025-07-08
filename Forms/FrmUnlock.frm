VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUnlock 
   Caption         =   "Unlock Record"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdUnlock 
      Caption         =   "Unlock Record"
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
      Left            =   10800
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of Locked Records"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VSFlex6DAOCtl.vsFlexGrid GridObservation 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5530
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
End
Attribute VB_Name = "FrmUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset

Private Sub CmdUnlock_Click()
    Conn.Execute "UPDATE COMPLAINS SET INUSE = 'FALSE' WHERE CARDNUMBER = '" & GridObservation.TextMatrix(GridObservation.Row, 0) & "'"
    Progress = 100
    
    Fill_LockedUsers
End Sub

Private Sub Form_Load()
    Fill_LockedUsers
    centerform Me
    Progress = 0
End Sub
Private Sub Fill_LockedUsers()
   ' On Error GoTo ErrorHandler
   KARI = GlbSysDate: Rcount = 0
   
    GridObservation.Clear
    GridObservation.Rows = 1
    GridObservation.Cols = 4
    For i = 0 To GridObservation.Cols - 1
        GridObservation.ColAlignment(i) = flexAlignCenterCenter
    Next i
    
    GridObservation.ColAlignment(1) = flexAlignCenterCenter
    
    GridObservation.Editable = True
    GridObservation.ColDataType(3) = flexDTBoolean
    GridObservation.ColWidth(1) = 3105
    GridObservation.ColWidth(2) = 3990
    GridObservation.FormatString = "CARD NUMBER|PATIENTS FULL NAME                                  .|BILLING COMPANY     |SELECT TO UNLOCK "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT PATIENT_DETAILS.* FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER  AND COMPLAINS.INUSE = '1' and visitdate = '" & Format(GlbSysDate, "DD MMM YYYY") & "'", Conn, adOpenDynamic, adLockOptimistic
        
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridObservation.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY
                        .MoveNext
                        Rcount = Rcount + 1
                    Wend
                End With
            End If
        TxtObservation = Rcount
    Exit Sub
'ErrorHandler:
  '  MsgBox Err.Description

End Sub

