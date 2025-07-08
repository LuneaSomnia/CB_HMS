VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAudit 
   Caption         =   "Audit Trail"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   Icon            =   "FrmAudit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Activity Trail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   11775
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7011
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
      Caption         =   "Audit Trail Selection Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10080
         TabIndex        =   15
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton OptSingleDate 
         Caption         =   "By Single Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1320
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptDateRange 
         Caption         =   "By Date Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTAudit 
         Height          =   315
         Left            =   8760
         TabIndex        =   6
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16908289
         CurrentDate     =   41110
      End
      Begin VB.ComboBox CboUser 
         Height          =   315
         Left            =   8760
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox CboScreen 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker DTStartDate 
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16908289
         CurrentDate     =   41110
      End
      Begin MSComCtl2.DTPicker DTEndDate 
         Height          =   315
         Left            =   7440
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16908289
         CurrentDate     =   41110
      End
      Begin VB.Label Label5 
         Caption         =   "End Date To Audit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Start  Date To Audit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1860
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   11640
         X2              =   120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label3 
         Caption         =   "Date To Audit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   5
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "User To Audit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Screen To Audit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim RsCombo As New ADODB.Recordset

Private Sub CmdSearch_Click()
    If CboScreen.Text = "" Then CboScreen = "ALL MODULES"
    If CboUser.Text = "" Then CboUser = "ALL USERS"
    
    POPULATE_Filter
End Sub

Private Sub Form_Load()
    'POPULATE COMBO FOR SCREENS
    RsCombo.Open "SELECT DISTINCT SCREEN FROM PROAUDIT", Conn, adOpenDynamic, adLockOptimistic
    'CboScreen.Clear
        With RsCombo
                CboScreen.AddItem "ALL MODULES"
            While .BOF = False And .EOF = False
                lvScreenName = FindRecord("SCREENS", "FORMNAME", "FORMID = '" & !Screen & "'")
                CboScreen.AddItem String(2 - Len(!Screen), "0") & !Screen + " - " + lvScreenName
                .MoveNext
            Wend
        End With
    RsCombo.Close
    
    'POPULATE COMBO FOR USERS
    RsCombo.Open "SELECT DISTINCT SYSTEMUSER FROM PROAUDIT", Conn, adOpenDynamic, adLockOptimistic
    'CboScreen.Clear
                CboUser.AddItem "ALL USERS"
        With RsCombo
            While .BOF = False And .EOF = False
                CboUser.AddItem !SYSTEMUSER
                .MoveNext
            Wend
        End With
    RsCombo.Close

    POPULATEGRID
    centerform Me
End Sub

Private Sub OptDateRange_Click()
    DTStartDate.Enabled = True
    DTEndDate.Enabled = True
    DTAudit.Enabled = False
End Sub

Private Sub OptSingleDate_Click()
    DTStartDate.Enabled = False
    DTEndDate.Enabled = False
    DTAudit.Enabled = True
End Sub
Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    'G.CellWidth G.Row, 3 = 20000
    G.FormatString = "USER| MODULE | AUDIT DATE | AUDIT TIME | ACTION  TAKEN  BY  USER                                                                                                      "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  PROAUDIT", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !SYSTEMUSER & vbTab & !Screen & vbTab & !Date & vbTab & !Time & vbTab & !Action
                    End With
                RsRecords.MoveNext
                Wend
            End If
        RsRecords.Close
    G.Editable = True
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
End Sub
Public Sub POPULATE_Filter()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "USER| MODULE | AUDIT DATE | AUDIT TIME | ACTION  TAKEN  BY  USER                                                                                                                  "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        If CboUser = "ALL USERS" And CboScreen <> "ALL MODULES" Then
            RsRecords.Open "SELECT * FROM  PROAUDIT WHERE SCREEN = '" & Val(GetID_NameFromCombo(CboScreen, 1)) & "'", Conn, adOpenStatic, adLockOptimistic
        ElseIf CboUser = "ALL USERS" And CboScreen = "ALL MODULES" Then
            RsRecords.Open "SELECT * FROM  PROAUDIT", Conn, adOpenStatic, adLockOptimistic
        ElseIf CboUser <> "ALL USERS" And CboScreen <> "ALL MODULES" Then
            RsRecords.Open "SELECT * FROM  PROAUDIT WHERE SCREEN = '" & Val(GetID_NameFromCombo(CboScreen, 1)) & "' AND SYSTEMUSER = '" & CboUser & "'", Conn, adOpenStatic, adLockOptimistic
        ElseIf CboUser <> "ALL USERS" And CboScreen = "ALL MODULES" Then
            RsRecords.Open "SELECT * FROM  PROAUDIT WHERE SYSTEMUSER = '" & CboUser & "'", Conn, adOpenStatic, adLockOptimistic
        End If
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !SYSTEMUSER & vbTab & !Screen & vbTab & !Date & vbTab & !Time & vbTab & !Action
                    End With
                RsRecords.MoveNext
                Wend
            End If
        RsRecords.Close
    G.Editable = True
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
End Sub

