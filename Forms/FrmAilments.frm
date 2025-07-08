VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmAilments 
   Caption         =   "Ailments Maintenance (Magonjwa)"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9105
   Icon            =   "FrmAilments.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Ailment Description Details"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7215
      Begin VB.TextBox TxtAilmentID 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox ChkReportMonthly 
         Caption         =   "Include Ailment in Goverment Monthly Reports"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   14
         Top             =   2520
         Width           =   4935
      End
      Begin VB.ComboBox CboCategory 
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox TxtDescription 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   1320
         Width           =   4935
      End
      Begin VB.CheckBox ChkReportWeekly 
         Caption         =   "Include Ailment in Goverment Weekly Reports"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   2160
         Width           =   4815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Ailment ID"
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
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7200
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ailment Category"
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
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Ailment Description"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List of Maintained Ailments/Diseases"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   7215
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   3015
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
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
   Begin VB.Frame Frame2 
      Caption         =   "Navigate"
      Height          =   6735
      Left            =   7440
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton CmdAdd 
         Caption         =   "New"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton CMdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmAilments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCombo As New ADODB.Recordset
Dim RsRecords As New ADODB.Recordset
Dim BlnEditing As Boolean
Private Sub cmdadd_Click()
    AddMode
    TxtAilmentID = "42"
    TxtAilmentID.Enabled = False
End Sub

Private Sub CMDCLOSE_Click()
    If CMdClose.Caption = "Cancel" Then
        ResetMode
    Else
        Unload Me
    End If
End Sub

Private Sub CmdDelete_Click()
Dim Resp
    Resp = MsgBox("Are you sure you wish to delete this record?", vbInformation + vbYesNo)
        If Resp = vbYes Then
            Conn.Execute "DELETE FROM DIAGNOSIS WHERE DIAGNOSISID = '" & TxtAilmentID & "'"
            MsgBox "Record Deleted Succesfully", vbInformation
            POPULATEGRID
            BlnEditing = False
            ResetMode
        Else
            MsgBox "Deletion aborted", vbInformation
        End If
    ResetMode
End Sub

Private Sub CmdEdit_Click()
    BlnEditing = True
    EditMode
End Sub

Private Sub CmdSave_Click()
    If TxtAilmentID = "" Then MsgBox "Please input the Ailment ID Number", vbExclamation: Exit Sub
    If TxtDescription = "" Then MsgBox "Please input the Ailment Description", vbExclamation: Exit Sub
    If CboCategory = "" Then MsgBox "Please Select the Ailment Category from the Drop-Down List", vbExclamation: Exit Sub
    
    Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM DIAGNOSIS WHERE DIAGNOSISDESCRIPTION = '" & TxtDescription & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If BlnEditing = True Then
                If .BOF = False And .EOF = False Then
                    '!DIAGNOSISID = TxtAilmentID
                    !DIAGNOSISCATEGORY = GetID_NameFromCombo(CboCategory, 1) 'Mid(CboCategory, 1, InStr(CboCategory, "-") - 1)
                    !DIAGNOSISDESCRIPTION = UCase(Replace(TxtDescription, "'", ""))
                    !REPORTWEEKLY = ChkReportWeekly.Value
                    !REPORTMONTHLY = ChkReportMonthly.Value
                    .Update
                    MsgBox "Billing Company Edited Successfully", vbInformation
                    BlnEditing = False
                    ResetMode
                End If
            Else
                If .BOF = False And .EOF = False Then
                    MsgBox "Billing Company with code " & TxtCompanyCode & " is already Maintained", vbExclamation, "Duplication"
                Else
                    .AddNew
                    '!DIAGNOSISID = TxtAilmentID
                    !DIAGNOSISCATEGORY = GetID_NameFromCombo(CboCategory, 1)
                    !DIAGNOSISDESCRIPTION = UCase(Replace(TxtDescription, "'", ""))
                    !REPORTWEEKLY = ChkReportWeekly.Value
                    !REPORTMONTHLY = ChkReportMonthly.Value
                    .Update
                    MsgBox "Ailment Maintained Successfully ", vbInformation
                    ResetMode
                End If
            End If
        End With
    POPULATEGRID
    ResetMode
    BlnEditing = False
End Sub

Private Sub Form_Load()
    'POPULATE COMBO FOR AILMENTS CATEGORY
    If RsCombo.State = 1 Then Set RsCombo = Nothing
    RsCombo.Open "SELECT DIAGNOSISCATEGORYID, DIAGNOSISCATEGORYDESC FROM DIAGNOSISCATEGORY", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboCategory.AddItem String(3 - Len(!DIAGNOSISCATEGORYID), "0") & !DIAGNOSISCATEGORYID & " - " & !DIAGNOSISCATEGORYDESC
                .MoveNext
            Wend
        End With
    RsCombo.Close
    centerform Me
    POPULATEGRID
End Sub
Public Sub AddMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmAilments
    ClearText FrmAilments
End Sub
Public Sub EditMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmAilments
End Sub
Public Sub ResetMode()
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    CMdClose.Caption = "Close"
    GreyOut FrmAilments
    ClearText FrmAilments
End Sub

Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "AILMENT CODE| AILMENT DESCRIPTION | IN WEEKLY REPORT | IN MONTHLY"
    'G.ColDataType(2) = flexDTBoolean
    If RsCombo.State = 1 Then Set RsCombo = Nothing
        RsCombo.Open "SELECT * FROM  DIAGNOSIS order by diagnosisid asc", Conn, adOpenStatic, adLockOptimistic
            If RsCombo.BOF = False And RsCombo.EOF = False Then
                While RsCombo.EOF = False
                    With RsCombo
                        G.AddItem !DIAGNOSISID & vbTab & !DIAGNOSISDESCRIPTION & vbTab & !REPORTWEEKLY & vbTab & !REPORTMONTHLY
                    End With
                RsCombo.MoveNext
                Wend
            End If
        RsCombo.Close
    G.Editable = True
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
End Sub

Private Sub G_SelChange()
    Dim lvCategoryID, lvCategoryDesc As String
    On Error GoTo ErrorHandler
        TxtAilmentID = G.TextMatrix(G.Row, 0)
        TxtDescription = G.TextMatrix(G.Row, 1)
        ChkReportWeekly.Value = G.TextMatrix(G.Row, 2)
        ChkReportMonthly.Value = G.TextMatrix(G.Row, 3)
        lvCategoryID = FindRecord("DIAGNOSIS", "DIAGNOSISCATEGORY", "DIAGNOSISID = '" & TxtAilmentID & "'")
        lvCategoryDesc = FindRecord("DIAGNOSISCATEGORY", "DIAGNOSISCATEGORYDESC", "DIAGNOSISCATEGORYID = '" & lvCategoryID & "'")
        CboCategory = String(3 - Len(lvCategoryID), "0") & lvCategoryID & " - " & lvCategoryDesc
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub TxtAilmentID_Change()
    ValidateDataType TxtAilmentID, 0, "FrmAilments", "TxtAilmentID"
End Sub

