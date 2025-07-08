VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmDosages 
   Caption         =   "Dossage Frequecy Maintenance"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmDosages.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Navigate"
      Height          =   5175
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton CmdAdd 
         Caption         =   "New"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton CMdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List of Dosage Frequencies"
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
      Top             =   1800
      Width           =   6255
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   3015
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
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
      Caption         =   "Dosage description"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.TextBox TxtDosageDesc 
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox TxtShortCode 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox CboServiceProvider 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Dosage Code"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Dosage Description"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Service Provider"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmDosages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim BlnEditing As Boolean
Dim StrDosageID As Integer

Public Sub AddMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    CmdClose.Caption = "Cancel"
    ReverseGreyOut FrmConsultation
    ClearText FrmConsultation
End Sub
Public Sub EditMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CmdClose.Caption = "Cancel"
    ReverseGreyOut FrmConsultation
End Sub
Public Sub ResetMode()
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    CmdClose.Caption = "Close"
    GreyOut FrmConsultation
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Public Sub POPULATEGRID()
On Error GoTo ERRORHANDLER
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "DOSAGE ID| DOSAGE DESCRIPTION         |DOSAGE SHORT CODE "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  DOSAGES", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem String(3 - Len(!DOSAGEID), "0") & !DOSAGEID & vbTab & !DOSAGE & vbTab & !DOSAGECODE
                    End With
                RsRecords.MoveNext
                Wend
            End If
        RsRecords.Close
    G.Editable = True
Exit Sub
ERRORHANDLER:
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmdadd_Click()
    AddMode
End Sub

Private Sub CMDCLOSE_Click()
    If CmdClose.Caption = "Cancel" Then
        ResetMode
    Else
        Unload Me
    End If
End Sub

Private Sub CmdDelete_Click()
On Error GoTo ERRORHANDLER
Dim Resp
    Resp = MsgBox("Are you sure you wish to delete this record?", vbInformation + vbYesNo)
        If Resp = vbYes Then
            Conn.Execute "DELETE FROM DOSAGES WHERE DOSAGEID = '" & G.TextMatrix(G.Row, 0) & "'"
            MsgBox "Record Deleted Succesfully", vbInformation
            POPULATEGRID
            BlnEditing = False
        Else
            MsgBox "Deletion aborted", vbInformation
        End If
    ResetMode
Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & Err.Description

End Sub

Private Sub CmdEdit_Click()
    EditMode
    BlnEditing = True
End Sub

Private Sub CmdSave_Click()
On Error GoTo ERRORHANDLER
    Set RsRecords = Nothing
    If BlnEditing = True Then
        RsRecords.Open "SELECT * FROM DOSAGES WHERE DOSAGEID = '" & StrDosageID & "'", Conn, adOpenStatic, adLockOptimistic
    Else
        RsRecords.Open "SELECT * FROM DOSAGES", Conn, adOpenStatic, adLockOptimistic
    End If
        With RsRecords
            If BlnEditing = True Then
            If RsRecords.State = 1 Then Set RsRecords = Nothing
                If .BOF = False And .EOF = False Then
                    !DOSAGE = TxtDosageDesc
                    !DOSAGECODE = TxtShortCode
                   .Update
                    MsgBox "Dosages Edited Successfully", vbInformation
                    BlnEditing = False
                End If
            Else
                    .AddNew
                    !DOSAGE = TxtDosageDesc
                    !DOSAGECODE = TxtShortCode
                    .Update
                    MsgBox "Dosage Successfully Maintained", vbInformation
            End If
        End With
    POPULATEGRID
    ResetMode
    BlnEditing = False
    Exit Sub
ERRORHANDLER:
MsgBox Err.Number & Err.Description
End Sub
Private Sub G_Click()
    On Error GoTo ERRORHANDLER
        CboServiceProvider.Text = G.TextMatrix(G.Row, 0)
        TxtDosageDesc = G.TextMatrix(G.Row, 1)
        TxtShortCode = G.TextMatrix(G.Row, 2)
        Exit Sub
ERRORHANDLER:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()

    'POPULATE GRID
'''    If RsRecords.State = 1 Then Set RsRecords = Nothing
'''    RsRecords.Open "SELECT * FROM SERVICE_PROVIDER", Conn, adOpenStatic, adLockOptimistic
'''        While RsRecords.EOF = False
'''            CboServiceProvider.AddItem RsRecords!PROVIDERID & " - " & RsRecords!PROVIDERNAME
'''            RsRecords.MoveNext
'''        Wend
'''    RsRecords.Close

    POPULATEGRID
    ResetMode
    centerform Me
End Sub


Private Sub G_DblClick()
    StrDosageID = G.TextMatrix(G.Row, 0)
End Sub
