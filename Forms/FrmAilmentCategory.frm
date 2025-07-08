VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmAilmentCategory 
   Caption         =   "Ailment Categories (Vikundi Vya Magonjwa)"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   Icon            =   "FrmAilmentCategory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Navigate"
      Height          =   5895
      Left            =   6360
      TabIndex        =   7
      Top             =   0
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
         Top             =   5280
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
   Begin VB.Frame Frame2 
      Caption         =   "Category"
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
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6255
      Begin VB.TextBox TxtCategoryName 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   5775
      End
      Begin VB.TextBox TxtCategoryID 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Ailment Category Description"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Category ID"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List Of Medicine Categories"
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
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   6255
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
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
End
Attribute VB_Name = "FrmAilmentCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim BlnEditing As Boolean

Public Sub AddMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmCategory
    ClearText FrmCategory
End Sub
Public Sub EditMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmCategory
End Sub
Public Sub ResetMode()
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    CMdClose.Caption = "Close"
    GreyOut FrmCategory
End Sub

Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "PRODUCT CATEGORY ID| PRODUCT CATEGORY NAME "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  DIAGNOSISCATEGORY", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !DIAGNOSISCATEGORYID & vbTab & !DIAGNOSISCATEGORYDESC
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

Private Sub cmdadd_Click()
    AddMode
    'GET LAST ID USED AND INCREMENT FOR NEXT RECORD
    TxtCategoryID = FindRecord("DIAGNOSISCATEGORY ORDER BY DIAGNOSISCATEGORYID DESC", "DIAGNOSISCATEGORYID") + 1
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
            Conn.Execute "DELETE FROM DIAGNOSISCATEGORY WHERE DIAGNOSISCATEGORYID = '" & TxtCategoryID & "'"
            MsgBox "Record Deleted Succesfully", vbInformation
            POPULATEGRID
            BlnEditing = False
        Else
            MsgBox "Deletion aborted", vbInformation
        End If
    ResetMode
End Sub

Private Sub CmdEdit_Click()
    EditMode
    BlnEditing = True
End Sub

Private Sub CmdSave_Click()
    'If Not IsNumeric(TxtCategoryID) Then MsgBox "Please enter a Valid Number On Category ID Field.", vbExclamation: Exit Sub
    If TxtCategoryName = "" Then MsgBox "Please enter Category Name before Saving", vbExclamation: Exit Sub
    
    Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM DIAGNOSISCATEGORY WHERE DIAGNOSISCATEGORYID = '" & TxtCategoryID & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If BlnEditing = True Then
                If .BOF = False And .EOF = False Then
                    '!PRODUCTGROUPID = Mid(CboServiceProvider, 1, 3)
                    !DIAGNOSISCATEGORYDESC = UCase(TxtCategoryName)
                    .Update
                    MsgBox "Product Category Edited Successfully", vbInformation
                    BlnEditing = False
                End If
            Else
                If .BOF = False And .EOF = False Then
                    MsgBox "Product Category with code " & TxtCompanyCode & " is already Maintained", vbExclamation, "Duplication not allowed"
                Else
                    .AddNew
                    '!PRODUCTGROUPID = TxtCategoryID
                    !DIAGNOSISCATEGORYDESC = UCase(TxtCategoryName)
                    .Update
                    MsgBox "Product Category Successfully Maintained", vbInformation
                End If
            End If
        End With
    POPULATEGRID
    ResetMode
    BlnEditing = False
End Sub
Private Sub G_Click()
    On Error GoTo ErrorHandler
        TxtCategoryID = G.TextMatrix(G.Row, 0)
        TxtCategoryName = G.TextMatrix(G.Row, 1)
        Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    POPULATEGRID
    ResetMode
    centerform Me
End Sub



Private Sub TxtCategoryID_Change()
    ValidateDataType TxtCategoryID, 0, "FrmAilmentCategory", "TxtCategoryID"
End Sub
