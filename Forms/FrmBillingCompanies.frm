VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmBillingCompanies 
   Caption         =   "Billing Companies Maintenance"
   ClientHeight    =   7455
   ClientLeft      =   4440
   ClientTop       =   2250
   ClientWidth     =   9120
   Icon            =   "FrmBillingCompanies.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9120
   Begin VB.Frame Frame3 
      Caption         =   "List Of Billing Companies"
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
      TabIndex        =   12
      Top             =   3840
      Width           =   7215
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   3015
         Left            =   120
         TabIndex        =   13
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
      Height          =   7215
      Left            =   7440
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton CMdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   6360
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
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "New"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Company Details"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox TxtContactPerson 
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox TxtAddress 
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   3000
         Width           =   3975
      End
      Begin VB.TextBox TxtTel 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   2400
         Width           =   3975
      End
      Begin VB.CheckBox ChkRecurrent 
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtCompanyName 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1200
         Width           =   5175
      End
      Begin VB.TextBox TxtCompanyCode 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label6 
         Caption         =   "Contact Person"
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
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
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
         TabIndex        =   16
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Telephone"
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
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Reccurrent Holiday"
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
         Left            =   6240
         TabIndex        =   4
         Top             =   1920
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Company Name"
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
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Company Code"
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
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmBillingCompanies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim RsHoliday As New ADODB.Recordset
Dim BlnEditing As Boolean

Public Sub AddMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    CmdClose.Caption = "Cancel"
    ReverseGreyOut FrmBillingCompanies
    ClearText FrmBillingCompanies
End Sub
Public Sub EditMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CmdClose.Caption = "Cancel"
    ReverseGreyOut FrmBillingCompanies
End Sub
Public Sub ResetMode()
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    CmdClose.Caption = "Close"
    GreyOut FrmBillingCompanies
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 4
    G.FormatString = "COMPANY CODE| BILLING COMPANY NAME | CONTACT PERSON | PHONE NUMBER "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  SERVICE_PROVIDER", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !COMPANYCODE & vbTab & !SERVICEPROVIDER & vbTab & !CONTACTPERSON & vbTab & !TelEPHONE
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
End Sub

Private Sub CMDCLOSE_Click()
    If CmdClose.Caption = "Cancel" Then
        ResetMode
    Else
        Unload Me
    End If
End Sub

Private Sub CmdDelete_Click()
Dim Resp
    Resp = MsgBox("Are you sure you wish to delete this record?", vbInformation + vbYesNo)
        If Resp = vbYes Then
            Conn.Execute "DELETE FROM SERVICE_PROVIDER WHERE COMPANYCODE = '" & TxtCompanyCode & "'"
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
    Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM SERVICE_PROVIDER WHERE COMPANYCODE = '" & TxtCompanyCode & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If BlnEditing = True Then
                If .BOF = False And .EOF = False Then
                    !COMPANYCODE = TxtCompanyCode
                    !SERVICEPROVIDER = TxtCompanyName
                    !CONTACTPERSON = TxtContactPerson
                    !TelEPHONE = TxtTel
                    .Update
                    MsgBox "Billing Company Edited Successfully", vbInformation
                    BlnEditing = False
                End If
            Else
                If .BOF = False And .EOF = False Then
                    MsgBox "Billing Company with code " & TxtCompanyCode & " is already Maintained", vbExclamation, "Duplication"
                Else
                    .AddNew
                    !COMPANYCODE = TxtCompanyCode
                    !SERVICEPROVIDER = TxtCompanyName
                    !CONTACTPERSON = TxtContactPerson
                    !TelEPHONE = TxtTel
                    .Update
                    MsgBox "Billing Company Maintained Successfully ", vbInformation
                End If
            End If
        End With
    POPULATEGRID
    ResetMode
    BlnEditing = False
End Sub

Private Sub Form_Load()
    POPULATEGRID
    ResetMode
    centerform Me
End Sub

Private Sub G_Click()
    On Error GoTo ErrorHandler
        TxtCompanyCode = G.TextMatrix(G.Row, 0)
        TxtCompanyName = G.TextMatrix(G.Row, 1)
        TxtContactPerson = G.TextMatrix(G.Row, 2)
        TxtTel = G.TextMatrix(G.Row, 3)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub TxtHolidayDate_Change()

End Sub

Private Sub TxtHolidayName_Change()

End Sub

Private Sub G_DblClick()
    GblServiceProviderID = TxtCompanyCode
    GblServiceProviderName = TxtCompanyName
    FrmPatients.CboBillingCompany.Text = GblServiceProviderID & " - " & TxtCompanyName
    Unload Me
End Sub
