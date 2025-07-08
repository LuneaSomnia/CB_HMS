VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmStockAudit 
   Caption         =   "Stock Update History"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14745
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmStockAudit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   14745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   6135
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selected Medicine Stock Details"
      Height          =   3855
      Left            =   6360
      TabIndex        =   20
      Top             =   3480
      Width           =   8295
      Begin VB.OptionButton OptHistory 
         Caption         =   "History Stock Updates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   3360
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton OptCurrent 
         Caption         =   "Current Stock Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   3360
         Width           =   2895
      End
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   2895
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5106
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
   Begin VB.Frame Frame5 
      Caption         =   "All Medicine In Category"
      Height          =   3255
      Left            =   6360
      TabIndex        =   18
      Top             =   120
      Width           =   8295
      Begin VSFlex6DAOCtl.vsFlexGrid GridLevels 
         Height          =   2895
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5106
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   6960
      TabIndex        =   16
      Top             =   4200
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "List of Previous Stock Updates"
      TabPicture(0)   =   "FrmStockAudit.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Current Medicine Stock Levels"
      TabPicture(1)   =   "FrmStockAudit.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Category and Medicine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox CboPrescriptionCategory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   3855
      End
      Begin VB.ComboBox CboDrugs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Medicine Category "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stock Update Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6135
      Begin VB.TextBox TxtReceivedBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox TxtDeliverdBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox TxtDeliveryDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox TxtAdditional 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TxtBeforeCount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Received By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Delivered By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2220
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Count Before Purchase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Delivery Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Additional Purchase Count"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1020
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmStockAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim RsCombo As New ADODB.Recordset

Private Sub CboDrugs_Click()
    OptCurrent.Value = True
    PopulateCurrentStock
End Sub

Private Sub CboPrescriptionCategory_click()
On Error GoTo ErrorHandler
    Dim lvPrescriptionCategoryID
    'POPULATE COMBO FOR PRESCRIPTION
    CboDrugs.Clear
    lvPrescriptionCategoryID = Mid(CboPrescriptionCategory, 1, 3)
    RsRecords.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' ORDER BY PRODUCTNAME", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboDrugs.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboDrugs.AddItem !PRODUCTID & " - " & !ProductName
                End If
                .MoveNext
            Wend
        End With
    RsRecords.Close
    FillMedicineLevels GetID_NameFromCombo(CboPrescriptionCategory, 1)
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
     centerform Me
    'POPULATE COMBO FOR PRESCRIPTION CATEGORY
    RsCombo.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY order by PRODUCTGROUP", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboPrescriptionCategory.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                .MoveNext
            Wend
        End With
    RsCombo.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub FillMedicineLevels(ByVal Category)
    RsRecords.Open "SELECT ProductID, ProductDescription, DeliveryDate, LastStockCount, DeliveredBy, ReceivedBy From STOCK_ENTRY where categoryid = '" & Category & "'", Conn, adOpenStatic, adLockBatchOptimistic
        GridLevels.Clear: GridLevels.Rows = 1
        GridLevels.FormatString = "PRODUCT DESCRIPTION | DELIVERY DATE | CURRENT STOCK | RECIEVED BY"
        With RsRecords
            While .EOF = False
                GridLevels.AddItem !PRODUCTID & " - " & !PRODUCTDESCRIPTION & vbTab & !Deliverydate & vbTab & !LastStockCount & vbTab & !ReceivedBy
                .MoveNext
            Wend
        End With
    RsRecords.Close
End Sub

Private Sub OptCurrent_Click()
    PopulateCurrentStock
End Sub
Private Sub PopulateCurrentStock()
    If CboDrugs = "" Then Exit Sub
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM STOCK_ENTRY WHERE PRODUCTID = '" & GetID_NameFromCombo(CboDrugs, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        Grid.Clear: Grid.Rows = 1: Grid.Cols = 5
        Grid.FormatString = " PRODUCT DESCRIPTION | PRODUCT PURCHASED COUNT | DELIVERY DATE | COUNT BEFORE PURCHASE | RECEIVED BY |"
        With RsRecords
            While .EOF = False
                Grid.AddItem !PRODUCTDESCRIPTION & vbTab & !PRODUCTCOUNT & vbTab & !Deliverydate & vbTab & !LastStockCount & vbTab & !ReceivedBy
                .MoveNext
            Wend
        End With
    RsRecords.Close
End Sub
Private Sub PopulateHistoryStock()
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM STOCK_ENTRY_HISTORY WHERE PRODUCTID = '" & GetID_NameFromCombo(CboDrugs, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        Grid.Clear: Grid.Rows = 1: Grid.Cols = 5
        Grid.FormatString = " PRODUCT DESCRIPTION | PRODUCT PURCHASED COUNT | DELIVERY DATE | COUNT BEFORE PURCHASE | RECEIVED BY |"
        With RsRecords
            While .EOF = False
                Grid.AddItem !PRODUCTDESCRIPTION & vbTab & !PRODUCTCOUNT & vbTab & !Deliverydate & vbTab & !LastStockCount & vbTab & !ReceivedBy
                .MoveNext
            Wend
        End With
    RsRecords.Close
End Sub
Private Sub OptHistory_Click()
    If CboDrugs = "" Then Exit Sub
    PopulateHistoryStock
End Sub

