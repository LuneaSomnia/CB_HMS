VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmNightShiftStock 
   Caption         =   "Night Shift Stock"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTabNight 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Current Night Stock"
      TabPicture(0)   =   "FrmNightShiftStock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "List of Night Sift Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   12135
         Begin VSFlex6DAOCtl.vsFlexGrid G 
            Height          =   5655
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   9975
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
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   12015
         Begin VB.TextBox TxtLastQuantity 
            Height          =   405
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox TxtCurrentStock 
            BackColor       =   &H80000009&
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   375
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CommandButton CmbDisburse 
            Caption         =   "Disburse Medicine"
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
            Left            =   9600
            TabIndex        =   10
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox TxtIssuedStock 
            Height          =   375
            Left            =   9600
            TabIndex        =   9
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox CboMedicine 
            Height          =   315
            Left            =   6840
            TabIndex        =   7
            Top             =   480
            Width           =   5055
         End
         Begin VB.ComboBox CboCategory 
            Height          =   315
            Left            =   2160
            TabIndex        =   6
            Top             =   480
            Width           =   4455
         End
         Begin MSComCtl2.DTPicker DTIssueDate 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55050241
            CurrentDate     =   41072
         End
         Begin VB.Label Label6 
            Caption         =   "Last Quantity Issued"
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
            Left            =   4320
            TabIndex        =   16
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label5 
            Caption         =   "Current Night Shift Stock"
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
            Left            =   6840
            TabIndex        =   12
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "Quantity to Issue ?"
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
            Left            =   9600
            TabIndex        =   8
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Type of Medicine Issued"
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
            Left            =   6840
            TabIndex        =   5
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "Medicine Category"
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
            Left            =   2160
            TabIndex        =   4
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Date of Issuing Medicine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "FrmNightShiftStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset

Private Sub CboCategory_Click()
On Error GoTo ErrorHandler
    Dim lvPrescriptionCategoryID As Long
    'POPULATE COMBO FOR DRUGS BY CATEGORY
    CboMedicine.Clear
    lvPrescriptionCategoryID = GetID_NameFromCombo(CboCategory, 1)
    RsRecords.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' order by productname asc", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboMedicine.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboMedicine.AddItem !PRODUCTID & " - " & !ProductName
                End If
                .MoveNext
            Wend
        End With
    RsRecords.Close
    POPULATE_FILTER lvPrescriptionCategoryID
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmbDisburse_Click()
On Error GoTo ErrorHandler
Dim LELIT, OLDSTOCK, lvNewStock, lvmed
Dim Resp
    lvmed = GetID_NameFromCombo(CboMedicine, 2)
            Resp = MsgBox("ARE YOU SURE YOU WANT TO INCREASE THE NIGHT SHIFT STOCK FOR  '" & lvmed & "'?", vbQuestion + vbYesNo)
            If Resp = vbNo Then TxtIssuedStock = 0: Exit Sub
                'CHECK IF STOCK RECORD ALREADY EXISTS AND UPDATE. IF NOT, INSERT NEW RECORD
                LELIT = FindRecord("NIGHT_STOCK_ENTRY", "PRODUCTID", "PRODUCTID = '" & GetID_NameFromCombo(CboMedicine, 1) & "'")
                
                If LELIT <> "" Then
                'BEFORE UPDATE, MOVE PREVIOUS STOCK RECORD TO HISTORY TABLE FOR REPORTING PURPOSES.
                    'ARCHIVE_STOCK lvProductID
                'GET LAST STOCK AND SUM UP WITH INCOMING STOCK COUNT
                    OLDSTOCK = FindRecord("NIGHT_STOCK_ENTRY", "QUANTITYISSUED", "PRODUCTID = '" & GetID_NameFromCombo(CboMedicine, 1) & "'")
                    lvNewStock = OLDSTOCK + TxtIssuedStock
                    Conn.Execute "UPDATE NIGHT_STOCK_ENTRY SET QUANTITYISSUED = '" & lvNewStock & "',IssuingUser = '" & UCase(GlbCurrentUser) & "' WHERE PRODUCTID = '" & GetID_NameFromCombo(CboMedicine, 1) & "'"
                Else
                
                Conn.Execute "INSERT INTO NIGHT_STOCK_ENTRY (PRODUCTCATEGORY,PRODUCTID,PRODUCTDESCRIPTION,QUANTITYISSUED,ISSUINGUSER,DATEISSUED)" & _
                             "Values('" & GetID_NameFromCombo(CboCategory, 1) & "','" & GetID_NameFromCombo(CboMedicine, 1) & "','" & GetID_NameFromCombo(CboMedicine, 2) & "','" & TxtIssuedStock & "','" & GlbCurrentUser & "','" & Format(GlbSysDate, "dd mmm yyyy") & "')"
                End If
                MsgBox "Stock Update for '" & GetID_NameFromCombo(CboMedicine, 2) & "' Has been Saved Succesfully", vbInformation, "Night Stock Inventory Update"
                CmbDisburse.Enabled = False
                POPULATEGRID
    Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
End Sub

Private Sub CboMedicine_Click()
    Dim lvSoldAtNight As String
    TxtCurrentStock = 0: TxtLastQuantity = 0: TxtIssuedStock = 0
    TxtLastQuantity = FindRecord("NIGHT_STOCK_ENTRY", "QUANTITYISSUED", "PRODUCTID = '" & GetID_NameFromCombo(CboMedicine, 1) & "'")
    If TxtLastQuantity.Text = "" Then GoTo SkipCalculation
    lvSoldAtNight = FindRecord("DRUG_SALES_REPORT", "QUANTITY", "PRODUCTID = '" & GetID_NameFromCombo(CboMedicine, 1) & "' AND SHIFT = '1'")
    If lvSoldAtNight = "" Then TxtCurrentStock = TxtLastQuantity: GoTo SkipCalculation
    TxtCurrentStock = (TxtLastQuantity - lvSoldAtNight)
SkipCalculation:
    If TxtCurrentStock = "" Then TxtCurrentStock = 0
    If TxtLastQuantity = "" Then TxtLastQuantity = 0
    CmbDisburse.Enabled = True
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    'POPULATE COMBO FOR PRESCRIPTION CATEGORY
    RsRecords.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY ORDER BY PRODUCTGROUP", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                CboCategory.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                'CboCategory2.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                'CboCategory3.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                .MoveNext
            Wend
        End With
    RsRecords.Close
    DTIssueDate = GlbSysDate
    centerform Me
    POPULATEGRID
    Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
End Sub
Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "PRODUCT CATEGORY ID| PRODUCT ID | PRODUCT DESCRIPTION     | QUANTITY | DATE ISSUED | ISSUING USER "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  NIGHT_STOCK_ENTRY", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !PRODUCTCATEGORY & vbTab & !PRODUCTID & vbTab & !PRODUCTDESCRIPTION & vbTab & !QUANTITYISSUED & vbTab & !DATEISSUED & vbTab & !ISSUINGUSER
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
Public Sub POPULATE_FILTER(ByVal Category)
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "PRODUCT CATEGORY ID| PRODUCT ID | PRODUCT DESCRIPTION     | QUANTITY | DATE ISSUED | ISSUING USER "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  NIGHT_STOCK_ENTRY WHERE PRODUCTCATEGORY = '" & Category & "'", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !PRODUCTCATEGORY & vbTab & !PRODUCTID & vbTab & !PRODUCTDESCRIPTION & vbTab & !QUANTITYISSUED & vbTab & !DATEISSUED & vbTab & !ISSUINGUSER
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

