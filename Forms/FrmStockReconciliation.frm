VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmStockReconciliation 
   Caption         =   "Stock Taking Reconciliation"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrstlRpt 
      Left            =   7320
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame5 
      Caption         =   "Discrepancy Count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   6120
      Width           =   8415
      Begin VB.CheckBox ChkDiscrepancy 
         Caption         =   "Display Discrepancy"
         Height          =   255
         Left            =   6360
         TabIndex        =   18
         Top             =   0
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox TxtDiscrepancy 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Stock Discrepancy"
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
         Left            =   1200
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   8415
      Begin VB.CommandButton CmdClearRecon 
         Caption         =   "Clear Old Reconciliation"
         Height          =   495
         Left            =   3360
         TabIndex        =   21
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton CmdReconciliationReport 
         Caption         =   "Reconciliation Report"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit Reconciliation"
         Height          =   495
         Left            =   6360
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Logical Count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   8415
      Begin VB.CheckBox ChkDisplayLogicCount 
         Caption         =   "Display Logical Count"
         Height          =   255
         Left            =   6360
         TabIndex        =   17
         Top             =   0
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox TxtLogicalCount 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Logical Stock Count"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Physical Count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   8415
      Begin VB.CommandButton CmdReconcile 
         Caption         =   "Reconcile"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6360
         TabIndex        =   19
         Top             =   300
         Width           =   1935
      End
      Begin VB.TextBox TxtPhysicalCount 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Physical Stock Count"
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
         Left            =   1080
         TabIndex        =   6
         Top             =   405
         UseMnemonic     =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.ComboBox CboCategory3 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   6135
      End
      Begin VB.ComboBox CboProduct3 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   6135
      End
      Begin VB.Label Label16 
         Caption         =   "Medicine Category ID "
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
         TabIndex        =   4
         Top             =   375
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "Medicine Description"
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
         Left            =   120
         TabIndex        =   3
         Top             =   975
         Width           =   1935
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid GridLevels 
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3413
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
Attribute VB_Name = "FrmStockReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Private Sub CboCategory3_click()
    Dim lvPrescriptionCategoryID As Long
    'POPULATE COMBO FOR DRUGS BY CATEGORY
    CboProduct3.Clear
    lvPrescriptionCategoryID = Mid(CboCategory3, 1, 3)
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT PRODUCTID, PRODUCTDESCRIPTION FROM STOCK_ENTRY WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' ORDER BY PRODUCTDESCRIPTION ASC", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboProduct3.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboProduct3.AddItem !PRODUCTID & " - " & !PRODUCTDESCRIPTION
                End If
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    FillMedicineLevels GetID_NameFromCombo(CboCategory3, 1)
End Sub

Private Sub CboProduct3_Click()
    TxtPhysicalCount = ""
    TxtLogicalCount = ""
    TxtDiscrepancy = ""
End Sub

Private Sub CmdClearRecon_Click()
Dim Resp
    Resp = MsgBox("Are you Sure you wish to clear Old Reconciliation? This will Delete the Reconciliation Report.", vbQuestion + vbYesNo)
    If Resp = vbYes Then
        Conn.Execute "DELETE FROM STOCK_RECONCILIATION"
        MsgBox "Reconciliation Report Deleted Succesfully", vbInformation
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdReconcile_Click()
    Dim lvLogicalStockCount As Integer
    lvLogicalStockCount = FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & GetID_NameFromCombo(CboProduct3.Text, 1) & "'")
    TxtLogicalCount = lvLogicalStockCount
    TxtDiscrepancy = (TxtLogicalCount - TxtPhysicalCount)
    
    'INSERT INTO OR UPDATE RECONCILIATION TABLE
    If FindRecord("STOCK_RECONCILIATION", "COUNT(PRODUCTID)", "PRODUCTID = '" & GetID_NameFromCombo(CboProduct3, 1) & "'") <> 0 Then
        Resp = MsgBox("Physical Stock has already been recorded for " & UCase(GetID_NameFromCombo(CboProduct3, 2)) & " Do you want to Overwrite?", vbQuestion + vbYesNo)
        If Resp = vbYes Then
            Conn.Execute "UPDATE STOCK_RECONCILIATION SET PHYSICALSTOCKCOUNT = '" & TxtPhysicalCount & "',DISCREPANCY = '" & TxtDiscrepancy & "' WHERE PRODUCTID = '" & GetID_NameFromCombo(CboProduct3.Text, 1) & "'"
        End If
    Else
    
        Conn.Execute "INSERT INTO STOCK_RECONCILIATION (CATEGORYID,PRODUCTID,PHYSICALSTOCKCOUNT,LOGICALSTOCKCOUNT,DISCREPANCY,RECONDATE)" & _
                 "VALUES ('" & GetID_NameFromCombo(CboCategory3, 1) & "','" & GetID_NameFromCombo(CboProduct3, 1) & "','" & TxtPhysicalCount & "', '" & lvLogicalStockCount & "','" & TxtDiscrepancy & "', '" & GlbSysDate & "')"
    End If
End Sub

Private Sub CmdReconciliationReport_Click()
On Error GoTo ErrorHandler
    If Conn.State <> 1 Then Exit Sub
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "RECONCILIATION REPORT"
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\ReconciliationReport.rpt"
                   .WindowTitle = StrCompanyName & " - " & "STOCK TAKE RECONCILIATION REPORT"
                   .Destination = 0
                   .WindowState = crptMaximized
                   .Action = 1
                End With
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub Form_Load()
    'POPULATE COMBO FOR PRESCRIPTION CATEGORY
    RsRecords.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY ORDER BY PRODUCTGROUP", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                CboCategory3.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                .MoveNext
            Wend
        End With
    RsRecords.Close
    centerform Me
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

Private Sub TxtPhysicalCount_Change()
    KISH = ValidateDataType_Advice(TxtPhysicalCount, 0)
    If KISH = False Then TxtPhysicalCount.Text = Mid(TxtPhysicalCount, 1, Len(TxtPhysicalCount) - 1)
    If TxtPhysicalCount.Text <> "" Then
        CmdReconcile.Enabled = True
    Else
        CmdReconcile.Enabled = False
    End If
End Sub
