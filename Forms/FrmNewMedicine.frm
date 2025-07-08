VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmNewMedicine 
   Caption         =   "New Medincine "
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   Icon            =   "FrmNewMedicine.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Medicine In this Category"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Width           =   8895
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   3255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5741
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
   Begin VB.Frame Frame3 
      Height          =   9255
      Left            =   9120
      TabIndex        =   20
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   880
         Width           =   1575
      End
      Begin VB.CommandButton CMdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   8640
         Width           =   1575
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1520
         Width           =   1575
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "New"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Medicine Maintenance Details Into the System"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.TextBox TxtNoofPills 
         Height          =   315
         Left            =   7920
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CboType 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   1764
         Width           =   2655
      End
      Begin VB.TextBox TxtProductID 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   828
         Width           =   6135
      End
      Begin VB.TextBox TxtPrice 
         Height          =   315
         Left            =   2520
         TabIndex        =   10
         Text            =   "0"
         Top             =   4440
         Width           =   2655
      End
      Begin VB.TextBox TxtReorderLevel 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "30"
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox TxtMinimumLevel 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "10"
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox TxtUnit 
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Text            =   "1"
         Top             =   2232
         Width           =   2655
      End
      Begin VB.TextBox TxtProductName 
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   1296
         Width           =   6135
      End
      Begin VB.ComboBox CboCategory2 
         Height          =   315
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   6135
      End
      Begin VB.ComboBox CboDosage 
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   2700
         Width           =   2655
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   3168
         Width           =   6135
      End
      Begin VB.TextBox TxtDuration 
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         Top             =   3600
         Width           =   2655
      End
      Begin VB.CommandButton CmdAddDosage 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   27
         Top             =   2700
         Width           =   615
      End
      Begin VB.Label lblNoofPills 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of Pills in Bottle"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   34
         Top             =   1800
         Visible         =   0   'False
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "  Optional (Inventory Details)"
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8760
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Medicine Type"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   1764
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Duration  (Days)"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   29
         Top             =   3636
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dosage Remarks"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   28
         Top             =   3168
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dosage"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   26
         Top             =   2700
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Category ID"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   19
         Top             =   360
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   18
         Top             =   1296
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Distribution Unit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   17
         Top             =   2232
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Minimum Level"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   16
         Top             =   4920
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Re Order Level"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5400
         TabIndex        =   15
         Top             =   4920
         Width           =   1575
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sale Price Per Unit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   14
         Top             =   4440
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Product  ID"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   13
         Top             =   828
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmNewMedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim BlnEditing As Boolean
Dim ArrDosage
Dim DosageNumber As Integer
Public Sub AddMode()
    ReverseGreyOut FrmNewMedicine
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    CMdClose.Caption = "Cancel"
    ClearText FrmNewMedicine
End Sub
Public Sub EditMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmNewMedicine
End Sub
Public Sub ResetMode()
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    CMdClose.Caption = "Close"
    GreyOut FrmNewMedicine
    ClearText FrmNewMedicine
End Sub

Private Sub CboCategory2_Click()
    POPULATEGRID
End Sub

Private Sub CboDosage_Click()
    DosageNumber = CboDosage.ListIndex
    StrArr = Split(ArrDosage, ",")
    TxtRemarks = FindRecord("DOSAGES", "DOSAGE", "DOSAGEID = '" & GetID_NameFromCombo(StrArr(DosageNumber), 1) & "'")
End Sub

Private Sub CboType_CLICK()
    If CboType.Text = "BOTTLE" Then
        lblNoofPills.Visible = True
        TxtNoofPills.Visible = True
    Else
        lblNoofPills.Visible = False
        TxtNoofPills.Visible = False
    End If
End Sub

Private Sub cmdadd_Click()
    On Error GoTo ErrorHandler
    lvcategory = CboCategory2.Text
    AddMode
    CboCategory2.Text = lvcategory
    'AddMode
    TxtProductID = "Automatically Updated"
    TxtMinimumLevel = 10
    TxtReorderLevel = 30
    TxtPrice = 0
    BlnEditing = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdAddDosage_Click()
    FrmDosages.Show
    PopulateDosage
End Sub

Private Sub CMDCLOSE_Click()
    If CMdClose.Caption = "Cancel" Then
        ResetMode
    Else
        Unload Me
    End If
End Sub

Private Sub CmdDelete_Click()
On Error GoTo ErrorHandler
    Conn.Execute "DELETE FROM PRODUCTS WHERE PRODUCTID = '" & TxtProductID & "'"
    MsgBox "Product Deleted Succesfully", vbInformation
    POPULATEGRID
    TempCategory = CboCategory2
    ResetMode
    CboCategory2 = TempCategory
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdEdit_Click()
    On Error GoTo ErrorHandler
    EditMode
    BlnEditing = True
    TxtMinimumLevel = 10
    TxtReorderLevel = 30
    If TxtPrice = "" Then TxtPrice = 0
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdSave_Click()
Dim lvcategory As String
On Error GoTo ErrorHandler
                If CboCategory2.Text = "" Then MsgBox "Please Select Category Before Saving", vbExclamation: Exit Sub
                If TxtProductName = "" Then MsgBox "Please Enter Product Name before Saving", vbExclamation: Exit Sub
                'CHECK IF PHARMACY IS PRESENT AND INSIST ON AMOUNT ENTRY
                If TxtPrice = "" Then MsgBox "Please Enter Amount before Saving", vbInformation: Exit Sub
                If TxtProductID = "Automatically Updated" Then
                    lvid = 0
                Else
                    lvid = TxtProductID
                End If
                Set RsRecords = Nothing
                RsRecords.Open "SELECT * FROM PRODUCTS WHERE PRODUCTid = '" & Trim(lvid) & "'", Conn, adOpenStatic, adLockOptimistic
                With RsRecords
                    If BlnEditing = True Then
                        If .BOF = False And .EOF = False Then
                                !CATEGORYID = Mid(CboCategory2, 1, InStr(CboCategory2, "-") - 1)
                                '!PRODUCTID = TxtProductID ' THIS FIELD IS CURRENTLY IDENTITY
                                !ProductName = UCase(Replace(TxtProductName, "'", " "))
                                !PRESCRIPTIONUNIT = TxtUnit
                                !MEDICINETYPE = CboType
                                If TxtNoofPills <> "" Then
                                    !COUNTPERBOTTLE = TxtNoofPills
                                End If
                                !DURATION = TxtDuration
                                !DOSAGE = CboDosage
                                !DosageRemarks = TxtRemarks
                                !MINLEVEL = TxtMinimumLevel
                                !REORDERLEVEL = TxtReorderLevel
                                !MAXLEVEL = TxtMinimumLevel
                                If TxtPrice = 0 Then
                                    Resp = MsgBox("Do you Want to Save without Specifying the Medicine Selling Price?", vbYesNo + vbQuestion)
                                    If Resp = vbNo Then Exit Sub
                                End If
                                !SALEPRICE = TxtPrice
                            .Update
                            POPULATEGRID
                            MsgBox "Product Edited Successfully", vbInformation
                            BlnEditing = False
                            lvcategory = CboCategory2.Text
                            ResetMode
                            CboCategory2.Text = lvcategory
                        End If
                    Else
                        If .BOF = False And .EOF = False Then
                            MsgBox "Medicine with Name " & TxtCompanyCode & " is already Maintained", vbExclamation, "Duplication"
                        Else
                            .AddNew
                                !CATEGORYID = Mid(CboCategory2, 1, InStr(CboCategory2, "-") - 1)
                                '!PRODUCTID = TxtProductID ' THIS FIELD IS CURRENTLY IDENTITY
                                !ProductName = UCase(Replace(TxtProductName, "'", " "))
                                !PRESCRIPTIONUNIT = TxtUnit
                                !MEDICINETYPE = CboType
                                If TxtNoofPills <> "" Then
                                    !COUNTPERBOTTLE = TxtNoofPills
                                End If
                                !DURATION = TxtDuration
                                !DOSAGE = CboDosage
                                !DosageRemarks = TxtRemarks
                                !MINLEVEL = TxtMinimumLevel
                                !REORDERLEVEL = TxtReorderLevel
                                !MAXLEVEL = TxtMinimumLevel
                                If TxtPrice = 0 Then
                                    Resp = MsgBox("Do you Want to Save without Specifying the Medicine Selling Price?", vbYesNo + vbQuestion)
                                    If Resp = vbNo Then Exit Sub
                                End If
                                !SALEPRICE = TxtPrice
                            .Update
                            POPULATEGRID
                            MsgBox "Product Maintained Successfully ", vbInformation
                            'ATTEMPT TO KEEP THE CATEGORY SELECTED DURING SAVING. ITS EASIER WHEN DOING BULK POSTING
                            lvcategory = CboCategory2.Text
                            ResetMode
                            CboCategory2.Text = lvcategory
                        End If
                    End If
                End With
            RsRecords.Close
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub TxtDuration_Change()
    ValidateDataType TxtDuration, 0, "FrmNewMedicine", "TxtDuration"
End Sub

Private Sub TxtProductName_LostFocus()
    TxtProductName = UCase(TxtProductName)
End Sub

Private Sub TxtUnit_Change()
    ValidateDataType TxtUnit, 0, "FrmProducts", "TxtUnit"
End Sub

Private Sub TxtPrice_Change()
    ValidateDataType TxtPrice, 0, "FrmProducts", "TxtPrice"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
        Dim lvPrescriptionCategoryID
        'POPULATE COMBO FOR PRESCRIPTION CATEGORY
        If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY ORDER BY PRODUCTGROUPID", Conn, adOpenDynamic, adLockOptimistic
        
            With RsRecords
                While .BOF = False And .EOF = False
                    CboCategory2.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                    .MoveNext
                Wend
            End With
        RsRecords.Close
        
        'POPULATE COMBO FOR MEDICINE TYPE
        If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM MEDICINE_TYPE ORDER BY MEDICINETYPE ASC", Conn, adOpenDynamic, adLockOptimistic
        
            With RsRecords
                While .BOF = False And .EOF = False
                    CboType.AddItem !MEDICINETYPE
                    .MoveNext
                Wend
            End With
        RsRecords.Close
        
        'POPULATE COMBO FOR DOSAGE
        PopulateDosage
        
        ResetMode
        
        centerform Me
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub
Private Sub PopulateDosage()
On Error GoTo ErrorHandler
        'POPULATE COMBO FOR DOSAGE
        If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM DOSAGES", Conn, adOpenDynamic, adLockOptimistic
            ArrDosage = ""
            CboDosage.Clear
            With RsRecords
                While .BOF = False And .EOF = False
                    If ArrDosage = "" Then
                        ArrDosage = String(3 - Len(!DOSAGEID), "0") & !DOSAGEID & " - " & !DOSAGECODE
                    Else
                        ArrDosage = ArrDosage + "," & String(3 - Len(!DOSAGEID), "0") & !DOSAGEID & " - " & !DOSAGECODE
                    End If
                    'Debug.Print ArrDosage
                    'CboDosage.AddItem String(3 - Len(!DOSAGEID), "0") & !DOSAGEID & " - " & !DOSAGECODE
                    CboDosage.AddItem !DOSAGECODE
                    .MoveNext
                Wend
            End With
        RsRecords.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
Private Sub G_Click()
On Error GoTo ErrorHandler
    TxtProductName = G.TextMatrix(G.Row, 1)
    TxtProductID = G.TextMatrix(G.Row, 0)
    CboType = G.TextMatrix(G.Row, 2)
    CboDosage = G.TextMatrix(G.Row, 3)
    TxtRemarks = G.TextMatrix(G.Row, 4)
    TxtDuration = G.TextMatrix(G.Row, 5)
    TxtUnit = G.TextMatrix(G.Row, 5)
    TxtPrice = G.TextMatrix(G.Row, 7)
    MousePointer = vbNormal
    TxtNoofPills.Visible = True
    lblNoofPills.Visible = True
    TxtNoofPills = G.TextMatrix(G.Row, 6)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub
Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "PRODUCT ID| PRODUCT DESCRIPTION| TYPE |DOSAGE| REMARKS | DURATION | NUMBER OF TABLETS | SALE PRICE PER UNIT "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  PRODUCTS WHERE CATEGORYID = '" & GetID_NameFromCombo(CboCategory2, 1) & "'", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !PRODUCTID & vbTab & !ProductName & vbTab & !MEDICINETYPE & vbTab & !DOSAGE & vbTab & !DosageRemarks & vbTab & !DURATION & vbTab & !COUNTPERBOTTLE & vbTab & !SALEPRICE
                    End With
                RsRecords.MoveNext
                Wend
            End If
        'Set RsRecords = Nothing
    G.Editable = True
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
End Sub

