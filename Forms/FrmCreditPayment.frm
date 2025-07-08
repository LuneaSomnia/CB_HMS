VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCreditPayment 
   Caption         =   "Credit Payments"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   Icon            =   "FrmCreditPayment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab CreditTab 
      Height          =   3735
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   6588
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Outstanding Credit Payments"
      TabPicture(0)   =   "FrmCreditPayment.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Partial Payments Received"
      TabPicture(1)   =   "FrmCreditPayment.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Notes"
      TabPicture(2)   =   "FrmCreditPayment.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LstNotes"
      Tab(2).Control(1)=   "CmdSaveNote"
      Tab(2).Control(2)=   "TxtNewNote"
      Tab(2).Control(3)=   "Label11"
      Tab(2).ControlCount=   4
      Begin VB.ListBox LstNotes 
         BackColor       =   &H00C0FFFF&
         Height          =   2400
         ItemData        =   "FrmCreditPayment.frx":035E
         Left            =   -74880
         List            =   "FrmCreditPayment.frx":0360
         TabIndex        =   41
         Top             =   360
         Width           =   8895
      End
      Begin VB.CommandButton CmdSaveNote 
         Caption         =   "Save New Note"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67800
         TabIndex        =   39
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox TxtNewNote 
         Height          =   615
         Left            =   -74880
         TabIndex        =   37
         Top             =   3000
         Width           =   6855
      End
      Begin VB.Frame Frame6 
         Caption         =   "Search Results"
         Height          =   3255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   8895
         Begin VB.TextBox TxtPaymentTotal 
            Height          =   375
            Left            =   6720
            TabIndex        =   33
            Top             =   2760
            Width           =   2055
         End
         Begin VB.TextBox TxtPaymentCount 
            Height          =   375
            Left            =   6720
            TabIndex        =   32
            Top             =   600
            Width           =   2055
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridPayments 
            Height          =   2895
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
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
         Begin VB.Label Label10 
            Caption         =   "Total Amount Paid"
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
            Left            =   6720
            TabIndex        =   36
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Label Label9 
            Caption         =   "Count of Paid Items"
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
            Left            =   6720
            TabIndex        =   35
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Search Results"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   8895
         Begin VB.TextBox TxtCreditEntries 
            Height          =   375
            Left            =   1560
            TabIndex        =   27
            Top             =   2760
            Width           =   2055
         End
         Begin VB.TextBox TxtCreditTotals 
            Height          =   375
            Left            =   6720
            TabIndex        =   26
            Top             =   2760
            Width           =   2055
         End
         Begin VSFlex6DAOCtl.vsFlexGrid Grid 
            Height          =   2415
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4260
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
         Begin VB.Label Label5 
            Caption         =   "Credit Entries"
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
            Left            =   120
            TabIndex        =   30
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Credit Amount"
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
            Left            =   4920
            TabIndex        =   29
            Top             =   2760
            Width           =   1455
         End
      End
      Begin VB.Label Label11 
         Caption         =   "New Notes for Partial Payment"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   2760
         Width           =   2895
      End
   End
   Begin VB.TextBox TxtReceivedAmount 
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   7515
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   9135
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton CmdAccept 
         Caption         =   "Ac&cept Payment"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Repayment"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   9135
      Begin VB.TextBox TxtPrescriptionID 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox TxtCreditAmount 
         Height          =   375
         Left            =   6840
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox TxtQuantity 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox TxtDescription 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1020
         Width           =   7215
      End
      Begin VB.TextBox TxtVisitDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Prescription ID"
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
         Left            =   5160
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Credit Amount"
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
         Left            =   5160
         TabIndex        =   15
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Quantity"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Visit Date"
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.OptionButton OptAll 
         Caption         =   "Display All Creditors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox TxtSearchCard 
         Height          =   405
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   5775
      End
      Begin VB.OptionButton OptCardNumber 
         Caption         =   "Display By Card Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
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
         Left            =   6360
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   7200
      Width           =   9135
      Begin VB.CommandButton CmdAddNote 
         Caption         =   "Add Notes"
         Height          =   495
         Left            =   1800
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptFullPayment 
         Caption         =   "Full Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptPartialPayment 
         Caption         =   "Partial Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3000
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Received Amount"
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
         Left            =   5040
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmCreditPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lvCrCardNumber As String
Dim lvCrVisitNumber As String
Dim RsSearch As New ADODB.Recordset

Private Sub CmdAccept_Click()
On Error GoTo ErrorHandler
    If OptFullPayment.Value = False And OptPartialPayment.Value = False Then MsgBox "Please Select FULL PAYMENT or PARTIAL PAYEMENT to Proceed", vbExclamation: Exit Sub
    If OptFullPayment.Value = True Then
        If TxtPrescriptionID = "" Then Exit Sub
        'Resp = (msgbox "Are you sure you wish to update these Credit transaction?"),vbYesNo
        Resp = MsgBox("Are you sure you wish to confirm the Full Payment for " & UCase(TxtDescription) & " as Received? ", vbExclamation + vbYesNo + vbDefaultButton2)
            If Resp = vbYes Then
                Conn.Execute "UPDATE PRESCRIPTION SET CASHAMOUNT = CREDITAMOUNT,CREDITAMOUNT =0,PAYDATE = '" & Date & "',PAYMENTMODE = 1,CASHIER = '" & GlbCurrentUser & "' WHERE PRESCRIPTIONID = '" & TxtPrescriptionID & "'"
                CmdSearch_Click
                MsgBox "Credit Payment Accepted", vbInformation
            Else
                MsgBox "Credit Payment Cancelled.", vbInformation
            End If
    End If
    
    If OptPartialPayment.Value = True Then
        If TxtPrescriptionID = "" Then Exit Sub
        If TxtReceivedAmount = "" Then MsgBox "Please Enter the Amount First, Received then Accept Payment", vbExclamation: Exit Sub
        Resp = MsgBox("Are you sure you wish to confirm the Partial Payment for " & UCase(TxtDescription) & " as Received? ", vbExclamation + vbYesNo + vbDefaultButton2)
            If Resp = vbYes Then
                'CREATE NEW RECORD TO SHOW THE PARTIAL PAYMENT TRANSACTION WITH CODE 003-(PARTIAL PAYMENT CODE)
                Conn.Execute "INSERT INTO PRESCRIPTION(CARDNUMBER,VISITNUMBER,VISITDATE,BILLINGCO,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT,PAYDATE,PAYMENTMODE,CASHIER,PAYMENTSTATUS)" & _
                             "VALUES ('" & Grid.TextMatrix(Grid.Row, 0) & "','" & Grid.TextMatrix(Grid.Row, 1) & "','" & TxtVisitDate & "','001','003','" & GetID_NameFromCombo(Grid.TextMatrix(Grid.Row, 2), 2) & "','0','" & TxtReceivedAmount & "','" & Format(GlbSysDate, "DD MMM YYYY") & "','1','" & GlbCurrentUser & "','1')"
                
                'NOW REDUCE THE PRINCIPLE CREDIT AMOUNT BY THE AMOUNT RECEIVED.
                Conn.Execute "UPDATE PRESCRIPTION SET CREDITAMOUNT = (" & CDbl(TxtCreditAmount) & " - " & CDbl(TxtReceivedAmount) & ") WHERE PRESCRIPTIONID = '" & TxtPrescriptionID & "'"
                CmdSearch_Click
                MsgBox "Credit Payment Accepted", vbInformation
            Else
                MsgBox "Credit Payment Cancelled.", vbInformation
            End If
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSaveNote_Click()
On Error GoTo ErrorHandler
    Dim lvNote As String
    If TxtNewNote = "" Then MsgBox "Please Enter Prescription Notes for Pharmacyst to Dispense before Saving", vbExclamation: Exit Sub
    lvNote = Format(GlbSysDate, "DD MMM YYYY") & " : " + TxtNewNote
    Conn.Execute "INSERT INTO CREDITORS_NOTES(NOTES,CARDNUMBER,VISITNUMBER)VALUES('" & lvNote & "','" & lvCrCardNumber & "','" & lvCrVisitNumber & "')"
    PopulateCreditorsNotes lvCrCardNumber, lvCrVisitNumber
    TxtNewNote = ""
    CmdSaveNote.Enabled = False
    CmdAccept.Enabled = True
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub
Private Sub PopulateCreditorsNotes(ByVal CardNo, VisitNo)
On Error GoTo ErrorHandler
    'TxtOldNotes = FindRecord("CREDITORS_NOTES", "NOTES", "CARDNUMBER = '" & lvCrCardNumber & "' AND VISITNUMBER = '" & lvCrVisitNumber & "'")
    LstNotes.Clear
    Dim RsNotes As New ADODB.Recordset
        RsNotes.Open "SELECT * FROM CREDITORS_NOTES WHERE CARDNUMBER = '" & lvCrCardNumber & "' AND VISITNUMBER = '" & lvCrVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsNotes
                While .EOF = False
                    LstNotes.AddItem !NOTES
                    .MoveNext
                Wend
            End With
        RsNotes.Close
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub
Private Sub CmdSearch_Click()
On Error GoTo ErrorHandler
    'If TxtSearchCard <> "" Then
    CreditTab.Tab = 0
        Dim SQLStatement As String
        KARI = GlbSysDate
        TxtCreditEntries = 0
        TxtCreditTotals = 0
        Grid.Clear
        Grid.Rows = 1
        Grid.Cols = 7
        Grid.ColAlignment(1) = flexAlignCenterCenter
        'Grid.ColDataType(7) = flexDTBoolean
        Grid.ColWidth(1) = 3105
        Grid.ColWidth(2) = 3990
        Grid.FormatString = "CARD NUMBER| DESCRIPTION |  DESCRIPTION  |   CREDIT AMOUNT   |   VISIT DATE "
        If RsSearch.State = adStateOpen Then RsSearch.Close
            If OptCardNumber.Value = True Then
                Grid.FormatString = "CARD NUMBER| VISIT NUMBVER | ITEM DESCRIPTION |  CREDIT AMOUNT     |  QUANTITY  |  VISIT DATE "
                SQLStatement = "SELECT * FROM PRESCRIPTION WHERE CARDNUMBER = '" & TxtSearchCard & "' AND PAYMENTMODE = '2' AND CREDITAMOUNT > 0 order by visitdate asc"
            ElseIf OptAll = True Then
                Grid.FormatString = "CARD NUMBER| VISIT NUMBER | ITEM DESCRIPTION       |  CREDIT AMOUNT  | QUANTITY  |  VISIT DATE "
                SQLStatement = "SELECT * FROM PRESCRIPTION WHERE  PAYMENTMODE = '2' and CREDITAMOUNT > 0 order by visitdate asc"
                'SQLStatement = "SELECT * FROM PATIENT_DETAILS WHERE PAYMENTMODE = '2' order by visitdate asc"
            End If
                If SQLStatement = "" Then Exit Sub
                RsSearch.Open SQLStatement, Conn, adOpenDynamic, adLockOptimistic
                    If RsSearch.BOF = False And RsSearch.EOF = False Then
                        With RsSearch
                            While Not .EOF
                                If OptCardNumber.Value = True Then
                                    Grid.AddItem !cardnumber & vbTab & !VISITNUMBER & vbTab & !CODE & " - " & !Description & vbTab & Format(!CreditAmount, "#,##0.#0") & vbTab & !Quantity & vbTab & !VisitDate & vbTab & !PRESCRIPTIONID
                                Else
                                    Grid.AddItem !cardnumber & vbTab & !VISITNUMBER & vbTab & !CODE & " - " & !Description & vbTab & Format(!CreditAmount, "#,##0.#0") & vbTab & !Quantity & vbTab & !VisitDate & vbTab & !PRESCRIPTIONID
                                End If
                                TxtCreditEntries = TxtCreditEntries + 1
                                TxtCreditTotals = Format(TxtCreditTotals + !CreditAmount, "#,##0.#0")
                                .MoveNext
                            Wend
                        End With
                    End If
                RsSearch.Close
    'End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdAddNote_Click()
    CreditTab.Tab = 2
End Sub

Private Sub CreditTab_Click(PreviousTab As Integer)
    PopulatePaidItems
End Sub

Private Sub Form_Load()
    centerform Me
    
    TxtNewNote = Format(GlbSysDate, "DD MMM YYYY") & " : "
End Sub

Private Sub Grid_Click()
On Error GoTo ErrorHandler
    lvCrCardNumber = CStr(Grid.TextMatrix(Grid.Row, 0))
    lvCrVisitNumber = Grid.TextMatrix(Grid.Row, 1)
    TxtDescription = Grid.TextMatrix(Grid.Row, 2)
    TxtQuantity = Grid.TextMatrix(Grid.Row, 4)
    TxtCreditAmount = Format(Grid.TextMatrix(Grid.Row, 3), "#,##0.#0")
    TxtVisitDate = Grid.TextMatrix(Grid.Row, 5)
    TxtPrescriptionID = Grid.TextMatrix(Grid.Row, 6)
    TxtReceivedAmount.SetFocus
    
    'POPULATES NOTES RELATED TO THE CARD NUMBER SELECTED
     PopulateCreditorsNotes lvCrCardNumber, lvCrVisitNumber
     
    'PREPARE FOR NEW NOTE
     TxtNewNote = "" 'Format(GlbSysDate, "DD MMM YYYY") & " : "
     CmdSaveNote.Enabled = True
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub
Private Sub PopulatePaidItems()
On Error GoTo ErrorHandler
        GridPayments.Clear
        If Grid.TextMatrix(Grid.Row, 0) = "" Then Exit Sub
        GridPayments.Rows = 1
        GridPayments.Cols = 7
        GridPayments.ColAlignment(1) = flexAlignCenterCenter
        'Grid.ColDataType(7) = flexDTBoolean
        GridPayments.ColWidth(1) = 3105
        GridPayments.ColWidth(2) = 3990
        TxtPaymentCount = 0: TxtPaymentTotal = 0
        GridPayments.FormatString = "CARD NUMBER| VISIT NUMBER | CODE / DESCRIPTION |  AMOUNT  |  PAY DATE "
        If RsSearch.State = adStateOpen Then RsSearch.Close
        RsSearch.Open "SELECT * FROM PRESCRIPTION WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' and VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "' AND CODE = '003'"
            With RsSearch
                While .EOF = False
                    GridPayments.AddItem !cardnumber & vbTab & !VISITNUMBER & vbTab & !CODE & " - " & !Description & vbTab & Format(!CashAmount, "#,##0.#0") & vbTab & !PAYDATE
                    TxtPaymentCount = TxtPaymentCount + 1
                    TxtPaymentTotal = TxtPaymentTotal + !CashAmount
                    .MoveNext
                Wend
                TxtPaymentTotal = Format(TxtPaymentTotal, "#,##0.#0")
            End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub OptFullPayment_Click()
    CmdAccept.Enabled = True
End Sub

Private Sub OptPartialPayment_Click()
    'CmdAccept.Enabled = True
    CreditTab.Tab = 2
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    PopulatePaidItems
End Sub

