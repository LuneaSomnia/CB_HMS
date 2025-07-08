VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmWarfarin 
   Caption         =   "WARFARIN CHART"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   Icon            =   "FrmWarfarin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Navigation"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   8160
      Width           =   14775
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save Chart"
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
         Left            =   12600
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Remove Item"
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
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Readings"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   14775
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   5895
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   10398
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
      Caption         =   "Warfarin details"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      Begin VB.TextBox TxtTargatINR 
         Height          =   315
         Left            =   2520
         TabIndex        =   18
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox TxtPTINR 
         Height          =   315
         Left            =   2520
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton CmdAccept 
         Caption         =   "Accept Readings"
         Height          =   495
         Left            =   12480
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TxtNextINR 
         Height          =   315
         Left            =   12960
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtNewDose 
         Height          =   315
         Left            =   8760
         TabIndex        =   2
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox TxtCurrentDose 
         Height          =   315
         Left            =   4740
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DTDate 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53280769
         CurrentDate     =   41319
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "TARGET INR :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Next INR Check"
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
         Left            =   12960
         TabIndex        =   10
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "PT/INR"
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
         Left            =   2520
         TabIndex        =   9
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "New Dose"
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
         Left            =   8760
         TabIndex        =   8
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Current Dose"
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
         Left            =   4800
         TabIndex        =   7
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
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
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmWarfarin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Private Sub CmdAccept_Click()
    G.Cols = 6
    G.FormatString = " DATE              . | PT / INR     . |     CURRENT DOSE                    | NEW DOSE                         |   NEXT INR CHECK        | RN/MD    ."
    
    'LOOP TO CHECK IF DATE HAS ALREADY BEEN USED.
    For J = 1 To G.Rows - 1
        If G.TextMatrix(G.Row, 0) = DTDate Then
            MsgBox "  " & UCase(Format(DTDate, "dd mmmm yyyy")) & "  has already been enterd. Duplicate Dates not allowed", vbExclamation: Exit Sub
        End If
    Next
    
    G.AddItem DTDate & vbTab & TxtPTINR & vbTab & TxtCurrentDose & vbTab & TxtNewDose & vbTab & TxtNextINR & vbTab & TxtNextINR
    
End Sub

Private Sub CmdSave_Click()

    Conn.Execute "DELETE FROM [WARFARIN CHART] WHERE CARDNUMBER = '11'"
    
    For i = 1 To G.Rows - 1
        Conn.Execute "INSERT INTO [WARFARIN CHART](CARDNUMBER,WARFARINDATE,PTINR,CURRENTDOSE,NEWDOSE,NEXTINRCHECK,RNMD) " & _
                     " VALUES( '11','" & G.TextMatrix(i, 0) & "','" & G.TextMatrix(i, 1) & "','" & G.TextMatrix(i, 2) & "','" & G.TextMatrix(i, 3) & "', '" & G.TextMatrix(i, 4) & "', '" & GlbCurrentUser & "')"
    Next
    
    MsgBox "Blood Warfarin Chart Details Saved Succesfully", vbInformation
End Sub

Private Sub Command1_Click()
    G.RemoveItem (G.Row)
End Sub

Private Sub Form_Load()
    centerform Me
    DTDate = GlbSysDate
    
    PopulateWarfarinChart lvOptionalCardNo
End Sub
Private Sub PopulateWarfarinChart(ByRef CardNo)
    G.Cols = 6
    G.FormatString = " DATE              . | PT / INR     . |     CURRENT DOSE                    | NEW DOSE                         |   NEXT INR CHECK        | RN/MD    ."
    'G.AddItem DTDate & vbTab & TxtPTINR & vbTab & TxtCurrentDose & vbTab & TxtNewDose & vbTab & TxtNextINR & vbTab & TxtNextINR
    
    RsRecords.Open "SELECT * FROM [WARFARIN CHART] where CARDNUMBER = '" & CardNo & "' order by warfarindate asc ", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            While .EOF = False
                G.AddItem !WARFARINDATE & vbTab & !PTINR & vbTab & !CURRENTDOSE & vbTab & !NEWDOSE & vbTab & !NEXTINRCHECK & vbTab & !RNMD
                .MoveNext
            Wend
        End With
    RsRecords.Close
End Sub
