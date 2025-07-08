VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBPChart 
   Caption         =   "Blood Pressure Chart"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   Icon            =   "FrmBPChart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   9360
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
      TabIndex        =   9
      Top             =   8160
      Width           =   9135
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
         Left            =   6840
         TabIndex        =   11
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
         TabIndex        =   10
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
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   9135
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   5895
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
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
      Caption         =   "Pressure"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
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
      Width           =   9135
      Begin VB.CommandButton CmdAccept 
         Caption         =   "Accept Readings"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TxtBPEvenings 
         Height          =   315
         Left            =   7320
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox TxtTimeEvenings 
         Height          =   315
         Left            =   4980
         TabIndex        =   14
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox TxtBPMornings 
         Height          =   315
         Left            =   7320
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtTimeMornings 
         Height          =   315
         Left            =   4980
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTDate 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   41319
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Evenings"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   18
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "BP"
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
         Left            =   7320
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Time"
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
         Left            =   5040
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mornings"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Time of Day"
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
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "BP"
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
         Left            =   7320
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Time"
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
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   495
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
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmBPChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Private Sub CmdAccept_Click()
    If TxtTimeMornings = "" Then MsgBox "Please Enter the Time Morning BP was taken before Saving Chart", vbInformation: Exit Sub
    If TxtTimeEvenings = "" Then MsgBox "Please Enter the Time Evening BP was taken before Saving Chart", vbInformation: Exit Sub
    If TxtBPMornings = "" Then MsgBox "Please Enter Morning BP reading before Saving Chart", vbInformation: Exit Sub
    If TxtBPEvenings = "" Then MsgBox "Please Enter Evening BP reading before Saving Chart", vbInformation: Exit Sub
    
    'LOOP TO CHECK IF DATE HAS ALREADY BEEN USED.
    For J = 1 To G.Rows - 1
        If G.TextMatrix(G.Row, 0) = DTDate Then
            MsgBox "Blood Pressure for  " & UCase(DTDate) & "  has already been enterd. Duplicate Dates not allowed", vbExclamation: Exit Sub
        End If
    Next
    
    G.AddItem DTDate & vbTab & "MORNING" & vbTab & TxtTimeMornings & vbTab & TxtBPMornings & vbTab & "EVENING" & vbTab & TxtTimeEvenings & vbTab & TxtBPEvenings
End Sub

Private Sub CmdSave_Click()

    Conn.Execute "DELETE FROM BLOOD_PRESSURE_CHART WHERE CARDNUMBER = '11'"
    
    For i = 1 To G.Rows - 1
        Conn.Execute "INSERT INTO BLOOD_PRESSURE_CHART(CARDNUMBER,READINGDATE,MORNINGREADINGTIME,MORNINGREADING,EVENINGREADINGTIME,EVENINGREADING) " & _
                     " VALUES( '11','" & G.TextMatrix(i, 0) & "','" & G.TextMatrix(i, 2) & "','" & G.TextMatrix(i, 3) & "','" & G.TextMatrix(i, 5) & "', '" & G.TextMatrix(i, 6) & "')"
    Next
    
    MsgBox "Blood Pressure Details Saved Succesfully", vbInformation
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    G.RemoveItem (G.Row)
End Sub

Private Sub Form_Load()
    centerform Me
    DTDate = GlbSysDate
    
    PopulateBPChart lvOptionalCardNo
End Sub
Private Sub PopulateBPChart(ByRef CardNo)
    G.Clear
    G.Cols = 7
    G.FormatString = "CAPTURE DATE | TIME OF DAY | TIME | BLOOD PRESSURE | TIME OF DAY | TIME | BLOOD PRESSURE"
        If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM BLOOD_PRESSURE_CHART where CARDNUMBER = '" & CardNo & "'", Conn, adOpenStatic, adLockOptimistic
            With RsRecords
                While .EOF = False
                    G.AddItem !READINGDATE & vbTab & "MORNING" & vbTab & !MORNINGREADINGTIME & vbTab & !MORNINGREADING & vbTab & "EVENING" & vbTab & !EVENINGREADINGTIME & vbTab & !EVENINGREADING
                    .MoveNext
                Wend
            End With
End Sub
