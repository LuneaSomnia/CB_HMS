VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPacemaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pacemaker"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
   Icon            =   "FrmPacemaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Comments"
      Height          =   3255
      Left            =   120
      TabIndex        =   56
      Top             =   6720
      Width           =   5655
      Begin VB.TextBox TxtNotes 
         Height          =   2895
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List of History"
      Height          =   3255
      Left            =   5880
      TabIndex        =   50
      Top             =   6720
      Width           =   6495
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   2895
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
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
   Begin VB.Frame Frame2 
      Caption         =   "Data"
      Height          =   3615
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   12255
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   6120
         TabIndex        =   55
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   2280
         TabIndex        =   53
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   10560
         TabIndex        =   49
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   10560
         TabIndex        =   47
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   10560
         TabIndex        =   46
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   10560
         TabIndex        =   45
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   6120
         TabIndex        =   40
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   2400
         TabIndex        =   38
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   2400
         TabIndex        =   37
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   2400
         TabIndex        =   36
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1560
         TabIndex        =   31
         Top             =   840
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         Format          =   138608641
         CurrentDate     =   41322
      End
      Begin VB.Label Label27 
         Caption         =   "High Rate Episodes :"
         Height          =   255
         Left            =   4440
         TabIndex        =   54
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Battery Life :"
         Height          =   255
         Left            =   1320
         TabIndex        =   52
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "APVP"
         Height          =   255
         Left            =   9960
         TabIndex        =   48
         Top             =   2835
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "APVS"
         Height          =   255
         Left            =   9960
         TabIndex        =   44
         Top             =   2385
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "ASVP"
         Height          =   255
         Left            =   9960
         TabIndex        =   43
         Top             =   1875
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "ASVS"
         Height          =   255
         Left            =   9960
         TabIndex        =   42
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Percentage Paced"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   41
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Impendance Lower rate :"
         Height          =   255
         Left            =   4200
         TabIndex        =   39
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "L Wave"
         Height          =   255
         Left            =   1680
         TabIndex        =   35
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "R Wave"
         Height          =   255
         Left            =   1680
         TabIndex        =   34
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "P Wave"
         Height          =   255
         Left            =   1680
         TabIndex        =   33
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Sensing :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Threshold :"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Date :"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Details"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12255
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   10200
         TabIndex        =   26
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   6000
         TabIndex        =   24
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   10200
         TabIndex        =   20
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   6000
         TabIndex        =   18
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   10200
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   6000
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   10200
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6000
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Format          =   138608641
         CurrentDate     =   41322
      End
      Begin VB.Label Label13 
         Caption         =   "Serial Number :"
         Height          =   255
         Left            =   8880
         TabIndex        =   25
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Model Number :"
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Lv Lead :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Serial Number :"
         Height          =   255
         Left            =   8880
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Model Number :"
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Rv Lead :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Serial Number :"
         Height          =   255
         Left            =   8880
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Model Number :"
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Atrial Lead :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Serial Number :"
         Height          =   255
         Left            =   8880
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Model Number :"
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Pacemaker :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date of Implant :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmPacemaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub Form_Load()
    centerform Me
End Sub
