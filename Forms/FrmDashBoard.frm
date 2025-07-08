VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FrmDashBoard 
   Caption         =   "Dashboard"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17130
   Icon            =   "FrmDashBoard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   17130
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrRefresh 
      Interval        =   20000
      Left            =   9120
      Top             =   4680
   End
   Begin TabDlg.SSTab TabDash 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16965
      _ExtentX        =   29924
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Doctors Waiting List"
      TabPicture(0)   =   "FrmDashBoard.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CboDoctors"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "All Waiting Lists"
      TabPicture(1)   =   "FrmDashBoard.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.ComboBox CboDoctors 
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   3255
      End
      Begin VB.Frame Frame6 
         Caption         =   "Doctors Queue"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9015
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   16575
         Begin VB.CommandButton Command1 
            Caption         =   "Revisits"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11280
            TabIndex        =   31
            Top             =   8400
            Width           =   1215
         End
         Begin VB.TextBox TxtWaiting 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   9480
            TabIndex        =   26
            Text            =   "0"
            Top             =   8400
            Width           =   1575
         End
         Begin VB.TextBox TxtSeen 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4200
            TabIndex        =   25
            Text            =   "0"
            Top             =   8400
            Width           =   1455
         End
         Begin VB.CommandButton CmdCKRefresh 
            Caption         =   "Refresh List(s)"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   14280
            TabIndex        =   22
            Top             =   8400
            Width           =   2175
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridDoc 
            Height          =   3615
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   16335
            _ExtentX        =   28813
            _ExtentY        =   6376
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   65535
            ForeColor       =   0
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   -2147483636
            BackColorAlternate=   49344
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
            Editable        =   -1  'True
            ShowComboButton =   -1  'True
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridDischarged 
            Height          =   3855
            Left            =   0
            TabIndex        =   23
            Top             =   4440
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   6800
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   65535
            ForeColor       =   0
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   -2147483636
            BackColorAlternate=   49344
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
         Begin VSFlex6DAOCtl.vsFlexGrid GridReturns 
            Height          =   3855
            Left            =   11280
            TabIndex        =   29
            Top             =   4440
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   6800
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   65535
            ForeColor       =   0
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   -2147483636
            BackColorAlternate=   49344
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
         Begin VB.Label Label7 
            Caption         =   "Return Patients of the day"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11280
            TabIndex        =   30
            Top             =   4200
            Width           =   4455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Patients/Visitors Waiting:"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            TabIndex        =   27
            Top             =   8400
            Width           =   3255
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Patients/Visitors Already Seen:"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   8400
            Width           =   3855
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   -74880
         TabIndex        =   17
         Top             =   8520
         Width           =   16695
         Begin VB.CommandButton CmdRefresh 
            Caption         =   "Refresh List (s)"
            Height          =   495
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton CmdExit 
            Caption         =   "E&xit"
            Height          =   495
            Left            =   14640
            TabIndex        =   18
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Observation Queue"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -74880
         TabIndex        =   13
         Top             =   600
         Width           =   8535
         Begin VB.TextBox TxtObservation 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   6600
            TabIndex        =   14
            Top             =   3360
            Width           =   1815
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridObservation 
            Height          =   3015
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   5318
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   65535
            ForeColor       =   16711680
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   16711680
            BackColorBkg    =   -2147483636
            BackColorAlternate=   49344
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
         Begin VB.Label Label1 
            Caption         =   "Record Count"
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
            Left            =   5160
            TabIndex        =   16
            Top             =   3480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cashier's Queue"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74880
         TabIndex        =   9
         Top             =   4560
         Width           =   8535
         Begin VB.TextBox TxtCashier 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   6600
            TabIndex        =   10
            Top             =   3480
            Width           =   1815
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridCashier 
            Height          =   3015
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   5318
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   65535
            ForeColor       =   16711680
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   16711680
            BackColorBkg    =   -2147483636
            BackColorAlternate=   49344
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
         Begin VB.Label Label3 
            Caption         =   "Record Count"
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
            Left            =   5160
            TabIndex        =   12
            Top             =   3600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Discharged List"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -66240
         TabIndex        =   5
         Top             =   4560
         Width           =   8055
         Begin VB.TextBox TxtPharmacy 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   6120
            TabIndex        =   6
            Top             =   3480
            Width           =   1815
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridPharmacy 
            Height          =   3015
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5318
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   65535
            ForeColor       =   16711680
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   16711680
            BackColorBkg    =   -2147483636
            BackColorAlternate=   49344
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
            Rows            =   0
            Cols            =   10
            FixedRows       =   0
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
         Begin VB.Label Label4 
            Caption         =   "Record Count"
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
            Left            =   4560
            TabIndex        =   8
            Top             =   3600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Doctors Queue"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -66240
         TabIndex        =   1
         Top             =   600
         Width           =   8055
         Begin VB.TextBox TxtDoctors 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   6120
            TabIndex        =   2
            Top             =   3360
            Width           =   1815
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridDoctors 
            Height          =   3015
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   5318
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   65535
            ForeColor       =   16711680
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   65535
            ForeColorSel    =   16711680
            BackColorBkg    =   -2147483636
            BackColorAlternate=   49344
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
         Begin VB.Label Label2 
            Caption         =   "Record Count"
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
            Left            =   4680
            TabIndex        =   4
            Top             =   3480
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "FrmDashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset
Dim RsLoopClients As New ADODB.Recordset
Dim Rcount As Integer

Private Sub CboDoctors_Change()
    Fill_CK_Grid
End Sub

Private Sub CboDoctors_Click()
    Fill_CK_Grid
End Sub

Private Sub CmdCKRefresh_Click()
    Fill_CK_Grid
    Fill_GridRETURNS
    PopulateDischarged_CK
    CmdRefresh_Click
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Function CALCULATE_AMOUNT(ByVal SCARDNUMBER As String, ByVal SVISITNUMBER As Double, CASH_OR_CREDIT) As String
On Error GoTo Errorhandler
    Dim RsAmount As New ADODB.Recordset
    If RsAmount.State = 1 Then Set RsAmount = Nothing
        Select Case CASH_OR_CREDIT
            Case 1
                RsAmount.Open "SELECT SUM(CASHAMOUNT) AS AMOUNT FROM PRESCRIPTION WHERE CARDNUMBER = '" & SCARDNUMBER & "' AND VISITNUMBER = '" & SVISITNUMBER & "'", Conn, adOpenStatic, adLockOptimistic
                If Not IsNumeric(RsAmount!amount) Then CALCULATE_AMOUNT = 0: Exit Function
            Case 2
                RsAmount.Open "SELECT SUM(CREDITAMOUNT) AS AMOUNT FROM PRESCRIPTION WHERE CARDNUMBER = '" & SCARDNUMBER & "' AND VISITNUMBER = '" & SVISITNUMBER & "'", Conn, adOpenStatic, adLockOptimistic
                If Not IsNumeric(RsAmount!amount) Then CALCULATE_AMOUNT = 0: Exit Function
        End Select
            With RsAmount
                CALCULATE_AMOUNT = !amount
            End With
    Exit Function
Errorhandler:
    MsgBox Err.Description
End Function

Private Sub FillObservation()
    On Error GoTo Errorhandler
   KARI = GlbSysDate: Rcount = 0
   
    GridObservation.Clear
    GridObservation.Rows = 1
    GridObservation.Cols = 9
    For i = 0 To GridObservation.Cols - 1
        GridObservation.ColAlignment(i) = flexAlignCenterCenter
    Next i
    
    GridObservation.ColAlignment(1) = flexAlignCenterCenter
 
    GridObservation.ColDataType(8) = flexDTBoolean
    GridObservation.ColWidth(1) = 3105
    GridObservation.ColWidth(2) = 3990
    GridObservation.FormatString = "CARD NUMBER|PATIENTS FULL NAME  |BILLING COMPANY     |ID NUMBER |BLOOD PRESSURE |WEIGHT | HEIGHT IN METERS |  BMI  |SELECT "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT PATIENT_DETAILS.* FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND COMPLAINS.TOOBSERVATION = '1'", Conn, adOpenDynamic, adLockOptimistic
        
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridObservation.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                        .MoveNext
                        Rcount = Rcount + 1
                    Wend
                End With
            End If
        TxtObservation = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description

End Sub
Private Sub FillWaitingRoom()
    On Error GoTo Errorhandler
   KARI = GlbSysDate: Rcount = 0
    GridDoctors.Clear
    GridDoctors.Rows = 1
    GridDoctors.Cols = 10
    GridDoctors.ColAlignment(1) = flexAlignCenterCenter
    GridDoctors.ColWidth(1) = 3105
    GridDoctors.ColWidth(3) = 4990
    GridDoctors.FormatString = "DOCTOR  | CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |BLOOD PRESSURE |WEIGHT | HEIGHT     |VISIT DATE  "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '0'", Conn, adOpenDynamic, adLockOptimistic
        'ADDED INUSE=0 SG 22012011
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND TODOCTORS = '1' AND INUSE = '0' AND COMPLAINS.CARDNUMBER <> '9999/99' AND DISMISSED = 'FALSE'", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        If Trim(!DOCTOR) = "Null" Then lvdoctor = "NONE" Else lvdoctor = !DOCTOR
                        GridDoctors.AddItem lvdoctor & vbTab & !CardNumber & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !BP & vbTab & !Weight & vbTab & !Height & vbTab & !VisitDate
                        'GridDoc.AddItem lvdoctor & vbTab & !CardNumber & vbTab & !VISITNUMBER & vbTab & !Surname & " " & !FirstName & " " & !SecondName & vbTab & !BillingCompany & vbTab & !BP & vbTab & !Weight & vbTab & !Height & vbTab & !VisitDate
                        .MoveNext
                        Rcount = Rcount + 1
                    Wend
                End With
            End If
        TxtDoctors = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description

End Sub
Private Sub Fill_GridRETURNS()
    On Error GoTo Errorhandler
    Dim TheCount As Integer
    KARI = GlbSysDate: Rcount = 0
    
    GridReturns.Clear
    GridReturns.Rows = 1
    'GridDoc.Cols = 2
    GridReturns.ColWidth(0) = 3000
    GridReturns.FormatString = "PATIENTS FULL NAME | CARD NUMBER"
    GridReturns.ColAlignment(1) = flexAlignCenterCenter
    GridReturns.ColWidth(1) = 3000
        'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '0'", Conn, adOpenDynamic, adLockOptimistic
        'ADDED INUSE=0 SG 22012011
            'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND (TODOCTORS = '1' OR TOOBSERVATION = '1') AND INUSE = '0' and dismissed = 'false'", Conn, adOpenStatic, adLockOptimistic
                
                'Added this Grid on 20012025 'Im reusing code I wrote in 2011 :) A whole 14 years ago. Im impressed with my younger self
                'LOOP TO ESTABLISHED RETURN PATIENTS - SELECT ALL
                If RsGrid.State = 1 Then Set RsGrid = Nothing
                RsGrid.Open "SELECT * FROM COMPLAINS WHERE VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "'", Conn, adOpenStatic, adLockOptimistic
                With RsGrid
                    While .EOF = False
                        'THEN LOOP AND SELECT ALL BY CARD NUMBER AND VISIT DATE
                        RsLoopClients.Open "SELECT * FROM COMPLAINS WHERE CARDNUMBER = '" & GlbCardNumber & "' AND VISITDATE = '" & GlbVisitDate & "'", Conn, adOpenStatic, adLockOptimistic
                        'Wakisema zi drop, use the doc=1 to check if one is visited and the other on isn't. You're Welcome
                            With RsLoopClients
                                While .EOF = False
                                    TheCount = TheCount + 1
                                    .MoveNext
                                Wend
                                
                            End With
                            
                            If TheCount < 1 Then
                                'POPULATE THE FULL NAME ON THE RETURNS GRID
                                lvFullName = FindRecord("PATIENT_DETAILS", "FULLNAME", "CARDNUMBER = '" & GlbCardNumber & "'") 'Potential Error
                                GridReturns.AddItem lvFullName & vbTab & GlbCardNumber
                            End If
                        RsLoopClients.Close
                        RsGrid.MoveNext
                    Wend
                End With
        End If
        'TxtWaiting = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description
   ' Resume
End Sub
Private Function TimeDiff(TimeIn As String, TimeNow As String)
    On Error GoTo Errorhandler
    Dim lvHour As Integer
    Dim lvMinutes As Integer
    
   'FORMAT THE DATE TO 24 HOURS
    TimeIn = Format(TimeIn, "mm/dd/yyyy HH:mm:ss")
    TimeNow = Format(TimeNow, "mm/dd/yyyy HH:mm:ss")
    
    lvHour = CDbl(Mid(TimeNow, 12, 2)) - CDbl(Mid(TimeIn, 12, 2))
    lvHours = String(2 - Len(Trim(lvHour)), "0") & CDbl(lvHour)
    lvMinutes = CDbl(Mid(TimeNow, 15, 2)) - CDbl(Mid(TimeIn, 15, 2))
    lvSeconds = CDbl(Mid(TimeNow, 18, 2)) - CDbl(Mid(TimeIn, 18, 2))
    lvMinute = String(2 - Len(Trim(Right(lvMinutes, 2))), "0") & CDbl(Right(lvMinutes, 2))
    TimeDiff = lvHours & ":" & Replace(lvMinute, "-", "")
    
    Exit Function
Errorhandler:
    MsgBox "Error in calculating Time difference.", vbExclamation
    Exit Function
    Resume
End Function
Private Sub PopulateCashiers()
    On Error GoTo Errorhandler
   Dim TOTAL_PATIENT_CASHAMOUNT As Double
   Dim RecCount As Integer
   KARI = GlbSysDate
    GridCashier.Clear
    GridCashier.Rows = 1
    GridCashier.Cols = 4
    GridCashier.ColAlignment(1) = flexAlignCenterCenter
    GridCashier.ColWidth(1) = 3105
    GridCashier.ColWidth(2) = 3990
    GridCashier.FormatString = "CARD NUMBER | VISIT NUMBER  |   PATIENTS FULL NAME   |   BILLING COMPANY     | CASH AMOUNT | CREDIT AMOUNT"
        If RsGrid.State = adStateOpen Then RsGrid.Close
'        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '1' AND PAID = '0'", Conn, adOpenDynamic, adLockOptimistic
         RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.TOCASHIER =  '1'", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        PATIENT_CASH_AMOUNT = CALCULATE_AMOUNT(!CardNumber, !VISITNUMBER, 1)
                        PATIENT_CREDIT_AMOUNT = CALCULATE_AMOUNT(!CardNumber, !VISITNUMBER, 2)
                        TOTAL_PATIENT_CASHAMOUNT = TOTAL_PATIENT_CASHAMOUNT + PATIENT_CASH_AMOUNT
                        TOTAL_PATIENT_CREDITAMOUNT = (Val(TOTAL_PATIENT_CREDITAMOUNT) + Val(PATIENT_CREDIT_AMOUNT))
                        GridCashier.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & PATIENT_CASH_AMOUNT & vbTab & PATIENT_CREDIT_AMOUNT
                        .MoveNext
                        RecCount = RecCount + 1
                    Wend
                End With
            End If
            TxtCashier = RecCount
            TxtTotalAmount = Format(TOTAL_PATIENT_CASHAMOUNT, "###,###.#0")
            TxtTotalCreditAmount = Format(TOTAL_PATIENT_CREDITAMOUNT, "###,###.#0")
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub PopulateDischarged()
    On Error GoTo Errorhandler
    Dim Rcount As Integer
   KARI = GlbSysDate
    GridPharmacy.Clear
    GridPharmacy.Rows = 1
    GridPharmacy.Cols = 9
    GridPharmacy.ColAlignment(1) = flexAlignCenterCenter
    GridPharmacy.ColWidth(1) = 3105
    GridPharmacy.ColWidth(2) = 3990
    GridPharmacy.FormatString = "DOCTOR     | CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING CO  |BP |WEIGHT | BMI     |VISIT DATE     | NURSE"
    'GridDischarged.FormatString = "DOCTOR     | CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME                 .|   BILLING COMPANY     |BLOOD PRESSURE |WEIGHT | HEIGHT     |VISIT DATE   |NURSE"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '1' AND TOPHARMACY = '1' AND DRUGS <> '1'", Conn, adOpenDynamic, adLockOptimistic
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND DISMISSED = 'True' AND COMPLAINS.CARDNUMBER <> '9999/99'", Conn, adOpenDynamic, adLockOptimistic

            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridPharmacy.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & !FirstName & " " & !SECONDNAME & " " & !SURNAME & vbTab & !BILLINGCOMPANY & vbTab & !BP & vbTab & !Weight & vbTab & !BMINDEX & vbTab & !VisitDate & vbTab & !Nurse
                        'GridDischarged.AddItem !Doctor & vbTab & lvCardNumber & vbTab & !VISITNUMBER & vbTab & lvNames & vbTab & lvBillingCo & vbTab & !BP & vbTab & !Weight & vbTab & !BMINDEX & vbTab & !VisitDate & vbTab & !Nurse
                        .MoveNext
                        Rcount = Rcount + 1
                        'TxtSeen = Rcount
                    Wend
                End With
            End If
            TxtPharmacy = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub PopulateDischarged_CK()
    On Error GoTo Errorhandler
    Dim Rcount As Integer
   KARI = GlbSysDate
    GridDischarged.Clear: Rcount = 0
    GridDischarged.Cols = 7
    'LstPrescription.Clear
    'ClearText FrmPharmacy
    GridDischarged.Rows = 1
    GridDischarged.FormatString = " CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME                 .|   TEXT - CO |DATE |NURSE |  PROCEEDURE "
     GridDischarged.ColWidth(6) = 5000
     GridDischarged.ColWidth(4) = 1500
       If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '1' AND TOPHARMACY = '1' AND DRUGS <> '1'", Conn, adOpenDynamic, adLockOptimistic
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND DISMISSED = 'True'", Conn, adOpenDynamic, adLockOptimistic

            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        If !CardNumber = "9999/99" Then
                            lvNames = FindRecord("NO_OFFICIAL_APPOINTMENTS", "NAMES", "VISITNUMBER = '" & !VISITNUMBER & "'")
                            lvCardNumber = "OFFICIAL"
                            lvBillingCo = !DOCTOR
                        Else
                            lvNames = !FirstName & "  " & !SECONDNAME & "  " & !SURNAME
                            lvCardNumber = !CardNumber
                            lvBillingCo = FindRecord("SERVICE_PROVIDER", "SERVICEPROVIDER", "COMPANYCODE = '" & !BILLINGCOMPANY & "'")
                        End If
                        'GridPharmacy.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & !FirstName & " " & !SECONDNAME & " " & !SURNAME & vbTab & !BILLINGCOMPANY & vbTab & !BP & vbTab & !Weight & vbTab & !BMINDEX & vbTab & !VisitDate & vbTab & !Nurse
                        GridDischarged.AddItem lvCardNumber & vbTab & !VISITNUMBER & vbTab & lvNames & vbTab & lvBillingCo & vbTab & !VisitDate & vbTab & !Nurse & vbTab & !NextProceedure
                        .MoveNext
                        Rcount = Rcount + 1
                        TxtSeen = Rcount
                    Wend
                End With
            End If
            'TxtPharmacy = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description
    Exit Sub
    Resume
End Sub

Private Sub CmdRefresh_Click()
    FillObservation
    FillWaitingRoom
    PopulateCashiers
    PopulateDischarged_CK
End Sub

Private Sub Command1_Click()
    Fill_GridRETURNS
End Sub

Private Sub Form_Load()
    Fill_CK_Grid
    FillObservation
    FillWaitingRoom
    PopulateCashiers
    PopulateDischarged
    PopulateDischarged_CK
    
   'POPULATE DOCTORS COMBO BOX
    If RsGrid.State = 1 Then Set RsGrid = Nothing
    CboDoctors.AddItem "<< ANY >>"
    RsGrid.Open "SELECT * FROM PROFILES WHERE RIGHTS = 'DOCTORS' AND ACCESS = 'TRUE'", Conn, adOpenStatic, adLockOptimistic
        While RsGrid.EOF = False
            CboDoctors.AddItem RsGrid!UserName
            RsGrid.MoveNext
        Wend
    RsGrid.Close
    
    GlbCurrentForm = EnumDashBoard
    centerform Me
    TabDash.Tab = 0
End Sub

Private Sub GridDischarged_SelChange()

End Sub

Private Sub GridDoc_Click()
On Error GoTo Errorhandler
    If GridDoc.TextMatrix(GridDoc.Row, 0) = "" Then Exit Sub
    If Not IsNumeric(GridDoc.TextMatrix(GridDoc.Row, 0)) Then Exit Sub
    If GridDoc.TextMatrix(GridDoc.Row, 0) <> 0 Then
        Resp = MsgBox("Insert Record to Discharged List ?", vbYesNo + vbDefaultButton2 + vbQuestion)
            If Resp = vbYes Then
            
                GlbDropNumber = GridDoc.TextMatrix(GridDoc.Row, 8)
                
                'COMPEL NURSE TO ENTER PROCEEDURE DONE BEFORE DISCHARGE
               If FindRecord("COMPLAINS", "NEXTPROCEEDURE", "VISITNUMBER = '" & GlbDropNumber & "'") = "" Then
                    
                    MsgBox "Please Enter Proceedure done before dropping patient", vbExclamation, "Mandatory Field"
                    GlbDropCancel = True
                    
                        FrmNextVisit.LblFullNames = GridDoc.TextMatrix(GridDoc.Row, 4)
                        FrmNextVisit.Show 1
                    
                        GoTo Drop
                    
                    'Fill_CK_Grid
                    Exit Sub
               End If
Drop:
                DoEvents
                If GlbDropCancel = False Then
                    Conn.Execute "UPDATE COMPLAINS SET TODOCTORS = 'false',TOOBSERVATION= 'FALSE', INUSE = 0, DISMISSED = 'TRUE', doctor = '" & GlbCurrentUser & "' WHERE  VISITNUMBER = '" & GlbDropNumber & "'"
                End If
                GlbDropCancel = False
                
            End If
    End If
    Fill_CK_Grid
    Fill_GridRETURNS
    PopulateDischarged_CK
    Exit Sub
Errorhandler:
    MsgBox Err.Description
    Exit Sub
    Resume
End Sub


Private Sub GridDoc_DblClick()
On Error Resume Next
    'LOAD THE OBSERVATION SCREEN WITH THE PATIENTS VITALS FOR THIS VISIT
    
    'SHOW PROCEEDURE FOR RETURN PATIENTS
        GlbDropView = True
                GlbDropNumber = GridDoc.TextMatrix(GridDoc.Row, 8)
                FrmNextVisit.LblFullNames = GridDoc.TextMatrix(GridDoc.Row, 4)
                FrmNextVisit.DtNextVisit = GlbSysDate
                FrmNextVisit.TxtProceedure = FindRecord("COMPLAINS", "NEXTPROCEEDURE", "VISITNUMBER = '" & GlbDropNumber & "'")
                FrmNextVisit.Show 1
        GlbDropView = False
End Sub


Private Sub TmrRefresh_Timer()
    Fill_CK_Grid
    PopulateDischarged
End Sub
