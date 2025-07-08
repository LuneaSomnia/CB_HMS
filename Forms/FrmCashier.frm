VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmCashier 
   Caption         =   "Cashier"
   ClientHeight    =   10185
   ClientLeft      =   5325
   ClientTop       =   1890
   ClientWidth     =   10785
   Icon            =   "FrmCashier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   10785
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh List"
      Height          =   495
      Left            =   4080
      TabIndex        =   49
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Frame Frame6 
      Caption         =   "Post Patient"
      Height          =   855
      Left            =   120
      TabIndex        =   43
      Top             =   8640
      Width           =   10575
      Begin VB.OptionButton OptLab 
         Caption         =   "TO LAB"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   7800
         TabIndex        =   52
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptPharmacy 
         Caption         =   "TO PHARMACY"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   6000
         TabIndex        =   48
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptDoctors 
         Caption         =   "TO DOCTORS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   47
         Top             =   280
         Width           =   1455
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO OBSERVATION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   46
         Top             =   280
         Width           =   1935
      End
      Begin VB.OptionButton OptConsultation 
         Caption         =   "TO CONSULTATION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   280
         Width           =   2055
      End
      Begin VB.CommandButton CmdPost 
         Caption         =   "POST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   44
         Top             =   240
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2535
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Payment Details"
      TabPicture(0)   =   "FrmCashier.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Discharge Settlement"
      TabPicture(1)   =   "FrmCashier.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "GridWard"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CboWard"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CboBed"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Option2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Option1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.OptionButton Option1 
         Caption         =   "Interim Statement"
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
         TabIndex        =   64
         Top             =   1800
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Final Statement"
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
         TabIndex        =   63
         Top             =   2190
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.ComboBox CboBed 
         Height          =   315
         Left            =   -74880
         TabIndex        =   56
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox CboWard 
         Height          =   315
         Left            =   -74880
         TabIndex        =   55
         Top             =   720
         Width           =   3015
      End
      Begin VB.Frame Frame5 
         Height          =   2055
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   10335
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   315
            Left            =   4920
            TabIndex        =   42
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton CmdChequeDetails 
            Caption         =   "..."
            Height          =   315
            Left            =   4920
            TabIndex        =   41
            Top             =   1155
            Width           =   375
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   7680
            TabIndex        =   40
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton CmdNew 
            Caption         =   "New"
            Height          =   495
            Left            =   5760
            TabIndex        =   39
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton CmdEdit 
            Caption         =   "Edit"
            Height          =   495
            Left            =   5760
            TabIndex        =   38
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "Delete"
            Height          =   495
            Left            =   5760
            TabIndex        =   37
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox TxtAmount 
            Height          =   315
            Left            =   1680
            TabIndex        =   34
            Top             =   1560
            Width           =   3135
         End
         Begin VB.CommandButton CmdPayReceived 
            Caption         =   "Payment Received"
            Height          =   495
            Left            =   7680
            TabIndex        =   33
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton CmdPayOffered 
            Caption         =   "Payment Offered"
            Height          =   495
            Left            =   7680
            TabIndex        =   32
            Top             =   240
            Width           =   2535
         End
         Begin VB.ComboBox CMBPayAllocation 
            Height          =   315
            Left            =   1680
            TabIndex        =   31
            Top             =   645
            Width           =   3615
         End
         Begin VB.ComboBox CMBPaymethod 
            Height          =   315
            Left            =   1680
            TabIndex        =   30
            Top             =   1155
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker DTPayment 
            Height          =   315
            Left            =   1680
            TabIndex        =   29
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            Format          =   49610753
            CurrentDate     =   39505
         End
         Begin VB.Label Label9 
            Caption         =   "Payment Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Payment Allocation"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Payment Date"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
      End
      Begin VSFlex6DAOCtl.vsFlexGrid GridWard 
         Height          =   1815
         Left            =   -71760
         TabIndex        =   54
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3201
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
      Begin VB.Label Label13 
         Caption         =   "BED"
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
         Left            =   -74880
         TabIndex        =   58
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "WARD"
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
         Left            =   -74880
         TabIndex        =   57
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton CMDPrintInvoice 
      Caption         =   "Print Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   9600
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5741
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Current Transactions"
      TabPicture(0)   =   "FrmCashier.frx":047A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Payments Drill Down"
      TabPicture(1)   =   "FrmCashier.frx":0496
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pharmacy Direct Sales"
      TabPicture(2)   =   "FrmCashier.frx":04B2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   59
         Top             =   360
         Width           =   10335
         Begin VB.ComboBox CboSaleNumber 
            Height          =   315
            Left            =   120
            TabIndex        =   62
            Top             =   720
            Width           =   2055
         End
         Begin VSFlex6DAOCtl.vsFlexGrid G 
            Height          =   2415
            Left            =   2280
            TabIndex        =   60
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
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
         Begin VB.Label Label14 
            Caption         =   "Sale Number"
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
            TabIndex        =   61
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Archive List"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   10335
         Begin VB.TextBox TxtCashierNotes 
            Height          =   495
            Left            =   1800
            TabIndex        =   65
            Top             =   2160
            Width           =   8415
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridDrillDown 
            Height          =   1815
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   3201
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
         Begin VB.Label Label17 
            Caption         =   "Cashier's Notes:"
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
            TabIndex        =   66
            Top             =   2280
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Current List"
         Height          =   2775
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   10335
         Begin VSFlex6DAOCtl.vsFlexGrid Grid 
            Height          =   2415
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
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
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   17
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   7800
      Width           =   10575
      Begin VB.CommandButton CmdCreditTill 
         Caption         =   "Credit Till"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   8880
         TabIndex        =   53
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalCreditAmount 
         Height          =   375
         Left            =   6240
         TabIndex        =   51
         Top             =   430
         Width           =   2295
      End
      Begin VB.TextBox TxtTotalAmount 
         Height          =   375
         Left            =   3240
         TabIndex        =   36
         Top             =   430
         Width           =   2295
      End
      Begin VB.TextBox TxtCount 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   430
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Total Credit Amount"
         Height          =   255
         Left            =   6240
         TabIndex        =   50
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Total Cash Amount"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Count of Records"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   160
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Card Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.TextBox TxtID 
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox TxtCardNumber 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox TxtSurname 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox TxtFirstName 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox TxtSecondName 
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   855
         Width           =   3255
      End
      Begin VB.ComboBox CboPaymentMode 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   345
         Width           =   4215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "ID Number"
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   " Card Number"
         Height          =   255
         Left            =   5640
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Second Name"
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Surname"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Payment Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CrstlRpt 
      Left            =   6840
      Top             =   9600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8520
      Top             =   5640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"FrmCashier.frx":04CE
      OLEDBString     =   $"FrmCashier.frx":055B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset
Dim RsRecords As New ADODB.Recordset
Dim StrCardNumber As String
Dim strVisitNumber As String
Dim ItemSelected As Integer
Private Sub AddMode()
    DTPayment.Enabled = False
    CMBPayAllocation.Enabled = False
    CMBPaymethod.Enabled = False
    CmdNew.Enabled = False
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdPayOffered.Enabled = True
    CmdPayReceived.Enabled = True
    'CmdExit.Caption = "Cancel"
    Dim CTL As Control
    For Each CTL In Me
        If TypeOf CTL Is ComboBox Then
            CTL.Enabled = True
        End If
    Next
End Sub
Private Sub EditMode()
    CmdNew.Enabled = False
    CmdEdit.Enabled = False
    CmdDelete.Enabled = True
    CmdPayOffered.Enabled = True
    CmdPayReceived.Enabled = True
    'CmdExit.Caption = "Cancel"
End Sub
Public Sub ManageProcessFlow(ActiveForm)
On Error GoTo ERRORHANDLER
    Dim RsControls As New ADODB.Recordset
    RsControls.Open "SELECT * FROM PROCESSFLOW WHERE SCREENID = '" & ActiveForm & "'", Conn, adOpenStatic, adLockOptimistic
        If RsControls.EOF = False Then
            With RsControls
                If !CONSULTATION = 1 Then OptConsultation.Item(1).Enabled = True
                If !OBSERVATION = 1 Then OptObservation.Item(2).Enabled = True
                If !DOCTORS = 1 Then OptDoctors.Item(3).Enabled = True
                If !PHARMACY = 1 Then OptPharmacy.Item(4).Enabled = True
                If !LAB = 1 Then OptLab.Item(5).Enabled = True
            End With
        End If
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub PopulateDrillDown()
On Error GoTo ERRORHANDLER
   Dim TOTAL_PATIENTAMOUNT As Double
   Dim RecCount As Integer
   Dim lvPayMode As String
   'Dim CAshMultiplyByUnits As Integer
   'RESET AMOUNT
   TxtTotalAmount = Format(0, "###,###.#0")
   TxtTotalCreditAmount = Format(0, "###,###.#0")
   KARI = GlbSysDate
    GridDrillDown.Clear
    GridDrillDown.Rows = 1
    GridDrillDown.Cols = 5
    GridDrillDown.ColAlignment(1) = flexAlignCenterCenter
    GridDrillDown.ColWidth(1) = 3105
    GridDrillDown.ColWidth(2) = 3990
    GridDrillDown.FormatString = "CARD NUMBER|VISIT NUMBER|CHARGES ID|DESCRIPTION|MODE OF PAYMENT|CASH AMOUNT |CREDIT AMOUNT"
        If RsGrid.State = adStateOpen Then RsGrid.Close
'        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '1' AND PAID = '0'", Conn, adOpenDynamic, adLockOptimistic
         'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.TOCASHIER =  '1'", Conn, adOpenDynamic, adLockOptimistic
         RsGrid.Open "SELECT PRESCRIPTION.CARDNUMBER,PRESCRIPTION.VISITNUMBER,PRESCRIPTION.PRESCRIPTIONID,PRESCRIPTION.DESCRIPTION,PRESCRIPTION.PAYMENTMODE,PRESCRIPTION.QUANTITY,PRESCRIPTION.CASHAMOUNT,PRESCRIPTION.CREDITAMOUNT FROM PRESCRIPTION WHERE CARDNUMBER = '" & StrCardNumber & "' AND VISITNUMBER = '" & strVisitNumber & "' AND PRESCRIPTION.PAYMENTSTATUS = '0'", Conn, adOpenStatic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        'PATIENT_AMOUNT = CALCULATE_AMOUNT(!CARDNUMBER, !VISITNUMBER)
                        'TOTAL_PATIENTAMOUNT = TOTAL_PATIENTAMOUNT + PATIENT_AMOUNT
                        If !PAYMENTMODE = 1 Then
                            lvPayMode = "CASH"
                        ElseIf !PAYMENTMODE = 2 Then
                            lvPayMode = "CREDIT"
                        ElseIf !PAYMENTMODE = 3 Then
                            lvPayMode = "CHEQHE"
                        End If
                        
                        GridDrillDown.AddItem !CARDNUMBER & vbTab & !VISITNUMBER & vbTab & !PRESCRIPTIONID & vbTab & !Description & vbTab & lvPayMode & vbTab & !CashAmount & vbTab & !CreditAmount
                                   RecCount = RecCount + 1
                        TxtTotalCreditAmount = TxtTotalCreditAmount + !CreditAmount
                        TOTAL_PATIENTAMOUNT = TOTAL_PATIENTAMOUNT + !CashAmount
                        .MoveNext
                Wend
                End With
            End If
            TxtCount = RecCount
            TxtTotalCreditAmount = Format(TxtTotalCreditAmount, "###,###.#0")
            TxtTotalAmount = Format(TOTAL_PATIENTAMOUNT, "###,###.#0")
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub ResetMode()
    DTPayment.Enabled = False
    CMBPayAllocation.Enabled = False
    CMBPaymethod.Enabled = False

    CmdNew.Enabled = True
    CmdEdit.Enabled = True
    CmdDelete.Enabled = True
    CmdPayOffered.Enabled = False
    CmdPayReceived.Enabled = False
    CmdExit.Caption = "Exit"
End Sub

Private Function CALCULATE_AMOUNT(ByVal SCARDNUMBER As String, ByVal SVISITNUMBER As Double, CASH_OR_CREDIT) As String
On Error GoTo ERRORHANDLER
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
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Function
Private Sub AllocateMonies(ByVal Cash1Credit2Cheque3 As Integer)
On Error GoTo ERRORHANDLER
    'VALIDATE FIELDS ARE NOT BLANK OR INVALID
    If TxtAmount = "" Then MsgBox "PLEASE ENTER PAYMENT AMOUNT BEFORE POSTING PAYMENT", vbExclamation: Exit Sub

    'INSERT INTO CHARGES TABLE INITIALLY CALLED PRESCRIPTION. CAPTION REMAINS AS SUCH. SG 22012011
    Select Case Cash1Credit2Cheque3
        Case 1
            Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT,PAYDATE,PAYMENTMODE,VISITDATE,BILLINGCO,CASHIER)" & _
                 "VALUES('" & Grid.TextMatrix(Grid.Row, 0) & "','" & Grid.TextMatrix(Grid.Row, 1) & "','" & Mid(CMBPayAllocation, 1, 3) & "','" & Mid(CMBPayAllocation, 6, Len(CMBPayAllocation)) & "','0', '" & TxtAmount & "','" & Format(DTPayment, "DD MMM YYYY") & "','" & Cash1Credit2Cheque3 & "'," & _
                 "'" & Format(DTPayment, "DD MMM YYYY") & "','" & Left(CboPaymentMode, 3) & "','" & StrCurrentUser & "')"
        Case 2
            Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,CODE,DESCRIPTION,QUANTITY,CREDITAMOUNT,PAYDATE,PAYMENTMODE,VISITDATE,BILLINGCO,CASHIER)" & _
                 "VALUES('" & Grid.TextMatrix(Grid.Row, 0) & "','" & Grid.TextMatrix(Grid.Row, 1) & "','" & Mid(CMBPayAllocation, 1, 3) & "','" & Mid(CMBPayAllocation, 6, Len(CMBPayAllocation)) & "','0', '" & TxtAmount & "','" & Format(DTPayment, "DD MMM YYYY") & "','" & Cash1Credit2Cheque3 & "'," & _
                 "'" & Format(DTPayment, "DD MMM YYYY") & "','" & Left(CboPaymentMode, 3) & "','" & StrCurrentUser & "')"
        Case 3
            Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT,PAYDATE,PAYMENTMODE,VISITDATE,BILLINGCO,CASHIER)" & _
                 "VALUES('" & Grid.TextMatrix(Grid.Row, 0) & "','" & Grid.TextMatrix(Grid.Row, 1) & "','" & Mid(CMBPayAllocation, 1, 3) & "','" & Mid(CMBPayAllocation, 6, Len(CMBPayAllocation)) & "','0', '" & TxtAmount & "','" & Format(DTPayment, "DD MMM YYYY") & "','" & Cash1Credit2Cheque3 & "'," & _
                 "'" & Format(DTPayment, "DD MMM YYYY") & "','" & Left(CboPaymentMode, 3) & "','" & StrCurrentUser & "')"
    End Select
    'CHECK IF THIS WAS A CHEQUE PAYMENT AND UPDATE THE CHEQUE AMOUNT ACCORDINGLY.
    If Left(CMBPaymethod, 2) = "03" Then
       Conn.Execute "UPDATE CHEQUE_PAYMENTS SET AMOUNT = " & TxtAmount & "  WHERE  CARDNUMBER='" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'"
    End If
    PopulateDrillDown
    MsgBox "Record Generated Succesfully", vbInformation
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub
Private Sub PopulateCashiers()
    On Error GoTo ERRORHANDLER
   Dim TOTAL_PATIENT_CASHAMOUNT As Double
   Dim RecCount As Integer
   KARI = GlbSysDate
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 4
    Grid.ColAlignment(1) = flexAlignCenterCenter
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER | VISIT NUMBER  |   PATIENTS FULL NAME   |   BILLING COMPANY     | CASH AMOUNT | CREDIT AMOUNT"
        If RsGrid.State = adStateOpen Then RsGrid.Close
'        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '1' AND PAID = '0'", Conn, adOpenDynamic, adLockOptimistic
         RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.TOCASHIER =  '1'", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        PATIENT_CASH_AMOUNT = CALCULATE_AMOUNT(!CARDNUMBER, !VISITNUMBER, 1)
                        PATIENT_CREDIT_AMOUNT = CALCULATE_AMOUNT(!CARDNUMBER, !VISITNUMBER, 2)
                        TOTAL_PATIENT_CASHAMOUNT = TOTAL_PATIENT_CASHAMOUNT + PATIENT_CASH_AMOUNT
                        TOTAL_PATIENT_CREDITAMOUNT = (Val(TOTAL_PATIENT_CREDITAMOUNT) + Val(PATIENT_CREDIT_AMOUNT))
                        Grid.AddItem !CARDNUMBER & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & PATIENT_CASH_AMOUNT & vbTab & PATIENT_CREDIT_AMOUNT
                        .MoveNext
                        RecCount = RecCount + 1
                    Wend
                End With
            End If
            TxtCount = RecCount
            TxtTotalAmount = Format(TOTAL_PATIENT_CASHAMOUNT, "###,###.#0")
            TxtTotalCreditAmount = Format(TOTAL_PATIENT_CREDITAMOUNT, "###,###.#0")
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub
Private Sub PopulateWardGrid(ByVal WardNo, ByVal BedNo)
    On Error GoTo ERRORHANDLER
   Dim lvPaitentName As String
   Dim RecCount As Integer
   KARI = GlbSysDate
    GridWard.Clear
    GridWard.Rows = 1
    GridWard.Cols = 4
    GridWard.ColAlignment(1) = flexAlignCenterCenter
    GridWard.ColWidth(1) = 3105
    GridWard.ColWidth(2) = 3990
    GridWard.FormatString = "CARD NUMBER | VISIT NUMBER  |   PATIENTS FULL NAME  "
        
        If RsGrid.State = adStateOpen Then RsGrid.Close
         'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.TOCASHIER =  '1'", Conn, adOpenDynamic, adLockOptimistic
         RsGrid.Open "SELECT CARDNUMBER, VISITNUMBER FROM INPATIENTS  WHERE WARDNUMBER = '" & Mid(CboWard, 1, 1) & "'  AND BEDNUMBER =  '" & Mid(CboBed, 1, 1) & "' ", Conn, adOpenStatic, adLockOptimistic
         
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        lvPaitentName = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & !CARDNUMBER & "'")
                        GridWard.AddItem !CARDNUMBER & vbTab & !VISITNUMBER & vbTab & lvPaitentName '& " " & !FIRSTNAME & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & PATIENT_CASH_AMOUNT & vbTab & PATIENT_CREDIT_AMOUNT
                        .MoveNext
                    Wend
                End With
            End If
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Description
End Sub

Private Sub CboBed_Click()
    PopulateWardGrid Mid(CboWard, 1, 1), Mid(CboBed, 1, 1)
End Sub

Private Sub CboSaleNumber_Click()
    POPULATEGRID CboSaleNumber
End Sub

Private Sub CboWard_Click()
Dim RsRecords As New ADODB.Recordset

    'POPULATE BEDS FROM WARDS
    CboBed.Clear
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT BEDNUMBER FROM BEDs_AVAILABILITY WHERE WARDNUMBER = '" & Mid(CboWard, 1, 1) & "' AND OCCUPIEDBY = 1", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboBed.AddItem RsRecords!BEDNUMBER '& " - " & RsRecords!WARDDESCRIPTION
            RsRecords.MoveNext
        Wend
    RsRecords.Close
End Sub

Private Sub CMBPayAllocation_Change()
    'TxtAmount = 0
End Sub

Private Sub CMBPayAllocation_Click()
TxtAmount = 0
End Sub

Private Sub CMBPaymethod_Click()
    If Left(CMBPaymethod, 2) = "03" Then
        CmdChequeDetails_Click
    End If
End Sub

Private Sub CmdCancel_Click()
    ResetMode
End Sub

Private Sub CmdChequeDetails_Click()
    StrDocCardNo = StrCardNumber
    StrDocVisitNumber = strVisitNumber
    FrmChequesMaintenance.Show
End Sub

Private Sub CmdCreditTill_Click()
On Error GoTo ERRORHANDLER
    'AUTOMATE SENDING TO WARD IF PAYMENT IS DEPOSIT ' TODO
    
    Select Case SSTab1.Tab
    Case 2
            If CboSaleNumber = "" Then MsgBox "Please Select Sale Number Before Crediting Till", vbInformation: Exit Sub
            Conn.Execute "INSERT INTO DRUGS_SALES SELECT * FROM PRE_DRUGS_SALES WHERE SALENUMBER = '" & CboSaleNumber & "'"
            MsgBox "Purchase Posted Succesfully", vbInformation
            Conn.Execute "DELETE FROM PRE_DRUGS_SALES WHERE SALENUMBER = '" & CboSaleNumber & "'"
            POPULATEGRID 0
            RePopulateSalesCombo
    
    Case Else
        If Grid.Rows < 2 Then Exit Sub
        
        If TxtCashierNotes = "" Then TxtCashierNotes = 0 'SITAKI IKUWE BLANK.
        'UPDATE PAYMENTSTATUS
        Conn.Execute "UPDATE PRESCRIPTION SET PAYDATE = '" & GlbSysDate & "', PAYMENTSTATUS = 1, CASHIER = '" & GlbCurrentUser & "',NOTES = '" & TxtCashierNotes & "' WHERE VISITNUMBER = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 1) & "' AND PAYMENTMODE = 1"
        MsgBox "Cash accepted and Credited to Till Succesfully", vbInformation
    End Select
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdDelete_Click()
    If SSTab1.Tab = 1 Then
    Dim Resp
    Resp = MsgBox("Are you Sure you wish to Delete this record?", vbQuestion + vbYesNo)
            If Resp = vbNo Then MsgBox "Deletion aborted", vbInformation: Exit Sub
        Conn.Execute "DELETE FROM PRESCRIPTION WHERE CARDNUMBER = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 0) & "' AND VISITNUMBER = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 1) & "' AND PRESCRIPTIONID = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 2) & "'"
            PopulateDrillDown
            MsgBox "Record Deleted Succesfully", vbInformation
            ResetMode
    Else
        'Conn.Execute "DELETE FROM PRESCRIPTION WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'"
            ResetMode
    End If
End Sub

Private Sub CmdEdit_Click()
    EditMode
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdNew_Click()
    If TxtCardNumber = "" Then MsgBox "Please Select Patient record from list below before adding Payment entries", vbExclamation: Exit Sub
    AddMode
End Sub

Private Sub CmdPayOffered_Click()
    AllocateMonies (2)
    ResetMode
End Sub

Private Sub CmdPayReceived_Click()
    AllocateMonies (1)
    ResetMode
End Sub

Private Sub CmdPost_Click()
    Dim StrCardNumber As String
    KARI = GlbSysDate
    StrCardNumber = Grid.TextMatrix(Grid.Row, 0)
        Select Case ItemSelected
       Case 1
            'To Consultation
                DUMMY = SendPatient(EnumConsultation, StrCardNumber, KARI)
                'FrmPatients.Show
                'Unload Me
        Case 2
            'To Observation
                DUMMY = SendPatient(EnumObservation, StrCardNumber, KARI)
                'FrmWaitingRoom.Show
                'Unload Me
       Case 3
            'To Doctors
                DUMMY = SendPatient(EnumDoctors, StrCardNumber, KARI)
                'FrmCashier.Show
                'Unload Me
        Case 4
            'To Pharmacy
                DUMMY = SendPatient(EnumPharmacy, StrCardNumber, KARI)
                'FrmPharmacy.Show
                'Unload Me
        Case 5
            'To Lab
                DUMMY = SendPatient(EnumLab, StrCardNumber, KARI)
                Conn.Execute "UPDATE COMPLAINS SET DOCTOR = 'FROM CASHIER',INUSE = '0'  WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITDATE = '" & Format(GlbSysDate, "DDMMMYYYY") & "'"
        End Select
        SSTab1.Tab = 0
        PopulateCashiers
        'GridDrillDown.Clear
End Sub

Private Sub CMDPrintInvoice_Click()
On Error GoTo ERRORHANDLER
'''    Dim Sconn As New ADODB.Connection
'''    Sconn.ConnectionString = "Provider=SQLOLEDB.1;Password=Today123;Persist Security Info=True;User ID=SA;Initial Catalog=OUTPATIENT;Data Source=CAP-D102A413A57"
'''    Sconn.Open
    If Conn.State <> 1 Then Exit Sub
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "Provider=SQLOLEDB.1;Password=CES123;Persist Security Info=True;User ID=SA;Initial Catalog=OUTPATIENT;Data Source=SYB-KEN-NB-002\SQL2005"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=Today123;DSQ=SYB-KEN-NB-002\SQL2005;"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\PAYMENTInvoice.rpt"
                   .WindowTitle = App.EXEName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                    .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 0) & "' and {PRESCRIPTION.VISITNUMBER} = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 1) & "'"
                   '.SelectionFormula = " {PRESCRIPTION.VISITNUMBER} = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 1) & "'"
                   If InStr(1, CrstlRpt.SelectionFormula, "VISIT NUMBER") > 1 Then
                        .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} =  '" & StrCardNumber & "'"
                        .SelectionFormula = " {PRESCRIPTION.VISITNUMBER} = '" & strVisitNumber & "'"
                    End If
                   .Destination = 0
                   .WindowState = crptMaximized
                   .Action = 1
                End With
Exit Sub
ERRORHANDLER:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CmdSave_Click()

'VALIDATE FIELDS ARE NOT BLANK OR INVALID
    'ToDo
    
'INSERT INTO CHARGES TABLE INITIALLY CALLED PRESCRIPTION. CAPTION REMAINS AS SUCH.
    Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,CODE,DESCRIPTION,QUANTITY,AMOUNT,PAYMENTMODE)" & _
    "VALUES('" & Grid.TextMatrix(Grid.Row, 0) & "','" & Grid.TextMatrix(Grid.Row, 1) & "','" & Mid(CMBPayAllocation, 1, 3) & "','" & Mid(CMBPayAllocation, 6, Len(CMBPayAllocation)) & "','0', '" & TxtAmount & "','" & Mid(CMBPaymethod, 1, 1) & "')"
End Sub

Private Sub Command1_Click()
    Load FrmConsultation
    FrmConsultation.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
Dim RsRecords As New ADODB.Recordset

    'SET DEFAULT TABS.
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    
    'POPULATE GRID
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM PAYMENT_MODES", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CMBPaymethod.AddItem String(2 - Len(RsRecords!PAYMENTCODE), "0") & RsRecords!PAYMENTCODE & " - " & RsRecords!PAYMENTDESCRIPTION
            RsRecords.MoveNext
        Wend
    RsRecords.Close
        
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM PAY_ALLOCATION", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CMBPayAllocation.AddItem RsRecords!PAYCODE & " - " & RsRecords!PAYDESCRIPTION
            RsRecords.MoveNext
        Wend
    RsRecords.Close
    
    'POPULATE WARDS COMBO
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM WARDS", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboWard.AddItem RsRecords!WARDNUMBER & " - " & RsRecords!WARDDESCRIPTION
            RsRecords.MoveNext
        Wend
    RsRecords.Close
        
    'POPULATE COMBO FOR PHARMACY SALES
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT DISTINCT SALENUMBER FROM PRE_DRUGS_SALES", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboSaleNumber.AddItem RsRecords!SaleNumber
            RsRecords.MoveNext
        Wend
    RsRecords.Close
        
    centerform Me
    PopulateCashiers
    POPULATEGRID 0
    ResetMode
    DTPayment.Value = GlbSysDate
    
    GlbCurrentForm = EnumCashier
    ManageProcessFlow EnumCashier
End Sub

Private Sub Grid_Click()
'''    StrCardNumber = Grid.TextMatrix(Grid.Row, 0)
'''    strVisitNumber = Grid.TextMatrix(Grid.Row, 1)
End Sub

Private Sub Grid_DblClick()
On Error GoTo ERRORHANDLER
    TxtCardNumber = Grid.TextMatrix(Grid.Row, 0)
    StrCardNumber = Grid.TextMatrix(Grid.Row, 0)
    strVisitNumber = Grid.TextMatrix(Grid.Row, 1)
    
    Dim RSCASHIER As New ADODB.Recordset
    If RSCASHIER.State = 1 Then Set RSCASHIER = Nothing
    RSCASHIER.Open "SELECT * FROM PATIENT_DETAILS WHERE CARDNUMBER = '" & TxtCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
    With RSCASHIER
        If .BOF = False And .EOF = False Then
            TxtFirstname = !FirstName
            TxtSecondName = !SECONDNAME
            TxtSurname = !SURNAME
            TxtID = !IDNUMBER
            CboPaymentMode = !BILLINGCOMPANY
        End If
    End With
Exit Sub
ERRORHANDLER:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub PrintInvoice_Click()
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   .Connect = Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\PAYMENTInvoice.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = " & Grid.TextMatrix(Grid.Row, 0) & ""
                   .SelectionFormula = " {PRESCRIPTION.VISITNUMBER} = " & Grid.TextMatrix(Grid.Row, 1) & ""
                   .Destination = 0
                   .Action = 1
                End With
End Sub

Private Sub OptAdmission_Click(Index As Integer)
    ItemSelected = OptAdmission.Item(Index).Index
End Sub

Private Sub GridDrillDown_DblClick()
    FrmDiscount.LblPatient = TxtCardNumber & " - " & TxtFirstname & "  " & TxtSecondName
    FrmDiscount.TxtDiscountItem = GridDrillDown.TextMatrix(GridDrillDown.Row, 2) & " - " & GridDrillDown.TextMatrix(GridDrillDown.Row, 3)
    FrmDiscount.TxtOriginalAmount = GridDrillDown.TextMatrix(GridDrillDown.Row, 5)
    
    FrmDiscount.Show 1
    PopulateDrillDown
End Sub

Private Sub GridWard_CLICK()
    StrCardNumber = GridWard.TextMatrix(GridWard.Row, 0)
    TxtCardNumber = StrCardNumber
    'Todo GET NAMES AND POPULATE ON TEXTBOXES
    '
    strVisitNumber = Trim(GridWard.TextMatrix(GridWard.Row, 1))
    SSTab1.Tab = 1
   ' PopulateDrillDown
End Sub

Private Sub OptConsultation_Click(Index As Integer)
    ItemSelected = OptObservation.Item(Index).Index
End Sub

Private Sub OptDoctors_Click(Index As Integer)
    ItemSelected = OptDoctors.Item(Index).Index
End Sub

Private Sub OptLab_Click(Index As Integer)
    ItemSelected = OptLab.Item(Index).Index
End Sub

Private Sub OptObservation_Click(Index As Integer)
    ItemSelected = OptObservation.Item(Index).Index
End Sub

Private Sub OptPharmacy_Click(Index As Integer)
    ItemSelected = OptPharmacy.Item(Index).Index
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    PopulateDrillDown
End Sub
Public Sub POPULATEGRID(ByVal SaleNumber As Double)
On Error GoTo ERRORHANDLER
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "CATEGORY ID|  PRODUCT CATEGORY NAME    | QUANTITY | AMOUNT"
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        If SaleNumber = 0 Then
            RsRecords.Open "SELECT * FROM  PRE_DRUGS_SALES WHERE DOCTOR0PHARMACY1 = '1' ", Conn, adOpenStatic, adLockOptimistic
        Else
            RsRecords.Open "SELECT * FROM  PRE_DRUGS_SALES WHERE DOCTOR0PHARMACY1 = '1' AND SALENUMBER = '" & CboSaleNumber & "'", Conn, adOpenStatic, adLockOptimistic
        End If
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !CATEGORYID & vbTab & !PRODUCTID & " - " & !PRODUCTDESCRIPTION & vbTab & !Quantity & vbTab & !amount
                    End With
                RsRecords.MoveNext
                Wend
            End If
        RsRecords.Close
    G.Editable = True
Exit Sub
ERRORHANDLER:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub
Private Sub RePopulateSalesCombo()
    'POPULATE COMBO FOR PHARMACY SALES
    CboSaleNumber.Clear
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT DISTINCT SALENUMBER FROM PRE_DRUGS_SALES", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboSaleNumber.AddItem RsRecords!SaleNumber
            RsRecords.MoveNext
        Wend
    RsRecords.Close
End Sub
