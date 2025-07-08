VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmTreatmentOld 
   Caption         =   "Doctors Diagnosis"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14160
   Icon            =   "FrmTreatmentOld.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame12 
      Height          =   9495
      Left            =   12120
      TabIndex        =   40
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton CmdFoodFrequency 
         Caption         =   "Food Frequency"
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
         TabIndex        =   74
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton CmdMeasurements 
         Caption         =   "View Measurements"
         Height          =   495
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdConclude 
         Caption         =   "Conclude Treatment"
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
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
         TabIndex        =   3
         Top             =   8880
         Width           =   1695
      End
      Begin VB.CommandButton CMDSchedule 
         Caption         =   "Schedule Visit"
         Height          =   495
         Left            =   120
         TabIndex        =   43
         Top             =   5520
         Width           =   1695
      End
      Begin VB.CommandButton CmdPrintReferral 
         Caption         =   "Print Referral"
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   7920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CmdPrescription 
         Caption         =   "Print Prescription"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   7320
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdClearHistory 
      Caption         =   "Clear History Data"
      Height          =   495
      Left            =   10320
      TabIndex        =   31
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame Frame9 
      Caption         =   "Patient History"
      Height          =   2175
      Left            =   120
      TabIndex        =   30
      Top             =   1200
      Width           =   11895
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
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
   End
   Begin Crystal.CrystalReport CrstlRpt 
      Left            =   6840
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame8 
      Caption         =   "Post Patient"
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   8880
      Width           =   11895
      Begin VB.OptionButton OptLab 
         Caption         =   "TO LAB"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   8760
         TabIndex        =   34
         Top             =   350
         Width           =   975
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
         Left            =   9960
         TabIndex        =   2
         Top             =   180
         Width           =   1815
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO CONSULTATION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   350
         Width           =   2175
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO OBSERVATION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   2630
         TabIndex        =   28
         Top             =   350
         Width           =   2055
      End
      Begin VB.OptionButton OptCashier 
         Caption         =   "TO CASHIER"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   5020
         TabIndex        =   1
         Top             =   350
         Width           =   1455
      End
      Begin VB.OptionButton OptPharmacy 
         Caption         =   "TO PHARMACY"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   6810
         TabIndex        =   27
         Top             =   350
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab DocTab 
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lab Test Request"
      TabPicture(0)   =   "FrmTreatmentOld.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "CmdLaboratory"
      Tab(0).Control(2)=   "Frame10"
      Tab(0).Control(3)=   "ChkOverride"
      Tab(0).Control(4)=   "CmdViewScan"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Diagnosis"
      TabPicture(1)   =   "FrmTreatmentOld.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdPharmacy"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "CmdRemoveDiagnosis"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Prescription"
      TabPicture(2)   =   "FrmTreatmentOld.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Doctors Assessment"
      TabPicture(3)   =   "FrmTreatmentOld.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Referral"
      TabPicture(4)   =   "FrmTreatmentOld.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   71
         Top             =   360
         Width           =   11655
         Begin VB.TextBox TxtAssesment 
            Height          =   2295
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   11415
         End
      End
      Begin VB.CommandButton CmdViewScan 
         Caption         =   "View Full Image"
         Height          =   375
         Left            =   -65640
         TabIndex        =   69
         Top             =   2040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame Frame6 
         Caption         =   "Referral Notes"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   11655
         Begin VB.CommandButton CmdReferral 
            Caption         =   "To Referral"
            Height          =   495
            Left            =   9720
            TabIndex        =   53
            Top             =   2040
            Width           =   1815
         End
         Begin VB.TextBox TxtReferal 
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   240
            Width           =   11415
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1680
            TabIndex        =   50
            Top             =   2100
            Width           =   7095
         End
         Begin VB.Label Label17 
            Caption         =   "Referral Hospital"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   2160
            Width           =   1335
         End
      End
      Begin VB.CommandButton CmdRemoveDiagnosis 
         Caption         =   "Remove Diagnosis"
         Height          =   495
         Left            =   -64800
         TabIndex        =   48
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   46
         Top             =   360
         Width           =   9975
         Begin VSFlex6DAOCtl.vsFlexGrid GridDiagnosis 
            Height          =   2295
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   4048
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
      Begin VB.CheckBox ChkOverride 
         Caption         =   "Generate Diagnosis and Prescription with Data from Laboratory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72240
         TabIndex        =   39
         Top             =   2760
         Width           =   6975
      End
      Begin VB.Frame Frame10 
         Caption         =   "Lab Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -68640
         TabIndex        =   32
         Top             =   360
         Width           =   5535
         Begin VB.TextBox TxtLabResults 
            Height          =   1935
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   240
            Width           =   5295
         End
         Begin VB.Image ImgPreview 
            Height          =   1995
            Left            =   120
            Picture         =   "FrmTreatmentOld.frx":04CE
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   5280
         End
      End
      Begin VB.CommandButton CmdPharmacy 
         Caption         =   "To Phramacy"
         Height          =   495
         Left            =   -66240
         TabIndex        =   23
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CommandButton CmdLaboratory 
         Caption         =   "To Laboatory"
         Height          =   495
         Left            =   -67080
         TabIndex        =   22
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Request"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   5895
         Begin VB.TextBox TxtLabRequest 
            Height          =   1935
            HideSelection   =   0   'False
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            ToolTipText     =   "Tests Requested From laboratory"
            Top             =   240
            Width           =   5655
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   11655
         Begin VB.CommandButton CmdRemoveDrug 
            Caption         =   "Remove Drug"
            Height          =   495
            Left            =   120
            TabIndex        =   68
            Top             =   1920
            Width           =   1695
         End
         Begin VB.ComboBox CboPrescriptionCategory 
            Height          =   315
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton CmdAddDrug 
            Caption         =   "Add Drug"
            Height          =   495
            Left            =   2040
            TabIndex        =   60
            Top             =   1920
            Width           =   1695
         End
         Begin VB.ComboBox CboPaymentMode 
            Height          =   315
            Left            =   120
            TabIndex        =   59
            Top             =   1440
            Width           =   3615
         End
         Begin VB.ComboBox CboDosages 
            Height          =   315
            Left            =   480
            TabIndex        =   58
            Top             =   2040
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox CboDrugs 
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   3615
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridPrescription 
            Height          =   2295
            Left            =   3840
            TabIndex        =   63
            Top             =   240
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4048
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
         Begin VB.TextBox TxtPrescription 
            Height          =   1335
            Left            =   5160
            TabIndex        =   56
            Top             =   840
            Width           =   4695
         End
         Begin VB.ListBox LstPrescription 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
            Height          =   420
            IntegralHeight  =   0   'False
            ItemData        =   "FrmTreatmentOld.frx":15858
            Left            =   4080
            List            =   "FrmTreatmentOld.frx":1585A
            MultiSelect     =   2  'Extended
            TabIndex        =   55
            Top             =   1800
            Width           =   5055
         End
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            Height          =   315
            Left            =   2880
            TabIndex        =   61
            Top             =   2040
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "Prescription Drug"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Prescription Category"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Payment Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Dosage"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   2040
            Visible         =   0   'False
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Complaints"
      Height          =   2295
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   11895
      Begin VB.CommandButton CmdAddComplains 
         Caption         =   "Add Complains"
         Height          =   495
         Left            =   10200
         TabIndex        =   70
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CmdAddDiagnosis 
         Caption         =   "Add Diagnosis"
         Height          =   495
         Left            =   10200
         TabIndex        =   45
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox CboDiagnosis 
         Height          =   315
         Left            =   6120
         TabIndex        =   36
         Top             =   1680
         Width           =   3615
      End
      Begin VB.ComboBox CboDiagnosisCategory 
         Height          =   315
         Left            =   1680
         TabIndex        =   35
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox TxtComplaints 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   9975
      End
      Begin VB.Label Label12 
         Caption         =   "DISEASE:"
         Height          =   255
         Left            =   5280
         TabIndex        =   38
         Top             =   1755
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Diagnosis Category"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observations"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.TextBox TxtBMI 
         Height          =   285
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DtCurrDate 
         Height          =   285
         Left            =   8280
         TabIndex        =   18
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   53215233
         CurrentDate     =   39415
      End
      Begin VB.TextBox TxtHeight 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtWeight 
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtBp 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtFirstname 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox TxtSecondName 
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label BMI 
         Alignment       =   2  'Center
         Caption         =   "BMI"
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
         Left            =   10080
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10320
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Date"
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
         Left            =   7680
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Height"
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
         Left            =   5400
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Weight"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Blood Pressure"
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
         Left            =   120
         TabIndex        =   11
         Top             =   705
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name"
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
         Left            =   480
         TabIndex        =   10
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Second Name"
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
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmTreatmentOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsfill As New ADODB.Recordset
Dim RsGrid As New ADODB.Recordset
Dim RsFilter As New ADODB.Recordset
Dim RsImageStore As New ADODB.Recordset
Dim KARI As Date
Dim ItemSelected As Integer
Dim StrArray
Dim lvDirectSalesNo As Integer
Public Sub DisableControls()
    Dim CTL As Control
    For Each CTL In Me.Controls
        If TypeOf CTL Is TextBox Then
            CTL.Locked = True
        End If
    Next
    'SELECTIVE DISABLING
        CmdLaboratory.Enabled = False
        CmdReferral.Enabled = False
        'CmdAdmission.Enabled = False
        CmdPharmacy.Enabled = False
        CmdPost.Enabled = False
End Sub
Public Sub ManageProcessFlow(ActiveForm)
On Error GoTo ErrorHandler
    Dim RsControls As New ADODB.Recordset
    RsControls.Open "SELECT * FROM PROCESSFLOW WHERE SCREENID = '" & ActiveForm & "'", Conn, adOpenStatic, adLockOptimistic
        If RsControls.EOF = False Then
            With RsControls
                If !CONSULTATION = 1 Then OptConsultation.Item(1).Enabled = True
                If !OBSERVATION = 1 Then OptObservation.Item(2).Enabled = True
                If !CASHIER = 1 Then OptCashier.Item(3).Enabled = True
                If !PHARMACY = 1 Then OptPharmacy.Item(4).Enabled = True
                If !LAB = 1 Then OptLab.Item(5).Enabled = True
            End With
        End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Sub ArraySplit(ByVal PrescriptionList As String)
    StrArray = Split(PrescriptionList, "^")
End Sub

Private Sub UpdateTreatment(ByVal InsertDetails As Boolean)
On Error GoTo ErrorHandler
Dim CashAmount As Integer
Dim CreditAmount As Integer
Dim lvSoldUnits As Integer
    If InsertDetails = False And ItemSelected <> 5 Then Exit Sub
    'If InsertDetails = False Then Exit Sub
    
            If CboDiagnosis = "" Then
                Conn.Execute "UPDATE COMPLAINS SET COMPLAINS = '" & TxtComplaints & "',LABREQUEST = '" & TxtLabRequest & "',PRESCRIPTION = '" & TxtPrescription & "',REFERRAL = '" & TxtReferal & "',ADMISSIONnumber = '',DOCTOR = '" & GlbCurrentUser & "' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
            Else
                Conn.Execute "UPDATE COMPLAINS SET COMPLAINS = '" & TxtComplaints & "',LABREQUEST = '" & TxtLabRequest & "',DIAGNOSIS = '" & CDbl(Mid(CboDiagnosis, 1, InStr(CboDiagnosis, "-") - 2)) & "',PRESCRIPTION = '" & TxtPrescription & "',REFERRAL = '" & TxtReferal & "',ADMISSIONNumber = '',DOCTOR = '" & GlbCurrentUser & "' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
            End If
    
    
            If GridPrescription.Rows > 1 Then
                For i = 1 To GridPrescription.Rows - 1
                    If GridPrescription.TextMatrix(i, 7) = 1 Then
                        CashAmount = GridPrescription.TextMatrix(i, 6): CreditAmount = 0
                    ElseIf GridPrescription.TextMatrix(i, 7) = 2 Then
                        CreditAmount = GridPrescription.TextMatrix(i, 6): CashAmount = 0
                    End If
                    
                    'DEDUCT FROM STOCK BEFORE INSERTING INTO PRESCRIPTION.
                    lvSoldUnits = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & GetID_NameFromCombo(GridPrescription.TextMatrix(1, 1), 1) & "'")
                    Deducted = DEDUCT_DRUG_FROM_STOCK(GridPrescription.TextMatrix(1, 1), GridPrescription.TextMatrix(1, 2), lvSoldUnits)
                    If Deducted <> True Then MsgBox "Problem Encountered when deducting Medicine from Stock. Transaction not Succesfull", vbExclamation: Exit Sub
                    Conn.Execute "INSERT INTO PRESCRIPTION  (CARDNUMBER, VISITNUMBER,BILLINGCO,VISITDATE,CODE,DESCRIPTION,QUANTITY,CREDITAMOUNT,CASHAMOUNT,PAYDATE,PAYMENTMODE,PAYMENTSTATUS)" & _
                    "VALUES('" & StrDocCardNo & "', '" & StrDocVisitNumber & "','" & Grid.TextMatrix(1, 3) & "','" & Format(GlbSysDate, "DD MMM YYYY") & "','" & GetID_NameFromCombo(GridPrescription.TextMatrix(i, 1), 1) & "', '" & GetID_NameFromCombo(GridPrescription.TextMatrix(i, 1), 2) & "','" & GridPrescription.TextMatrix(i, 2) & "' ,'" & (CreditAmount * GridPrescription.TextMatrix(i, 2)) & "','" & (CashAmount * GridPrescription.TextMatrix(i, 2)) & "','" & Format(GlbSysDate, "DD MMM YYYY") & "','" & GridPrescription.TextMatrix(i, 7) & "','0')"
                    
                    'INSERT SHEDULE FOR MEDICINE TOP UP REMINDER
                    Conn.Execute "INSERT INTO PRESCRIPTION_SCHEDULE (CARDNUMBER,VISITNUMBER,VISITDATE,PRODUCTCODE,MEDICINECOUNT)" & _
                                 "VALUES('" & StrDocCardNo & "', '" & StrDocVisitNumber & "','" & Format(GlbSysDate, "DD MMM YYYY") & "','" & GetID_NameFromCombo(GridPrescription.TextMatrix(i, 1), 1) & "','" & GridPrescription.TextMatrix(i, 2) & "')"
                Next
            Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = " & lvDirectSalesNo & " + 1 WHERE ITEMNAME = 'DIRECTSALES'"
            End If
            
            'CLEAR SCREEN FOR DOC AFTER POSTING. TODO
            
''''    If LstPrescription.ListCount > 0 Then
''''        For i = 0 To LstPrescription.ListCount - 1
''''            If InStr(LstPrescription.List(i), "^") < 1 Then GoTo NextRecord
''''            ArraySplit LstPrescription.List(i)
''''            DRUGAMOUNT = FindRecord("PRODUCTS", "SALEPRICE", "PRODUCTID = '" & Mid(LstPrescription.List(i), 1, InStr(LstPrescription.List(i), "-") - 1) & "'")
''''            'DrugQuantity = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & Mid(LstPrescription.List(i), 1, InStr(LstPrescription.List(i), "-") - 1) & "'")
''''                If Trim(StrArray(2)) = "001" Then     'CASH PAYMENT
''''                    Conn.Execute "INSERT INTO PRESCRIPTION  (CARDNUMBER, VISITNUMBER,BILLINGCO,VISITDATE,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT,PAYDATE,PAYMENTMODE) VALUES('" & StrDocCardNo & "', '" & StrDocVisitNumber & "','" & Grid.TextMatrix(1, 3) & "','" & Format(Date, "DD MMM YYYY") & "','" & Mid(LstPrescription.List(i), 1, InStr(LstPrescription.List(i), "-") - 1) & "', '" & Trim(Mid(Replace(Replace(StrArray(0), ".", ""), "'", " "), InStr(StrArray(0), "-") + 1, Len(StrArray(0)))) & "','" & StrArray(3) & "' ,'" & DRUGAMOUNT * StrArray(3) & "','" & Format(Date, "DD MMM YYYY") & "','" & Mid(StrArray(2), 1, 3) & "')"
''''                Else ' CREDIT FACILITY
''''                    Conn.Execute "INSERT INTO PRESCRIPTION  (CARDNUMBER, VISITNUMBER,BILLINGCO,VISITDATE,CODE,DESCRIPTION,QUANTITY,CREDITAMOUNT,PAYMENTMODE) VALUES('" & StrDocCardNo & "', '" & StrDocVisitNumber & "','" & Grid.TextMatrix(1, 3) & "','" & Format(Date, "DD MMM YYYY") & "','" & Mid(LstPrescription.List(i), 1, InStr(LstPrescription.List(i), "-") - 1) & "', '" & Trim(Mid(Replace(Replace(StrArray(0), ".", ""), "'", " "), InStr(StrArray(0), "-") + 1, Len(StrArray(0)))) & "','" & StrArray(3) & "' ,'" & DRUGAMOUNT * StrArray(3) & "','" & Mid(StrArray(2), 1, 3) & "')"
''''                End If
''''NextRecord:
''''        Next i
''''    End If
    'MsgBox "Treatment Details Saved Succesfully", vbInformation
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
   ' Resume
End Sub
Private Sub FillHistory()
    On Error GoTo ErrorHandler
   KARI = GlbSysDate
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 7
    Grid.ColAlignment(1) = flexAlignCenterCenter
    'Grid.ColDataType(7) = flexDTBoolean
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |   VISIT DATE "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.CARDNUMBER = '" & StrDocCardNo & "' ORDER BY VISITDATE DESC", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        BILLINGNAME = FindRecord("SERVICE_PROVIDER", "SERVICEPROVIDER", "COMPANYCODE = '" & !BILLINGCOMPANY & "'")
                        Grid.AddItem !CARDNUMBER & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY + " - " + BILLINGNAME & vbTab & !IDNUMBER & vbTab & !VisitDate
                        .MoveNext
                    Wend
                End With
            End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Private Sub FillHistoryFilter()
On Error GoTo ErrorHandler
   Dim RsHistory As New ADODB.Recordset
   Dim RsPrescription As New ADODB.Recordset
    LstPrescription.Clear
    StrDocVisitNumber = Grid.TextMatrix(Grid.Row, 1)
    RsHistory.Open "SELECT * FROM COMPLAINS WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        If RsHistory.BOF = False And RsHistory.EOF = False Then
        
            If Not IsNull(RsHistory!COMPLAINS) = False Then
                    TxtComplaints = ""
                Else
                    TxtComplaints = RsHistory!COMPLAINS
            End If
            If IsNull(RsHistory!DIAGNOSIS) = True Then
                    GridDiagnosis.Clear: GridPrescription.Rows = 1
                Else
                    TxtDiagnosis = RsHistory!DIAGNOSIS
            End If
            If IsNull(RsHistory!PRESCRIPTION) = True Then
                    GridPrescription.Clear
                Else
                    TxtPrescription = RsHistory!PRESCRIPTION
            End If
                        'PRESCRIPTIONS NI MORE THAN ONE
                            RsPrescription.Open "SELECT * FROM PRESCRIPTION WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "' AND CODE <> '001'", Conn, adOpenStatic, adLockOptimistic
                                GridPrescription.Clear: GridPrescription.Rows = 1
                                GridPrescription.FormatString = "    CODE    |    DESCRIPTION    |   QUANTITY"
                                With RsPrescription
                                    While .BOF = False And .EOF = False
                                        LstPrescription.AddItem !CODE & " - " & !Description
                                        GridPrescription.AddItem !CODE & vbTab & !Description & vbTab & !Quantity
                                        .MoveNext
                                    Wend
                                End With
                            RsPrescription.Close
                        'END
            If IsNull(RsHistory!LABREQUEST) = True Then
                    TxtLabRequest = ""
                Else
                    TxtLabRequest = RsHistory!LABREQUEST
            End If
            If IsNull(RsHistory!LABRESULTS) = True Then
                    TxtLabResults = ""
                Else
                    TxtLabResults = RsHistory!LABRESULTS
            End If
            If IsNull(RsHistory!PRESCRIPTION) = True Then
                    TxtReferal = ""
                Else
                    TxtReferal = RsHistory!REFERRAL
            End If
'            If Not IsNull(RsHistory!ADMISSION) Then
'                TxtAdmission = RsHistory!ADMISSION
'            End If
        End If
    RsHistory.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboDiagnosisCategory_Click()
On Error GoTo ErrorHandler
    Dim lvDiagnosisCategoryID As Long
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    CboDiagnosis.Clear
    lvDiagnosisCategoryID = Mid(CboDiagnosisCategory, 1, 3)
    RsFilter.Open "SELECT DIAGNOSISID, DIAGNOSISDESCRIPTION FROM DIAGNOSIS WHERE DIAGNOSISCATEGORY = '" & lvDiagnosisCategoryID & "'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsFilter
            While .BOF = False And .EOF = False
                CboDiagnosis.AddItem !DIAGNOSISID & " - " & !DIAGNOSISDESCRIPTION
                .MoveNext
            Wend
        End With
    RsFilter.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboPrescriptionCategory_Click()
On Error GoTo ErrorHandler
    Dim lvPrescriptionCategoryID
    'POPULATE COMBO FOR PRESCRIPTION
    CboDrugs.Clear
    lvPrescriptionCategoryID = Mid(CboPrescriptionCategory, 1, 3)
    RsFilter.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' order by productname", Conn, adOpenDynamic, adLockOptimistic
    
        With RsFilter
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboDrugs.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboDrugs.AddItem !PRODUCTID & " - " & !ProductName
                End If
                .MoveNext
            Wend
        End With
    RsFilter.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboWard_Click()
On Error GoTo ErrorHandler
    Dim lvWardNumber
    'POPULATE COMBO FOR PRESCRIPTION
    CboBed.Clear
    lvWardNumber = Mid(CboWard, 1, 3)
    RsFilter.Open "SELECT BEDNUMBER FROM BEDS_AVAILABILITY WHERE WARDNUMBER = '" & Val(lvWardNumber) & "' AND OCCUPIEDBY = '0'", Conn, adOpenDynamic, adLockOptimistic
        With RsFilter
            While .BOF = False And .EOF = False
                    CboBed.AddItem "BED No" & " - " & !BEDNUMBER
                .MoveNext
            Wend
        End With
    RsFilter.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub ChkOverride_Click()
    If ChkOverride.Value = 1 Then
        CmdPost.Enabled = True
    End If
End Sub

Private Sub CmdAddComplains_Click()
    FrmComplains.Show 1
End Sub

Private Sub CmdAddDrug_Click()
On Error GoTo ErrorHandler

Dim lvUnit As Integer
Dim lvPrice As Integer
Dim lvStock As Boolean
Dim lvDosage As String
Dim lvDosageRemarks As String
Dim lvDays As Integer
    'NILIKUWA NATRY LOGIC FULANI HAPA BUT IMEKATAA. FUCK IT
    Dim KARI As String
    Dim MADOTS As String
    'If CboDosages = "" Then MsgBox "Please select the Dosage Frequency before adding the drug to prescription List", vbInformation: Exit Sub
    If CboPaymentMode = "" Then MsgBox "Please select the Mode of Paymebt before adding the drug to prescription List", vbInformation: Exit Sub
    SOLO = Len(CboDrugs)
    MADOTS = String(SOLO - Len(LELIT), ".") & LELIT
    MADOTS = SOLO
    If CboDrugs.Text = "" Then Exit Sub
    
    'TAKE QUANTITY OF MEDICINE FROM DOCTOR
    Dim Pos, StrProductID As String
    Pos = InStr(CboDrugs, "-")
    StrProductID = Left(CboDrugs, Pos - 2)
    FrmQuantity.TxtUnit = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & StrProductID & "'")
    lvUnit = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & StrProductID & "'")
    FrmQuantity.TxtPrice = FindRecord("PRODUCTS", "SALEPRICE", "PRODUCTID = '" & StrProductID & "'")
    lvPrice = FindRecord("PRODUCTS", "SALEPRICE", "PRODUCTID = '" & StrProductID & "'")
    lvDosage = FindRecord("PRODUCTS", "DOSAGE", "PRODUCTID = '" & StrProductID & "'")
    lvDosageRemarks = FindRecord("PRODUCTS", "DOSAGEREMARKS", "PRODUCTID = '" & StrProductID & "'")
    lvDays = FindRecord("PRODUCTS", "DURATION", "PRODUCTID = '" & StrProductID & "'")
    
   ' FrmQuantity.Show vbModal
   'QUANTITY = DOSAGE * NUMBER OF DAYS. IF IT IS CAPSULE OR TABLET THEN DO NOT MULTIPLY.
   lvDrugType = FindRecord("PRODUCTS", "MEDICINETYPE", "PRODUCTID = '" & StrProductID & "'")
   If InStr(1, lvDrugType, "TABLET") = 1 Then 'CALCULATE IF ITS TABLET
        GlbUnitQuantity = (Mid(lvDosage, 1, 1) * Right(lvDosage, 1)) * lvDays
   ElseIf InStr(1, lvDrugType, "CAPSULE") = 1 Then  'CALCULATE IF ITS CAPSULE
        GlbUnitQuantity = (Mid(lvDosage, 1, 1) * Right(lvDosage, 1)) * lvDays
   Else  ' DO NOT CALCULATE IF ITS SYRUP OR CREAM OR POWDER OR OTHER
         GlbUnitQuantity = 1
   End If
   
    If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'ExcludePharmacy'") = 0 Then
        'CHECK IF MEDICINE IS IN STOCK
        If FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & StrProductID & "'") = "" Then
            MsgBox GetID_NameFromCombo(CboDrugs.Text, 2) & "  Is NOT Available In Stock And Therefore Cannot Be Issued.", vbExclamation, "Inventory"
            Exit Sub
        End If
        
'''        'CHECK IF QUANTITY REQUESTED IS ABOVE STOCK LEVEL AND PROMPT. TO DO.
'''        If (FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & StrProductID & "'")) > GlbUnitQuantity Then
'''            lvZilizosalia = FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & StrProductID & "'")
'''            MsgBox "Quantity Requested for " & GetID_NameFromCombo(CboDrugs, 2) & " is above Current Stock Levels. Only " & lvZilizosalia & " Units are left in Stock", vbExclamation
'''            Exit Sub
'''        End If
    End If

    'INSERT MEDICINE INTO PRE SALES TABLE AND DISPLAY ON GRID. DO NOT DEDUCT AT THIS POINT
    Conn.Execute "INSERT INTO PRE_DRUGS_SALES(SALENUMBER,CATEGORYID,PRODUCTID,PRODUCTDESCRIPTION,DISTRIBUTIONUNIT,QUANTITY,AMOUNT,PAYDATE,SOLDBY,DOCTOR0PHARMACY1,PAYMENTMODE,DOSAGE,DOSAGEREMARKS,DAYS)" & _
                 "VALUES ('" & lvDirectSalesNo & "','" & GetID_NameFromCombo(CboPrescriptionCategory, 1) & "','" & GetID_NameFromCombo(Replace(CboDrugs, "'", " "), 1) & "','" & GetID_NameFromCombo(Replace(CboDrugs, "'", " "), 2) & "','" & lvUnit & "','" & GlbUnitQuantity & "','" & lvPrice & "','" & GlbSysDate & "','" & GlbCurrentUser & "','0', '" & Val(GetID_NameFromCombo(CboPaymentMode, 1)) & "','" & lvDosage & "','" & lvDosageRemarks & "','" & lvDays & "')"
    
    'POPULATE INSERTED DATA TO GRID
    PopulatePrescriptionGRID lvDirectSalesNo
    
'''    'DEDUCT MEDICINE FROM STOCK ONCE PRESCRIBED TO PATIENT
'''    If DEDUCT_DRUG_FROM_STOCK(CboDrugs, GlbUnitQuantity) <> True Then Exit Sub
'''    KISH = String(73 - MADOTS, ".") & " ^" + CboDosages
'''
'''    LstPrescription.AddItem CboDrugs + KISH + " ^" & Mid(CboPaymentMode, 1, 3) + " ^" & GlbUnitQuantity
    
    CmdPost.Enabled = True
    Debug.Print KISH
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
 '   Resume
End Sub
Public Sub PopulatePrescriptionGRID(ByVal SaleNumber)
On Error GoTo ErrorHandler
    GridPrescription.Clear: GridPrescription.Rows = 1: GridPrescription.Cols = 2
    GridPrescription.FormatString = "CATEGORY ID|  PRESCRIPTION NAME    | QUANTITY |DAYS| DOSAGE | DOSAGE REMARKS | AMOUNT | PAY MODE"
    'G.ColDataType(2) = flexDTBoolean
    If rsfill.State = 1 Then Set rsfill = Nothing
        rsfill.Open "SELECT * FROM  PRE_DRUGS_SALES WHERE SALENUMBER = '" & SaleNumber & "' AND DOCTOR0PHARMACY1 = '0' AND SOLDBY = '" & GlbCurrentUser & "'", Conn, adOpenStatic, adLockOptimistic
            If rsfill.BOF = False And rsfill.EOF = False Then
                While rsfill.EOF = False
                    With rsfill
                        GridPrescription.AddItem !CATEGORYID & vbTab & !PRODUCTID & " - " & !PRODUCTDESCRIPTION & vbTab & !Quantity & vbTab & !DAYS & vbTab & !DOSAGE & vbTab & !DosageRemarks & vbTab & !amount & vbTab & !PAYMENTMODE
                    End With
                rsfill.MoveNext
                Wend
            End If
        rsfill.Close
    GridPrescription.Editable = False
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
    ''Resume
End Sub

Private Sub CmdAdmission_Click()
On Error GoTo ErrorHandler
'''''    UpdateTreatment (True)
'''''    'UPDATE TREATMENT DIRECTION
'''''    Conn.Execute "UPDATE COMPLAINS SET TOADMISSION = '1' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "'"
'''''    'UPDATE OCCUPIED BED
'''''    Conn.Execute "UPDATE BEDS SET"
Dim DepositNotPaid As Integer
Dim DepositAmount As Integer
    DepositNotPaid = 0
    DepositAmount = -20000
    If StrDocCardNo = "" Then MsgBox "No Card Number", vbInformation: Exit Sub
        'INSERT INTO INPATIENTS
        Conn.Execute "INSERT INTO INPATIENTS (CARDNUMBER,VISITNUMBER,WARDNUMBER,BEDNUMBER,DOCTOR,ADMISSIONDATE,PAYMENTSTATUS)" & _
                     "VALUES('" & StrDocCardNo & "','" & Trim(StrDocVisitNumber) & "','" & Val(Mid(CboWard, 1, 3)) & "','" & Right(CboBed, 1) & "','" & TxtDoctor & "','" & Format(GlbSysDate, "DD MMM YYYY") & "','" & DepositNotPaid & "')"
                     
        'INSERT DOCTORS NOTES
        Conn.Execute "INSERT INTO NURSE_DOC_NOTES (CARDNUMBER,VISITNUMBER,DOCTORSNOTES)" & _
                     "VALUES('" & StrDocCardNo & "','" & StrDocVisitNumber & "','" & TxtAdmissionNotes & "')"
                     
        'SET BED NUMBER AS OCCUPIED
        Conn.Execute "UPDATE BEDS_AVAILABILITY SET OCCUPIEDby = 1 WHERE WARDNUMBER = '" & Val(Mid(CboWard, 1, 3)) & "' AND BEDNUMBER = '" & Right(CboBed, 1) & "'"
        
        'UPDATE PATIENT TO CASHIER
        DMY = SendPatient(EnumWard, StrDocCardNo, GlbSysDate)
        
        'CREATE DEPOSIT PAYMENT RECORD TO CASHIER
        Conn.Execute "INSERT INTO PRESCRIPTION(CARDNUMBER,VISITNUMBER,BILLINGCO,DESCRIPTION,CASHAMOUNT)" & _
        "VALUES('" & StrDocCardNo & "','" & StrDocVisitNumber & "','001','WARD ADMISSION DEPOSIT','" & DepositAmount & "')"
        
    MsgBox "Patient Details Posted to Admission Succesfully", vbInformation
        CmdAdmission.Enabled = False
        LblAdmit.Visible = True
        ClearText FrmTreatment
        LblAdmit.Caption = "Admission for " & TxtFirstname & " Succesfull"
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub CmdClearHistory_Click()
On Error GoTo ErrorHandler
    Dim CTL As Control
    Dim KARI As Form
    Set KARI = Me
        For Each CTL In KARI.Controls
            If TypeOf CTL Is TextBox Then
                CTL.Locked = False
                'CTL.Text = "" 'DONT CLEAR TEXT BOXES.
            ElseIf TypeOf CTL Is ComboBox Then
                CTL.Text = ""
            End If
        Next
        
    
    'SELECTIVE CLEARING OF TEXT BOXES
    TxtComplaints = ""
    TxtLabRequest = ""
    TxtLabResults = ""
    GridPrescription.Clear: GridPrescription.Rows = 1
    GridDiagnosis.Clear: GridDiagnosis.Rows = 1
    CmdPost.Enabled = True
   'ReverseGreyOut FrmTreatment
   
   'PUT DEFAULT VALUES
    CboPaymentMode.Text = "001 - CASH"
    
    'GET THE LAST VISIT NUMBER AFTER VIEWING HISTORY INFORMATION
    Dim lvLastrow As Long
                    For i = 1 To Grid.Rows - 1
                        lvLastrow = i
                    Next
    StrDocVisitNumber = Grid.TextMatrix(lvLastrow, 1)
 
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub CmdDelete_Click()
    Dim Resp
    Resp = MsgBox("Are you Sure you wish to Delete this record?", vbQuestion + vbYesNo)
            If Resp = vbNo Then MsgBox "Deletion aborted", vbInformation: Exit Sub
        Conn.Execute "DELETE FROM PRESCRIPTION WHERE CARDNUMBER = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 0) & "' AND VISITNUMBER = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 1) & "' AND PRESCRIPTIONID = '" & GridDrillDown.TextMatrix(GridDrillDown.Row, 2) & "'"
            PopulateDrillDown
            MsgBox "Record Deleted Succesfully", vbInformation
            ResetMode
        'Conn.Execute "DELETE FROM PRESCRIPTION WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'"
            ResetMode
End Sub

Private Sub CmdExit_Click()
On Error GoTo ErrorHandler
    If TxtFirstname <> "" And TxtBMI <> "" And TxtHeight <> "" Then
        'Release document to be accessed by other system users
        Conn.Execute "UPDATE COMPLAINS SET INUSE = 'False' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "'"
    End If
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub CmdConclude_Click()
Dim Resp
    Resp = MsgBox("Please confirm that you wish to conclude treatment for '" & TxtFirstname & "'", vbExclamation + vbYesNo)
        If Resp = vbYes Then
            
        End If
End Sub

Private Sub CmdFoodFrequency_Click()
    If StrDocCardNo <> "" Then lvFoodCardNo = StrDocCardNo: lvFoodVisitNo = StrDocVisitNumber
    FrmFoodFrequency.Show 1
End Sub

Private Sub CmdLaboratory_Click()
On Error GoTo ErrorHandler
    UpdateTreatment (True)
    'UPDATE TREATMENT DIRECTION
    Conn.Execute "UPDATE COMPLAINS SET TOLABOROTARY = '1' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "'"
    
    MsgBox "Patient Details Posted to Laboratory Succesfully", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub CmdPayOffered_Click()
    AllocateMonies 2
End Sub

Private Sub CmdPayReceived_Click()
    AllocateMonies 1
End Sub

Private Sub CmdMeasurements_Click()
    BlnViewMeasurements = True
    FrmObservation.Show 1
End Sub

Private Sub CmdPharmacy_Click()
    UpdateTreatment (True)
    'UPDATE TREATMENT DIRECTION
    Conn.Execute "UPDATE COMPLAINS SET TOPHARMACY = '1' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "'"
    
    MsgBox "Patient Details Posted to Pharmacy Succesfully", vbInformation
End Sub

Private Sub CmdPost_Click()
On Error GoTo ErrorHandler
    Dim PostDetails As Boolean
    KARI = GlbSysDate
    PostDetails = True
    If ItemSelected = 0 Then MsgBox "Please select Location to send This entry before Posting", vbInformation: Exit Sub
    If TxtComplaints = "" Then
        Resp = MsgBox("Are You Sure you want to Post this Record without entering the Complains ?", vbYesNo)
        If Resp = vbNo Then Exit Sub
        'PostDetails = False
    End If
'    If LstPrescription.ListCount < 1 Then
'        Resp = MsgBox("Are You Sure you want to Post this Record without Adding Prescription ?", vbYesNo)
'        If Resp = vbNo Then Exit Sub
'        PostDetails = False
'    End If
    
    UpdateTreatment (PostDetails)
    
    'Release document to be accessed by other system users
    Conn.Execute "UPDATE COMPLAINS SET INUSE = 'False' WHERE CARDNUMBER = ' " & StrDocCardNo & " ' AND VISITDATE = ' " & Format(KARI, "DD MMM YYYY") & " '"
    Select Case ItemSelected
   Case 1
        'To Consultation
        DMY = SendPatient(EnumConsultation, StrDocCardNo, Format(KARI, "DD MMM YYYY"))
         'FrmPatients.Show
         'Unload Me
    Case 2
        'To Observation
        DMY = SendPatient(EnumObservation, StrDocCardNo, Format(KARI, "DD MMM YYYY"))
        'FrmObservation.Show
        'Unload Me
    Case 3
        'To Cashier
        DMY = SendPatient(EnumCashier, StrDocCardNo, Format(KARI, "DD MMM YYYY"))
        'FrmCashier.Show
        'Unload Me
    Case 4
        'To Pharmacy
        DMY = SendPatient(EnumPharmacy, StrDocCardNo, Format(KARI, "DD MMM YYYY"))
        'FrmPharmacy.Show
        'Unload Me
    Case 5
        DMY = SendPatient(EnumLab, StrDocCardNo, Format(KARI, "DD MMM YYYY"))
    End Select
    
    'CLEAR GRID AND ALL CONTROLS.
        ClearText FrmTreatment
        Grid.Clear: Grid.Rows = 1
        GridPrescription.Clear
    CmdClearHistory_Click
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
   ' Resume
End Sub

Private Sub CmdPrescription_Click()
On Error GoTo ErrorHandler
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=Today123;DSQ=SYB-KEN-NB-002\SQL2005;"
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\PRESCRIPTION.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT PRESCRIPTION REPORT"
                   .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & StrDocCardNo & "' AND {PRESCRIPTION.VISITNUMBER} = '" & StrDocVisitNumber & "' AND {PRESCRIPTION.CODE} <> '001'"
                   '.SelectionFormula = "{PRESCRIPTION.VISITDATE} = '" & Format(DtCurrDate, "YYYY,MM,DD") & "'"
                   '.SelectionFormula = " {COMPLAINS.VISITNUMBER} = '" & StrDocVisitNumber & "'"
                   .Destination = 0
                   .Action = 1
                End With
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CmdPrintReferral_Click()
On Error GoTo ErrorHandler
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   .Connect = Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\REFERRAL.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT REFERRAL REPORT"
                   '.SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & Grid.TextMatrix(Grid.Row, 0) & "'"
                   '.SelectionFormula = " {PRESCRIPTION.VISITNUMBER} = '" & Grid.TextMatrix(Grid.Row, 1) & "'"
                   .Destination = 0
                   .Action = 1
                End With
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CmdReferral_Click()
    UpdateTreatment (True)
    'UPDATE TREATMENT DIRECTION
    Conn.Execute "UPDATE COMPLAINS SET TOREFERRAL = '1',INUSE = 'False' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "'"
    
    MsgBox "Patient Details Posted to Referral Succesfully", vbInformation
End Sub

Private Sub CmdRemoveDiagnosis_Click()
On Error GoTo ErrorHandler
    'DELETE RECORD FROM DOC_DIAGNOSIS TABLE
    Conn.Execute "DELETE FROM DOC_DIAGNOSIS WHERE DIAGNOSISID = '" & GetID_NameFromCombo(GridDiagnosis.TextMatrix(GridDiagnosis.Row, 2), 1) & "' AND CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
    'LOAD DIAGNOSIS TO GRID
    PopuateDiagnosis StrDocCardNo, StrDocVisitNumber
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub CmdRemoveDrug_Click()
On Error GoTo ErrorHandler
    'Dim pos As String
    'pos = InStr(LstPrescription.Text, "-")
    'RETURN DRUG TO STOCK.
    
    'RETURN_DRUG_TO_STOCK LstPrescription.Text, ddd
    'LstPrescription.RemoveItem (LstPrescription.ListIndex)
    Conn.Execute "DELETE FROM PRE_DRUGS_SALES WHERE PRODUCTID = '" & GetID_NameFromCombo(GridPrescription.TextMatrix(GridPrescription.Row, 1), 1) & "'"
    PopulatePrescriptionGRID lvDirectSalesNo
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub


Private Sub CMDSchedule_Click()
    FrmScheduleVisit.Show 1
End Sub

Private Sub CmdAddDiagnosis_Click()
On Error GoTo ErrorHandler
    DocTab.Tab = 1
    'INSERT RECORD INTO DOC_DIAGNOSIS TABLE
    Conn.Execute "INSERT INTO DOC_DIAGNOSIS (DIAGNOSISID,CARDNUMBER,VISITNUMBER,VISITDATE)VALUES('" & GetID_NameFromCombo(CboDiagnosis, 1) & "','" & StrDocCardNo & "','" & StrDocVisitNumber & "','" & GlbSysDate & "')"
    'LOAD DIAGNOSIS TO GRID
    PopuateDiagnosis StrDocCardNo, StrDocVisitNumber
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Public Sub PopuateDiagnosis(ByVal CardNo, VisitNo)
On Error GoTo ErrorHandler
    Dim lvDescription As String
    GridDiagnosis.Clear: GridDiagnosis.Rows = 1: GridDiagnosis.Cols = 2
    GridDiagnosis.FormatString = "CARD NUMBER |  VISIT NUMBER    | DIAGNOSIS ID ID AND DESCRIPTION"
    'G.ColDataType(2) = flexDTBoolean
    If rsfill.State = 1 Then Set rsfill = Nothing
        rsfill.Open "SELECT * FROM  DOC_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
            If rsfill.BOF = False And rsfill.EOF = False Then
                While rsfill.EOF = False
                    With rsfill
                        lvDescription = FindRecord("DIAGNOSIS", "DIAGNOSISDESCRIPTION", "DIAGNOSISID = '" & !DIAGNOSISID & "'")
                        GridDiagnosis.AddItem !CARDNUMBER & vbTab & !VISITNUMBER & vbTab & !DIAGNOSISID + " - " + lvDescription
                    End With
                rsfill.MoveNext
                Wend
            End If
        rsfill.Close
    GridDiagnosis.Editable = False
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
    ''Resume
End Sub

Private Sub CmdViewScan_Click()
    On Error GoTo ErrorHandler
    FrmFullImage.ImgFull = ImgPreview.Picture
    FrmFullImage.Show
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command4_Click()
    Load FrmDosages
    FrmDosages.Show
End Sub


Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim RsCombo As New ADODB.Recordset
Dim RsPrescription  As New ADODB.Recordset
    centerform Me
    KARI = GlbSysDate
    
    'VARIABLE FOR PRESCRIPTION PER PATIENT
    lvDirectSalesNo = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'DirectSales'")
    
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    RsCombo.Open "SELECT DIAGNOSISCATEGORYID, DIAGNOSISCATEGORYDESC FROM DIAGNOSISCATEGORY ORDER BY DIAGNOSISCATEGORYDESC", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboDiagnosisCategory.AddItem String(3 - Len(!DIAGNOSISCATEGORYID), "0") & !DIAGNOSISCATEGORYID & " - " & !DIAGNOSISCATEGORYDESC
                .MoveNext
            Wend
        End With
    RsCombo.Close
    
    'POPULATE COMBO FOR PRESCRIPTION CATEGORY
    RsCombo.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY ORDER BY PRODUCTGROUP", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboPrescriptionCategory.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                .MoveNext
            Wend
        End With
    RsCombo.Close
    
    'POPULATE COMBO FOR DOSAGES
    RsCombo.Open "SELECT DOSAGEID, DOSAGE FROM DOSAGES", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboDosages.AddItem String(3 - Len(!DOSAGEID), "0") & !DOSAGEID & " - " & !DOSAGE
                .MoveNext
            Wend
        End With
    RsCombo.Close
    
    'POPULATE COMBO PAYMENT MODES
    If RsCombo.State = 1 Then Set RsCombo = Nothing
    RsCombo.Open "SELECT * FROM PAYMENT_MODES", Conn, adOpenStatic, adLockOptimistic
        While RsCombo.EOF = False
            CboPaymentMode.AddItem String(3 - Len(RsCombo!PAYMENTCODE), "0") & RsCombo!PAYMENTCODE & " - " & RsCombo!PAYMENTDESCRIPTION
            'CMBPaymethod.AddItem String(3 - Len(RsCombo!PAYMENTCODE), "0") & RsCombo!PAYMENTCODE & " - " & RsCombo!PAYMENTDESCRIPTION
            RsCombo.MoveNext
        Wend
    RsCombo.Close
    
''''    'POPULATE COMBO FOR SERVICES
''''    If RsCombo.State = 1 Then Set RsCombo = Nothing
''''    RsCombo.Open "SELECT * FROM PAY_ALLOCATION", Conn, adOpenStatic, adLockOptimistic
''''        While RsCombo.EOF = False
''''            CMBPayAllocation.AddItem RsCombo!PAYCODE & " - " & RsCombo!PAYDESCRIPTION
''''            RsCombo.MoveNext
''''        Wend
''''    RsCombo.Close
    
''''    'POPULATE COMBO FOR WARDS
''''    RsCombo.Open "SELECT DISTINCT WARDS.WARDNUMBER, WARDS.WardDescription AS WARDS FROM WARDS INNER JOIN BEDS_AVAILABILITY ON WARDS.WardNumber = BEDS_AVAILABILITY.WardNumber WHERE (BEDS_AVAILABILITY.OccupiedBy = 0)", Conn, adOpenDynamic, adLockOptimistic
''''        With RsCombo
''''            While .BOF = False And .EOF = False
''''                CboWard.AddItem String(3 - Len(!WARDNUMBER), "0") & !WARDNUMBER & " - " & !WARDS
''''                .MoveNext
''''            Wend
''''        End With
''''    RsCombo.Close
    
    'FILL DETAILS FROM THE TWO VARIABLES.
    Select Case BlnHISTORY
        Case False
        
            If rsfill.State = 1 Then Set rsfill = Nothing
            rsfill.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CardNumber = COMPLAINS.CardNumber AND COMPLAINS.CardNumber = '" & StrDocCardNo & "' AND COMPLAINS.VisitDate = '" & Format(StrDocVisitDate, "DDMMMYYYY") & "'", Conn, adOpenStatic, adLockOptimistic
                If rsfill.BOF = False And rsfill.EOF = False Then
                    With rsfill
                    TxtFirstname = !SURNAME & "  " & !FirstName
                    TxtSecondName = !SECONDNAME
                    TxtBp = !BP
                    TxtBMI = !BMINDEX
                    TxtWeight = !Weight
                    TxtHeight = !Height
                    DtCurrDate = !VisitDate
                    End With
                Else
                    'THIS CASE IS HIGHLY UNLICKELY, I WILL NOT EVEN BOTHER WRITTING CODE FOR IT.
                End If
        Case Else
            If rsfill.State = 1 Then Set rsfill = Nothing
            rsfill.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CardNumber = COMPLAINS.CardNumber AND COMPLAINS.CardNumber = '" & StrDocCardNo & "' AND COMPLAINS.VisitDate = '" & Format(StrDocVisitDate, "DDMMMYYYY") & "'", Conn, adOpenStatic, adLockOptimistic
                If rsfill.BOF = False And rsfill.EOF = False Then
                    With rsfill
                        TxtFirstname = !SURNAME & "  " & !FirstName
                        TxtSecondName = !SECONDNAME
                        TxtBp = !BP
                        TxtWeight = !Weight
                        TxtHeight = !Height
                        DtCurrDate = !VisitDate
                        TxtComplaints = !COMPLAINS & vbNullString
                        TxtDiagnosis = !DIAGNOSIS & vbNullString
                        'PRESCRIPTIONS NI MORE THAN ONE
                            RsPrescription.Open "SELECT * FROM PRESCRIPTION WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
                                With RsPrescription
                                    While .BOF = False And .EOF = False
                                        LstPrescription.AddItem !CODE & " - " & !Description
                                        .MoveNext
                                    Wend
                                End With
                            RsPrescription.Close
                        'END
                        
                        TxtReferal = !REFERRAL & vbNullString
                        'TxtAdmission = !ADMISSION & vbNullString
                        BlnHISTORY = False & vbNullString
                        
                        'DisableControls
                    End With
                Else
                    'THIS CASE IS HIGHLY UNLICKELY, I WILL NOT EVEN BOTHER WRITTING CODE FOR IT.
                End If
        End Select
    If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'ExcludELaboratory'") = 1 Then
        CmdViewScan.Visible = True
        ImgPreview.Visible = True
        TxtLabResults.Visible = False
    End If
    FillHistory
    centerform Me
    ManageProcessFlow EnumDoctors
    DocTab.Tab = 2
    CboPaymentMode.Text = "001 - CASH"
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Public Function LoadPictureFromDB(ByRef rs As ADODB.Recordset, ByVal fldName As String, ByRef Image1 As Object, Optional ByVal strFileName As String)

    On Error GoTo ErrorHandler
    
    'If Recordset is Empty, Then Exit
    If rs Is Nothing Then
        'GoTo procNoPicture
    End If
    
    Dim strTempFileName As String
    'strTempFileName = GetParam("IMAGETEMPFILE")
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    
    strStream.Write rs.Fields(fldName).Value

    If strFileName = "" Then
        strFileName = "c:\Temp.bmp"
    End If
    strStream.SaveToFile strFileName, adSaveCreateOverWrite
    ImgPreview.Picture = LoadPicture(strFileName)
    On Error Resume Next
    'Image1.DisplayBlankImage Image1.Width, Image1.Height, 800, 600, 2
    'Image1.Display
    If Err.Number <> 0 Then
        'Image1.DisplayBlankImage Image1.Width, Image1.Height, 800, 600, 2
        'Image1.Refresh
        ImgPreview.Picture = LoadPicture(App.Path & "\NoImage.bmp")
        LoadPictureFromDB = False
    Else
        LoadPictureFromDB = True
    End If
    On Error GoTo 0
    'Kill ("C:\Temp.tif")
    'LoadPictureFromDB = True

Exit Function
ErrorHandler:
    MsgBox Err.Description
    'SystemErrorHandler Err.Number, Err.Description
End Function

Private Sub OptAdmission_Click(Index As Integer)
    ItemSelected = OptAdmission
End Sub

Private Sub OptDoctors_Click(Index As Integer)
    ItemSelected = OptDoctors.Item(Index).Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHandler
''''    While LstPrescription.Text <> ""
''''        CmdRemoveDrug_Click
''''    Wend
    Conn.Execute "UPDATE COMPLAINS SET INUSE = 0 WHERE CARDNUMBER = '" & StrDocCardNo & "'"
    Conn.Execute "DELETE FROM PRE_DRUGS_SALES WHERE SALENUMBER = '" & lvDirectSalesNo & "' AND DOCTOR0PHARMACY1='0'"
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub Grid_DblClick()
On Error GoTo ErrorHandler
    'ClearText FrmTreatment
    FillHistoryFilter
    DisableControls
    If Not IsNumeric(StrDocVisitNumber) Then StrDocVisitNumber = Grid.TextMatrix(Grid.Row, 1)
    
    'POPULATE DIAGNOSIS
    PopuateDiagnosis StrDocCardNo, StrDocVisitNumber
    
    'POPULATE LAB SCAN RESULTS
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM LAB_SCAN WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        With RsImageStore
            If RsImageStore.EOF = False Then
                LoadPictureFromDB RsImageStore, "SCANIMAGE", ImgPreview, D
            Else
                ImgPreview = LoadPicture(App.Path & "\NoImage.bmp")
            End If
        End With
    RsImageStore.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub GridPrescription_DblClick()
    If GridPrescription.Row = 0 Then Exit Sub
    GlbMedEditID = GetID_NameFromCombo(GridPrescription.TextMatrix(GridPrescription.Row, 1), 1)
    FrmMedEdit.TxtMedName = GetID_NameFromCombo(GridPrescription.TextMatrix(GridPrescription.Row, 1), 2)
    FrmMedEdit.TxtType = FindRecord("PRODUCTS", "MEDICINETYPE", "PRODUCTID = '" & GetID_NameFromCombo(GridPrescription.TextMatrix(GridPrescription.Row, 1), 1) & "'")
    FrmMedEdit.TxtDays = FindRecord("PRODUCTS", "DURATION", "PRODUCTID = '" & GetID_NameFromCombo(GridPrescription.TextMatrix(GridPrescription.Row, 1), 1) & "'")
    FrmMedEdit.CboDosage = GridPrescription.TextMatrix(GridPrescription.Row, 4)
    FrmMedEdit.TxtRemarks = GridPrescription.TextMatrix(GridPrescription.Row, 5)
    FrmMedEdit.Show 1
    PopulatePrescriptionGRID lvDirectSalesNo
End Sub

Private Sub OptCashier_Click(Index As Integer)
    ItemSelected = OptCashier.Item(Index).Index
End Sub

Private Sub Option1_Click()
    ItemSelected = 5
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


Private Sub TxtLabRequest_DblClick()
    FrmLabParameters.Show 1
End Sub
Private Sub AllocateMonies(ByVal Cash1Credit2Cheque3 As Integer)
On Error GoTo ErrorHandler
    'VALIDATE FIELDS ARE NOT BLANK OR INVALID
    If TxtAmount = "" Then MsgBox "PLEASE ENTER PAYMENT AMOUNT BEFORE POSTING PAYMENT", vbExclamation: Exit Sub

    'INSERT INTO CHARGES TABLE INITIALLY CALLED PRESCRIPTION. CAPTION REMAINS AS SUCH. SG 22012011
    Select Case Cash1Credit2Cheque3
        Case 1
            Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT,PAYDATE,PAYMENTMODE,VISITDATE,BILLINGCO,CASHIER)" & _
                 "VALUES('" & StrDocCardNo & "','" & StrDocVisitNumber & "','" & Mid(CMBPayAllocation, 1, 3) & "','" & Mid(CMBPayAllocation, 6, Len(CMBPayAllocation)) & "','1', '" & TxtAmount & "','" & Format(DTPayment, "DD MMM YYYY") & "','" & Cash1Credit2Cheque3 & "'," & _
                 "'" & Format(GlbSysDate, "DD MMM YYYY") & "','" & Left(CMBPaymethod, 3) & "','" & StrCurrentUser & "')"
        Case 2
            Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,CODE,DESCRIPTION,QUANTITY,CREDITAMOUNT,PAYDATE,PAYMENTMODE,VISITDATE,BILLINGCO,CASHIER)" & _
                 "VALUES('" & StrDocCardNo & "','" & StrDocVisitNumber & "','" & Mid(CMBPayAllocation, 1, 3) & "','" & Mid(CMBPayAllocation, 6, Len(CMBPayAllocation)) & "','1', '" & TxtAmount & "','" & Format(DTPayment, "DD MMM YYYY") & "','" & Cash1Credit2Cheque3 & "'," & _
                 "'" & Format(GlbSysDate, "DD MMM YYYY") & "','" & Left(CMBPaymethod, 3) & "','" & StrCurrentUser & "')"
        Case 3
            Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT,PAYDATE,PAYMENTMODE,VISITDATE,BILLINGCO,CASHIER)" & _
                 "VALUES('" & StrDocCardNo & "','" & StrDocVisitNumber & "','" & Mid(CMBPayAllocation, 1, 3) & "','" & Mid(CMBPayAllocation, 6, Len(CMBPayAllocation)) & "','1', '" & TxtAmount & "','" & Format(DTPayment, "DD MMM YYYY") & "','" & Cash1Credit2Cheque3 & "'," & _
                 "'" & Format(GlbSysDate, "DD MMM YYYY") & "','" & Left(CMBPaymethod, 3) & "','" & StrCurrentUser & "')"
    End Select
    'CHECK IF THIS WAS A CHEQUE PAYMENT AND UPDATE THE CHEQUE AMOUNT ACCORDINGLY.
    If Left(CMBPaymethod, 2) = "03" Then
       Conn.Execute "UPDATE CHEQUE_PAYMENTS SET AMOUNT = " & TxtAmount & "  WHERE  CARDNUMBER='" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'"
    End If
    PopulateDrillDown
    'MsgBox "Service Transaction Generated Succesfully", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub PopulateDrillDown()
On Error GoTo ErrorHandler
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
    GridDrillDown.FormatString = "CARD No|VISIT No|CHARGES ID|DESCRIPTION|MODE OF PAYMENT|CASH AMOUNT |CREDIT AMOUNT"
        If RsGrid.State = adStateOpen Then RsGrid.Close
'        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '1' AND PAID = '0'", Conn, adOpenDynamic, adLockOptimistic
         'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.TOCASHIER =  '1'", Conn, adOpenDynamic, adLockOptimistic
         RsGrid.Open "SELECT PRESCRIPTION.CARDNUMBER,PRESCRIPTION.VISITNUMBER,PRESCRIPTION.PRESCRIPTIONID,PRESCRIPTION.DESCRIPTION,PRESCRIPTION.PAYMENTMODE,PRESCRIPTION.QUANTITY,PRESCRIPTION.CASHAMOUNT,PRESCRIPTION.CREDITAMOUNT FROM PRESCRIPTION WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "' AND PRESCRIPTION.PAYMENTSTATUS = '0'", Conn, adOpenStatic, adLockOptimistic
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
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub ResetMode()
    CMBPayAllocation.Text = ""
    CMBPaymethod.Text = ""
    
End Sub
