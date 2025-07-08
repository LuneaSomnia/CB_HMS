VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmFoodFrequency 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Food Frequency Questionaire"
   ClientHeight    =   9810
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   15975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Food Frequency Questionaire"
      TabPicture(0)   =   "FrmFoodFrequency.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Cooking Oil / Alcohol / Tobacco Drugs"
      TabPicture(1)   =   "FrmFoodFrequency.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame10 
         Caption         =   "Exercise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   58
         Top             =   6120
         Width           =   10695
         Begin VB.TextBox TxtExerciseAweek 
            Height          =   285
            Left            =   7320
            TabIndex        =   78
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton CmdRemoveExercise 
            Caption         =   "Remove"
            Height          =   495
            Left            =   8520
            TabIndex        =   65
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton CmdAddExericise 
            Caption         =   "Add "
            Height          =   495
            Left            =   8520
            TabIndex        =   64
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox CboExercise 
            Height          =   315
            Left            =   1920
            TabIndex        =   62
            Top             =   360
            Width           =   2775
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridExercise 
            Height          =   1695
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2990
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
         Begin VB.Label Label18 
            Caption         =   "How many times a week"
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
            Left            =   5040
            TabIndex        =   79
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Type of Exercise"
            Height          =   255
            Left            =   360
            TabIndex        =   61
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Alcohol / Tobacco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   15495
         Begin VB.TextBox TxtAlcoholaWeek 
            Height          =   285
            Left            =   12840
            TabIndex        =   76
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TxtAlcoholAday 
            Height          =   285
            Left            =   9360
            TabIndex        =   74
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox TxtSmokesAday 
            Height          =   285
            Left            =   9360
            TabIndex        =   72
            Top             =   840
            Width           =   735
         End
         Begin VB.ComboBox CboTobacco 
            Height          =   315
            Left            =   4320
            TabIndex        =   60
            Top             =   840
            Width           =   3375
         End
         Begin VB.CommandButton CmdRemoveAlcohol 
            Caption         =   "Remove"
            Height          =   495
            Left            =   8520
            TabIndex        =   57
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton CmdAddAlcoholTobacco 
            Caption         =   "Add "
            Height          =   495
            Left            =   8520
            TabIndex        =   56
            Top             =   1320
            Width           =   2055
         End
         Begin VB.ComboBox CboAlcohol 
            Height          =   315
            Left            =   4320
            TabIndex        =   54
            Top             =   360
            Width           =   3375
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridAlcoholTobacco 
            Height          =   1695
            Left            =   120
            TabIndex        =   55
            Top             =   1320
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2990
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
         Begin VB.Frame Frame11 
            Height          =   1095
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   2535
            Begin VB.CheckBox ChkNonDrinker 
               Caption         =   "Non Drinker"
               BeginProperty Font 
                  Name            =   "Garamond"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   82
               Top             =   240
               Width           =   2175
            End
            Begin VB.CheckBox ChkNonSmoker 
               Caption         =   "Non Smoker"
               BeginProperty Font 
                  Name            =   "Garamond"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               TabIndex        =   81
               Top             =   720
               Width           =   2175
            End
         End
         Begin VB.Label Label17 
            Caption         =   "How many times a week"
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
            Left            =   10320
            TabIndex        =   77
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label20 
            Caption         =   "How many a day"
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
            Left            =   7800
            TabIndex        =   75
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label21 
            Caption         =   "Number a day"
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
            Left            =   7800
            TabIndex        =   73
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Type of Tobacco"
            Height          =   255
            Left            =   2520
            TabIndex        =   59
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Type of Alcohol"
            Height          =   255
            Left            =   2520
            TabIndex        =   53
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Cooking Oils"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   46
         Top             =   3480
         Width           =   10695
         Begin VB.CommandButton CmdRemoveCookingOil 
            Caption         =   "Remove"
            Height          =   495
            Left            =   8520
            TabIndex        =   51
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton CmdAddCookingOils 
            Caption         =   "Add "
            Height          =   495
            Left            =   8520
            TabIndex        =   50
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox CboCookingOils 
            Height          =   315
            Left            =   1920
            TabIndex        =   48
            Top             =   360
            Width           =   5775
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridCookingOils 
            Height          =   1695
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2990
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
            Alignment       =   1  'Right Justify
            Caption         =   "Type Of Oil"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Bedtime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -67080
         TabIndex        =   39
         Top             =   5940
         Width           =   7695
         Begin VB.CommandButton CmdRemoveBedtime 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   71
            Top             =   2160
            Width           =   615
         End
         Begin VB.CommandButton CmdAddBedtime 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   42
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TxtBedTime 
            Height          =   285
            Left            =   5520
            TabIndex        =   41
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox CboBedTime 
            Height          =   315
            Left            =   1200
            TabIndex        =   40
            Top             =   360
            Width           =   3135
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridBedtime 
            Height          =   1695
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2990
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
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Quantity"
            Height          =   195
            Left            =   4440
            TabIndex        =   45
            Top             =   405
            Width           =   975
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Item"
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lunch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   32
         Top             =   5940
         Width           =   7695
         Begin VB.CommandButton CmdRemoveSnack 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6960
            TabIndex        =   68
            Top             =   2160
            Width           =   615
         End
         Begin VB.CommandButton CmdAddLunch 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   35
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TxtLunchQuantity 
            Height          =   285
            Left            =   5640
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox CboLunch 
            Height          =   315
            Left            =   1200
            TabIndex        =   33
            Top             =   360
            Width           =   3375
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridLunch 
            Height          =   1695
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2990
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
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Quantity"
            Height          =   195
            Left            =   4680
            TabIndex        =   38
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Item"
            Height          =   315
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dinner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -67080
         TabIndex        =   25
         Top             =   3180
         Width           =   7695
         Begin VB.CommandButton CmdRemoveDinner 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   70
            Top             =   2160
            Width           =   615
         End
         Begin VB.CommandButton CmdAddDinner 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   28
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TxtDinner 
            Height          =   285
            Left            =   5520
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox CboDinner 
            Height          =   315
            Left            =   1200
            TabIndex        =   26
            Top             =   360
            Width           =   3135
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridDinner 
            Height          =   1695
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2990
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
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Quantity"
            Height          =   195
            Left            =   4440
            TabIndex        =   31
            Top             =   405
            Width           =   975
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Item"
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Afternoon Snack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -67080
         TabIndex        =   18
         Top             =   420
         Width           =   7695
         Begin VB.CommandButton CmdRemoveAfternoonSnack 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   69
            Top             =   2160
            Width           =   615
         End
         Begin VB.CommandButton CmdAddAfternoonSnack 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   21
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TxtAfternoonSnack 
            Height          =   285
            Left            =   5520
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox CboAfternoonSnack 
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   3135
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridAfternoonSnack 
            Height          =   1695
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2990
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Quantity"
            Height          =   300
            Left            =   4440
            TabIndex        =   24
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Item"
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Morning Snack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   11
         Top             =   3180
         Width           =   7695
         Begin VB.CommandButton CmdRemoveMorningSnack 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6960
            TabIndex        =   67
            Top             =   2160
            Width           =   615
         End
         Begin VB.CommandButton CmdAddMorningSnack 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   14
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TxtMorningSnackQuantity 
            Height          =   285
            Left            =   5520
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox CboMorningSnack 
            Height          =   315
            Left            =   1080
            TabIndex        =   12
            Top             =   360
            Width           =   3495
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridMorningSnack 
            Height          =   1695
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2990
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Quantity"
            Height          =   195
            Left            =   4680
            TabIndex        =   17
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Item"
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Breakfast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   4
         Top             =   420
         Width           =   7695
         Begin VB.CommandButton CmdRemoveBreakfast 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6960
            TabIndex        =   66
            Top             =   2160
            Width           =   615
         End
         Begin VB.ComboBox CboBreakfast 
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox TxtBreakfastQuantity 
            Height          =   285
            Left            =   5760
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton CmdAddBreakfast 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6960
            TabIndex        =   5
            Top             =   840
            Width           =   615
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridBreakfast 
            Height          =   1695
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   2990
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
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Item"
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Quantity"
            Height          =   195
            Left            =   4920
            TabIndex        =   9
            Top             =   405
            Width           =   735
         End
      End
      Begin VB.Image Image1 
         Height          =   6735
         Left            =   10920
         Picture         =   "FrmFoodFrequency.frx":0038
         Top             =   3600
         Width           =   17490
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   8880
      Width           =   15735
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   13560
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   10800
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Menu MnuFoodMaintenance 
      Caption         =   "Food Maintenance"
      Begin VB.Menu mnuFoodtype 
         Caption         =   "Add or Remove Food Type"
      End
   End
End
Attribute VB_Name = "FrmFoodFrequency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim RsRetrieve As New ADODB.Recordset

Public Sub PopulateLunch(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '3'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridLunch.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
                GridLunch.Cols = 2
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridLunch.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Public Sub PopulateAfternoonSnack(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '4'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridAfternoonSnack.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
                GridAfternoonSnack.Cols = 2
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridAfternoonSnack.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Public Sub PopulateDinner(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '5'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridDinner.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
                GridDinner.Cols = 2
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridDinner.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Public Sub PopulateBedtime(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '6'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridBedtime.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
                GridBedtime.Cols = 2
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridBedtime.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Private Sub PopulateCookingOil(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '7'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridCookingOils.FormatString = "      TYPE OF COOKING OIL                                "
                GridCookingOils.Cols = 1
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridCookingOils.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Private Sub PopulateAlcoholTobacco(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '8'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
            GridAlcoholTobacco.FormatString = "TYPE OF ALCOHOL / CIGARETTE | HOW MANY A DAY  - (HOW MANY A WEEK)  "
            GridAlcoholTobacco.Cols = 2
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridAlcoholTobacco.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Private Sub PopulateExercise(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '9'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridExercise.FormatString = "      TYPE OF EXERCISE                        |    HOW MANY TIMES A WEEK         "
                GridExercise.Cols = 2
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridExercise.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Private Sub ChkNonDrinker_Click()
    If ChkNonDrinker.Value = 1 Then
        CboAlcohol.Enabled = False: CboAlcohol.Text = ""
        TxtAlcoholAday.Enabled = False: TxtAlcoholAday = ""
        TxtAlcoholaWeek.Enabled = False: TxtAlcoholaWeek = ""
        GridAlcoholTobacco.Clear
        GridAlcoholTobacco.Rows = 1
    Else
        CboAlcohol.Enabled = True: CboAlcohol.Text = ""
        TxtAlcoholAday.Enabled = True: TxtAlcoholAday = ""
        TxtAlcoholaWeek.Enabled = True: TxtAlcoholaWeek = ""
        GridAlcoholTobacco.Clear
        GridAlcoholTobacco.Rows = 1
    End If
End Sub

Private Sub ChkNonSmoker_Click()
    If ChkNonSmoker.Value = 1 Then
        CboTobacco.Enabled = False: CboTobacco = ""
        TxtSmokesAday.Enabled = False: TxtSmokesAday = ""
        GridAlcoholTobacco.Clear
        GridAlcoholTobacco.Rows = 1
    Else
        CboTobacco.Enabled = True: CboTobacco = ""
        TxtSmokesAday.Enabled = True: TxtSmokesAday = ""
        GridAlcoholTobacco.Clear
        GridAlcoholTobacco.Rows = 1
    End If
End Sub

Private Sub CmdAddAfternoonSnack_Click()
    If TxtAfternoonSnack = "" Then MsgBox "Please Input Quantity Before Adding to List!", vbExclamation: Exit Sub
    GridAfternoonSnack.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridAfternoonSnack.Cols = 2
    GridAfternoonSnack.AddItem CboAfternoonSnack & vbTab & TxtAfternoonSnack
End Sub

Private Sub CmdAddAlcoholTobacco_Click()
    If ChkNonDrinker.Value = 0 Then
        If CboAlcohol = "" Then MsgBox "Please Select Alcohol Type Before Adding to List!", vbExclamation: Exit Sub
        If TxtAlcoholAday = "" Then MsgBox "Please Input Alcohol-A-Day Before Adding to List!", vbExclamation: Exit Sub
        If TxtAlcoholaWeek = "" Then MsgBox "Please Input Alcohol-A-Week Before Adding to List!", vbExclamation: Exit Sub
    End If
    If ChkNonSmoker.Value = 0 Then
        If TxtSmokesAday = "" Then MsgBox "Please Input Number of cigarettes a day Before Adding to List!", vbExclamation: Exit Sub
    End If
    
    If ChkNonDrinker.Value = 1 And ChkNonSmoker.Value = 0 Then
        GridAlcoholTobacco.FormatString = "TYPE OF TOBACCO                |  NUMBER OF CIGARETTES A DAY   "
        GridAlcoholTobacco.Cols = 2
        GridAlcoholTobacco.AddItem CboTobacco & vbTab & TxtSmokesAday
    ElseIf ChkNonDrinker.Value = 0 And ChkNonSmoker.Value = 1 Then
        GridAlcoholTobacco.FormatString = " TYPE OF ALCOHOL               | HOW MANY A DAY    | HOW MANY A WEEK   "
        GridAlcoholTobacco.Cols = 2
        GridAlcoholTobacco.AddItem CboAlcohol & vbTab & TxtAlcoholAday & vbtbat & TxtAlcoholaWeek
    ElseIf ChkNonDrinker.Value = 0 And ChkNonSmoker.Value = 0 Then
        GridAlcoholTobacco.FormatString = "TYPE OF ALCOHOL / CIGARETTE | HOW MANY A DAY  - (HOW MANY A WEEK)  "
        GridAlcoholTobacco.Cols = 2
        'GridAlcoholTobacco.AddItem CboAlcohol & vbTab & TxtAlcoholAday & vbTab & TxtAlcoholaWeek & vbTab & CboTobacco & vbTab & TxtSmokesAday
        GridAlcoholTobacco.AddItem CboAlcohol & vbTab & TxtAlcoholAday & " - " & TxtAlcoholaWeek
        GridAlcoholTobacco.AddItem CboTobacco & vbTab & TxtSmokesAday
    End If
End Sub

Private Sub CmdAddBedtime_Click()
    If TxtBedTime = "" Then MsgBox "Please Input Quantity Before Adding to List!", vbExclamation: Exit Sub
    GridBedtime.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridBedtime.Cols = 2
    GridBedtime.AddItem CboBedTime & vbTab & TxtBedTime
End Sub

Private Sub CmdAddBreakfast_Click()
    If TxtBreakfastQuantity = "" Then MsgBox "Please Input Quantity Before Adding to List!", vbExclamation: Exit Sub
    GridBreakfast.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridBreakfast.Cols = 2
    GridBreakfast.AddItem CboBreakfast & vbTab & TxtBreakfastQuantity
End Sub

Private Sub CmdAddCookingOils_Click()
    GridCookingOils.FormatString = "      TYPE OF COOKING OIL                                "
    GridCookingOils.Cols = 1

    GridCookingOils.AddItem CboCookingOils
End Sub

Private Sub CmdAddDinner_Click()
    If TxtDinner = "" Then MsgBox "Please Input Quantity Before Adding to List!", vbExclamation: Exit Sub
    GridDinner.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridDinner.Cols = 2
    GridDinner.AddItem CboDinner & vbTab & TxtDinner
End Sub

Private Sub CmdAddExericise_Click()
    GridExercise.FormatString = "      TYPE OF EXERCISE                        |    HOW MANY TIMES A WEEK         "
    GridExercise.Cols = 2
    GridExercise.AddItem CboExercise & vbTab & TxtExerciseAweek
End Sub

Private Sub CmdAddLunch_Click()
    If TxtLunchQuantity = "" Then MsgBox "Please Input Quantity Before Adding to List!", vbExclamation: Exit Sub
    GridLunch.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridLunch.Cols = 2

    GridLunch.AddItem CboLunch & vbTab & TxtLunchQuantity
End Sub

Private Sub CmdAddMorningSnack_Click()
    If TxtMorningSnackQuantity = "" Then MsgBox "Please Input Quantity Before Adding to List!", vbExclamation: Exit Sub
    GridMorningSnack.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridMorningSnack.Cols = 2
    GridMorningSnack.AddItem CboMorningSnack & vbTab & TxtMorningSnackQuantity
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdRemoveAfternoonSnack_Click()
    GridAfternoonSnack.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridAfternoonSnack.Cols = 2
    GridAfternoonSnack.RemoveItem (GridAfternoonSnack.Row)
End Sub

Private Sub CmdRemoveAlcohol_Click()
    'GridCookingOils.FormatString = "      TYPE OF COOKING OIL                                "
    'GridAlcoholTobacco.Cols = 1
    GridAlcoholTobacco.RemoveItem (GridAlcoholTobacco.Row)
End Sub

Private Sub CmdRemoveBedtime_Click()
    GridBedtime.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridBedtime.Cols = 2
    GridBedtime.RemoveItem (GridBedtime.Row)
End Sub

Private Sub CmdRemoveBreakfast_Click()
    GridBreakfast.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridBreakfast.Cols = 2
    GridBreakfast.RemoveItem (GridBreakfast.Row)
End Sub

Private Sub CmdRemoveCookingOil_Click()
    GridCookingOils.FormatString = "      TYPE OF COOKING OIL                                "
    GridCookingOils.Cols = 1
    GridCookingOils.RemoveItem (GridCookingOils.Row)
End Sub

Private Sub CmdRemoveDinner_Click()
    GridDinner.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridDinner.Cols = 2
    GridDinner.RemoveItem (GridDinner.Row)
End Sub

Private Sub CmdRemoveExercise_Click()
    'GridCookingOils.FormatString = "      TYPE OF COOKING OIL                                "
    'GridCookingOils.Cols = 1
    GridExercise.RemoveItem (GridExercise.Row)
End Sub

Private Sub CmdRemoveMorningSnack_Click()
    GridMorningSnack.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridMorningSnack.Cols = 2

    GridMorningSnack.RemoveItem (GridMorningSnack.Row)
End Sub

Private Sub CmdRemoveLunch_Click()
    GridLunch.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
    GridLunch.Cols = 2

    GridLunch.RemoveItem (GridLunch.Row)
End Sub

Private Sub CmdSave_Click()
    'CmdSave.Enabled = False
    SaveMealsToDB
End Sub

Private Sub Command10_Click()

End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    PopulateFoodComboBoxes
    PopulateDataGrids
    centerform Me
End Sub
Public Sub PopulateDataGrids()
    PopulateBreakfast lvFoodCardNo, lvFoodVisitNo
    PopulateMorningSnack lvFoodCardNo, lvFoodVisitNo
    PopulateLunch lvFoodCardNo, lvFoodVisitNo
    PopulateAfternoonSnack lvFoodCardNo, lvFoodVisitNo
    PopulateDinner lvFoodCardNo, lvFoodVisitNo
    PopulateBedtime lvFoodCardNo, lvFoodVisitNo
    PopulateCookingOil lvFoodCardNo, lvFoodVisitNo
    PopulateAlcoholTobacco lvFoodCardNo, lvFoodVisitNo
    PopulateExercise lvFoodCardNo, lvFoodVisitNo
End Sub
Public Sub PopulateBreakfast(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '1'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridBreakfast.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
                GridBreakfast.Cols = 2
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridBreakfast.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Public Sub PopulateMorningSnack(CardNo As String, VisitNo As Integer)
    Dim lvMeal As String
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND MEALTYPE = '2'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridMorningSnack.FormatString = "      TYPE OF MEAL                        |    QUANTITY MEASURED IN CUPS         "
                GridMorningSnack.Cols = 2
                While .EOF = False
                    lvMeal = FindRecord("MEAL_TYPES", "MEALDESCRIPTION", "MEALCODE = '" & !MEALID & "'")
                    GridMorningSnack.AddItem String(3 - Len(!MEALID), "0") & !MEALID & " - " & UCase(lvMeal) & vbTab & !Quantity
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub

Public Sub PopulateFoodComboBoxes()
    'POPULATE COMBO FOR BREAKFAST MEALS
    RsRecords.Open "SELECT * FROM MEAL_TYPES WHERE MEALID = '1'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboBreakfast.Clear
            While .BOF = False And .EOF = False
                CboBreakfast.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    'POPULATE COMBO FOR MORNING SNACK
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '2'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboMorningSnack.Clear
            While .BOF = False And .EOF = False
                CboMorningSnack.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close

    'POPULATE COMBO FOR LUNCH
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '3'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboLunch.Clear
            While .BOF = False And .EOF = False
                CboLunch.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    'POPULATE COMBO FOR AFTERNOON SNACK
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '4'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboAfternoonSnack.Clear
            While .BOF = False And .EOF = False
                CboAfternoonSnack.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    'POPULATE COMBO FOR DINNER
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '5'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboDinner.Clear
            While .BOF = False And .EOF = False
                CboDinner.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    'POPULATE COMBO FOR BEDTIME
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '6'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboBedTime.Clear
            While .BOF = False And .EOF = False
                CboBedTime.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    
    
    'POPULATE COMBO FOR COOKING OILS
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '7'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboCookingOils.Clear
            While .BOF = False And .EOF = False
                CboCookingOils.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    'POPULATE COMBO FOR ALCOHOL
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '8'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboAlcohol.Clear
            While .BOF = False And .EOF = False
                CboAlcohol.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    'POPULATE COMBO FOR TOBACCO
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '9'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboTobacco.Clear
            While .BOF = False And .EOF = False
                CboTobacco.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    'POPULATE COMBO FOR EXERCISE
    RsRecords.Open "SELECT * FROM MEAL_TYPES  WHERE MEALID = '10'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            CboExercise.Clear
            While .BOF = False And .EOF = False
                CboExercise.AddItem String(3 - Len(!MEALCODE), "0") & !MEALCODE & " - " & UCase(!MEALDESCRIPTION)
                .MoveNext
            Wend
        End With
    RsRecords.Close
End Sub


Private Sub mnuFoodtype_Click()
    FrmCategoryMeals.Show 1
    PopulateFoodComboBoxes
End Sub

Private Sub SaveMealsToDB()
    '1.
        SaveBreakfastToDB
    '2.
        SaveMorningSnack
    '3.
        SaveLunch
    '4.
        SaveAfternoonSnack
    '5.
        SaveDinner
    '6.
        SaveBedTime
    '7.
        SaveCookingOils
    '8.
        SaveAlcoholandTobacco
    '9.
        SaveExercise
        
End Sub
Private Sub SaveBreakfastToDB()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo
    
    RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '1'"
                If GridBreakfast.Rows > 1 Then
                    For i = 1 To GridBreakfast.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 1
                            !MEALID = Val(GetID_NameFromCombo(GridBreakfast.TextMatrix(i, 0), 1))
                            !Quantity = GridBreakfast.TextMatrix(i, 1)
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridBreakfast.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !MEALTYPE = 1
                        !MEALID = Val(GetID_NameFromCombo(GridBreakfast.TextMatrix(i, 0), 1))
                        !Quantity = GridBreakfast.TextMatrix(i, 1)
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub
Private Sub SaveMorningSnack()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo
    
    RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '2'"
                If GridMorningSnack.Rows > 1 Then
                    For i = 1 To GridMorningSnack.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 2
                            !MEALID = Val(GetID_NameFromCombo(GridMorningSnack.TextMatrix(i, 0), 1))
                            !Quantity = GridMorningSnack.TextMatrix(i, 1)
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridMorningSnack.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !MEALTYPE = 2
                        !MEALID = Val(GetID_NameFromCombo(GridMorningSnack.TextMatrix(i, 0), 1))
                        !Quantity = GridMorningSnack.TextMatrix(i, 1)
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub

Private Sub SaveLunch()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo
    
    RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '3'"
                If GridLunch.Rows > 1 Then
                    For i = 1 To GridLunch.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 3
                            !MEALID = Val(GetID_NameFromCombo(GridLunch.TextMatrix(i, 0), 1))
                            !Quantity = GridLunch.TextMatrix(i, 1)
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridLunch.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !MEALTYPE = 3
                        !MEALID = Val(GetID_NameFromCombo(GridLunch.TextMatrix(i, 0), 1))
                        !Quantity = GridLunch.TextMatrix(i, 1)
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub
Private Sub SaveAfternoonSnack()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo
    
    RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '4'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '4'"
                If GridAfternoonSnack.Rows > 1 Then
                    For i = 1 To GridAfternoonSnack.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 4
                            !MEALID = Val(GetID_NameFromCombo(GridAfternoonSnack.TextMatrix(i, 0), 1))
                            !Quantity = GridAfternoonSnack.TextMatrix(i, 1)
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridAfternoonSnack.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !MEALTYPE = 4
                        !MEALID = Val(GetID_NameFromCombo(GridAfternoonSnack.TextMatrix(i, 0), 1))
                        !Quantity = GridAfternoonSnack.TextMatrix(i, 1)
                    .Update
                Next i
            End If
        End With
    RsRecords.Close

End Sub
Public Sub SaveDinner()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo
    
    RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '5'"
                If GridDinner.Rows > 1 Then
                    For i = 1 To GridDinner.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 5
                            !MEALID = Val(GetID_NameFromCombo(GridDinner.TextMatrix(i, 0), 1))
                            !Quantity = GridDinner.TextMatrix(i, 1)
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridDinner.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !MEALTYPE = 5
                        !MEALID = Val(GetID_NameFromCombo(GridDinner.TextMatrix(i, 0), 1))
                        !Quantity = GridDinner.TextMatrix(i, 1)
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub
Private Sub SaveBedTime()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo
    
    RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
               Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '6'"
               If GridBedtime.Rows > 1 Then
                    For i = 1 To GridBedtime.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 6
                            !MEALID = Val(GetID_NameFromCombo(GridBedtime.TextMatrix(i, 0), 1))
                            !Quantity = GridBedtime.TextMatrix(i, 1)
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridBedtime.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !MEALTYPE = 6
                        !MEALID = Val(GetID_NameFromCombo(GridBedtime.TextMatrix(i, 0), 1))
                        !Quantity = GridBedtime.TextMatrix(i, 1)
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub

Public Sub SaveCookingOils()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo
    
    RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '7'"
                If GridCookingOils.Rows > 1 Then
                    For i = 1 To GridCookingOils.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 7
                            !MEALID = Val(GetID_NameFromCombo(GridCookingOils.TextMatrix(i, 0), 1))
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridCookingOils.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !MEALTYPE = 7
                        !MEALID = Val(GetID_NameFromCombo(GridCookingOils.TextMatrix(i, 0), 1))
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub
Private Sub SaveAlcoholandTobacco()
        Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
        TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo

        RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsRecords
                If .EOF = False Then
                    Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '8'"
                    If GridAlcoholTobacco.Rows > 1 Then
                        For i = 1 To GridAlcoholTobacco.Rows - 1
                            .AddNew
                                !CardNumber = TempCardNumber
                                !VISITNUMBER = Val(TempVisitNumber)
                                !MEALTYPE = 8
                                !MEALID = Val(GetID_NameFromCombo(GridAlcoholTobacco.TextMatrix(i, 0), 1))
                                !Quantity = GridAlcoholTobacco.TextMatrix(i, 1)
                            .Update
                        Next
                    End If
                Else
                    For i = 1 To GridAlcoholTobacco.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 8
                            !MEALID = Val(GetID_NameFromCombo(GridAlcoholTobacco.TextMatrix(i, 0), 1))
                            !Quantity = GridAlcoholTobacco.TextMatrix(i, 1)
                        .Update
                    Next i
                End If
            End With
        RsRecords.Close
End Sub
Private Sub SaveExercise()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = lvFoodCardNo: TempVisitNumber = lvFoodVisitNo
    
    RsRecords.Open "SELECT * FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                Conn.Execute "DELETE FROM MEALS WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "' AND MEALTYPE = '9'"
                If GridExercise.Rows > 1 Then
                    For i = 1 To GridExercise.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !MEALTYPE = 9
                            !MEALID = Val(GetID_NameFromCombo(GridExercise.TextMatrix(i, 0), 1))
                            !Quantity = GridExercise.TextMatrix(i, 1)
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridExercise.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !MEALTYPE = 9
                        !MEALID = Val(GetID_NameFromCombo(GridExercise.TextMatrix(i, 0), 1))
                        !Quantity = GridExercise.TextMatrix(i, 1)
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub
