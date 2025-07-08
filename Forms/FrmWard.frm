VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmWard 
   Caption         =   "Ward"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16020
   Icon            =   "FrmWard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   16020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   6375
      Left            =   120
      TabIndex        =   39
      Top             =   2880
      Width           =   7095
      Begin TabDlg.SSTab SSTab2 
         Height          =   6015
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   10610
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Nurse's Notes"
         TabPicture(0)   =   "FrmWard.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "CmdSaveNurseNotes"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame11"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Doctor's Notes"
         TabPicture(1)   =   "FrmWard.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CmdSaveDocNotes"
         Tab(1).Control(1)=   "Frame8"
         Tab(1).ControlCount=   2
         Begin VB.CommandButton CmdSaveDocNotes 
            Caption         =   "Save Notes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -70080
            TabIndex        =   56
            Top             =   5400
            Width           =   1815
         End
         Begin VB.Frame Frame11 
            Caption         =   "Nurse Notes"
            Height          =   4815
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   6615
            Begin VB.TextBox TxtNurseNotes 
               Height          =   1815
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   48
               Top             =   2880
               Width           =   6375
            End
            Begin VSFlex6DAOCtl.vsFlexGrid GridNotesNurse 
               Height          =   2295
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   6375
               _ExtentX        =   11245
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
            Begin VB.Label Label8 
               Caption         =   "New Notes"
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   2640
               Width           =   855
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Nurse Notes"
            Height          =   4575
            Left            =   -74880
            TabIndex        =   43
            Top             =   360
            Width           =   6855
            Begin VB.TextBox Text1 
               Height          =   3615
               Left            =   240
               TabIndex        =   44
               Text            =   "Text1"
               Top             =   840
               Width           =   6495
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   300
               Left            =   5280
               TabIndex        =   45
               Top             =   300
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   529
               _Version        =   393216
               Format          =   16908289
               CurrentDate     =   40747
            End
            Begin VB.Label Label7 
               Caption         =   "Nurse Notes By Date"
               Height          =   255
               Left            =   3495
               TabIndex        =   46
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Doctor Notes"
            Height          =   4935
            Left            =   -74880
            TabIndex        =   42
            Top             =   360
            Width           =   6615
            Begin VB.TextBox TxtDoctorNotes 
               Height          =   2055
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   58
               Top             =   240
               Width           =   6375
            End
            Begin VSFlex6DAOCtl.vsFlexGrid GridNotesDoc 
               Height          =   2175
               Left            =   120
               TabIndex        =   53
               Top             =   2640
               Width           =   6375
               _ExtentX        =   11245
               _ExtentY        =   3836
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
               Caption         =   "New Notes"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   2400
               Width           =   855
            End
         End
         Begin VB.CommandButton CmdSaveNurseNotes 
            Caption         =   "Save Notes"
            Height          =   495
            Left            =   4920
            TabIndex        =   41
            Top             =   5280
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame5 
      Height          =   6375
      Left            =   7320
      TabIndex        =   17
      Top             =   2880
      Width           =   6975
      Begin TabDlg.SSTab SSTab1 
         Height          =   6015
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   10610
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Lab Request"
         TabPicture(0)   =   "FrmWard.frx":047A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label9"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "GridLab"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame6"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame7"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "CmdLabRequest"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "CboLabTests"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "CmdRemoveTest"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "CmdAddTests"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Medicine Request"
         TabPicture(1)   =   "FrmWard.frx":0496
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame9"
         Tab(1).ControlCount=   1
         Begin VB.CommandButton CmdAddTests 
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
            Height          =   300
            Left            =   6240
            TabIndex        =   64
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton CmdRemoveTest 
            Caption         =   "Remove From List"
            Height          =   495
            Left            =   240
            TabIndex        =   63
            Top             =   5400
            Width           =   1815
         End
         Begin VB.ComboBox CboLabTests 
            Height          =   315
            Left            =   2040
            TabIndex        =   60
            Top             =   480
            Width           =   3975
         End
         Begin VB.Frame Frame9 
            Height          =   5535
            Left            =   -74880
            TabIndex        =   23
            Top             =   360
            Width           =   6495
            Begin VB.CommandButton CmdPostMedicine 
               Caption         =   "Post"
               Height          =   495
               Left            =   5040
               TabIndex        =   49
               Top             =   4920
               Width           =   1335
            End
            Begin VB.CommandButton CmdAddDrug 
               Caption         =   "Add To List"
               Height          =   495
               Left            =   120
               TabIndex        =   37
               Top             =   4920
               Width           =   1455
            End
            Begin VB.CommandButton CmdRemoveDrug 
               Caption         =   "Remove From List"
               Height          =   495
               Left            =   1680
               TabIndex        =   36
               Top             =   4920
               Width           =   1455
            End
            Begin VB.ComboBox CboCategory3 
               Height          =   315
               Left            =   1800
               TabIndex        =   29
               Top             =   240
               Width           =   4455
            End
            Begin VB.ComboBox CboProduct3 
               Height          =   315
               Left            =   1800
               TabIndex        =   28
               Top             =   840
               Width           =   4455
            End
            Begin VB.TextBox TxtUnit2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1800
               TabIndex        =   27
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox TxtPrice2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1800
               TabIndex        =   26
               Top             =   2040
               Width           =   1095
            End
            Begin VB.TextBox TxtQuantity 
               Height          =   375
               Left            =   4680
               TabIndex        =   25
               Top             =   1440
               Width           =   1575
            End
            Begin VB.TextBox TxtTotalAmount 
               Enabled         =   0   'False
               Height          =   375
               Left            =   4680
               TabIndex        =   24
               Top             =   2040
               Width           =   1575
            End
            Begin VSFlex6DAOCtl.vsFlexGrid G 
               Height          =   2295
               Left            =   120
               TabIndex        =   38
               Top             =   2520
               Width           =   6255
               _ExtentX        =   11033
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
            Begin VB.Label Label16 
               Caption         =   "Medicine Category ID "
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label17 
               Caption         =   "Medicine Description"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label18 
               Caption         =   "Distribution Unit"
               Height          =   255
               Left            =   480
               TabIndex        =   33
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label Label19 
               Caption         =   "Price Per Unit"
               Height          =   255
               Left            =   600
               TabIndex        =   32
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label20 
               Caption         =   "Quantity To Issue"
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
               Left            =   3000
               TabIndex        =   31
               Top             =   1560
               Width           =   1575
            End
            Begin VB.Label Label21 
               Caption         =   "Total Amount"
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
               Left            =   3240
               TabIndex        =   30
               Top             =   2040
               Width           =   1215
            End
         End
         Begin VB.CommandButton CmdLabRequest 
            Caption         =   "Send Request"
            Height          =   495
            Left            =   4800
            TabIndex        =   22
            Top             =   5400
            Width           =   1815
         End
         Begin VB.Frame Frame7 
            Caption         =   "Results"
            Height          =   2655
            Left            =   3360
            TabIndex        =   20
            Top             =   840
            Width           =   3255
            Begin VB.TextBox TxtLabResults 
               Height          =   2295
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   21
               Top             =   240
               Width           =   3015
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Request"
            Height          =   2655
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   3135
            Begin VB.ListBox LstLabTests 
               Height          =   2310
               ItemData        =   "FrmWard.frx":04B2
               Left            =   120
               List            =   "FrmWard.frx":04B4
               Style           =   1  'Checkbox
               TabIndex        =   62
               Top             =   240
               Width           =   2895
            End
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridLab 
            Height          =   1695
            Left            =   120
            TabIndex        =   59
            Top             =   3600
            Width           =   6495
            _ExtentX        =   11456
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
            Caption         =   "Test To be Conducted"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   480
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Occupant Details"
      Height          =   2775
      Left            =   7320
      TabIndex        =   15
      Top             =   120
      Width           =   6975
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   2415
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
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
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   7095
      Begin VB.TextBox TxtDays 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TxtDoctor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtAdmissionDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TxtBedNumber 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox TxtCardNumber 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Doctor"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Admission Date"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Bed Number"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Card Number"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Ward"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.ComboBox CboBedNumber 
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox CboWard 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label6 
         Caption         =   "Bed Number"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Ward Name/Number"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Exit"
      Height          =   9135
      Left            =   14400
      TabIndex        =   50
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton CmdDischarge 
         Caption         =   "Discharge"
         Height          =   735
         Left            =   120
         TabIndex        =   55
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Close"
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   8280
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmWard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCombo As New ADODB.Recordset
Dim RsFilter As New ADODB.Recordset
Dim RsGrid As New ADODB.Recordset
Dim lvDirectSalesNo As Long
Dim lvCardNumber As String
Dim lvVisitNumber As String

Private Sub CboBedNumber_Click()
On Error GoTo ErrorHandler
    lvCardNumber = FindRecord("INPATIENTS", "CARDNUMBER", "BEDNUMBER = '" & Right(CboBedNumber, 1) & "' AND WARDNUMBER = '" & Val(Mid(CboWard, 1, 3)) & "'")
    lvVisitNumber = Trim(FindRecord("INPATIENTS", "VISITNUMBER", "BEDNUMBER = '" & Right(CboBedNumber, 1) & "' AND WARDNUMBER = '" & Val(Mid(CboWard, 1, 3)) & "'"))
    FillGrid lvCardNumber
    FillNotesGrid lvCardNumber, lvVisitNumber
    FillGridLabResults (lvCardNumber)
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboCategory3_Click()
    'POPULATE COMBO FOR DRUGS BY CATEGORY
    CboProduct3.Clear
    lvPrescriptionCategoryID = Mid(CboCategory3, 1, 3)
    RsCombo.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = '" & lvPrescriptionCategoryID & "'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboProduct3.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboProduct3.AddItem !PRODUCTID & " - " & !ProductName
                End If
                .MoveNext
            Wend
        End With
    RsCombo.Close
End Sub

Private Sub CboLabTests_Click()
    LstLabTests.AddItem CboLabTests.Text
End Sub

Private Sub CboProduct3_Click()
    Dim Pos, StrProductID As String
    Pos = InStr(CboProduct3, "-")
    StrProductID = Left(CboProduct3, Pos - 2)
    TxtUnit2 = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & StrProductID & "'")
    TxtPrice2 = FindRecord("PRODUCTS", "SALEPRICE", "PRODUCTID = '" & StrProductID & "'")
End Sub

Private Sub CboWard_Click()
On Error GoTo ErrorHandler
    Dim lvWardNumber
    'POPULATE COMBO FOR PRESCRIPTION
    CboBedNumber.Clear
    lvWardNumber = Val(Mid(CboWard, 1, 3))
    RsFilter.Open "SELECT BEDNUMBER FROM INPATIENTS WHERE WARDNUMBER = '" & Val(lvWardNumber) & "'", Conn, adOpenDynamic, adLockOptimistic
        With RsFilter
            While .BOF = False And .EOF = False
                    CboBedNumber.AddItem "BED No" & " - " & !BEDNUMBER
                .MoveNext
            Wend
        End With
    RsFilter.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "CATEGORY ID|  PRODUCT CATEGORY NAME    | QUANTITY | AMOUNT"
    'G.ColDataType(2) = flexDTBoolean
    If RsGrid.State = 1 Then Set RsGrid = Nothing
        RsGrid.Open "SELECT * FROM  PRE_DRUGS_SALES WHERE SALENUMBER = '" & lvDirectSalesNo & "' ", Conn, adOpenStatic, adLockOptimistic
            If RsGrid.BOF = False And RsGrid.EOF = False Then
                While RsGrid.EOF = False
                    With RsGrid
                        G.AddItem !CATEGORYID & vbTab & !PRODUCTID + " - " + !PRODUCTDESCRIPTION & vbTab & !Quantity & vbTab & !amount
                    End With
                RsGrid.MoveNext
                Wend
            End If
        RsGrid.Close
    G.Editable = True
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdAddDrug_Click()
    On Error GoTo ErrorHandler
    lvDirectSalesNo = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'DirectSales'")
    Conn.Execute "INSERT INTO PRE_DRUGS_SALES(SALENUMBER,CATEGORYID,PRODUCTID,PRODUCTDESCRIPTION,DISTRIBUTIONUNIT,QUANTITY,AMOUNT,PAYDATE,SOLDBY)" & _
                 "VALUES ('" & Trim(lvDirectSalesNo) & "','" & GetID_NameFromCombo(CboCategory3, 1) & "','" & GetID_NameFromCombo(Replace(CboProduct3, "'", " "), 1) & "','" & GetID_NameFromCombo(Replace(CboProduct3, "'", " "), 2) & "','" & TxtUnit2 & "','" & TxtQuantity & "','" & TxtTotalAmount & "','" & GlbSysDate & "','" & GlbCurrentUser & "')"
    POPULATEGRID
    TxtQuantity = ""
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub
'''Public Function GetID_NameFromCombo(ByVal Combo, ID_or_Name)
'''    On Error GoTo ErrorHandler
'''    Dim Pos As String
'''    Pos = InStr(Combo, "-")
'''    If ID_or_Name = 1 Then
'''        GetID_NameFromCombo = Left(Combo, Pos - 2)
'''    Else
'''        GetID_NameFromCombo = Mid(Combo, Pos + 2, Len(Combo))
'''    End If
'''Exit Function
'''ErrorHandler:
'''    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
'''End Function

Private Sub CmdAddTests_Click()
    FrmLabParameters.Show
End Sub

Private Sub CmdDischarge_Click()
    'conn.Execute "UPDATE COMPLAINS SET
    MsgBox "WIP", vbInformation, "Discharge Succesfull"
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdLabRequest_Click()
On Error GoTo ErrorHandler
    Dim TestsAmount As Integer
    Dim lvTestList As String
    
    If CmdLabRequest.Caption = "New Request" Then
        LstLabTests.Clear
        TxtLabResults = ""
        CmdLabRequest.Caption = "Send Request"
    Exit Sub
    End If
    
    If LstLabTests.ListCount < 1 Then Exit Sub
    'GET THE TESTS TO BE CONDUCTED
    For j = 0 To LstLabTests.ListCount - 1
        If j = 0 Then
            lvTestList = LstLabTests.List(j) & Chr(13)
        Else
            lvTestList = lvTestList + LstLabTests.List(j) & Chr(13)
        End If
    Next
    If LstLabTests.ListCount < 1 Then MsgBox "Please enter the lab request before sending to LAB.", vbInformation: Exit Sub
    Conn.Execute "INSERT INTO COMPLAINS (CARDNUMBER,VISITDATE,BILLINGCOMPANY,LABREQUEST,TOLABORATORY,ADMISSIONNUMBER,DOCTOR)VALUES ('000', '" & Format(GlbSysDate, "DD MMM YYYY") & "','001','" & lvTestList & "','1','" & TxtCardNumber & "','" & GlbCurrentUser & "')"
    
    'LOOP THROUGH THE TESTS TO CALCULATE TOTAL AMOUNT
    For i = 0 To LstLabTests.ListCount - 1
            TestsAmount = TestsAmount + FindRecord("LABTESTPARAMETERS", "AMOUNT", "TESTID = '" & GetID_NameFromCombo(CboLabTests, 1) & "'")
    Next
    
    Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,VISITDATE,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT)" & _
    "VALUES('" & TxtCardNumber & "','" & lvVisitNumber & "','" & GlbSysDate & "','002','Lab Test(s)','1','" & TestsAmount & "')"
    
    MsgBox "Lab Request Sent Succesfully", vbInformation
    CmdLabRequest.Enabled = False
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CmdPostMedicine_Click()
On Error GoTo ErrorHandler
Dim lvVisitNo As String
    'INSERT INTO PRESCRIPTION AND DELETE FROM PRE-DRUG SALES THEN SEND TO PHARMACY
            If G.Rows > 1 Then
            
                If G.Rows <= 1 Then MsgBox "Please Select Prescription before Posting Drugs.", vbInformation: Exit Sub
                Conn.Execute "INSERT INTO COMPLAINS (CARDNUMBER,VISITDATE,BILLINGCOMPANY,TOPHARMACY,ADMISSIONNUMBER,DOCTOR)" & _
                "VALUES ('000', '" & Format(GlbSysDate, "DD MMM YYYY") & "','001','1','" & TxtCardNumber & "','" & GlbCurrentUser & "')"
                
                lvVisitNo = FindRecord("COMPLAINS", "VISITNUMBER", "ADMISSIONNUMBER = '" & TxtCardNumber & "' AND CARDNUMBER = '000' ORDER BY VISITNUMBER DESC")
                
                For i = 1 To G.Rows - 1
                    Conn.Execute "INSERT INTO PRESCRIPTION  (CARDNUMBER, VISITNUMBER,BILLINGCO,VISITDATE,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT,PAYDATE,PAYMENTMODE,PAYMENTSTATUS)" & _
                    "VALUES('" & StrDocCardNo & "', '" & lvVisitNo & "','001','" & Format(GlbSysDate, "DD MMM YYYY") & "','" & GetID_NameFromCombo(G.TextMatrix(i, 1), 1) & "', '" & GetID_NameFromCombo(G.TextMatrix(i, 1), 2) & "','" & G.TextMatrix(i, 2) & "','" & CashAmount & "','" & Format(GlbSysDate, "DD MMM YYYY") & "','1','0')"
                Next
            Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = " & lvDirectSalesNo & " + 1 WHERE ITEMNAME = 'DIRECTSALES'"
            End If
            CmdPostMedicine.Enabled = False
        MsgBox "Patient Drugs Posted Succesfully", vbInformation
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CmdRemove_Click()

End Sub

Private Sub CmdRemoveTest_Click()
On Error GoTo ErrorHandler
    If LstLabTests.ListIndex = -1 Then Exit Sub
    LstLabTests.RemoveItem (LstLabTests.ListIndex)
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CmdSaveDocNotes_Click()
    If CmdSaveDocNotes.Caption = "Add Note" Then
        TxtDoctorNotes.Text = ""
        CmdSaveDocNotes.Caption = "Save Notes"
        Exit Sub
    End If
    If TxtDoctorNotes = "" Then Exit Sub
    'lvVisitNo = GridNotesDoc.TextMatrix(GridNotesDoc.Row, 2)
    Conn.Execute "INSERT INTO NURSE_DOC_NOTES(CARDNUMBER,VISITNUMBER,NOTESDATE,DOCTORSNOTES)" & _
    "VALUES('" & TxtCardNumber & "','" & lvVisitNumber & "','" & GlbSysDate & "','" & TxtDoctorNotes & "')"
    
    CmdSaveDocNotes.Caption = "Save Notes"
    FillNotesGrid TxtCardNumber, lvVisitNo
End Sub

Private Sub CmdSaveNurseNotes_Click()
    If CmdSaveNurseNotes.Caption = "Add Note" Then
        TxtNurseNotes.Text = ""
        CmdSaveNurseNotes.Caption = "Save Notes"
        Exit Sub
    End If
    If TxtNurseNotes = "" Then Exit Sub
    'lvVisitNo = GridNotesNurse.TextMatrix(GridNotesNurse.Row, 2)
    Conn.Execute "INSERT INTO NURSE_DOC_NOTES(CARDNUMBER,VISITNUMBER,NOTESDATE,NURSESNOTES)" & _
    "VALUES('" & TxtCardNumber & "','" & lvVisitNumber & "','" & GlbSysDate & "','" & Replace(TxtNurseNotes, "'", " ") & "')"
    
    CmdSaveNurseNotes.Caption = "Save Notes"
    FillNotesGrid TxtCardNumber, lvVisitNumber
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
GlbCurrentForm = EnumWard
On Error GoTo ErrorHandler
    'POPULATE COMBO FOR WARDS
    RsCombo.Open "SELECT DISTINCT WARDS.WARDNUMBER, WARDS.WardDescription AS WARDS FROM WARDS INNER JOIN BEDS_AVAILABILITY ON WARDS.WardNumber = BEDS_AVAILABILITY.WardNumber WHERE (BEDS_AVAILABILITY.OccupiedBy = 0)", Conn, adOpenDynamic, adLockOptimistic
        With RsCombo
            While .BOF = False And .EOF = False
                CboWard.AddItem String(3 - Len(!WARDNUMBER), "0") & !WARDNUMBER & " - " & !WARDS
                .MoveNext
            Wend
        End With
    RsCombo.Close
    
    'POPULATE COMBO FOR PRESCRIPTION CATEGORY
    RsCombo.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboCategory3.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                .MoveNext
            Wend
        End With
    RsCombo.Close
    
    'POPULATE COMBO FOR LAB TESTS
    RsCombo.Open "SELECT TESTID, TESTDESCRIPTION FROM LABTESTPARAMETERS", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboLabTests.AddItem String(3 - Len(!TESTID), "0") & !TESTID & " - " & !TESTDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close

    centerform Me
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Private Sub FillGrid(lvCardNumber)
    On Error GoTo ErrorHandler
  ' KARI = GlbSysDate
  
  StrDocCardNo = lvCardNumber
  
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 7
    Grid.ColAlignment(1) = flexAlignCenterCenter
    'Grid.ColDataType(7) = flexDTBoolean
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "BED NUMBER | CARD NUMBER | ADMISSION DATE |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |   VISIT DATE "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN INPATIENTS ON PATIENT_DETAILS.CARDNUMBER = INPATIENTS.CARDNUMBER AND INPATIENTS.CARDNUMBER = '" & StrDocCardNo & "'", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !BEDNUMBER & vbTab & !cardnumber & vbTab & !ADMISSIONDATE & vbTab & !SURNAME & " " & !FIRSTNAME & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                        TxtCardNumber = !cardnumber
                        TxtAdmissionDate = Format(!ADMISSIONDATE, "DD MMM YYYY")
                        TxtBedNumber = GetID_NameFromCombo(CboBedNumber, 0)
                        'TxtDoctorNotes = !DOCTORSNOTES
                        .MoveNext
                    Wend
                End With
            End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
    'Resume
End Sub
Private Sub FillNotesGrid(lvCardNumber, lvVisitNumber)
    On Error GoTo ErrorHandler
  ' KARI = GlbSysDate
  
  StrDocCardNo = lvCardNumber
  
    GridNotesDoc.Clear
    GridNotesNurse.Clear
    GridNotesNurse.Rows = 1
    GridNotesDoc.Rows = 1
    GridNotesDoc.Cols = 4
    GridNotesNurse.Cols = 4
    GridNotesDoc.ColAlignment(1) = flexAlignCenterCenter
    'Grid.ColDataType(7) = flexDTBoolean
    GridNotesDoc.ColWidth(1) = 3105
    GridNotesDoc.ColWidth(2) = 3990
    GridNotesNurse.FormatString = "DATE OF NOTES | VISIT NUMBER | CARD NUMBER | NURSE'S NOTES   "
    GridNotesDoc.FormatString = "DATE OF NOTES | CARD NUMBER | VISIT NUMBER | DOCTOR'S NOTES    "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT * FROM NURSE_DOC_NOTES WHERE CARDNUMBER = '" & lvCardNumber & "' AND VISITNUMBER = '" & lvVisitNumber & "'", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        If !NURSESNOTES <> "" Then
                            GridNotesNurse.AddItem !NotesDate & vbTab & !cardnumber & vbTab & !VISITNUMBER & vbTab & !NURSESNOTES
                        End If
                        If !DOCTORSNOTES <> "" Then
                            GridNotesDoc.AddItem !NotesDate & vbTab & !cardnumber & vbTab & !VISITNUMBER & vbTab & !DOCTORSNOTES
                        End If
                        .MoveNext
                    Wend
                End With
            End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
    'Resume
End Sub
Private Sub FillGridLabResults(lvCardNumber)
    On Error GoTo ErrorHandler
  ' KARI = GlbSysDate
  
  StrDocCardNo = lvCardNumber
  
    GridLab.Clear
    GridLab.Rows = 1
    GridLab.Cols = 4
    GridLab.ColAlignment(1) = flexAlignCenterCenter
    'Grid.ColDataType(7) = flexDTBoolean
    GridLab.ColWidth(1) = 3105
    GridLab.ColWidth(2) = 3990
    GridLab.FormatString = " CARD NUMBER | LAB REQUEST |  LAB RESULTS  |"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT ADMISSIONNUMBER,LABREQUEST,LABRESULTS FROM COMPLAINS WHERE CARDNUMBER = '000' AND ADMISSIONNUMBER = '" & TxtCardNumber & "'"
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridLab.AddItem !ADMISSIONNUMBER & vbTab & !LABREQUEST & vbTab & !LABRESULTS
                        .MoveNext
                    Wend
                End With
            End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
    'Resume
End Sub

Private Sub GridLab_Click()
    Dim ArrRequest
    LstLabTests.Clear: TxtLabResults = ""
    ArrRequest = Split(GridLab.TextMatrix(GridLab.Row, 1), Chr(13))
    
    For i = 0 To UBound(ArrRequest) - 1
        LstLabTests.AddItem ArrRequest(i)
    Next
    TxtLabResults = GridLab.TextMatrix(GridLab.Row, 2)
    CmdLabRequest.Caption = "New Request"
End Sub


Private Sub GridNotesDoc_Click()
    CmdSaveDocNotes.Caption = "Add Note"
    TxtDoctorNotes = GridNotesDoc.TextMatrix(GridNotesDoc.Row, 3)
End Sub

Private Sub GridNotesNurse_SelChange()
    CmdSaveNurseNotes.Caption = "Add Note"
    TxtNurseNotes = GridNotesNurse.TextMatrix(GridNotesNurse.Row, 3)
End Sub

Private Sub TxtQuantity_Change()
    If TxtPrice2 = "" Or TxtUnit2 = "" Or TxtQuantity = "" Then TxtTotalAmount = "": Exit Sub
    TxtTotalAmount = (TxtPrice2 / TxtUnit2) * TxtQuantity
End Sub
