VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FrmObservationOld 
   Caption         =   "Observation Patients"
   ClientHeight    =   9210
   ClientLeft      =   4290
   ClientTop       =   2970
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmObservationOld.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   15105
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   36
      Top             =   4560
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   6376
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Medical History"
      TabPicture(0)   =   "FrmObservationOld.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TxtMedicalHistory"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "GridMedicalHistory"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OptGridMedical"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptText"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Current Medication"
      TabPicture(1)   =   "FrmObservationOld.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "OptCurrentMedText"
      Tab(1).Control(1)=   "OptCurrentMedication"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(3)=   "GridCurrentMedication"
      Tab(1).Control(4)=   "TxtCurrentMedication"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Family Medical History"
      TabPicture(2)   =   "FrmObservationOld.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtFamilyHistory"
      Tab(2).Control(1)=   "GridFamilyHistory"
      Tab(2).Control(2)=   "Frame8"
      Tab(2).Control(3)=   "OptFamilyHistory"
      Tab(2).Control(4)=   "OptFamilyHistText"
      Tab(2).ControlCount=   5
      Begin VB.OptionButton OptFamilyHistText 
         Caption         =   "Additional Notes"
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
         Height          =   270
         Left            =   -65280
         TabIndex        =   73
         Top             =   3045
         Width           =   2295
      End
      Begin VB.OptionButton OptFamilyHistory 
         Caption         =   "List of Conditions"
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
         Height          =   270
         Left            =   -69240
         TabIndex        =   72
         Top             =   3045
         Width           =   2535
      End
      Begin VB.Frame Frame8 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   61
         Top             =   360
         Width           =   4575
         Begin VB.CommandButton Command14 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            TabIndex        =   78
            Top             =   1440
            Width           =   375
         End
         Begin VB.ComboBox CboRelationship 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmObservationOld.frx":0496
            Left            =   1200
            List            =   "FrmObservationOld.frx":04B5
            TabIndex        =   77
            Top             =   1440
            Width           =   2775
         End
         Begin VB.ComboBox CboDiagnosis2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   3855
         End
         Begin VB.ComboBox CboDiagnosisCategory2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   3855
         End
         Begin VB.CommandButton CmdAddFamilyHistory 
            Caption         =   "Add Condition"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            TabIndex        =   65
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton CmdRemoveFamilyHistory 
            Caption         =   "Remove Condition"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   64
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton Command9 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            TabIndex        =   63
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Command8 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            TabIndex        =   62
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "Relationship"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Condition/Disease"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Condition Category"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.OptionButton OptCurrentMedText 
         Caption         =   "Additional Notes"
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
         Height          =   315
         Left            =   -65040
         TabIndex        =   60
         Top             =   3045
         Width           =   2295
      End
      Begin VB.OptionButton OptCurrentMedication 
         Caption         =   "List of Medication"
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
         Height          =   315
         Left            =   -69240
         TabIndex        =   59
         Top             =   3045
         Width           =   2535
      End
      Begin VB.OptionButton OptText 
         Caption         =   "Additional Notes"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   270
         Left            =   9840
         TabIndex        =   56
         Top             =   3045
         Width           =   2295
      End
      Begin VB.OptionButton OptGridMedical 
         Caption         =   "List of Conditions"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   270
         Left            =   5640
         TabIndex        =   55
         Top             =   3045
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.Frame Frame7 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   48
         Top             =   360
         Width           =   4575
         Begin VB.CommandButton Command13 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            TabIndex        =   75
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command12 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            TabIndex        =   74
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton CmdAddMedication 
            Caption         =   "Add Condition"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            TabIndex        =   58
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton CmdRemoveMedication 
            Caption         =   "Remove Condition"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   57
            Top             =   1920
            Width           =   1695
         End
         Begin VB.ComboBox CboDrugs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Width           =   3855
         End
         Begin VB.ComboBox CboPrescriptionCategory 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   49
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label19 
            Caption         =   "Prescription Category"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label18 
            Caption         =   "Prescription Drug"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2535
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   4575
         Begin VB.CommandButton Command5 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            TabIndex        =   54
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Command4 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            TabIndex        =   53
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton CmdRemoveCondition 
            Caption         =   "Remove Condition"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton CmdAddCondition 
            Caption         =   "Add Condition"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            TabIndex        =   45
            Top             =   1920
            Width           =   1695
         End
         Begin VB.ComboBox CboDiagnosisCategory 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   3855
         End
         Begin VB.ComboBox CboDiagnosis 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   3855
         End
         Begin VB.Label Label17 
            Caption         =   "Condition Category"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "Condition/Disease"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   1215
         End
      End
      Begin VSFlex6DAOCtl.vsFlexGrid GridFamilyHistory 
         Height          =   2415
         Left            =   -70200
         TabIndex        =   70
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
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
      Begin VB.TextBox TxtFamilyHistory 
         Height          =   2055
         Left            =   -70200
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   480
         Width           =   7815
      End
      Begin VSFlex6DAOCtl.vsFlexGrid GridCurrentMedication 
         Height          =   2415
         Left            =   -70200
         TabIndex        =   47
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
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
      Begin VB.TextBox TxtCurrentMedication 
         BackColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   -70200
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   7815
      End
      Begin VSFlex6DAOCtl.vsFlexGrid GridMedicalHistory 
         Height          =   2415
         Left            =   4800
         TabIndex        =   44
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
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
      Begin VB.TextBox TxtMedicalHistory 
         Height          =   2055
         Left            =   4800
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   480
         Width           =   7815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Navigation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   13080
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton Command3 
         Caption         =   "Update Current Medication"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   81
         Top             =   6480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update Family History"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   80
         Top             =   5640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Update Medical History "
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   240
         TabIndex        =   79
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh List"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton FrmExit 
         Caption         =   "E&xit"
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
         Left            =   120
         TabIndex        =   4
         Top             =   8400
         Width           =   1695
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Measurements"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   12855
      Begin VB.TextBox TxtWeight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9120
         MaxLength       =   4
         TabIndex        =   35
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TxtRightArmBP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   33
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TxtLeftArmBP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   31
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox TxtHeight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11640
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxtHip 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9120
         MaxLength       =   3
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtWaist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtWrist 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   23
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtLHgrip 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9120
         MaxLength       =   3
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtRHgrip 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox CboDorminantHand 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmObservationOld.frx":050B
         Left            =   2040
         List            =   "FrmObservationOld.frx":0515
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label10 
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
         Left            =   7920
         TabIndex        =   34
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Right Arm BP"
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
         Left            =   4080
         TabIndex        =   32
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Left Arm BP"
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
         Left            =   720
         TabIndex        =   30
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
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
         Left            =   10920
         TabIndex        =   28
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Hip Circumference"
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
         Left            =   7320
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Waist Circumference"
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
         Left            =   3840
         TabIndex        =   24
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Wrist Circumference"
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
         TabIndex        =   22
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Left Hand Grip"
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
         Left            =   7560
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Right Hand Grip"
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
         Left            =   3960
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Dorminant hand"
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
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BMI Calculator"
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
      Left            =   11280
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      Caption         =   "Post Patient"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   8280
      Width           =   12855
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
         Index           =   0
         Left            =   9360
         TabIndex        =   13
         Top             =   240
         Width           =   1215
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
         Left            =   7350
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptCashier 
         Caption         =   "TO CASHIER"
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
         Left            =   5340
         TabIndex        =   10
         Top             =   280
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
         Index           =   2
         Left            =   2970
         TabIndex        =   9
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
         Left            =   480
         TabIndex        =   8
         Top             =   280
         Width           =   2055
      End
      Begin VB.CommandButton CmdPost 
         Caption         =   "POST"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10680
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Doctors Waiting List"
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
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Patient for Observation"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
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
   End
End
Attribute VB_Name = "FrmObservationOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset
Dim RsCombo As New ADODB.Recordset
Dim RsFilter As New ADODB.Recordset
Dim RsRecords As New ADODB.Recordset
Dim RsRetrieve As New ADODB.Recordset
Dim ItemSelected As Integer
Dim lvCardNumber As String
Dim lvVisitNumber As Integer

Public Sub InsertCurrentMedication()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = Grid.TextMatrix(Grid.Row, 0): TempVisitNumber = Grid.TextMatrix(Grid.Row, 1)
    
    RsRecords.Open "SELECT * FROM CURRENT_MEDICATION WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                If GridCurrentMedication.Rows > 1 Then
                    Conn.Execute "DELETE FROM CURRENT_MEDICATION WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'"
                    For i = 1 To GridCurrentMedication.Rows - 1
                        .AddNew
                            !CardNumber = Trim(TempCardNumber)
                            !VISITNUMBER = Val(TempVisitNumber)
                            !PRESCRIPTIONCATEGORYID = Val(GetID_NameFromCombo(GridCurrentMedication.TextMatrix(i, 0), 1))
                            !PRESCRIPTIONID = Val(GetID_NameFromCombo(GridCurrentMedication.TextMatrix(i, 1), 1))
                            !ADDTEXT = TxtCurrentMedication
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridCurrentMedication.Rows - 1
                    .AddNew
                        !CardNumber = Trim(TempCardNumber)
                        !VISITNUMBER = Val(TempVisitNumber)
                        !PRESCRIPTIONCATEGORYID = GetID_NameFromCombo(GridCurrentMedication.TextMatrix(i, 0), 1)
                        !PRESCRIPTIONID = GetID_NameFromCombo(GridCurrentMedication.TextMatrix(i, 1), 1)
                        !ADDTEXT = TxtCurrentMedication
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub

Public Sub ManageProcessFlow(ActiveForm)
    Dim RsControls As New ADODB.Recordset
    RsControls.Open "SELECT * FROM PROCESSFLOW WHERE SCREENID = '" & ActiveForm & "'", Conn, adOpenStatic, adLockOptimistic
        If RsControls.EOF = False Then
            With RsControls
                If !CONSULTATION = 1 Then OptConsultation.Item(1).Enabled = True
                If !DOCTORS = 1 Then OptDoctors.Item(2).Enabled = True
                If !CASHIER = 1 Then OptCashier.Item(3).Enabled = True
                If !PHARMACY = 1 Then OptPharmacy.Item(4).Enabled = True
                If !LAB = 1 Then OptLab.Item(5).Enabled = True
            End With
        End If
End Sub

Private Sub PopulateMedicalHistory()
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM MEDICAL_HISTORY WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridMedicalHistory.FormatString = "   DESCRIPTION OF MEDICAL HISTORY     "
                GridMedicalHistory.Cols = 1
                While .EOF = False
                    GridMedicalHistory.AddItem FindRecord("DIAGNOSIS", "DIAGNOSISDESCRIPTION", "DIAGNOSISID = '" & !DIAGNOSISID & "'")
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Private Sub PopulateFamilyMedicalHistory()
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM FAMILY_MEDICAL_HISTORY WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridFamilyHistory.FormatString = "   DESCRIPTION OF MEDICAL HISTORY     | RELATIONSHIP         "
                GridFamilyHistory.Cols = 2
                While .EOF = False
                    GridFamilyHistory.AddItem FindRecord("DIAGNOSIS", "DIAGNOSISDESCRIPTION", "DIAGNOSISID = '" & !DIAGNOSISID & "'") & vbTab & !Relationship
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Private Sub PopulateCurrentMedication()
    If StrDocCardNo = "" Then Exit Sub
    If RsRetrieve.State = 1 Then Set RsRetrieve = Nothing
    RsRetrieve.Open "SELECT * FROM CURRENT_MEDICATION WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        If RsRetrieve.EOF = False Then
            With RsRetrieve
                GridCurrentMedication.FormatString = "   DESCRIPTION OF CURRENT MEDICATION     "
                GridCurrentMedication.Cols = 1
                While .EOF = False
                    GridCurrentMedication.AddItem FindRecord("PRODUCTS", "PRODUCTNAME", "PRODUCTID = '" & !PRESCRIPTIONID & "'")
                    .MoveNext
                Wend
            End With
        End If
    RsRetrieve.Close
End Sub
Private Sub CboDiagnosisCategory_Click()
On Error GoTo Errorhandler
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
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboDiagnosisCategory2_Click()
On Error GoTo Errorhandler
    Dim lvDiagnosisCategoryID As Long
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    CboDiagnosis.Clear
    lvDiagnosisCategoryID = Mid(CboDiagnosisCategory2, 1, 3)
    RsFilter.Open "SELECT DIAGNOSISID, DIAGNOSISDESCRIPTION FROM DIAGNOSIS WHERE DIAGNOSISCATEGORY = '" & lvDiagnosisCategoryID & "'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsFilter
            While .BOF = False And .EOF = False
                CboDiagnosis2.AddItem !DIAGNOSISID & " - " & !DIAGNOSISDESCRIPTION
                .MoveNext
            Wend
        End With
    RsFilter.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboPrescriptionCategory_click()
On Error GoTo Errorhandler
    Dim lvPrescriptionCategoryID
    'POPULATE COMBO FOR PRESCRIPTION
    CboDrugs.Clear
    lvPrescriptionCategoryID = Mid(CboPrescriptionCategory, 1, 3)
    If RsFilter.State = 1 Then Set RsFilter = Nothing
    RsFilter.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' order by productname", Conn, adOpenStatic, adLockOptimistic
    
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
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub ChkAlcohol_Click()
    If ChkAlcohol.Value = 1 Then
        TxtAlcoholDays.Enabled = True
        TxtAlcoholaWeek.Enabled = True
    Else
        TxtAlcoholDays.Enabled = False
        TxtAlcoholaWeek.Enabled = False
        End If
End Sub

Private Sub ChkSmokingStatus_Click()
    If ChkSmokingStatus.Value = 1 Then
        TxtSmokingDays.Enabled = True
    Else
        TxtSmokingDays.Enabled = False
    End If
End Sub

Private Sub CmdAddCondition_Click()
    GridMedicalHistory.FormatString = "CATEGORY                              | CATEGORY DESCRIPTION                                        "
    GridMedicalHistory.Cols = 2
    GridMedicalHistory.AddItem CboDiagnosisCategory & vbTab & CboDiagnosis
End Sub

Private Sub CmdAddFamilyHistory_Click()
    GridFamilyHistory.FormatString = "CATEGORY                                 | CATEGORY DESCRIPTION                                        | RELATIONSHIP "
    GridFamilyHistory.Cols = 3
    GridFamilyHistory.AddItem CboDiagnosisCategory2 & vbTab & CboDiagnosis2 & vbTab & CboRelationship
End Sub


Private Sub CmdAddMedication_Click()
    GridCurrentMedication.FormatString = "PRESCRIPTION CATEGORY                              | PRESCRIPTION DESCRIPTION                                        "
    GridCurrentMedication.Cols = 2
    GridCurrentMedication.AddItem CboPrescriptionCategory & vbTab & CboDrugs
End Sub

Private Sub CmdRefresh_Click()
    FillObservation
End Sub

Private Sub OptAdmission_Click(Index As Integer)
    ItemSelected = OptAdmission.Item(Index).Index
End Sub

Private Sub CmdRemoveCondition_Click()
    GridMedicalHistory.FormatString = "CATEGORY                              | CATEGORY DESCRIPTION                                        "
    GridMedicalHistory.Cols = 2
    GridMedicalHistory.RemoveItem (GridMedicalHistory.Row)
End Sub

Private Sub CmdRemoveFamilyHistory_Click()
    GridFamilyHistory.FormatString = "CATEGORY                              | CATEGORY DESCRIPTION                                        | RELATIONSHIP"
    GridFamilyHistory.Cols = 3
    GridFamilyHistory.RemoveItem GridFamilyHistory.Row
End Sub


Private Sub CmdRemoveMedication_Click()
    GridCurrentMedication.FormatString = "PRESCRIPTION CATEGORY                              | PRESCRIPTION DESCRIPTION                                        "
    GridCurrentMedication.Cols = 2
    GridCurrentMedication.RemoveItem (GridCurrentMedication.Row)
End Sub

Private Sub Command1_Click()
On Error Resume Next
    'Kill App.Path & "\bmi_calc.exe"
    Shell App.Path & "\bmi_calc.exe", vbNormalFocus
End Sub

Private Sub Command11_Click()

End Sub

Private Sub Command15_Click()
    InsertMedicalHistory
End Sub

Private Sub Command2_Click()
    InsertFamilyMedicalHistory
End Sub

Private Sub Command3_Click()
    InsertCurrentMedication
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    BlnViewMeasurements = False
End Sub

Private Sub OptConsultation_Click(Index As Integer)
    ItemSelected = OptConsultation.Item(Index).Index
End Sub

Private Sub OptCurrentMedication_Click()
    If OptCurrentMedication.Value = True Then
        GridCurrentMedication.Visible = True
        TxtCurrentMedication.Visible = False
    End If
End Sub

Private Sub OptCurrentMedText_Click()
    If OptCurrentMedText.Value = True Then
        GridCurrentMedication.Visible = False
        TxtCurrentMedication.Visible = True
    End If
End Sub

Private Sub OptDoctors_Click(Index As Integer)
    ItemSelected = OptDoctors.Item(Index).Index
End Sub

Private Sub OptObservation_Click(Index As Integer)
    ItemSelected = OptObservation.Item(Index).Index
End Sub

Private Sub OptGrid_Click()

End Sub

Private Sub OptFamilyHistory_Click()
    If OptFamilyHistory.Value = True Then
        GridFamilyHistory.Visible = True
        TxtFamilyHistory.Visible = False
    End If
End Sub

Private Sub OptFamilyHistText_Click()
    If OptFamilyHistText.Value = True Then
        GridFamilyHistory.Visible = False
        TxtFamilyHistory.Visible = True
    End If
    
End Sub

Private Sub OptGridMedical_Click()
    If OptGridMedical.Value = True Then
        GridMedicalHistory.Visible = True
        TxtMedicalHistory.Visible = False
    Else
        GridMedicalHistory.Visible = False
        TxtMedicalHistory.Visible = True
    End If
End Sub

Private Sub Option4_Click()

End Sub

Private Sub OptPharmacy_Click(Index As Integer)
    ItemSelected = OptPharmacy.Item(Index).Index
End Sub

Private Sub CmdPost_Click()
On Error GoTo Errorhandler
Dim StrCardNumber As String
KARI = GlbSysDate
'Dim lvVisitNumber As Integer

  
    'NEW UPDATE FUNCTIONS
        InsertMedicalHistory
        InsertCurrentMedication
        InsertFamilyMedicalHistory
    
    Conn.Execute "UPDATE COMPLAINS SET BP = '" & Grid.TextMatrix(Grid.Row, 5) & "',WEIGHT = '" & TxtWeight & "',HEIGHT = '" & TxtHeight & "',BMINDEX = '" & Grid.TextMatrix(Grid.Row, 7) & "',NURSE = '" & GlbCurrentUser & "',INUSE = '0' WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "'"

    'GET THE PATIENT'S LATEST VISIT NUMBER
    lvVisitNumber = FindRecord("COMPLAINS", "VISITNUMBER", "CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' ORDER BY VISITNUMBER DESC")

    Conn.Execute "INSERT INTO OBSERVATION_MEASUREMENTS (CARDNUMBER,VISITNUMBER,DORMINANTHAND,Righthandgrip,lefthandgrip,WristCircumference,WaistCircumference,HipCircumference,LeftArmBP,RightArmBP,Weight,Height)" & _
                 "VALUES ('" & Grid.TextMatrix(Grid.Row, 0) & "','" & lvVisitNumber & "','" & CboDorminantHand & "','" & TxtRHgrip & "','" & TxtLHgrip & "','" & TxtWrist & "','" & TxtWaist & "','" & TxtHip & "','" & TxtLeftArmBP & "','" & TxtRightArmBP & "','" & TxtWeight & "','" & TxtHeight & "')"
                 
    lvCardNumber = Grid.TextMatrix(Grid.Row, 0)

    StrCardNumber = lvCardNumber
    Select Case ItemSelected
   Case 1
            DUMMY = SendPatient(EnumConsultation, StrCardNumber, KARI)
            'FrmPatients.Show
            'Unload Me
    Case 2
        'To Doctor
            DUMMY = SendPatient(EnumDoctors, StrCardNumber, KARI)
            If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'NurseDoctorRolesCombined'") = 1 Then
                FrmWaitingRoom.Show
            End If
            Unload Me
   Case 3
        'To Cashier
            DUMMY = SendPatient(EnumCashier, StrCardNumber, KARI)
            'FrmCashier.Show
            'Unload Me
    Case 4
        'To Pharmacy
            DUMMY = SendPatient(EnumPharmacy, StrCardNumber, KARI)
            'FrmPharmacy.Show
            'Unload Me
    End Select
    FillObservation
    
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
    'Resume
End Sub

Private Sub cmdUpdate_Click()
KARI = GlbSysDate
    'Conn.Execute "INSERT INTO COMPLAINS (CARDNUMBER,BP,WEIGHT,HEIGHT,VISITDATE,COMPLAINS,DIAGNOSIS,PRESCRIPTION,ADMISSION,REFERRAL,NURSE,DOCTOR,OBSERVED)VALUES " & _
                 "('" & Grid.TextMatrix(1, 0) & "','" & Grid.TextMatrix(1, 4) & "','" & Grid.TextMatrix(1, 5) & "','" & Grid.TextMatrix(1, 6) & "', '" & Format(KARI, "DDMMMYYYY") & "','','','','','','NANCY','','1')"
    
    Conn.Execute "UPDATE COMPLAINS SET BP = '" & Grid.TextMatrix(1, 4) & "',WEIGHT = '" & Grid.TextMatrix(1, 5) & "',HEIGHT = '" & Grid.TextMatrix(1, 6) & "',BMINDEX = '" & Grid.TextMatrix(1, 7) & "' WHERE CARDNUMBER = '" & Grid.TextMatrix(1, 0) & "'"
    MsgBox "Observations Updated Succesfully", vbInformation
    FillObservation
End Sub

Private Sub Form_Load()
    centerform Me
    FillObservation
    SSTab1.Tab = 0
    
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    RsCombo.Open "SELECT DIAGNOSISCATEGORYID, DIAGNOSISCATEGORYDESC FROM DIAGNOSISCATEGORY ORDER BY DIAGNOSISCATEGORYDESC", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboDiagnosisCategory.AddItem String(3 - Len(!DIAGNOSISCATEGORYID), "0") & !DIAGNOSISCATEGORYID & " - " & !DIAGNOSISCATEGORYDESC
                CboDiagnosisCategory2.AddItem String(3 - Len(!DIAGNOSISCATEGORYID), "0") & !DIAGNOSISCATEGORYID & " - " & !DIAGNOSISCATEGORYDESC
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
    
    
  'ASSIGN FORM NAME AS CURRENT FORM
    GlbCurrentForm = EnumObservation
    ManageProcessFlow EnumObservation
End Sub

Private Sub FrmExit_Click()
    Unload Me
End Sub
Private Sub FillObservation()
   ' On Error GoTo ErrorHandler
   KARI = GlbSysDate
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 9
    For i = 0 To Grid.Cols - 1
        Grid.ColAlignment(i) = flexAlignCenterCenter
    Next i
    
    Grid.ColAlignment(1) = flexAlignCenterCenter
 
    Grid.ColDataType(8) = flexDTBoolean
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER| VISIT NUMBER | PATIENTS FULL NAME                           |BILLING COMPANY     |ID NUMBER |BLOOD PRESSURE |WEIGHT | HEIGHT IN METERS |  BMI  |SELECT "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        If BlnViewMeasurements = True Then
            RsGrid.Open "SELECT PATIENT_DETAILS.*, COMPLAINS.VISITNUMBER AS VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND COMPLAINS.CARDNUMBER = '" & StrDocCardNo & "'", Conn, adOpenDynamic, adLockOptimistic
            PopulateMeasurements
            PopulateMedicalHistory
            PopulateCurrentMedication
            PopulateFamilyMedicalHistory
        Else
            RsGrid.Open "SELECT PATIENT_DETAILS.*,COMPLAINS.VISITNUMBER AS VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND COMPLAINS.TOOBSERVATION = '1'", Conn, adOpenDynamic, adLockOptimistic
        End If
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                        .MoveNext
                    Wend
                End With
            End If
    Exit Sub
Errorhandler:
    MsgBox Err.Description

End Sub
Private Sub PopulateMeasurements()
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM OBSERVATION_MEASUREMENTS WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                CboDorminantHand = !DORMINANTHAND
                TxtRHgrip = !RIGHTHANDGRIP
                TxtLHgrip = !LEFTHANDGRIP
                TxtWaist = !WAISTCIRCUMFERENCE
                TxtWrist = !WRISTCIRCUMFERENCE
                TxtHip = !HIPCIRCUMFERENCE
                TxtLeftArmBP = !LEFTARMBP
                TxtRightArmBP = !RIGHTARMBP
                TxtWeight = !Weight
                TxtHeight = !Height
            End With
        End If
End Sub
    
Private Sub Grid_Click()
Dim i As Double

    For i = 0 To Grid.Rows - 1
        If Grid.TextMatrix(i, 9) = "SELECT" Then GoTo SKIPHEADER
        If Grid.TextMatrix(i, 9) = "" Then Grid.TextMatrix(i, 9) = 0
        If Grid.TextMatrix(i, 9) = -1 Then
             Dim StrCardNumber, strVisitNumber, StrPatientName, StrBillingCompany, StrIDNumber, StrBP, StrWeight, StrHeight
             StrCardNumber = Grid.TextMatrix(i, 0): lvCardNumber = Grid.TextMatrix(i, 0)
             strVisitNumber = Grid.TextMatrix(i, 1): lvVisitNumber = Grid.TextMatrix(i, 1)
             StrPatientName = Grid.TextMatrix(i, 2)
             StrBillingCompany = Grid.TextMatrix(i, 3)
             StrIDNumber = Grid.TextMatrix(i, 4)
             StrBP = Grid.TextMatrix(i, 5)
             StrWeight = Grid.TextMatrix(i, 6)
             StrHeight = Grid.TextMatrix(i, 7)
             
             Grid.Clear: Grid.Col = 1: Grid.Rows = 1
             Grid.FormatString = "CARD NUMBER|   PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |BLOOD PRESSURE |WEIGHT | HEIGHT |SELECT "
             Grid.AddItem StrCardNumber & vbTab & strVisitNumber & vbTab & StrPatientName & vbTab & StrBillingCompany & vbTab & StrIDNumber & vbTab & StrBP & vbTab & StrWeight & vbTab & StrHeight
             Exit For
        End If
SKIPHEADER:
    Next
End Sub

Private Sub OptText_Click()
    If OptText.Value = True Then
        TxtMedicalHistory.Visible = True
        GridMedicalHistory.Visible = False
    End If
End Sub

Public Sub InsertMedicalHistory()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = Grid.TextMatrix(Grid.Row, 0): TempVisitNumber = Grid.TextMatrix(Grid.Row, 1)
    
    RsRecords.Open "SELECT * FROM MEDICAL_HISTORY WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                If GridMedicalHistory.Rows > 1 Then
                    Conn.Execute "DELETE FROM MEDICAL_HISTORY WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'"
                    For i = 1 To GridMedicalHistory.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !DIAGNOSISCATEGORYID = Val(GetID_NameFromCombo(GridMedicalHistory.TextMatrix(i, 0), 1))
                            !DIAGNOSISID = Val(GetID_NameFromCombo(GridMedicalHistory.TextMatrix(i, 1), 1))
                            !ADDTEXT = TxtMedicalHistory
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridMedicalHistory.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !DIAGNOSISCATEGORYID = GetID_NameFromCombo(GridMedicalHistory.TextMatrix(i, 0), 1)
                        !DIAGNOSISID = GetID_NameFromCombo(GridMedicalHistory.TextMatrix(i, 1), 1)
                        !ADDTEXT = TxtMedicalHistory
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub
Public Sub InsertFamilyMedicalHistory()
    Dim TempCardNumber As String:  Dim TempVisitNumber As Integer
    TempCardNumber = Grid.TextMatrix(Grid.Row, 0): TempVisitNumber = Grid.TextMatrix(Grid.Row, 1)
    
    RsRecords.Open "SELECT * FROM FAMILY_MEDICAL_HISTORY WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                If GridFamilyHistory.Rows > 1 Then
                    Conn.Execute "DELETE FROM FAMILY_MEDICAL_HISTORY WHERE CARDNUMBER = '" & TempCardNumber & "' AND VISITNUMBER = '" & TempVisitNumber & "'"
                    For i = 1 To GridFamilyHistory.Rows - 1
                        .AddNew
                            !CardNumber = TempCardNumber
                            !VISITNUMBER = Val(TempVisitNumber)
                            !DIAGNOSISCATEGORYID = Val(GetID_NameFromCombo(GridFamilyHistory.TextMatrix(i, 0), 1))
                            !DIAGNOSISID = Val(GetID_NameFromCombo(GridFamilyHistory.TextMatrix(i, 1), 1))
                            !Relationship = GridFamilyHistory.TextMatrix(i, 2)
                            !ADDTEXT = TxtFamilyHistory
                        .Update
                    Next
                End If
            Else
                For i = 1 To GridFamilyHistory.Rows - 1
                    .AddNew
                        !CardNumber = TempCardNumber
                        !VISITNUMBER = Val(TempVisitNumber)
                        !DIAGNOSISCATEGORYID = GetID_NameFromCombo(GridFamilyHistory.TextMatrix(i, 0), 1)
                        !DIAGNOSISID = GetID_NameFromCombo(GridFamilyHistory.TextMatrix(i, 1), 1)
                        !Relationship = GridFamilyHistory.TextMatrix(i, 2)
                        !ADDTEXT = TxtFamilyHistory
                    .Update
                Next i
            End If
        End With
    RsRecords.Close
End Sub

Private Sub TxtHeight_Change()
    KISH = ValidateDataType_Advice(TxtHeight, 0)
    If KISH = False Then TxtHeight.Text = Mid(TxtHeight, 1, Len(TxtHeight) - 1)
End Sub

Private Sub TxtHip_Change()
    KISH = ValidateDataType_Advice(TxtHip, 0)
    If KISH = False Then TxtHip.Text = Mid(TxtHip, 1, Len(TxtHip) - 1)
End Sub

Private Sub TxtLeftArmBP_Change()
    KISH = ValidateDataType_Advice(TxtLeftArmBP, 0)
    If KISH = False Then TxtLeftArmBP.Text = Mid(TxtLeftArmBP, 1, Len(TxtLeftArmBP) - 1)
End Sub


Private Sub TxtLHgrip_Change()
    KISH = ValidateDataType_Advice(TxtLHgrip, 0)
    If KISH = False Then TxtLHgrip.Text = Mid(TxtLHgrip, 1, Len(TxtLHgrip) - 1)
End Sub

Private Sub TxtRHgrip_Change()
    KISH = ValidateDataType_Advice(TxtRHgrip, 0)
    If KISH = False Then TxtRHgrip.Text = Mid(TxtRHgrip, 1, Len(TxtRHgrip) - 1)
End Sub


Private Sub TxtRightArmBP_Change()
    KISH = ValidateDataType_Advice(TxtRightArmBP, 0)
    If KISH = False Then TxtRightArmBP.Text = Mid(TxtRightArmBP, 1, Len(TxtRightArmBP) - 1)
End Sub

Private Sub TxtWaist_Change()
    KISH = ValidateDataType_Advice(TxtWaist, 0)
    If KISH = False Then TxtWaist.Text = Mid(TxtWaist, 1, Len(TxtWaist) - 1)
End Sub

Private Sub TxtWeight_Change()
    KISH = ValidateDataType_Advice(TxtWeight, 0)
    If KISH = False Then TxtWeight.Text = Mid(TxtWeight, 1, Len(TxtWeight) - 1)
End Sub

Private Sub TxtWrist_Change()
    KISH = ValidateDataType_Advice(TxtWrist, 0)
    If KISH = False Then TxtWrist.Text = Mid(TxtWrist, 1, Len(TxtWrist) - 1)
End Sub


