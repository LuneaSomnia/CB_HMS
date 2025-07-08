VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FrmTreatment 
   ClientHeight    =   10200
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15990
   Icon            =   "FrmTreatment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   15990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame14 
      Height          =   830
      Left            =   13920
      TabIndex        =   80
      Top             =   9270
      Width           =   1935
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
         TabIndex        =   81
         Top             =   240
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   15795
      _ExtentX        =   27861
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "History of illness"
      TabPicture(0)   =   "FrmTreatment.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TxtSave"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "CAD Risk Factors"
      TabPicture(1)   =   "FrmTreatment.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Previous Medical History"
      TabPicture(2)   =   "FrmTreatment.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label11"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "SSTab2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text7"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text8"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text9"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Check22"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Check23"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Command5"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Social History"
      TabPicture(3)   =   "FrmTreatment.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Physical Examination"
      TabPicture(4)   =   "FrmTreatment.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame17"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame18"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame19"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame20"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Diagnosis"
      TabPicture(5)   =   "FrmTreatment.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame13"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "SSTab3"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Tests"
      TabPicture(6)   =   "FrmTreatment.frx":04EA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "SSTab4"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin TabDlg.SSTab SSTab4 
         Height          =   4575
         Left            =   120
         TabIndex        =   182
         Top             =   600
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   8070
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   14737632
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Lab Result Image"
         TabPicture(0)   =   "FrmTreatment.frx":0506
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame16"
         Tab(0).Control(1)=   "Frame15"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Test Results Images"
         TabPicture(1)   =   "FrmTreatment.frx":0522
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame22"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame22 
            Caption         =   "Imaging"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   4095
            Left            =   120
            TabIndex        =   189
            Top             =   360
            Width           =   15255
            Begin VB.ListBox LstTests 
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2610
               Left            =   120
               TabIndex        =   201
               Top             =   840
               Width           =   3495
            End
            Begin VB.OptionButton OptNone 
               Caption         =   "Option1"
               Height          =   195
               Left            =   3240
               TabIndex        =   200
               Top             =   480
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox TxtLabResults 
               Height          =   3615
               Left            =   3720
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   199
               Top             =   360
               Width           =   11415
            End
            Begin VB.TextBox TxtEcho 
               Height          =   3615
               Left            =   3720
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   198
               Top             =   360
               Width           =   11415
            End
            Begin VB.CheckBox ChkViewImage 
               Caption         =   "View Image"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   375
               Left            =   480
               TabIndex        =   197
               Top             =   360
               Width           =   2055
            End
            Begin VB.OptionButton OptEcho 
               Caption         =   "Echo"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   196
               Top             =   3120
               Width           =   1215
            End
            Begin VB.OptionButton OptLabResults 
               Caption         =   "Lab Results"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   195
               Top             =   2400
               Width           =   2175
            End
            Begin VB.TextBox TxTXray 
               Height          =   3615
               Left            =   3720
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   194
               Top             =   360
               Width           =   11415
            End
            Begin VB.TextBox TxtUltraSound 
               Height          =   3375
               Left            =   3840
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   193
               Top             =   360
               Width           =   11175
            End
            Begin VB.TextBox TxTCtScan 
               Height          =   3495
               Left            =   3840
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   192
               Top             =   360
               Width           =   11175
            End
            Begin VB.TextBox TxTMRI 
               Height          =   3615
               Left            =   3840
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   191
               Top             =   360
               Width           =   11295
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Save Tests"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1800
               TabIndex        =   190
               Top             =   3480
               Width           =   1815
            End
            Begin VB.OptionButton OptUltrasountImage 
               Caption         =   "ULTRASOUND Image"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   1
               Top             =   945
               Width           =   2655
            End
            Begin VB.OptionButton OptCTScanImage 
               Caption         =   "CT SCAN Image"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   3
               Top             =   1440
               Width           =   2655
            End
            Begin VB.OptionButton OptMRIImage 
               Caption         =   "MRI Image"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   4
               Top             =   2880
               Width           =   2535
            End
            Begin VB.OptionButton OptXRayImage 
               Caption         =   "X-RAY Image"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   2
               Top             =   1920
               Width           =   1935
            End
         End
         Begin VB.Frame Frame15 
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
            Height          =   4095
            Left            =   -67200
            TabIndex        =   186
            Top             =   360
            Width           =   7575
            Begin VB.TextBox TxtLabResultsOld 
               Height          =   3735
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   188
               Top             =   240
               Width           =   7335
            End
            Begin VB.CommandButton CmdViewScan 
               Caption         =   "View Full Image"
               Height          =   435
               Left            =   960
               TabIndex        =   187
               Top             =   3000
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.Image ImgPreview 
               Height          =   3315
               Left            =   120
               Picture         =   "FrmTreatment.frx":053E
               Stretch         =   -1  'True
               Top             =   240
               Visible         =   0   'False
               Width           =   3240
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Lab Request"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   -74880
            TabIndex        =   183
            Top             =   480
            Width           =   6735
            Begin VB.ListBox LstLabTests 
               Height          =   3375
               Left            =   120
               TabIndex        =   185
               Top             =   240
               Width           =   6375
            End
            Begin VB.CheckBox ChkLabTest 
               Caption         =   "Select Lab Test"
               Height          =   255
               Left            =   1680
               TabIndex        =   184
               Top             =   0
               Width           =   1455
            End
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   4575
         Left            =   -70080
         TabIndex        =   159
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8070
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         ForeColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "ICD-10"
         TabPicture(0)   =   "FrmTreatment.frx":158C8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame21"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Prescription"
         TabPicture(1)   =   "FrmTreatment.frx":158E4
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame25"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Assesment and Plan"
         TabPicture(2)   =   "FrmTreatment.frx":15900
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame26"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame26 
            Caption         =   "Assesment && Plan"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   -74880
            TabIndex        =   175
            Top             =   360
            Width           =   10575
            Begin VB.CommandButton CmdSaveAssesmentPlan 
               Caption         =   "Save Assesment && Plan"
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
               Left            =   7800
               TabIndex        =   180
               Top             =   3360
               Width           =   2295
            End
            Begin VB.TextBox TxTAssesment 
               Height          =   1815
               Left            =   1680
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   177
               Top             =   240
               Width           =   8775
            End
            Begin VB.TextBox TxtPlan 
               Height          =   1815
               Left            =   1680
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   176
               Top             =   2160
               Width           =   8775
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               Caption         =   "ASSESMENT :"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   179
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "PLAN :"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   178
               Top             =   2280
               Width           =   615
            End
         End
         Begin VB.Frame Frame25 
            Caption         =   "Prescription Details"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   120
            TabIndex        =   170
            Top             =   360
            Width           =   10575
            Begin VB.CommandButton CmdSavePrescription 
               Caption         =   "Save Prescription"
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
               Left            =   7920
               TabIndex        =   173
               Top             =   3360
               Width           =   2175
            End
            Begin VB.TextBox TxtPrescription 
               Height          =   3615
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   171
               Top             =   360
               Width           =   10335
            End
         End
         Begin VB.Frame Frame21 
            Caption         =   "ICD 10 Clasification"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   4215
            Left            =   -75000
            TabIndex        =   160
            Top             =   360
            Width           =   10695
            Begin VB.Frame Frame23 
               Height          =   3975
               Left            =   120
               TabIndex        =   161
               Top             =   240
               Width           =   10575
               Begin VB.ComboBox CboICD10Category 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   167
                  Top             =   405
                  Width           =   9135
               End
               Begin VB.ListBox LstICD10 
                  Height          =   2205
                  Left            =   120
                  TabIndex        =   166
                  Top             =   1560
                  Width           =   5655
               End
               Begin VB.ComboBox CboICD10SubCategory 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   165
                  Top             =   1005
                  Width           =   9135
               End
               Begin VB.Frame Frame24 
                  Caption         =   "Selected Disease Codes"
                  Height          =   2655
                  Left            =   5880
                  TabIndex        =   163
                  Top             =   1320
                  Width           =   4695
                  Begin VB.CommandButton CmdSaveDiagnosis 
                     Caption         =   "Save Diagnosis"
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
                     Left            =   2640
                     TabIndex        =   172
                     Top             =   1920
                     Width           =   1935
                  End
                  Begin VB.ListBox LstSelectedICD10 
                     Height          =   2205
                     Left            =   120
                     TabIndex        =   164
                     Top             =   240
                     Width           =   4455
                  End
               End
               Begin VB.CommandButton CboRemove 
                  Caption         =   "Remove From List"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   735
                  Left            =   9360
                  TabIndex        =   162
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Label Label49 
                  Caption         =   "Select Disease Category"
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
                  TabIndex        =   169
                  Top             =   120
                  Width           =   2775
               End
               Begin VB.Label Label51 
                  Caption         =   "Select Disease Sub - Category"
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
                  TabIndex        =   168
                  Top             =   720
                  Width           =   2775
               End
            End
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save Medical History"
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
         Left            =   -62520
         TabIndex        =   152
         Top             =   3300
         Width           =   2175
      End
      Begin VB.CommandButton TxtSave 
         Caption         =   "Save Partial History of Presenting illness"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -62160
         TabIndex        =   150
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Mamogram"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -64200
         TabIndex        =   143
         Top             =   4740
         Width           =   1455
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Pap smear"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -68640
         TabIndex        =   142
         Top             =   4740
         Width           =   1335
      End
      Begin VB.Frame Frame20 
         Caption         =   "Abdomen"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -67200
         TabIndex        =   127
         Top             =   3660
         Width           =   7815
         Begin VB.TextBox TxtGut 
            Height          =   975
            Left            =   4200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   149
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox TxtAbdomen 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   147
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label47 
            Caption         =   "GUT"
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
            Left            =   3720
            TabIndex        =   148
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Respiratory"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   -74880
         TabIndex        =   126
         Top             =   3660
         Width           =   7575
         Begin VB.TextBox TxtCNS 
            Height          =   975
            Left            =   4200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   146
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox TxtRespiratory 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   144
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label46 
            Caption         =   "CNS"
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
            Left            =   3720
            TabIndex        =   145
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Extremities"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   96
         Top             =   1620
         Width           =   15375
         Begin VB.CommandButton Command7 
            Caption         =   "Save Physical Examination"
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
            Left            =   12360
            TabIndex        =   154
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox TxtPEOthers 
            Height          =   735
            Left            =   6000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   125
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox TxtMurmurs 
            Height          =   1485
            Left            =   11640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   124
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox TxtSThreeSfour 
            Height          =   285
            Left            =   9480
            TabIndex        =   122
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox TxTSOneSTwo 
            Height          =   285
            Left            =   9480
            TabIndex        =   120
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox TxtApicalImpulse 
            Height          =   285
            Left            =   9480
            TabIndex        =   118
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox TxtRightPT 
            Height          =   285
            Left            =   4680
            TabIndex        =   113
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox TxtRightPA 
            Height          =   285
            Left            =   4680
            TabIndex        =   112
            Top             =   680
            Width           =   1095
         End
         Begin VB.TextBox TxtRightDP 
            Height          =   285
            Left            =   4680
            TabIndex        =   111
            Top             =   1120
            Width           =   1095
         End
         Begin VB.TextBox TxtRightFA 
            Height          =   285
            Left            =   4680
            TabIndex        =   110
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtLeftPT 
            Height          =   285
            Left            =   1800
            TabIndex        =   105
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox TxtLeftPA 
            Height          =   285
            Left            =   1800
            TabIndex        =   104
            Top             =   680
            Width           =   1095
         End
         Begin VB.TextBox TxtLeftDP 
            Height          =   285
            Left            =   1800
            TabIndex        =   103
            Top             =   1120
            Width           =   1095
         End
         Begin VB.TextBox TxtLeftFA 
            Height          =   285
            Left            =   1800
            TabIndex        =   102
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label37 
            Caption         =   "MURMURS"
            Height          =   255
            Left            =   11640
            TabIndex        =   123
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label36 
            Caption         =   "S3  S4"
            Height          =   255
            Left            =   9480
            TabIndex        =   121
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label35 
            Caption         =   "S1  S2"
            Height          =   255
            Left            =   9480
            TabIndex        =   119
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label34 
            Caption         =   "APICAL IMPULSE"
            Height          =   255
            Left            =   9480
            TabIndex        =   117
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label33 
            Caption         =   "CARDIAC"
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
            Left            =   8520
            TabIndex        =   116
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label32 
            Caption         =   "OTHERS"
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
            Left            =   6000
            TabIndex        =   115
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "PULSE RIGHT"
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
            Left            =   3000
            TabIndex        =   114
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "PT"
            Height          =   255
            Left            =   4200
            TabIndex        =   109
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "DP"
            Height          =   255
            Left            =   4200
            TabIndex        =   108
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "PA"
            Height          =   255
            Left            =   4320
            TabIndex        =   107
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "FA"
            Height          =   255
            Left            =   4200
            TabIndex        =   106
            Top             =   285
            Width           =   375
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "PT"
            Height          =   255
            Left            =   1320
            TabIndex        =   101
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "DP"
            Height          =   255
            Left            =   1320
            TabIndex        =   100
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "PA"
            Height          =   255
            Left            =   1320
            TabIndex        =   99
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "FA"
            Height          =   255
            Left            =   1320
            TabIndex        =   98
            Top             =   285
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "PULSE LEFT"
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
            Left            =   120
            TabIndex        =   97
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Higher Functions"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   -74760
         TabIndex        =   85
         Top             =   420
         Width           =   15375
         Begin VB.TextBox TxtAdenopathy 
            Height          =   315
            Left            =   10800
            TabIndex        =   95
            Top             =   600
            Width           =   4455
         End
         Begin VB.CheckBox ChkThyromegally 
            Caption         =   "THYROMEGALLY"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8280
            TabIndex        =   93
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox ChkBruits 
            Caption         =   "BRUITS"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   92
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox ChkJVP 
            Caption         =   "JVP"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   91
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox TxtOropharynx 
            Height          =   315
            Left            =   1560
            TabIndex        =   89
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox TxtHead 
            Height          =   315
            Left            =   1560
            TabIndex        =   87
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label21 
            Caption         =   "ADENOPATHY"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10800
            TabIndex        =   94
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label20 
            Caption         =   "NECK"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   90
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "OROPHARYNX"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "HEAD"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   86
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame13 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   82
         Top             =   540
         Width           =   4695
         Begin VB.CommandButton CmdAccuteOther 
            Caption         =   "Save Acute && Other"
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
            Left            =   2160
            TabIndex        =   174
            Top             =   3840
            Width           =   2055
         End
         Begin VB.TextBox TxtOther 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   157
            Top             =   2640
            Width           =   4455
         End
         Begin VB.TextBox TxtAccute 
            Height          =   1575
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   156
            Top             =   600
            Width           =   4455
         End
         Begin VB.Label Label16 
            Caption         =   "ACUTE"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "OTHER"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   2280
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Social History"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4575
         Left            =   -74640
         TabIndex        =   61
         Top             =   420
         Width           =   15135
         Begin VB.Frame Frame11 
            Caption         =   "Review of Systems"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3375
            Left            =   240
            TabIndex        =   66
            Top             =   1080
            Width           =   14775
            Begin VB.CommandButton Command6 
               Caption         =   "Save Social History"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   12360
               TabIndex        =   153
               Top             =   1080
               Width           =   2175
            End
            Begin VB.CheckBox ChkPreviousColonoscopy 
               Caption         =   "PREVIOUS COLONOSCOPY"
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
               Left            =   7080
               TabIndex        =   78
               Top             =   840
               Width           =   2415
            End
            Begin VB.CheckBox ChkSnorring 
               Caption         =   "SNORRING"
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
               Left            =   7080
               TabIndex        =   77
               Top             =   360
               Width           =   1695
            End
            Begin VB.CheckBox ChkParaesthesias 
               Caption         =   "PARAESTHESIAS"
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
               Left            =   4080
               TabIndex        =   76
               Top             =   1800
               Width           =   1815
            End
            Begin VB.CheckBox ChkProstatism 
               Caption         =   "PROSTATISM"
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
               Left            =   4080
               TabIndex        =   75
               Top             =   1320
               Width           =   1575
            End
            Begin VB.CheckBox ChkDyspepsia 
               Caption         =   "DYSPEPSIA"
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
               Left            =   4080
               TabIndex        =   74
               Top             =   840
               Width           =   1815
            End
            Begin VB.CheckBox ChkWheeze 
               Caption         =   "WHEEZE"
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
               Left            =   4080
               TabIndex        =   73
               Top             =   360
               Width           =   1215
            End
            Begin VB.CheckBox ChkHeadache 
               Caption         =   "CNS:    HEADACHE"
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
               Left            =   360
               TabIndex        =   72
               Top             =   1800
               Width           =   2415
            End
            Begin VB.CheckBox ChkNocturia 
               Caption         =   "GIT:    NOCTURIA"
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
               Left            =   360
               TabIndex        =   71
               Top             =   1320
               Width           =   2415
            End
            Begin VB.CheckBox ChkConstipation 
               Caption         =   "GIT:    CONSTIPATION"
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
               Left            =   360
               TabIndex        =   70
               Top             =   840
               Width           =   2415
            End
            Begin VB.CheckBox ChkCough 
               Caption         =   "RESPIRATORY:    COUGH"
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
               Left            =   360
               TabIndex        =   69
               Top             =   360
               Width           =   2415
            End
            Begin VB.TextBox TxtMusculoskeletal 
               Height          =   1455
               Left            =   7080
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   68
               Top             =   1800
               Width           =   7455
            End
            Begin VB.Label Label15 
               Caption         =   "MUSCULOSKELETAL:"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7080
               TabIndex        =   67
               Top             =   1560
               Width           =   2295
            End
         End
         Begin VB.TextBox TxtOccupation 
            Height          =   285
            Left            =   7800
            TabIndex        =   65
            Top             =   480
            Width           =   4335
         End
         Begin VB.ComboBox CboMaritalStatus 
            Height          =   315
            ItemData        =   "FrmTreatment.frx":1591C
            Left            =   2880
            List            =   "FrmTreatment.frx":1592C
            TabIndex        =   63
            Top             =   480
            Width           =   2175
         End
         Begin VB.Line Line2 
            X1              =   240
            X2              =   14760
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label14 
            Caption         =   "Occupation"
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
            Left            =   6600
            TabIndex        =   64
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Marital Status"
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
            Left            =   1440
            TabIndex        =   62
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -62400
         TabIndex        =   58
         Top             =   4740
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -67200
         TabIndex        =   56
         Top             =   4740
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -71280
         TabIndex        =   54
         Top             =   4740
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -73560
         TabIndex        =   52
         Top             =   4740
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3735
         Left            =   -74640
         TabIndex        =   43
         Top             =   600
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   6588
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Previous Cardiac Evaluation"
         TabPicture(0)   =   "FrmTreatment.frx":15951
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame5"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Past Medical History"
         TabPicture(1)   =   "FrmTreatment.frx":1596D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Past Surgical History"
         TabPicture(2)   =   "FrmTreatment.frx":15989
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame7"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame7 
            Height          =   3135
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   14775
            Begin VB.TextBox TxtSurgicalHistory 
               Height          =   2775
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               Top             =   240
               Width           =   14535
            End
         End
         Begin VB.Frame Frame6 
            Height          =   3135
            Left            =   -74880
            TabIndex        =   46
            Top             =   360
            Width           =   14775
            Begin VB.TextBox TxtMedicalHistory 
               BackColor       =   &H00FFFFFF&
               Height          =   2775
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   47
               Top             =   240
               Width           =   14535
            End
         End
         Begin VB.Frame Frame5 
            Height          =   3135
            Left            =   -74880
            TabIndex        =   44
            Top             =   360
            Width           =   14775
            Begin VB.TextBox TxtCardiacEvaluation 
               Height          =   2775
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   45
               Top             =   240
               Width           =   14535
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CAD Risk Factors"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4695
         Left            =   -74880
         TabIndex        =   33
         Top             =   420
         Width           =   15495
         Begin VB.ComboBox CboFamilyMember 
            Height          =   315
            ItemData        =   "FrmTreatment.frx":159A5
            Left            =   10320
            List            =   "FrmTreatment.frx":159C1
            TabIndex        =   133
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox TxtFamilyCondition 
            Height          =   885
            Left            =   10320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   131
            Top             =   1080
            Width           =   4695
         End
         Begin VB.TextBox TxtLastSmoked 
            Height          =   285
            Left            =   2760
            TabIndex        =   129
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox ChkSmokedPreviously 
            Caption         =   "SMOKED PREVIOUSLY"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   128
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox ChkSmokingCurrently 
            Caption         =   "SMOKING CURRENTLY"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   42
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox ChkFamilyHistory 
            Caption         =   "FAMILY HISTORY"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8040
            TabIndex        =   41
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox ChkChronicKidney 
            Caption         =   "CHRONIC KIDNEY DISEASES"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5040
            TabIndex        =   40
            Top             =   1380
            Width           =   2535
         End
         Begin VB.CheckBox ChkAlcohol 
            Caption         =   "ALCOHOL"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   39
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox ChkDiabetesMellitus 
            Caption         =   "DIABETES MELLITUS/IFG"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5040
            TabIndex        =   38
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox ChkDyslipidaemia 
            Caption         =   "DYSLIPIDAEMIA"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   37
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox ChkHypertension 
            Caption         =   "HYPERTENSION"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   36
            Top             =   1260
            Width           =   1695
         End
         Begin VB.Frame Frame4 
            Caption         =   "Over the Counter  Medication / Suppliments"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   2655
            Left            =   240
            TabIndex        =   34
            Top             =   1920
            Width           =   15135
            Begin VB.CommandButton Command1 
               Caption         =   "Save Cad Risk Factors"
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
               Left            =   12120
               TabIndex        =   151
               Top             =   1920
               Width           =   2295
            End
            Begin VB.TextBox TxtOverTheCounter 
               BackColor       =   &H00FFFFFF&
               Height          =   2175
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Top             =   360
               Width           =   14775
            End
         End
         Begin VB.Label Label40 
            Caption         =   "Family Condition"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10320
            TabIndex        =   134
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label39 
            Caption         =   "Family Member"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10320
            TabIndex        =   132
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label38 
            Caption         =   "Last Smoked"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   130
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "History of presenting illness"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4695
         Left            =   -74760
         TabIndex        =   31
         Top             =   420
         Width           =   12375
         Begin VB.TextBox TxtHistoryOfIllness 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   360
            Width           =   12135
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Year Done"
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
         Left            =   -62400
         TabIndex        =   57
         Top             =   4500
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Year Done"
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
         Left            =   -67200
         TabIndex        =   55
         Top             =   4500
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71760
         TabIndex        =   53
         Top             =   4740
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Para "
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
         Left            =   -74160
         TabIndex        =   51
         Top             =   4740
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "PAST OBSTETRICS AND GYNAECOLOGICAL HISTORY"
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
         Left            =   -74520
         TabIndex        =   50
         Top             =   4380
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observations"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   12855
      Begin VB.TextBox TxtBloodSugar 
         Height          =   285
         Left            =   11760
         TabIndex        =   138
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtSurname 
         Height          =   285
         Left            =   9720
         TabIndex        =   60
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox TxtSecondName 
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox TxtFirstname 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox TxtBp 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TxtHR 
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtSPO2 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TxtTemperature 
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label45 
         Caption         =   "C"
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
         Left            =   9720
         TabIndex        =   141
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label44 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   140
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "mmHg STANDING"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   139
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label43 
         Caption         =   "BLOOD SUGAR"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   137
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label42 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7350
         TabIndex        =   136
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "b/min"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   135
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8640
         TabIndex        =   59
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Second Name"
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
         Left            =   4560
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name"
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
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "BP   SITTING"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   705
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "RR"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "SPO2"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   24
         Top             =   720
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10320
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label BMI 
         Alignment       =   2  'Center
         Caption         =   "TEMP"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Post Patient"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   9240
      Width           =   13695
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
         Left            =   8160
         TabIndex        =   15
         Top             =   400
         Width           =   2175
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
         Left            =   6000
         TabIndex        =   14
         Top             =   400
         Width           =   2055
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
         Left            =   3240
         TabIndex        =   13
         Top             =   400
         Width           =   2655
      End
      Begin VB.OptionButton OptConsultation 
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
         Left            =   360
         TabIndex        =   12
         Top             =   400
         Width           =   2655
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
         Left            =   11760
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
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
         Height          =   210
         Index           =   5
         Left            =   10440
         TabIndex        =   10
         Top             =   400
         Width           =   1335
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Patient Previous Visits Ordered by Visit Number"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   12855
      Begin VB.CheckBox ChkShowingHistory 
         Caption         =   "Showing History"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   6480
         TabIndex        =   155
         Top             =   0
         Width           =   1935
      End
      Begin VB.CheckBox ChkClearHistory 
         Caption         =   "Show Current Visit Information"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   9240
         TabIndex        =   30
         Top             =   0
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
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
   End
   Begin VB.Frame Frame12 
      Height          =   3855
      Left            =   13080
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton CmdAttachDocs 
         Caption         =   "Attach Documents"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   181
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton CmdConclude 
         Caption         =   "Save && C&onclude Treatment"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   158
         Top             =   3120
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Phase ll Results"
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
         Left            =   240
         TabIndex        =   79
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton CMDSchedule 
         Caption         =   "Schedule Visit"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton CmdMeasurements 
         Caption         =   "View Measurements"
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmTreatment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsfill As New ADODB.Recordset
Dim RsInsert As New ADODB.Recordset
Dim RsGrid As New ADODB.Recordset
Dim RsCombo As New ADODB.Recordset
Dim RsMeasurements As New ADODB.Recordset
Dim RsPrescription As New ADODB.Recordset
Dim ItemSelected As Integer

Private Sub Populate_Tests(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    
    'POPULATE LAB TEST
    TxtLabResults = ""
      RsMeasurements.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '1'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                    TxtLabResults = !TESTDESCRIPTION
                    .MoveNext
            End With
        End If
    RsMeasurements.Close
    
    'POPULATE XRAY
    TxTXray = ""
    RsMeasurements.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '2'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                'While .EOF = False
                    TxTXray = !TESTDESCRIPTION
                    .MoveNext
                'Wend
            End With
        End If
    RsMeasurements.Close
    
    'POPULATE ULTRASOUND
    TxtUltraSound = ""
    RsMeasurements.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '3'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                'While .EOF = False
                    TxtUltraSound = !TESTDESCRIPTION
                    .MoveNext
                'Wend
            End With
        End If
    RsMeasurements.Close
    
    'POPULATE CT SCAN
    TxTCtScan = ""
      RsMeasurements.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '4'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                'While .EOF = False
                    'If Len(!DIAGNOSISDESCRIPTION) = 2 Then KARI = 0 + Right(!DIAGNOSISDESCRIPTION, 1)
                    'KISH = FindRecord("ICD10_CODES", "DISEASEDESCRIPTION", "ICD10CHARACTER = '" & Left(!DIAGNOSISDESCRIPTION, 1) & "' And ICD10NUMBER = '" & Right(!DIAGNOSISDESCRIPTION, 2) & "'")
                    TxTCtScan = !TESTDESCRIPTION
                    .MoveNext
                'Wend
            End With
        End If
    RsMeasurements.Close
    
    'POPULATE MRI
    TxTMRI = ""
      RsMeasurements.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '5'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                'While .EOF = False
                    'If Len(!DIAGNOSISDESCRIPTION) = 2 Then KARI = 0 + Right(!DIAGNOSISDESCRIPTION, 1)
                    'KISH = FindRecord("ICD10_CODES", "DISEASEDESCRIPTION", "ICD10CHARACTER = '" & Left(!DIAGNOSISDESCRIPTION, 1) & "' And ICD10NUMBER = '" & Right(!DIAGNOSISDESCRIPTION, 2) & "'")
                    TxTMRI = !TESTDESCRIPTION
                    .MoveNext
                'Wend
            End With
        End If
    RsMeasurements.Close
    
    'POPULATE ECHO
    TxtEcho = ""
      RsMeasurements.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '6'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                'While .EOF = False
                    'If Len(!DIAGNOSISDESCRIPTION) = 2 Then KARI = 0 + Right(!DIAGNOSISDESCRIPTION, 1)
                    'KISH = FindRecord("ICD10_CODES", "DISEASEDESCRIPTION", "ICD10CHARACTER = '" & Left(!DIAGNOSISDESCRIPTION, 1) & "' And ICD10NUMBER = '" & Right(!DIAGNOSISDESCRIPTION, 2) & "'")
                    TxtEcho = !TESTDESCRIPTION
                    .MoveNext
                'Wend
            End With
        End If
    RsMeasurements.Close
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub PopulateAssesment_and_Plan(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM DOCTOR_ASSESMENT_AND_PLAN WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                TxTAssesment = !ASSESMENT
                TxtPlan = !THEPLAN
            End With
        Else
            TxTAssesment = ""
            TxtPlan = ""
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub Save_Assesment_and_Plan(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_ASSESMENT_AND_PLAN WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_ASSESMENT_AND_PLAN WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'"
            End If
               ' For i = 0 To LstAccute.ListCount - 1
                    .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !ASSESMENT = TxTAssesment
                        !THEPLAN = TxtPlan
                    .Update
                    Beep
                    Beep
               ' Next i
        End With
    RsInsert.Close
    Exit Sub
Errorhandler:
        MsgBox Err.Description
        Exit Sub
       ' Resume
End Sub

Private Sub Save_HistoryOfIllness()
On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM DOCTOR_HISTORY_OF_ILLNESS WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !HISTORYOFILLNESS = TxtHistoryOfIllness
                    .Update
            End With
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub ShowCurrentVisit()
On Error GoTo Errorhandler
        If rsfill.State = 1 Then Set rsfill = Nothing
        rsfill.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CardNumber = COMPLAINS.CardNumber AND COMPLAINS.CardNumber = '" & StrDocCardNo & "' AND COMPLAINS.VisitDate = '" & Format(StrDocVisitDate, "DD MMM YYYY") & "'", Conn, adOpenStatic, adLockOptimistic
            If rsfill.BOF = False And rsfill.EOF = False Then
                    With rsfill
                    TxtFirstName = !FirstName
                    TxtSecondName = !SECONDNAME
                    TxtSurname = !SURNAME
                    TxtBp = !BP
                    TxtBMI = !BMINDEX
                    TxtWeight = !Weight
                    TxtHeight = !Height
                    DtCurrDate = !VisitDate
                    End With
            Else
                'THIS CASE IS HIGHLY UNLICKELY, I WILL NOT EVEN BOTHER WRITTING CODE FOR IT.
            End If
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Public Sub ShowText(OptIndex)
On Error GoTo Errorhandler
    Select Case OptIndex
        Case 1 ' LABTEST
            TxtUltraSound.Visible = False
            TxTXray.Visible = False
            TxTCtScan.Visible = False
            TxTMRI.Visible = False
            TxtLabResults.Visible = True
            TxtEcho.Visible = False
        Case 2 ' XRAY
            TxtUltraSound.Visible = False
            TxTXray.Visible = True
            TxTCtScan.Visible = False
            TxTMRI.Visible = False
            TxtLabResults.Visible = False
            TxtEcho.Visible = False
        Case 3 ' ULTRASOUND
            TxtUltraSound.Visible = True
            TxTXray.Visible = False
            TxTCtScan.Visible = False
            TxTMRI.Visible = False
            TxtLabResults.Visible = False
            TxtEcho.Visible = False
        Case 4 ' CT SCAN
            TxtUltraSound.Visible = False
            TxTXray.Visible = False
            TxTCtScan.Visible = True
            TxTMRI.Visible = False
            TxtLabResults.Visible = False
            TxtEcho.Visible = False
        Case 5 ' MRI
            TxtUltraSound.Visible = False
            TxTXray.Visible = False
            TxTCtScan.Visible = False
            TxTMRI.Visible = True
            TxtLabResults.Visible = False
            TxtEcho.Visible = False
    End Select
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub CboICD10Category_Click()
On Error GoTo Errorhandler
    'POPULATE COMBO FOR DISEASE SUB-CATEGORY
    CboICD10SubCategory.Clear
    LstICD10.Clear
    RsCombo.Open "SELECT DISEASESUBCATEGORY, SUBCATEGORYDESCRIPTION FROM ICD10_SUBCATEGORY WHERE DISEASECATEGORY = '" & GetID_NameFromCombo(CboICD10Category, 1) & "' ORDER BY DISEASESUBCATEGORY", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboICD10SubCategory.AddItem String(3 - Len(!DiseaseSubCategory), "0") & !DiseaseSubCategory & " - " & !SUBCATEGORYDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub CboICD10SubCategory_Click()
    PopulateICD10_CODES GetID_NameFromCombo(CboICD10Category, 1), CboICD10SubCategory.Text
End Sub


Private Sub PopulateICD10_CODES(ByRef DiseaseCategory, ByRef DiseaseSubCategory)
On Error GoTo Errorhandler
    'POPULATE LIST VIEW FOR MAIN CATEGORY
    LstICD10.Clear
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    If RsCombo.State = 1 Then Set RsCombo = Nothing
    RsCombo.Open "SELECT ICD10CHARACTER,ICD10NUMBER,DISEASEDESCRIPTION FROM ICD10_CODES WHERE DISEASECATEGORY = '" & DiseaseCategory & "' AND DISEASESUBCATEGORY = '" & GetID_NameFromCombo(DiseaseSubCategory, 1) & "' ORDER BY ICD10NUMBER", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                LstICD10.AddItem UCase(!ICD10CHARACTER) & String(2 - Len(!ICD10NUMBER), "0") & !ICD10NUMBER & " - " & !DISEASEDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Public Sub ManageProcessFlow(ActiveForm)
On Error GoTo Errorhandler
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
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub CboOther_Click()
    LstOther.AddItem CboOther
End Sub

Private Sub CboRemove_Click()
    If LstSelectedICD10.ListIndex < 0 Then Exit Sub
    LstSelectedICD10.RemoveItem LstSelectedICD10.ListIndex
End Sub


Private Sub CboAccute_Click()
    LstAccute.AddItem CboAccute
End Sub


Private Sub ChkClearHistory_Click()
On Error GoTo Errorhandler
    ClearText FrmTreatment
    ShowCurrentVisit
    'ChkClearHistory.Value = 1
    CmdPost.Enabled = True
    ChkShowingHistory.Caption = "Showing Current"
Exit Sub
Errorhandler:
    MsgBox Err.Description
    End Sub

Private Sub ChkLabTest_Click()
    If ChkLabTest.Value = 1 Then
        FrmLabParameters.Show 1
    End If
    ChkLabTest.Value = 0
End Sub

Private Sub ChkXray_Click()
    FrmTests.Show 1
End Sub

Private Sub ChkViewImage_Click()
    OptNone.Value = True
End Sub

Private Sub CmdAccuteOther_Click()
    Save_AcuteDiagnosis StrDocCardNo, StrDocVisitNumber
    Save_OtherDiagnosis StrDocCardNo, StrDocVisitNumber
    Save_ICD10Clasification StrDocCardNo, StrDocVisitNumber
    Save_Prescription StrDocCardNo, StrDocVisitNumber
    Save_Assesment_and_Plan StrDocCardNo, StrDocVisitNumber
End Sub

Private Sub CmdAttachDocs_Click()
    FrmAttachments.TxtPatientsNames = TxtFirstName & " " & TxtSecondName & " " & TxtSurname
    FrmAttachments.Show
    
End Sub

Private Sub CmdConclude_Click()
On Error GoTo Errorhandler
Dim Resp
    Resp = MsgBox("Please confirm that you wish to conclude treatment for '" & TxtFirstName & "'", vbExclamation + vbYesNo)
        If Resp = vbYes Then
            
        'FIRST OF ALL, SAVE EVERYTHING INCASE THE DOCTOR FORGETS.
        
            UpdateHistoryOfIllness StrDocCardNo, StrDocVisitNumber
            Update_CAD_RiskFactors StrDocCardNo, StrDocVisitNumber
            UpdatePrevious_Medical_History StrDocCardNo, StrDocVisitNumber
            Update_Social_History StrDocCardNo, StrDocVisitNumber
            Update_Physical_Examination StrDocCardNo, StrDocVisitNumber
            
            Save_DoctorTests StrDocCardNo, StrDocVisitNumber
            Save_AcuteDiagnosis StrDocCardNo, StrDocVisitNumber
            Save_OtherDiagnosis StrDocCardNo, StrDocVisitNumber
            Save_ICD10Clasification StrDocCardNo, StrDocVisitNumber
            Save_Prescription StrDocCardNo, StrDocVisitNumber
            Save_Assesment_and_Plan StrDocCardNo, StrDocVisitNumber
        
        'THEN NOW CONCLUDE TREATMENT
        
           Conn.Execute "UPDATE COMPLAINS SET TODOCTORS = 'false', INUSE = 0,TOOBSERVATION = 0, DISMISSED = 'TRUE', doctor = '" & GlbCurrentUser & "' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
           ClearText FrmTreatment
           Grid.Clear: Grid.Rows = 1
        End If
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub CmdMeasurements_Click()
    BlnViewMeasurements = True
    FrmObservation.Show 1
End Sub

Private Sub CmdPost_Click()
On Error GoTo Errorhandler
    If OptConsultation.Item(1).Value = True Then GoTo PostOnly ' do not save .. return to Consultation.
    UpdateHistoryOfIllness StrDocCardNo, StrDocVisitNumber
    Update_CAD_RiskFactors StrDocCardNo, StrDocVisitNumber
    UpdatePrevious_Medical_History StrDocCardNo, StrDocVisitNumber
    Update_Social_History StrDocCardNo, StrDocVisitNumber
    Update_Physical_Examination StrDocCardNo, StrDocVisitNumber
    Save_AcuteDiagnosis StrDocCardNo, StrDocVisitNumber
    Save_OtherDiagnosis StrDocCardNo, StrDocVisitNumber
    Save_ICD10Clasification StrDocCardNo, StrDocVisitNumber
    Save_DoctorTests StrDocCardNo, StrDocVisitNumber
    
    MsgBox "Doctors details saved Succesfully", vbInformation
    
    Conn.Execute "UPDATE COMPLAINS set INUSE = 0,  DOCTOR = '" & GlbCurrentUser & "' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
    
PostOnly:
    
         Select Case ItemSelected
        Case 1
                 DUMMY = SendPatient(EnumConsultation, StrDocCardNo, GlbSysDate)
                 'FrmPatients.Show
                 'Unload Me
         Case 2
             'To Doctor
                 DUMMY = SendPatient(EnumDoctors, StrDocCardNo, GlbSysDate)
                 If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'NurseDoctorRolesCombined'") = 1 Then
                     FrmWaitingRoom.Show
                 End If
                 Unload Me
        Case 3
             'To Cashier
                 DUMMY = SendPatient(EnumCashier, StrDocCardNo, GlbSysDate)
                 'FrmCashier.Show
                 'Unload Me
         Case 4
             'To Pharmacy
                 DUMMY = SendPatient(EnumPharmacy, StrDocCardNo, GlbSysDate)
                 'FrmPharmacy.Show
                 'Unload Me
         End Select
Exit Sub
Errorhandler:
    MsgBox Err.Description
    End Sub

Private Sub CmdRemoveAccute_Click()
    LstAccute.RemoveItem LstAccute.ListIndex
End Sub

Private Sub CmdRemoveOther_Click()
    LstOther.RemoveItem LstOther.ListIndex
End Sub

Private Sub CmdSaveAssesmentPlan_Click()

'    UpdateHistoryOfIllness StrDocCardNo, StrDocVisitNumber
'    Update_CAD_RiskFactors StrDocCardNo, StrDocVisitNumber
'    UpdatePrevious_Medical_History StrDocCardNo, StrDocVisitNumber
'    Update_Social_History StrDocCardNo, StrDocVisitNumber
'    Update_Physical_Examination StrDocCardNo, StrDocVisitNumber
'
'
'    Save_AcuteDiagnosis StrDocCardNo, StrDocVisitNumber
'    Save_OtherDiagnosis StrDocCardNo, StrDocVisitNumber
'    Save_ICD10Clasification StrDocCardNo, StrDocVisitNumber
'    Save_Prescription StrDocCardNo, StrDocVisitNumber
    Save_Assesment_and_Plan StrDocCardNo, StrDocVisitNumber
    
'    Save_DoctorTests StrDocCardNo, StrDocVisitNumber
End Sub

Private Sub CmdSaveDiagnosis_Click()
    Save_AcuteDiagnosis StrDocCardNo, StrDocVisitNumber
    Save_OtherDiagnosis StrDocCardNo, StrDocVisitNumber
    Save_ICD10Clasification StrDocCardNo, StrDocVisitNumber
    Save_Prescription StrDocCardNo, StrDocVisitNumber
    Save_Assesment_and_Plan StrDocCardNo, StrDocVisitNumber
End Sub

Private Sub CmdSavePrescription_Click()
    Save_AcuteDiagnosis StrDocCardNo, StrDocVisitNumber
    Save_OtherDiagnosis StrDocCardNo, StrDocVisitNumber
    Save_ICD10Clasification StrDocCardNo, StrDocVisitNumber
    Save_Prescription StrDocCardNo, StrDocVisitNumber
    Save_Assesment_and_Plan StrDocCardNo, StrDocVisitNumber
End Sub
Private Sub Save_Prescription(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    
    Conn.Execute "DELETE FROM PRESCRIPTION_TEXT WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'"
    
    Conn.Execute "INSERT INTO PRESCRIPTION_TEXT (CARDNUMBER,VISITNUMBER,PRESCRIPTIONTEXT)VALUES('" & StrDocCardNo & "','" & StrDocVisitNumber & "','" & TxtPrescription & "')"
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub CMDSchedule_Click()
    FrmScheduleVisit.Show 1
End Sub

Private Sub Save_DoctorTests(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    
    'INSERT LAB TEST REQUESTS
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and TESTTYPE = '1'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '1'"
            End If
                'For i = 0 To LstLabTests.ListCount - 1
                    .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !TESTTYPE = 1      'Number 1 Assigned to Lab Test. 2=Xray, 3=UltrSound, 4=CTScan, 5=MRI
                        !TESTDESCRIPTION = TxtLabResults
                    .Update
                'Next i
        End With
    RsInsert.Close
    
    'INSERT X-RAY TEST RESULTS
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and TESTTYPE = '2'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '2'"
            End If
                   .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !TESTTYPE = 2      'Number 1 Assigned to Lab Test. 2=Xray, 3=UltrSound, 4=CTScan, 5=MRI
                        !TESTDESCRIPTION = TxTXray
                    .Update
        End With
    RsInsert.Close
    
    'INSERT UNTRASOUND TEST RESULTS
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and TESTTYPE = '3'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '3'"
            End If
                   .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !TESTTYPE = 3      'Number 1 Assigned to Lab Test. 2=Xray, 3=UltrSound, 4=CTScan, 5=MRI
                        !TESTDESCRIPTION = TxtUltraSound
                    .Update
        End With
    RsInsert.Close
    
    'INSERT CT-SCAN TEST RESULTS
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and TESTTYPE = '4'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '4'"
            End If
                   .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !TESTTYPE = 4      'Number 1 Assigned to Lab Test. 2=Xray, 3=UltrSound, 4=CTScan, 5=MRI
                        !TESTDESCRIPTION = TxTCtScan
                    .Update
        End With
    RsInsert.Close
    
    'INSERT MRI TEST RESULTS
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and TESTTYPE = '5'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '5'"
            End If
                   .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !TESTTYPE = 5      'Number 1 Assigned to Lab Test. 2=Xray, 3=UltrSound, 4=CTScan, 5=MRI
                        !TESTDESCRIPTION = TxTMRI
                    .Update
        End With
    RsInsert.Close
    
    'INSERT ECHO TEST RESULTS
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and TESTTYPE = '6'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '6'"
            End If
                   .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !TESTTYPE = 6      'Number 1 Assigned to Lab Test. 2=Xray, 3=UltrSound, 4=CTScan, 5=MRI
                        !TESTDESCRIPTION = TxTMRI
                    .Update
        End With
    RsInsert.Close
    
''''    'INSERT LAB TEST RESULTS
''''    If RsInsert.State = 1 Then Set RsInsert = Nothing
''''    RsInsert.Open "SELECT * FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and TESTTYPE = ''", Conn, adOpenStatic, adLockOptimistic
''''        With RsInsert
''''            If .EOF = False Then
''''                Conn.Execute "DELETE FROM DOCTOR_TESTS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND TESTTYPE = '5'"
''''            End If
''''                   .AddNew
''''                        !CardNumber = CardNo
''''                        !VISITNUMBER = VisitNo
''''                        !TESTTYPE = 6      'Number 1 Assigned to Lab Test. 2=Xray, 3=UltrSound, 4=CTScan, 5=MRI
''''                        !TESTDESCRIPTION = TxtLabResults
''''                    .Update
''''        End With
''''    RsInsert.Close
    Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub

Private Sub Command3_Click()
    Save_DoctorTests StrDocCardNo, StrDocVisitNumber
End Sub

Private Sub Command5_Click()
    UpdatePrevious_Medical_History StrDocCardNo, StrDocVisitNumber
End Sub

Private Sub CmdExit_Click()
On Error GoTo Errorhandler
    If TxtFirstName <> "" And TxtBMI <> "" And TxtHeight <> "" Then
        'Release document to be accessed by other system users
        Conn.Execute "UPDATE COMPLAINS SET INUSE = 'False' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITDATE = '" & Format(GlbSysDate, "DD MMM YYYY") & "'"
    End If
    Unload Me
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub Command1_Click()
    Update_CAD_RiskFactors StrDocCardNo, StrDocVisitNumber
End Sub

Private Sub Command6_Click()
    Update_Social_History StrDocCardNo, StrDocVisitNumber
End Sub

Private Sub Command7_Click()
    Update_Physical_Examination StrDocCardNo, StrDocVisitNumber
End Sub


Private Sub Form_Load()
On Error GoTo Errorhandler
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    
        'FILL DETAILS FROM THE TWO VARIABLES.
    Select Case BlnHISTORY
    
        Case False
            If rsfill.State = 1 Then Set rsfill = Nothing
            rsfill.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CardNumber = COMPLAINS.CardNumber AND COMPLAINS.CardNumber = '" & StrDocCardNo & "' AND COMPLAINS.VisitDate = '" & Format(StrDocVisitDate, "DD MMM YYYY") & "'", Conn, adOpenStatic, adLockOptimistic
                If rsfill.BOF = False And rsfill.EOF = False Then
                    With rsfill
                    TxtFirstName = !FirstName
                    TxtSecondName = !SECONDNAME
                    TxtSurname = !SURNAME
                    TxtBp = !BP
                    TxtBMI = !BMINDEX
                    TxtWeight = !Weight
                    TxtHeight = !Height
                    DtCurrDate = !VisitDate
                    End With
                Else
                    'THIS CASE IS HIGHLY UNLICKELY, I WILL NOT EVEN BOTHER WRITTING CODE FOR IT.
                End If
                
        'FETCH A FEW VITAL READINGS AND POPULATE ON DOCTORS SCREEN. SG 20072023
        If rsfill.State = 1 Then Set rsfill = Nothing
        rsfill.Open "SELECT * FROM OBSERVATION_CLINIC_REVIEW WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With rsfill
                If .EOF = False Then
                    TxtBp = !BP
                    TxtHR = !RR
                    TxtSPO2 = !SPO2
                    TxtTemperature = !Temperature
                    TxtBloodSugar = !SUGAR
                End If
            End With
        rsfill.Close
        
        Case Else
            If rsfill.State = 1 Then Set rsfill = Nothing
            rsfill.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CardNumber = COMPLAINS.CardNumber AND COMPLAINS.CardNumber = '" & StrDocCardNo & "' AND COMPLAINS.VisitDate = '" & Format(StrDocVisitDate, "DDMMMYYYY") & "'", Conn, adOpenStatic, adLockOptimistic
                If rsfill.BOF = False And rsfill.EOF = False Then
                    With rsfill
                        TxtFirstName = !SURNAME & "  " & !FirstName
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
                                        'LstPrescription.AddItem !CODE & " - " & !Description
                                        .MoveNext
                                    Wend
                                End With
                            RsPrescription.Close
                        'END
                        
                        TxtReferal = !REFERRAL & vbNullString
                        'TxtAdmission = !ADMISSION & vbNullString
                        BlnHISTORY = False & vbNullString
                    End With
                Else
                    'THIS CASE IS HIGHLY UNLICKELY, I WILL NOT EVEN BOTHER WRITTING CODE FOR IT.
                End If
        End Select
        
    'POPULATE COMBO FOR DISEASE CATEGORY
    RsCombo.Open "SELECT CATEGORYID, CATEGORYDESCRIPTION FROM ICD10_CATEGORY ORDER BY CATEGORYID ASC", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboICD10Category.AddItem String(3 - Len(!CATEGORYID), "0") & !CATEGORYID & " - " & !CATEGORYDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close
    
    'LOAD TESTS TO SCAN LIST
    If RsCombo.State = 1 Then Set RsCombo = Nothing
    RsCombo.Open "SELECT ID, DESCRIPTION FROM ListOFScanTests ORDER BY ID ASC", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                LstTests.AddItem String(3 - Len(!ID), "0") & !ID & " - " & !Description
                .MoveNext
            Wend
        End With
    RsCombo.Close
    
    
    FillHistory
    centerform Me
    
    'ASSIGN FORM NAME AS CURRENT FORM
    GlbCurrentForm = EnumDoctors
    ManageProcessFlow EnumDoctors
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub PopulateHistoryOfPresentedIllnessXXX(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM OBSERVATION_MEDICAL_HISTORY WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                ChkDiabetes = !DIABETES
                ChkHypertension = !HYPERTENSION
                ChkHighCholestrol = !HIGHCHOLESTROL
                ChkGout = !GOUT
                ChkAsthma = !ASTHMA
                ChkCancer = !CANCER
                ChkPheumoniaVaccine = !PHEUMONIAVACCINE
                ChkHepatitisBvaccine = !HEPATITISBVACCINE
                ChkPeniciline = !PENICILINE
                ChkSulphurDrugs = !SULPHURDRUGS
                TxtOtherAllergies = !OTHERS
            End With
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub PopulateCAD_RiskFactors(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM DOCTOR_CAD_RISK_FACTORS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                'MsgBox "Start Retrieval"
                If !SMOKINGCURRENTLY <> "" Then ChkSmokingCurrently = !SMOKINGCURRENTLY
                If !SMOKEDPREVIOUSLY <> "" Then ChkSmokedPreviously = !SMOKEDPREVIOUSLY
                If !LASTSMOKED <> "" Then TxtLastSmoked.Text = !LASTSMOKED
                If !HYPERTENSION <> "" Then ChkHypertension = !HYPERTENSION
                If !DYSLIPIDAEMIA <> "" Then ChkDyslipidaemia = !DYSLIPIDAEMIA
                'MsgBox "Mid Retrieval"
                If !DIABETESMELLITUS <> "" Then ChkDiabetesMellitus = !DIABETESMELLITUS
                If !CHRONICKIDNEYDISEASE <> "" Then ChkChronicKidney = !CHRONICKIDNEYDISEASE
                If !ALCOHOL <> "" Then ChkAlcohol = !ALCOHOL
                If !FAMILYMEMBER <> "" Then CboFamilyMember = !FAMILYMEMBER
                If !FAMILYCONDITION <> "" Then TxtFamilyCondition = !FAMILYCONDITION
                'MsgBox "Second Last"
                If !OVERTHECOUNTERMEDICATION = "" Then
                    'DO NOTHING
                ElseIf Not IsNull(!OVERTHECOUNTERMEDICATION) Then
                    TxtOverTheCounter = !OVERTHECOUNTERMEDICATION
                End If
                'MsgBox "End"
            End With
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub PopulatePrevious_MedicalHistory(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    If RsMeasurements.State = 1 Then Set RsMeasurements = Nothing
    RsMeasurements.Open "SELECT * FROM DOCTOR_PAST_MEDICAL_HISTORY WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                TxtCardiacEvaluation = !PREVIOUSCARDIACEVALUATION
                TxtMedicalHistory = !PASTMEDICALHISTORY
                TxtSurgicalHistory = !PASTSURGICALHISTORY
            End With
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub PopulateSocial_History(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM DOCTOR_SOCIAL_HISTORY WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                CboMaritalStatus = !MARITALSTATUS
                TxtOccupation = !OCCUPATION
                'MsgBox "Check Boxes Begin"
                ChkCough = !COUGH
                ChkConstipation = !CONSTIPATION
                ChkNocturia = !NOCTURIA
                ChkHeadache = !HEADACHE
                ChkWheeze = !WHEEZE
                ChkDyspepsia = !DYSPEPSIA
                ChkSnorring = !SNORRING
                ChkPreviousColonoscopy = !PREVIOUSCOLONOSCOPY
                TxtMusculoskeletal = !MUSCuLOSKELETAL
                'MsgBox "Check Boxes End"
            End With
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub PopulatePhysical_Examination(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM DOCTOR_PHYSICAL_EXAMINATION WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                'MsgBox "Physical Begin"
                If !HEAD <> "" Then TxtHead = !HEAD
                If OROPHARYNX <> "" Then TxtOropharynx = !OROPHARYNX
                If !JVP <> "" Then ChkJVP = !JVP
                If !BRUITS <> "" Then ChkBruits = !BRUITS
                If !THYROMEGALLY <> "" Then ChkThyromegally = !THYROMEGALLY
                If !ADENOPATHY <> "" Then TxtAdenopathy = !ADENOPATHY
                If !LEFTFA <> "" Then TxtLeftFA = !LEFTFA
                If !LEFTPA <> "" Then TxtLeftPA = !LEFTPA
                If !LEFTDP <> "" Then TxtLeftDP = !LEFTDP
                If !LEFTPT <> "" Then TxtLeftPT = !LEFTPT
                'MsgBox "Physical Mid"
                If !RIGHTFA <> "" Then TxtRightFA = !RIGHTFA
                If !RIGHTPA <> "" Then TxtRightPA = !RIGHTPA
                If !RIGHTDP <> "" Then TxtRightDP = !RIGHTDP
                If !RIGHTPT <> "" Then TxtRightPT = !RIGHTPT
                If !OTHERS <> "" Then TxtPEOthers = !OTHERS
                If !APICALIMPULSE <> "" Then TxtApicalImpulse = !APICALIMPULSE
                If !S1S2 <> "" Then TxTSOneSTwo = !S1S2
                If !S3S4 <> "" Then TxtSThreeSfour = !S3S4
                'MsgBox "Pre Murmars"
                If !MURMURS <> "" Then TxtMurmurs = !MURMURS
                If !RESPIRATORY <> "" Then TxtRespiratory = !RESPIRATORY
                If !CNS <> "" Then TxtCNS = !CNS
                If !ABDOMEN <> "" Then TxtAbdomen = !ABDOMEN
                If !GUT <> "" Then TxtGut = !GUT
                'MsgBox "Physical End"
            End With
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub Populate_Diagnosis(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    'POPULATE ACCUTE DIAGNOSIS
    'LstAccute.Clear
    'MsgBox "Diagnosis Begin"
    RsMeasurements.Open "SELECT * FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND DIAGNOSISTYPE = '1'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                While .EOF = False
                    TxtAccute = !DIAGNOSISDESCRIPTION
                    .MoveNext
                Wend
            End With
        End If
    RsMeasurements.Close
    
    'POPULATE OTHER DIAGNOSIS
    'LstOther.Clear
    RsMeasurements.Open "SELECT * FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND DIAGNOSISTYPE = '2'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                While .EOF = False
                    TxtOther = !DIAGNOSISDESCRIPTION
                    .MoveNext
                Wend
            End With
        End If
    RsMeasurements.Close
    
    'POPULATE ICD10 CLASSIFICATION
    LstSelectedICD10.Clear
      RsMeasurements.Open "SELECT * FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND DIAGNOSISTYPE = '3'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                While .EOF = False
                    'If Len(!DIAGNOSISDESCRIPTION) = 2 Then KARI = 0 + Right(!DIAGNOSISDESCRIPTION, 1)
                    KISH = FindRecord("ICD10_CODES", "DISEASEDESCRIPTION", "ICD10CHARACTER = '" & Left(!DIAGNOSISDESCRIPTION, 1) & "' And ICD10NUMBER = '" & Right(!DIAGNOSISDESCRIPTION, 2) & "'")
                    LstSelectedICD10.AddItem !DIAGNOSISDESCRIPTION & " - " & KISH
                    .MoveNext
                Wend
            End With
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
    
End Sub
Private Sub PopulateHistoryOfPresentedIllness(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM DOCTOR_HISTORY_OF_ILLNESS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                TxtHistoryOfIllness = !HISTORYOFILLNESS
            End With
        Else
            TxtHistoryOfIllness = ""
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub PopulatePrescription(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM PRESCRIPTION_TEXT WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                TxtPrescription = !PRESCRIPTIONTEXT
            End With
        Else
            TxtPrescription = ""
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Public Sub UpdateHistoryOfIllness(ByRef CardNo, ByRef VisitNo)
On Error GoTo Errorhandler
    'If lvObservationCardNumber = "" Then Exit Sub
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM DOCTOR_HISTORY_OF_ILLNESS WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !HISTORYOFILLNESS = TxtHistoryOfIllness
                    .Update
                    Beep
            End With
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Public Sub Update_CAD_RiskFactors(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM DOCTOR_CAD_RISK_FACTORS WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !SMOKINGCURRENTLY = ChkSmokingCurrently
                        !SMOKEDPREVIOUSLY = ChkSmokedPreviously
                        !LASTSMOKED = TxtLastSmoked
                        !HYPERTENSION = ChkHypertension
                        !DYSLIPIDAEMIA = ChkDyslipidaemia
                        !DIABETESMELLITUS = ChkDiabetesMellitus
                        !CHRONICKIDNEYDISEASE = ChkChronicKidney
                        !ALCOHOL = ChkAlcohol
                        !FAMILYMEMBER = CboFamilyMember
                        !FAMILYCONDITION = TxtFamilyCondition
                        !OVERTHECOUNTERMEDICATION = TxtOverTheCounter
                    .Update
                    Beep
            End With
Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub
Public Sub UpdatePrevious_Medical_History(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM DOCTOR_PAST_MEDICAL_HISTORY WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !PREVIOUSCARDIACEVALUATION = TxtCardiacEvaluation
                        !PASTMEDICALHISTORY = TxtMedicalHistory
                        !PASTSURGICALHISTORY = TxtSurgicalHistory
                    .Update
                    Beep
            End With
Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub
Public Sub Update_Social_History(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM DOCTOR_SOCIAL_HISTORY WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !MARITALSTATUS = CboMaritalStatus
                        !OCCUPATION = TxtOccupation
                        !COUGH = ChkCough
                        !CONSTIPATION = ChkConstipation
                        !NOCTURIA = ChkNocturia
                        !HEADACHE = ChkHeadache
                        !WHEEZE = ChkWheeze
                        !DYSPEPSIA = ChkDyspepsia
                        !SNORRING = ChkSnorring
                        !PREVIOUSCOLONOSCOPY = ChkPreviousColonoscopy
                        !MUSCuLOSKELETAL = TxtMusculoskeletal
                    .Update
                    Beep
            End With
Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub
Public Sub Update_Physical_Examination(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM DOCTOR_PHYSICAL_EXAMINATION WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !HEAD = TxtHead
                        !OROPHARYNX = TxtOropharynx
                        !JVP = ChkJVP
                        !BRUITS = ChkBruits
                        !THYROMEGALLY = ChkThyromegally
                        !ADENOPATHY = TxtAdenopathy
                        If TxtLeftFA = "" Then TxtLeftFA = 0 Else !LEFTFA = TxtLeftFA
                        If TxtLeftPA = "" Then TxtLeftPA = 0 Else !LEFTPA = TxtLeftPA
                        If TxtLeftDP = "" Then TxtLeftDP = 0 Else !LEFTDP = TxtLeftDP
                        If TxtLeftPT = "" Then TxtLeftPT = 0 Else !LEFTPT = TxtLeftPT
                        If TxtRightFA = "" Then TxtRightFA = 0 Else !RIGHTFA = TxtRightFA
                        If TxtRightPA = "" Then TxtRightPA = 0 Else !RIGHTPA = TxtRightPA
                        If TxtRightDP = "" Then TxtRightDP = 0 Else !RIGHTDP = TxtRightDP
                        If TxtRightPT = "" Then TxtRightPT = 0 Else !RIGHTPT = TxtRightPT
                        If TxtPEOthers = "" Then TxtPEOthers = 0 Else !OTHERS = TxtPEOthers
                        If TxtApicalImpulse = "" Then TxtApicalImpulse = 0 Else !APICALIMPULSE = TxtApicalImpulse
                        If TxTSOneSTwo = "" Then TxTSOneSTwo = 0 Else !S1S2 = TxTSOneSTwo
                        If TxtSThreeSfour = "" Then TxtSThreeSfour = 0 Else !S3S4 = TxtSThreeSfour
                        If TxtMurmurs = "" Then TxtMurmurs = 0 Else !MURMURS = TxtMurmurs
                        If TxtRespiratory = "" Then TxtRespiratory = 0 Else !RESPIRATORY = TxtRespiratory
                        If TxtCNS = "" Then TxtCNS = 0 Else !CNS = TxtCNS
                        If TxtAbdomen = "" Then TxtAbdomen = 0 Else !ABDOMEN = TxtAbdomen
                        If TxtGut = "" Then TxtGut = 0 Else !GUT = TxtGut
                    .Update
                        Beep
                        Beep
                        'MsgBox "Physical Examination saved succesfully", vbInformation
            End With
Exit Sub
Errorhandler:
        MsgBox Err.Description
        Exit Sub
        Resume
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If TxtFirstName <> "" And TxtBMI <> "" And TxtHeight <> "" Then
        'Release document to be accessed by other system users
        Conn.Execute "UPDATE COMPLAINS SET INUSE = 'False' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITDATE = '" & Format(GlbSysDate, "DD MMM YYYY") & "'"
    End If
End Sub


Private Sub Grid_Click()
On Error GoTo Errorhandler
    Dim lvViewCardNo As String
    Dim lvViewVisitNo As String
    
    lvViewCardNo = Grid.TextMatrix(Grid.Row, 0)
    lvViewVisitNo = Grid.TextMatrix(Grid.Row, 1)
    
    'MsgBox "pass 1"
    'VISIT NUMBER FOR SAVING IMAGES
    GlbImageVisitNumber = Grid.TextMatrix(Grid.Row, 1)
    'MsgBox "pass 2"
    
    ClearText FrmTreatment
    ShowCurrentVisit
    ChkShowingHistory.Value = 1
    ChkShowingHistory.Caption = "Showing History"
    
    'MsgBox "Pass 3"
    Populate_Bio_Vital_Details lvViewCardNo, lvViewVisitNo
    PopulateHistoryOfPresentedIllness lvViewCardNo, lvViewVisitNo
    PopulateCAD_RiskFactors lvViewCardNo, lvViewVisitNo
    PopulatePrevious_MedicalHistory lvViewCardNo, lvViewVisitNo
    PopulateSocial_History lvViewCardNo, lvViewVisitNo
    PopulatePhysical_Examination lvViewCardNo, lvViewVisitNo
    Populate_Diagnosis lvViewCardNo, lvViewVisitNo
    Populate_Tests lvViewCardNo, lvViewVisitNo
    PopulatePrescription lvViewCardNo, lvViewVisitNo
    PopulateAssesment_and_Plan lvViewCardNo, lvViewVisitNo
    
    ChkClearHistory.Value = 0
    CmdPost.Enabled = False
    Exit Sub
Errorhandler:
    MsgBox Err.Description
    Exit Sub
    Resume
End Sub
Private Sub Populate_Bio_Vital_Details(lvCardNumber, lvVisitNumber)
            If rsfill.State = 1 Then Set rsfill = Nothing
            rsfill.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CardNumber = COMPLAINS.CardNumber AND COMPLAINS.CardNumber = '" & lvCardNumber & "' AND COMPLAINS.Visitnumber = '" & lvVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
                If rsfill.BOF = False And rsfill.EOF = False Then
                    With rsfill
                    TxtFirstName = !FirstName
                    TxtSecondName = !SECONDNAME
                    TxtSurname = !SURNAME
                    TxtBp = !BP
                    TxtBMI = !BMINDEX
                    TxtWeight = !Weight
                    TxtHeight = !Height
                    
                    DtCurrDate = !VisitDate
                    End With
                Else
                    'THIS CASE IS HIGHLY UNLICKELY, I WILL NOT EVEN BOTHER WRITTING CODE FOR IT.
                End If
            rsfill.Close
            
        'FETCH A FEW VITAL READINGS AND POPULATE ON DOCTORS SCREEN. SG 20072023
        If rsfill.State = 1 Then Set rsfill = Nothing
        rsfill.Open "SELECT * FROM OBSERVATION_CLINIC_REVIEW WHERE CARDNUMBER = '" & lvCardNumber & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With rsfill
                If .EOF = False Then
                    TxtBp = !BP
                    TxtHR = !RR
                    TxtSPO2 = !SPO2
                    TxtTemperature = !Temperature
                    TxtBloodSugar = !SUGAR
                End If
            End With
        rsfill.Close
End Sub
Private Sub LstICD10_Click()
 LstSelectedICD10.AddItem LstICD10  'GetID_NameFromCombo(LstLabTests, 2)
End Sub

Private Sub TxtLabRequest_DblClick()
    FrmLabParameters.Show 1
End Sub




Private Sub LstTests_DblClick()
On Error GoTo Errorhandler
    GlbTestImageType = CDbl(GetID_NameFromCombo(LstTests.Text, 1))
    If ChkViewImage.Value = 1 Then
        FrmTests.Show 1
    End If
    ShowText GlbTestImageType
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Data Vault System Ltd"
       Exit Sub
    Resume
End Sub


Private Sub OptCashier_Click(Index As Integer)
    ItemSelected = OptCashier.Item(Index).Index
End Sub

Private Sub OptConsultation_Click(Index As Integer)
    ItemSelected = OptConsultation.Item(Index).Index
End Sub

Private Sub OptCTScanImage_Click(Index As Integer)
On Error GoTo Errorhandler
    GlbTestImageType = 4
    If ChkViewImage.Value = 1 Then
        FrmTests.Show 1
    End If
    ShowText 4
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Data Vault System Ltd"
       Exit Sub
    Resume
End Sub

Private Sub OptEcho_Click()
    GlbTestImageType = 6
    If ChkViewImage.Value = 1 Then
        FrmTests.Show 1
    End If
    ShowText 6
End Sub

'''Private Sub OptCTScanImage()
'''    GlbTestImageType = 3
'''    If ChkViewImage.Value = 1 Then
'''        FrmTests.Show 1
'''    End If
'''End Sub
'''
Private Sub OptLab_Click(Index As Integer)
    ItemSelected = OptLab.Item(Index).Index
End Sub

Private Sub OptLabResults_Click()
    GlbTestImageType = 1
    If ChkViewImage.Value = 1 Then
        FrmTests.Show 1
    End If
    ShowText 1
End Sub

Private Sub OptMRIImage_Click()
    GlbTestImageType = 5
    If ChkViewImage.Value = 1 Then
        FrmTests.Show 1
    End If
    ShowText 5
End Sub

Private Sub OptObservation_Click(Index As Integer)
    ItemSelected = OptObservation.Item(Index).Index
End Sub

Private Sub OptPharmacy_Click(Index As Integer)
    ItemSelected = OptPharmacy.Item(Index).Index
End Sub

Private Sub OptUltrasountImage_Click()
On Error GoTo Errorhandler
    GlbTestImageType = 3
    If ChkViewImage.Value = 1 Then
        FrmTests.Show 1
    End If
    ShowText 3
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Data Vault System Ltd"
       Exit Sub
    Resume
End Sub

Private Sub OptXRayImage_Click()
    GlbTestImageType = 2
    If ChkViewImage.Value = 1 Then
        FrmTests.Show 1
    End If
    ShowText 2
End Sub

Private Sub TxtSave_Click()
    UpdateHistoryOfIllness StrDocCardNo, StrDocVisitNumber
End Sub
Private Sub FillHistory()
    On Error GoTo Errorhandler
   KARI = GlbSysDate
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 7
    Grid.ColAlignment(1) = flexAlignCenterCenter
    'Grid.ColDataType(7) = flexDTBoolean
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |   VISIT DATE |   DOCTOR  ."
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.CARDNUMBER = '" & StrDocCardNo & "' ORDER BY VISITNUMBER DESC", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        BILLINGNAME = FindRecord("SERVICE_PROVIDER", "SERVICEPROVIDER", "COMPANYCODE = '" & !BILLINGCOMPANY & "'")
                        Grid.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY + " - " + BILLINGNAME & vbTab & !IDNUMBER & vbTab & !VisitDate & vbTab & !DOCTOR
                        .MoveNext
                    Wend
                End With
            End If
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Private Sub Save_AcuteDiagnosis(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND DIAGNOSISTYPE = '1'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND DIAGNOSISTYPE = '1'"
            End If
               ' For i = 0 To LstAccute.ListCount - 1
                    .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !DIAGNOSISTYPE = 1      'Number 1 Assigned to Accute Diagnosis. 2=other, 3=ICD10 Classifications
                        !DIAGNOSISDESCRIPTION = TxtAccute
                    .Update
                    Beep
               ' Next i
        End With
    RsInsert.Close
    Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub
Private Sub Save_OtherDiagnosis(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and DIAGNOSISTYPE = '2'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND DIAGNOSISTYPE = '2'"
            End If
                'For i = 0 To LstOther.ListCount - 1
                    .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !DIAGNOSISTYPE = 2      'Number 1 Assigned to Accute Diagnosis. 2=other, 3=ICD10 Classifications
                        !DIAGNOSISDESCRIPTION = TxtOther
                    .Update
               ' Next i
        End With
    RsInsert.Close
    Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub
Private Sub Save_ICD10Clasification(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If RsInsert.State = 1 Then Set RsInsert = Nothing
    RsInsert.Open "SELECT * FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' and DIAGNOSISTYPE = '3'", Conn, adOpenStatic, adLockOptimistic
        With RsInsert
            If .EOF = False Then
                Conn.Execute "DELETE FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "' AND DIAGNOSISTYPE = '3'"
            End If
                For i = 0 To LstSelectedICD10.ListCount - 1
                    .AddNew
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !DIAGNOSISTYPE = 3      'Number 1 Assigned to Accute Diagnosis. 2=other, 3=ICD10 Classifications
                        !DIAGNOSISDESCRIPTION = GetID_NameFromCombo(LstSelectedICD10.List(i), 1)
                    .Update
                    Beep
                Next i
        End With
    RsInsert.Close
    Exit Sub
Errorhandler:
        MsgBox Err.Description
        'Resume
End Sub

