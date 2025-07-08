VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmProducts 
   Caption         =   " "
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12450
   Icon            =   "FrmProducts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Navigate"
      Height          =   6855
      Left            =   10800
      TabIndex        =   14
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton CmdEditStock 
         Caption         =   "EDIT STOCK COUNT"
         Height          =   735
         Left            =   120
         TabIndex        =   55
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "New"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton CMdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab TBStock 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Stock Inventory Update"
      TabPicture(0)   =   "FrmProducts.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pharmacy Direct Sales"
      TabPicture(1)   =   "FrmProducts.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Stock Entry Summary"
         Height          =   6135
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   10215
         Begin VB.ComboBox CboCategory 
            Height          =   315
            Left            =   4440
            TabIndex        =   13
            Top             =   600
            Width           =   5535
         End
         Begin VB.TextBox TxtProductCount 
            Height          =   375
            Left            =   4440
            TabIndex        =   12
            Top             =   2640
            Width           =   5535
         End
         Begin VB.ComboBox CboProduct 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            TabIndex        =   11
            Top             =   1260
            Width           =   5535
         End
         Begin VB.TextBox TxtDeliverdBy 
            Height          =   375
            Left            =   4440
            TabIndex        =   10
            Top             =   4080
            Width           =   5535
         End
         Begin VB.Label LblLastDeliveryBy 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4440
            TabIndex        =   21
            Top             =   3360
            Width           =   5535
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Last Delivery By:"
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
            Left            =   1035
            TabIndex        =   20
            Top             =   3400
            Width           =   2865
            WordWrap        =   -1  'True
         End
         Begin VB.Line Line1 
            X1              =   4080
            X2              =   4080
            Y1              =   600
            Y2              =   5160
         End
         Begin VB.Label LblReceivedBy 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4440
            TabIndex        =   9
            Top             =   4800
            Width           =   5535
         End
         Begin VB.Label LblProductCount 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   375
            Left            =   4440
            TabIndex        =   8
            Top             =   1920
            Width           =   5535
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Received By"
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
            Left            =   1560
            TabIndex        =   7
            Top             =   4800
            Width           =   2340
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Current Delivery By:"
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
            Left            =   660
            TabIndex        =   6
            Top             =   4100
            Width           =   3240
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Incoming Product Count"
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
            Left            =   135
            TabIndex        =   5
            Top             =   2700
            Width           =   3810
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "In Stock Product Count"
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
            TabIndex        =   4
            Top             =   2000
            Width           =   3675
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Product ID"
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
            Left            =   1680
            TabIndex        =   3
            Top             =   1300
            Width           =   2265
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Category ID"
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
            Left            =   1560
            TabIndex        =   2
            Top             =   600
            Width           =   2370
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Purchase To Cashier"
         TabPicture(0)   =   "FrmProducts.frx":047A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame6"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Purchase From Cashier"
         TabPicture(1)   =   "FrmProducts.frx":0496
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame7"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame7 
            Caption         =   "Submit"
            Height          =   855
            Left            =   -74880
            TabIndex        =   46
            Top             =   5160
            Width           =   10095
            Begin VB.CommandButton CmdSubmitted 
               Caption         =   "Drugs Submitted"
               Height          =   495
               Left            =   8040
               TabIndex        =   53
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Purchases Where Money Has Been Received By Cashier"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Left            =   -74880
            TabIndex        =   44
            Top             =   480
            Width           =   10095
            Begin VB.TextBox TxtCardNumber 
               Height          =   315
               Left            =   7920
               TabIndex        =   57
               Top             =   0
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.CheckBox ChkFree 
               Caption         =   "Medicine Given Out For Free"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   5280
               TabIndex        =   56
               Top             =   0
               Width           =   2415
            End
            Begin VB.TextBox TxtReceiptNumber 
               Height          =   315
               Left            =   7920
               MaxLength       =   20
               TabIndex        =   52
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox TxtCustomerName 
               Height          =   315
               Left            =   3600
               TabIndex        =   51
               Top             =   600
               Width           =   3975
            End
            Begin VB.ComboBox CboSaleNumber 
               Height          =   315
               Left            =   1440
               TabIndex        =   50
               Top             =   600
               Width           =   1935
            End
            Begin VSFlex6DAOCtl.vsFlexGrid GridPaid 
               Height          =   3495
               Left            =   120
               TabIndex        =   45
               Top             =   1080
               Width           =   9855
               _ExtentX        =   17383
               _ExtentY        =   6165
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
               Caption         =   "Receipt Number"
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
               TabIndex        =   54
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label8 
               Caption         =   "Customer Name"
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
               Left            =   3600
               TabIndex        =   49
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label Label7 
               Caption         =   "Sale Number:"
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
               TabIndex        =   48
               Top             =   600
               Width           =   1215
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Cash Due"
            Height          =   975
            Left            =   8160
            TabIndex        =   42
            Top             =   5040
            Width           =   2055
            Begin VB.TextBox TxtCashDue 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Width           =   1815
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Execute"
            Height          =   975
            Left            =   120
            TabIndex        =   40
            Top             =   5040
            Width           =   7935
            Begin VB.CommandButton CmdToCasheir 
               Caption         =   "Send To Cashier"
               Height          =   615
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   2295
            End
            Begin VB.CommandButton CmdPurchase 
               Caption         =   "Purchase"
               Enabled         =   0   'False
               Height          =   615
               Left            =   5520
               TabIndex        =   41
               Top             =   240
               Visible         =   0   'False
               Width           =   2295
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Medicine Description"
            Height          =   4695
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   10095
            Begin VB.CommandButton CmdRemove 
               Caption         =   "Remove From List"
               Height          =   615
               Left            =   8040
               TabIndex        =   32
               Top             =   3960
               Width           =   1935
            End
            Begin VB.CommandButton CmdAddDrug 
               Caption         =   "Add To List"
               Height          =   615
               Left            =   8040
               TabIndex        =   31
               Top             =   3000
               Width           =   1935
            End
            Begin VB.PictureBox Picture1 
               Height          =   2535
               Left            =   8040
               Picture         =   "FrmProducts.frx":04B2
               ScaleHeight     =   2475
               ScaleWidth      =   1875
               TabIndex        =   30
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox TxtTotalAmount 
               Height          =   375
               Left            =   5880
               TabIndex        =   29
               Top             =   2280
               Width           =   1935
            End
            Begin VB.TextBox TxtQuantity 
               Height          =   375
               Left            =   5880
               TabIndex        =   28
               Top             =   1680
               Width           =   1935
            End
            Begin VB.TextBox TxtPrice2 
               Height          =   375
               Left            =   2040
               TabIndex        =   27
               Top             =   2280
               Width           =   1695
            End
            Begin VB.TextBox TxtUnit2 
               Height          =   375
               Left            =   2040
               TabIndex        =   26
               Top             =   1680
               Width           =   1695
            End
            Begin VB.ComboBox CboProduct3 
               Height          =   315
               Left            =   2040
               TabIndex        =   25
               Top             =   1080
               Width           =   5775
            End
            Begin VB.ComboBox CboCategory3 
               Height          =   315
               Left            =   2040
               TabIndex        =   24
               Top             =   480
               Width           =   5775
            End
            Begin VSFlex6DAOCtl.vsFlexGrid G 
               Height          =   1695
               Left            =   120
               TabIndex        =   33
               Top             =   2880
               Width           =   7695
               _ExtentX        =   13573
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
               Left            =   4200
               TabIndex        =   39
               Top             =   2280
               Width           =   1215
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
               Left            =   4080
               TabIndex        =   38
               Top             =   1800
               Width           =   1575
            End
            Begin VB.Label Label19 
               Caption         =   "Price Per Unit"
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
               Left            =   600
               TabIndex        =   37
               Top             =   2300
               Width           =   1335
            End
            Begin VB.Label Label18 
               Caption         =   "Distribution Unit"
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
               Left            =   480
               TabIndex        =   36
               Top             =   1700
               Width           =   1455
            End
            Begin VB.Label Label17 
               Caption         =   "Medicine Description"
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
               Left            =   120
               TabIndex        =   35
               Top             =   1095
               Width           =   1935
            End
            Begin VB.Label Label16 
               Caption         =   "Medicine Category ID "
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
               Left            =   240
               TabIndex        =   34
               Top             =   495
               Width           =   1695
            End
         End
      End
   End
End
Attribute VB_Name = "FrmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim lvResult As String
Dim lvDirectSalesNo As Long
Public Sub AddMode()
    ReverseGreyOut FrmProducts
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    CMdClose.Caption = "Cancel"
    ClearText FrmProducts
End Sub
Public Sub EditMode()
    ReverseGreyOut FrmProducts
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CMdClose.Caption = "Cancel"
End Sub
Public Sub ResetMode()
    GreyOut FrmProducts
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    CMdClose.Caption = "Close"
    'REQUIREMENT TO KEEP THE LAST CATEGORY SELECTED. Requested by Gemma from Wellness.
    lvcategory = CboCategory
    ClearText FrmProducts
    CboCategory = lvcategory
End Sub

Private Sub CboCategory_Click()
On Error GoTo ErrorHandler
    Dim lvPrescriptionCategoryID As Long
    'POPULATE COMBO FOR DRUGS BY CATEGORY
    CboProduct.Clear
    lvPrescriptionCategoryID = Mid(CboCategory, 1, 3)
    RsRecords.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' order by PRODUCTNAME ASC", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboProduct.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboProduct.AddItem !PRODUCTID & " - " & !ProductName
                End If
                .MoveNext
            Wend
        End With
    RsRecords.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CboCategory3_Click()
On Error GoTo ErrorHandler
    Dim lvPrescriptionCategoryID As Long
    'POPULATE COMBO FOR DRUGS BY CATEGORY
    CboProduct3.Clear
    lvPrescriptionCategoryID = Mid(CboCategory3, 1, 3)
    RsRecords.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' ORDER BY PRODUCTNAME ASC", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboProduct3.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboProduct3.AddItem !PRODUCTID & " - " & !ProductName
                End If
                .MoveNext
            Wend
        End With
    RsRecords.Close
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CboProduct_Click()
    RsRecords.Open "SELECT LASTSTOCKCOUNT,DELIVEREDBY,RECEIVEDBY FROM STOCK_ENTRY WHERE CATEGORYID = '" & Mid(CboCategory, 1, 3) & "' AND PRODUCTID = '" & Mid(CboProduct, 1, InStr(CboProduct, "-") - 1) & "' ORDER BY STOCKID DESC", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If .EOF = False Then
                LblProductCount = !LastStockCount
                LblReceivedBy = !ReceivedBy
                LblLastDeliveryBy = !DELIVEREDBY
                .MoveNext
            Else
                MsgBox "No previous orders have been made for '" & CboProduct & "' Fields will be blank", vbInformation
                LblProductCount = ""
                LblReceivedBy = ""
                LblLastDeliveryBy = ""
            End If
        End With
    RsRecords.Close
End Sub

Private Sub CboProduct3_Click()
    Dim Pos, StrProductID As String
    Pos = InStr(CboProduct3, "-")
    StrProductID = Left(CboProduct3, Pos - 2)
    TxtUnit2 = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & StrProductID & "'")
    TxtPrice2 = FindRecord("PRODUCTS", "SALEPRICE", "PRODUCTID = '" & StrProductID & "'")
End Sub

Private Sub Cmd_Click()

End Sub

Private Sub CboSaleNumber_Click()
    POPULATEPaidGRID CboSaleNumber
End Sub

Private Sub ChkFree_Click()
    If ChkFree.Value = 1 Then
        TxtCardNumber.Visible = True
        TxtCustomerName.Enabled = False: TxtCustomerName = "MEDICINE GIVEN OUT FOR FREE"
        TxtReceiptNumber.Enabled = False: TxtReceiptNumber = 1
    Else
        TxtCardNumber.Visible = False
        TxtCustomerName.Enabled = True: TxtCustomerName = ""
        TxtReceiptNumber.Enabled = True: TxtReceiptNumber = ""
    End If
End Sub

Private Sub cmdadd_Click()
    On Error GoTo ErrorHandler
    AddMode
    lvDirectSalesNo = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'DirectSales'")
    If TBStock.Tab = 1 Then CmdSave.Enabled = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdAddDrug_Click()
    On Error GoTo ErrorHandler
    If lvDirectSalesNo = 0 Then lvDirectSalesNo = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'DirectSales'")
    If CboCategory3.Text = "" Then Exit Sub
    If CboProduct3.Text = "" Then Exit Sub
    
    'CHECK IF MEDICINE IS IN STOCK
    If Val(FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & GetID_NameFromCombo(CboProduct3, 1) & "'")) < Val(TxtQuantity) Then
        MsgBox GetID_NameFromCombo(CboProduct3.Text, 2) & "  Is NOT Available In Stock OR is not Enough for this Dosage and Therefore Cannot Be Issued.", vbExclamation, "Inventory"
        Exit Sub
    End If
        
    'DEDUCT DRUG FROM STOCK
    DEDUCT_DRUG_FROM_STOCK CboProduct3, TxtQuantity, TxtUnit2
    
    Conn.Execute "INSERT INTO PRE_DRUGS_SALES(SALENUMBER,CATEGORYID,PRODUCTID,PRODUCTDESCRIPTION,DISTRIBUTIONUNIT,QUANTITY,AMOUNT,PAYDATE,SOLDBY,DOCTOR0PHARMACY1,PAYMENTMODE,DRUGSUBMITTED)" & _
                 "VALUES ('" & lvDirectSalesNo & "','" & GetID_NameFromCombo(CboCategory3, 1) & "','" & GetID_NameFromCombo(Replace(CboProduct3, "'", " "), 1) & "','" & GetID_NameFromCombo(Replace(CboProduct3, "'", " "), 2) & "','" & TxtUnit2 & "','" & TxtQuantity & "','" & TxtTotalAmount & "','" & GlbSysDate & "','" & GlbCurrentUser & "','1','0','0')"
    POPULATEGRID
    TxtQuantity = ""
    TxtCashDue = TxtCashDue + TxtTotalAmount
    'Conn.Execute "UPDATE"
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
  '  Resume
End Sub
Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "CATEGORY ID|  PRODUCT CATEGORY NAME    | QUANTITY | AMOUNT"
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  PRE_DRUGS_SALES WHERE SALENUMBER = '" & lvDirectSalesNo & "' ", Conn, adOpenStatic, adLockOptimistic
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
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Public Function GetID_NameFromCombo(ByVal Combo, ID_or_Name)
    On Error GoTo ErrorHandler
    Dim Pos As String
    Pos = InStr(Combo, "-")
    If ID_or_Name = 1 Then
        GetID_NameFromCombo = Left(Combo, Pos - 2)
    Else
        GetID_NameFromCombo = Mid(Combo, Pos + 2, Len(Combo))
    End If
Exit Function
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Function
Private Sub CMDCLOSE_Click()
    On Error GoTo ErrorHandler
    If CMdClose.Caption = "Cancel" Then
        ResetMode
    Else
        If G.Row >= 1 Then
            'MsgBox "Please remove the drugs not purchased before closing", vbInformation
            MsgBox "Medicine Sent to Cashier for Cash Collection", vbInformation
            Unload Me
        Else
            Unload Me
        End If
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdDelete_Click()
    On Error GoTo ErrorHandler
    Resp = MsgBox("Deleting this product will also delete the Stock Entered for this Medicine. Are you sure you want to Continue?", vbYesNo + vbExclamation)
    If Resp = vbNo Then Exit Sub
    
    'CHECK IF USER IS SYSTEM ADMINISTRATOR. IF NOT, THEN DENY RIGHT TO DELETE STOCK.
    If VerifyAccess(GlbCurrentUser, "System Administrator") <> True Then MsgBox "You do not have Sufficient Privileges to Delete Stock. Contact System Administrator", vbExclamation: Exit Sub
    
    'DELETE STOCK FOR THE MEDEICINE BEFORE DELETING THE MEDICINE
    Conn.Execute "DELETE FROM STOCK_ENTRY WHERE CATEGORYID = '" & Val(GetID_NameFromCombo(CboCategory, 1)) & "' AND PRODUCTID = '" & GetID_NameFromCombo(CboProduct, 1) & "'"

    'DELETE THE MEDICINE
    Conn.Execute "DELETE FROM PRODUCTS WHERE CATEGORYID = '" & Val(GetID_NameFromCombo(CboCategory, 1)) & "' AND PRODUCTID = '" & GetID_NameFromCombo(CboProduct, 1) & "'"
    MsgBox "Product Deleted Succesfully", vbInformation
    ResetMode
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation

End Sub

Private Sub CmdEdit_Click()
    EditMode
    If TBStock.Tab = 1 Then CmdSave.Enabled = False
End Sub

Private Sub CmdEditStock_Click()
    On Error GoTo ErrorHandler
    If VerifyAccess(GlbCurrentUser, "Pharmacy") = False Then MsgBox "You do not have Sufficient Privileges to access this Report", vbExclamation: CmdEditStock.Enabled = False: Exit Sub
    FrmExitStock.CboCategory = CboCategory.Text
    FrmExitStock.CboProduct = CboProduct.Text
    FrmExitStock.TxtCurrentStock = LblProductCount
    FrmExitStock.Show
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdRemove_Click()
On Error GoTo ErrorHandler
    RETURN_DRUG_TO_STOCK G.TextMatrix(G.Row, 1), G.TextMatrix(G.Row, 2)
    Conn.Execute "DELETE FROM PRE_DRUGS_SALES WHERE CATEGORYID = '" & G.TextMatrix(G.Row, 0) & "' AND PRODUCTID = '" & GetID_NameFromCombo(G.TextMatrix(G.Row, 1), 1) & "'"
POPULATEGRID
    MsgBox "Removed", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdSave_Click()
On Error GoTo ErrorHandler
    Dim lvcategory, lvCategoryName, lvProductID, lvDescription As String
        Select Case TBStock.Tab
        
            Case 0
                If TxtProductCount = "" Then MsgBox "Please Enter the Number of Drugs received before Saving record", vbExclamation: Exit Sub
                If TxtDeliverdBy = "" Then MsgBox "Please Enter the Name of Supplier before Saving record", vbExclamation: Exit Sub
                lvcategory = Mid(CboCategory, 1, 3): lvProductID = Mid(CboProduct, 1, InStr(CboProduct, "-") - 1): lvDescription = Mid(CboProduct, InStr(CboProduct, "-") + 1, Len(CboProduct))
                
                'CHECK IF STOCK RECORD ALREADY EXISTS AND UPDATE. IF NOT, INSERT NEW RECORD
                LELIT = FindRecord("STOCK_ENTRY", "PRODUCTID", "PRODUCTID = '" & lvProductID & "'")
                
                If LELIT <> "" Then
                'BEFORE UPDATE, MOVE PREVIOUS STOCK RECORD TO HISTORY TABLE FOR REPORTING PURPOSES.
                    ARCHIVE_STOCK lvProductID
                'GET LAST STOCK AND SUM UP WITH INCOMING STOCK COUNT
                    OLDSTOCK = FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & lvProductID & "'")
                    Conn.Execute "UPDATE STOCK_ENTRY SET PRODUCTCOUNT = '" & TxtProductCount & "',LASTSTOCKCOUNT = '" & TxtProductCount + OLDSTOCK & "',DELIVEREDBY = '" & UCase(TxtDeliverdBy) & "' WHERE PRODUCTID = '" & lvProductID & "'"
                Else
                
                Conn.Execute "INSERT INTO STOCK_ENTRY (CATEGORYID,PRODUCTID,PRODUCTDESCRIPTION,PRODUCTCOUNT,DELIVERYDATE,LASTSTOCKCOUNT,DELIVEREDBY,RECEIVEDBY)" & _
                             "Values('" & lvcategory & "','" & lvProductID & "','" & Trim(lvDescription) & "','" & TxtProductCount & "','" & Format(GlbSysDate, "dd mmm yyyy") & "','" & TxtProductCount & "','" & Replace(UCase(TxtDeliverdBy), "'", "") & "','" & GlbCurrentUser & "')"
                End If
                MsgBox "Stock Update for '" & lvDescription & "' Has been Saved Succesfully", vbInformation, "Stock Inventory Update"
                lvCategoryName = CboCategory.Text
                    ResetMode
                CboCategory.Text = lvCategoryName
                
            Case 1
                If CboCategory2.Text = "" Then MsgBox "Please Select Category Before Saving", vbExclamation: Exit Sub
                If TxtProductName = "" Then MsgBox "Please Enter Product Name before Saving", vbExclamation: Exit Sub
                If TxtPrice = "" Then MsgBox "Please Enter Amount before Saving", vbInformation: Exit Sub
                Set RsRecords = Nothing
                RsRecords.Open "SELECT * FROM PRODUCTS WHERE PRODUCTNAME = '" & Trim(TxtProductName) & "'", Conn, adOpenStatic, adLockOptimistic
                With RsRecords
                    If BlnEditing = True Then
                        If .BOF = False And .EOF = False Then
                                !CATEGORYID = Mid(CboCategory, 1, InStr(CboCategory, "-") - 1)
                                '!PRODUCTID = TxtProductID ' THIS FIELD IS CURRENTLY IDENTITY
                                !ProductName = UCase(Replace(TxtProductName, "'", " "))
                                !PRESCRIPTIONUNIT = TxtUnit
                                !MINLEVEL = TxtMinimumLevel
                                !REORDERLEVEL = TxtReorderLevel
                                !MAXLEVEL = TxtMinimumLevel
                                !SALEPRICE = TxtPrice
                            .Update
                            MsgBox "Product Edited Successfully", vbInformation
                            BlnEditing = False
                            'lvCategoryName = CboCategory.Text
                            ResetMode
                            'CboCategory.Text = lvCategoryName
                        End If
                    Else
                        If .BOF = False And .EOF = False Then
                            MsgBox "Medicine with Name " & TxtCompanyCode & " is already Maintained", vbExclamation, "Duplication"
                        Else
                            .AddNew
                                !CATEGORYID = Mid(CboCategory2, 1, InStr(CboCategory2, "-") - 1)
                                '!PRODUCTID = TxtProductID ' THIS FIELD IS CURRENTLY IDENTITY
                                !ProductName = UCase(Replace(TxtProductName, "'", " "))
                                !PRESCRIPTIONUNIT = TxtUnit
                                !MINLEVEL = TxtMinimumLevel
                                !REORDERLEVEL = TxtReorderLevel
                                !MAXLEVEL = TxtMinimumLevel
                                !SALEPRICE = TxtPrice
                            .Update
                            MsgBox "Product Maintained Successfully ", vbInformation
                            'lvCategoryName = CboCategory.Text
                            ResetMode
                            'CboCategory.Text = lvCategoryName
                        End If
                    End If
                End With
            RsRecords.Close
        End Select
Exit Sub
ErrorHandler:
    MsgBox Err.Description + vbCrLf + vbCrLf + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
    'Resume
End Sub
Private Sub ARCHIVE_STOCK(STOCKID)
On Error GoTo ErrorHandler
    Dim RsArchive As New ADODB.Recordset
    Conn.Execute "INSERT INTO STOCK_ENTRY_HISTORY SELECT  CategoryID, ProductID, ProductDescription, ProductCount, DeliveryDate, LastStockCount, DeliveredBy, ReceivedBy From STOCK_ENTRY WHERE ProductID = '" & STOCKID & "'"
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdPurchase_Click()
On Error GoTo ErrorHandler
    Conn.Execute "INSERT INTO DRUGS_SALES SELECT * FROM PRE_DRUGS_SALES WHERE SALENUMBER = '" & lvDirectSalesNo & "'"
    MsgBox "Purchase Posted Succesfully", vbInformation
    Conn.Execute "DELETE FROM PRE_DRUGS_SALES"
    POPULATEGRID
Exit Sub
ErrorHandler:
MsgBox Err.Description, vbExclamation
End Sub

Private Sub CmdSubmitted_Click()
    If CboSaleNumber.Text = "" Then MsgBox "Please Select the Sale Number Before Submitting Drugs", vbExclamation: Exit Sub
    If ChkFree.Value = 0 Then
        If TxtCustomerName.Text = "" Then MsgBox "Please Enter the Customer Name as It appears on the Receipt Before Proceeding", vbExclamation: Exit Sub
        If TxtReceiptNumber.Text = "" Then MsgBox "Please Enter the Receipt Number as it appears on the Receipt Before Proceeding", vbExclamation: Exit Sub
    End If
    If ChkFree.Value = 1 And TxtCardNumber = "" Then MsgBox "Please Enter the Patient Card Number Before Proceeding", vbExclamation: TxtCardNumber.SetFocus: Exit Sub
    TxtCustomerName.Text = "Medicine Given out For Free - Card No " & Trim(TxtCardNumber)
    TxtReceiptNumber = 1
    
    'UPDATE DRUGS_SALES TO SHOW DRUGSUBMITTED = TRUE
    Conn.Execute "UPDATE DRUGS_SALES SET DRUGSUBMITTED = 1,PAYMENTMODE = 1,CUSTOMERNAME = '" & UCase(TxtCustomerName) & "',RECEIPTNUMBER = '" & TxtReceiptNumber & "' WHERE SALENUMBER = '" & CboSaleNumber & "'"
    MsgBox "Drugs Submitted and Deducted from Stock", vbInformation
    RePopulateSalesCombo
    GridPaid.Clear: GridPaid.Rows = 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdToCasheir_Click()
    'IF ITS NIGHT SHIFT, THEN ASK FOR RECEIPT NUMBER BEFORE SAVING
    If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'DayShift_0_NightShift_1'") = 1 Then
        FrmReceiptNumber.TxtSaleNumber = lvDirectSalesNo
        FrmReceiptNumber.Show 1
        If BlnReceiptDetails = False Then Exit Sub
    End If
    ResetMode
    G.Clear: G.Rows = 1
    Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = " & lvDirectSalesNo & " + 1 WHERE ITEMNAME = 'DIRECTSALES'"
    MsgBox "Transactions have been Sent to the Casheir for Payment", vbInformation
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    Dim lvPrescriptionCategoryID
    'IF THERE IS NO INHOUSE PHARMACY, THEN DISABLE ALL CONTROLS. DONT BOTHER LOADING ANYTHING.
    If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'ExcludePharmacy'") = 1 Then
        centerform Me
        GreyOut FrmProducts
        DisableAllButtons FrmProducts
        Exit Sub
    End If
    
    'POPULATE COMBO FOR PRESCRIPTION CATEGORY
    RsRecords.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY ORDER BY PRODUCTGROUP", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                CboCategory.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                'CboCategory2.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                CboCategory3.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                .MoveNext
            Wend
        End With
    RsRecords.Close
    ResetMode
    RePopulateSalesCombo
    POPULATEPaidGRID 0
    centerform Me
    ChkFree.Value = 0
    CmdEditStock.Visible = False
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub
Private Sub RePopulateSalesCombo()
    'POPULATE COMBO FOR PHARMACY SALES
    CboSaleNumber.Clear
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT DISTINCT SALENUMBER FROM DRUGS_SALES where drugsubmitted = '0'", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboSaleNumber.AddItem RsRecords!SaleNumber
            RsRecords.MoveNext
        Wend
    RsRecords.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If G.Row >= 1 Then
        MsgBox "Please remove Drugs from list before exiting", vbInformation: Exit Sub
    End If
End Sub

 Private Sub TBStock_Click(PreviousTab As Integer)
    If TBStock.Tab = 0 Then
        CmdEditStock.Visible = True
    Else
        CmdEditStock.Visible = False
    End If
End Sub
Private Sub TxtProductCount_Change()
    ValidateDataType TxtProductCount, 0, "FrmProducts", "TxtProductCount"
End Sub

Private Sub TxtQuantity_Change()
    If TxtPrice2 = "" Or TxtUnit2 = "" Or TxtQuantity = "" Then TxtTotalAmount = "": Exit Sub
    TxtTotalAmount = (TxtPrice2 / TxtUnit2) * TxtQuantity
End Sub
Public Sub POPULATEPaidGRID(ByVal SaleNumber As Double)
On Error GoTo ErrorHandler
    GridPaid.Clear: GridPaid.Rows = 1: GridPaid.Cols = 2
    GridPaid.FormatString = "CATEGORY ID|  PRODUCT CATEGORY NAME    | QUANTITY | AMOUNT"
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        If SaleNumber = 0 Then
            RsRecords.Open "SELECT * FROM  PRE_DRUGS_SALES WHERE DOCTOR0PHARMACY1 = '1' ", Conn, adOpenStatic, adLockOptimistic
        Else
            RsRecords.Open "SELECT * FROM  DRUGS_SALES WHERE DOCTOR0PHARMACY1 = '1' AND SALENUMBER = '" & CboSaleNumber & "'", Conn, adOpenStatic, adLockOptimistic
        End If
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        GridPaid.AddItem !CATEGORYID & vbTab & !PRODUCTID & " - " & !PRODUCTDESCRIPTION & vbTab & !Quantity & vbTab & !amount
                    End With
                RsRecords.MoveNext
                Wend
            End If
        RsRecords.Close
    GridPaid.Editable = True
Exit Sub
ErrorHandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub


