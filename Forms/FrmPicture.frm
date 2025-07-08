VERSION 5.00
Begin VB.Form FrmPicture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exploded Picture"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "FrmPicture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmPicture.frx":0442
   ScaleHeight     =   8685
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Pic"
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.PictureBox Picture1 
         Height          =   1815
         Left            =   4560
         ScaleHeight     =   1755
         ScaleWidth      =   1875
         TabIndex        =   1
         Top             =   3360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image ImgPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DragMode        =   1  'Automatic
         Height          =   7815
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   9735
      End
   End
End
Attribute VB_Name = "FrmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsImageStore As New ADODB.Recordset
Private Sub Form_Load()
On Error GoTo ErrorHandler
    centerform Me
    
    TxtLabTest = ""
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM PATIENT_PICTURES WHERE CARDNUMBER = '" & FrmPatients.TxtCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsImageStore
            If .EOF = False Then
                LoadPictureFromDB RsImageStore, "PICTURE", ImgPreview, D
            End If
        End With
    RsImageStore.Close
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
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
    MsgBox Err.Description, vbExclamation, "Image Problem"
    'SystemErrorHandler Err.Number, Err.Description
End Function


