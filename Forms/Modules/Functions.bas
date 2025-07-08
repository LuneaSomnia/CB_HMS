Attribute VB_Name = "Functions"
Public Conn As New ADODB.Connection
Public StrDocCardNo As String
Public lvOptionalCardNo As String
Public lvOptionalVisitNo As String
Public lvFoodCardNo As String
Public lvFoodVisitNo As Integer
Public StrDocVisitNumber As String
Public StrDocBMI As String
Public StrDocVisitDate As String
Public StrLabAmount As String
Public StrPharmCardNumber As String
Public StrPharmVisitDate As Date
Public BlnHISTORY As Boolean
Public SendCardnumber As String
Public SendVisitDate As Date
Public StrLoggedInUser As String
Public GlbUnitQuantity As String
Public GlbPasswordResetLogin As Boolean
Public GlbMedEditID As String
Public GlbSysDate As Date
Public DBUser As String
Public DBPassword As String
Public BlnEditingComplains As Boolean
Public BlnEditingLabTests As Boolean
Public BlnReceiptDetails As Boolean
Public BlnViewMeasurements As Boolean
Public GlbDropNumber As Long
Public GlbDropCancel As Boolean
Public GlbDropView As Boolean
Public GlbTestImageType As Integer
Public GlbImageVisitNumber As Long
'************* For Images - SCAN AND SAVE.
Global Const conChunkSize = 2048       '2048 increment  multiples  depending the image sizes
Private pInititialized As Boolean
Private pTable(0 To 255) As Long
Dim nFragmentOffset As Integer
Dim CHUNK() As Byte

Public Sub CheckLicense()
Dim LvExpiryDate As String
    LvExpiryDate = GetSetting("DVSHMS", "LICENSE", "EXPIRY")
        If LvExpiryDate = "" Then
            SaveSetting "DVSHMS", "LICENSE", "EXPIRY", GlbSysDate + 30
        Else
            If LvExpiryDate < GlbSysDate Then
                MsgBox "Your Software License has expired" & vbCrLf & _
                        "Please consult the System administrator or System Vendor", vbCritical, "LICENCE EXPIRED !!!"
                End
            End If
        End If
End Sub

Public Function LoadImageFromFileToDB(ByVal gsFileName As String, rsConnect As ADODB.Recordset, FieldName As String, ByVal lngImgSize As Long) As Boolean
    On Error GoTo Errorhandler
    Close nhandle
    nhandle = 1
    Open gsFileName For Binary Access Read As nhandle
        lSize = lngImgSize
'        lSize = LOF(nHandle)
'        If nHandle = 0 Then
'            Close nHandle
'        End If
        'lSize
        lChunks = lSize \ conChunkSize
        nFragmentOffset = lSize Mod conChunkSize
        
        'rsConnect("ImgMyKey") = lKey
        
        ReDim CHUNK(nFragmentOffset)
        Get nhandle, , CHUNK()
        rsConnect(FieldName).AppendChunk CHUNK()
        'rsConnect("ImageFileBin").AppendChunk Chunk()
        ReDim CHUNK(conChunkSize)
        lOffset = nFragmentOffset
        For i = 1 To lChunks
            Get nhandle, , CHUNK()
            rsConnect(FieldName).AppendChunk CHUNK()
            lOffset = lOffset + conChunkSize
            txtByteCount = lOffset
            'DoEvents
        Next
        Close nhandle
        Exit Function
        
Errorhandler:
MsgBox Err.Number & " " & Err.Description
    Exit Function
    Resume
End Function
Public Function FindFileRows(ByVal strFileName As String) As Long
   On Error GoTo Errorhandler
   Dim lngFileNo As Long
   Dim lngLineCount As Long
   Dim strBuff As String
 lngFileNo = FreeFile(0)
 Open strFileName For Input As lngFileNo
   Do Until EOF(lngFileNo)
      Line Input #lngFileNo, strBuff
      If strBuff <> "" Then
            lngLineCount = lngLineCount + 1
      End If
   Loop
 Close lngFileNo
 'FindFileRows =lngLineCount
 'arr = Split(strBuff, Chr(9))
 'FindFileRows = UBound(arr)
 FindFileRows = lngLineCount
 Exit Function
Errorhandler:
MsgBox Err.Number & " " & Err.Description
End Function

Public Function FindRecord(TableName As String, LookupField As String, Optional WhereCondition As String = "") As Variant
    On Error GoTo Errorhandler:
    Dim strSQL As String
      Dim adoRS As New ADODB.Recordset
      
    LookupValue = Null
    strSQL = "SELECT " & LookupField & " FROM " & TableName
        If Trim(WhereCondition) <> "" Then strSQL = strSQL & " WHERE " & WhereCondition
        adoRS.Open strSQL, Conn, adOpenStatic, adLockOptimistic
          If Not adoRS.EOF Then
            FindRecord = adoRS.Fields(0).Value
          Else
            FindRecord = ""
          End If
        adoRS.Close
        Set adoRS = Nothing
    Exit Function
Errorhandler:
MsgBox Err.Description

End Function
Public Function SendPatient(Section As Integer, SendCardnumber As String, ByVal VisitDate As Date, Optional VisitNo As Long)
        Select Case Section
            Case 1
                'To Consultation
                    If VisitNo = 0 Then
                        Conn.Execute "UPDATE COMPLAINS SET TOCONSULTATION = '1',TODOCTORS = '0',TOPHARMACY = '0',TOCASHIER = '0',TOOBSERVATION = '0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITDATE = '" & Format(VisitDate, "DDMMMYYYY") & "'"
                        MDIMain.SSBar.Panels(3).Text = "Record Succesfully Posted  to Consultation"
                    Else
                        'SENDING ONLY ONE VISIT NUMBER AND NOT ALL FROM THE SAME CARD NUMBER.
                        Conn.Execute "UPDATE COMPLAINS SET TOCONSULTATION = '1',TODOCTORS = '0',TOPHARMACY = '0',TOCASHIER = '0',TOOBSERVATION = '0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITNUMBER = '" & VisitNo & "'"
                        MDIMain.SSBar.Panels(3).Text = "Record Succesfully Posted  to Consultation"
                    End If
            Case 2
                 'To Observation
                        Conn.Execute "UPDATE COMPLAINS SET TOOBSERVATION = '1',TOCONSULTATION = '0',TODOCTORS = '0',TOPHARMACY = '0',TOCASHIER = '0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITDATE = '" & Format(VisitDate, "DD MMMM YYYY") & "'"
                        MDIMain.SSBar.Panels(3).Text = "Record Succesfully Posted  to Observation"
             Case 3
                 'To Doctor
                    If VisitNo = 0 Then
                        Conn.Execute "UPDATE COMPLAINS SET TODOCTORS = '1', TOCONSULTATION = '0', TOOBSERVATION = '0', TOPHARMACY = '0', TOCASHIER = '0',TOLABORATORY = '0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITDATE = '" & Format(VisitDate, "DDMMMYYYY") & "'"
                        MDIMain.SSBar.Panels(3).Text = "Record Succesfully Posted  to Doctors"
                    Else
                        Conn.Execute "UPDATE COMPLAINS SET TODOCTORS = '1', TOCONSULTATION = '0', TOOBSERVATION = '0', TOPHARMACY = '0', TOCASHIER = '0',TOLABORATORY = '0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITNUMBER = '" & VisitNo & "'"
                        MDIMain.SSBar.Panels(3).Text = "Record Succesfully Posted  to Doctors"
                    End If
             Case 4
                 'To Pharmacy
                     Conn.Execute "UPDATE COMPLAINS SET TOPHARMACY = '1',TOCONSULTATION = '0',TOOBSERVATION = '0',TODOCTORS = '0',TOCASHIER = '0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITDATE = '" & Format(VisitDate, "DDMMMYYYY") & "'"
                     MDIMain.SSBar.Panels(3).Text = "Record Succesfully Posted  to Pharmacy"
            Case 5
                 'To Casheir
                     Conn.Execute "UPDATE COMPLAINS SET TOCASHIER = '1', TOCONSULTATION = '0',TOOBSERVATION = '0',TODOCTORS='0',TOPHARMACY = '0',TOLABORATORY = '0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITDATE = '" & Format(VisitDate, "DDMMMYYYY") & "'"
                    MDIMain.SSBar.Panels(3).Text = "Record Succesfully Posted  to Cashier"
            Case 6
                 'To Laboratory
                     Conn.Execute "UPDATE COMPLAINS SET TOLABORATORY = '1', TOCONSULTATION = '0',TOOBSERVATION = '0',TODOCTORS='0',TOPHARMACY = '0',TOCASHIER ='0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITDATE = '" & Format(VisitDate, "DDMMMYYYY") & "'"
                    MDIMain.SSBar.Panels(3).Text = "Record Succesfully Posted  to Laboratory"
            Case 11
                'To Admission
                     Conn.Execute "UPDATE COMPLAINS SET TOADMISSION = '1', TOCONSULTATION = '0',TOOBSERVATION = '0',TODOCTORS='0',TOPHARMACY = '0',TOCASHIER ='0' WHERE CARDNUMBER = '" & SendCardnumber & "' AND VISITDATE = '" & Format(VisitDate, "DDMMMYYYY") & "'"
                    MDIMain.SSBar.Panels(3).Text = "Card Number " & SendCardnumber & " Succesfully Posted  to Admission"
        End Select
            MsgBox MDIMain.SSBar.Panels(3).Text, vbInformation, "Posted !!"
            
        GlbCardNumber = ""
        Beep
        SendPatient = True
End Function
Public Sub ValidateDataType(CurrentText As String, ExpectedDataType_0_Num_1_Text As Integer, CurrentForm As String, TxTName As String)
On Error GoTo Errorhandler
        Select Case ExpectedDataType_0_Num_1_Text
            Case 0
            
                'CHECK FOR CHARACTERS IN THE STRING
                ' IF LENGTH IS GREATER THAN ONE AND THERE ARE NO NUMBERS, IT IS A STRING ONLY
                If Len(CurrentText) >= 1 Then
                    If Not IsNumeric(CurrentText) Then
                           MsgBox "Only Numbers are allowed in this field. Please re enter", vbInformation
                    End If
                End If
            Case 1
            
                For i = 0 To 9
                    'CHECK FOR NUMBERS IN THE STRING
                    If InStr(1, CurrentText, i) <> 0 Then
                            MsgBox "Only Characters are allowed in this field. Please re enter", vbInformation
                    Exit For
                    End If
                Next i
        End Select
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Public Function ValidateDataType_Advice(CurrentTextBox, ExpectedDataType_0_Num_1_Text As Integer) As Boolean
On Error GoTo Errorhandler
        ValidateDataType_Advice = True
        Select Case ExpectedDataType_0_Num_1_Text
            Case 0
            
                'CHECK FOR CHARACTERS IN THE STRING
                ' IF LENGTH IS GREATER THAN ONE AND THERE ARE NO NUMBERS, IT IS A STRING ONLY
                If Len(CurrentTextBox) >= 1 Then
                    If Not IsNumeric(CurrentTextBox) Then
                           MsgBox "Only Numbers are allowed in this field. Please re enter", vbInformation
                           ValidateDataType_Advice = False
                    End If
                End If
            Case 1
            
                For i = 0 To 9
                    'CHECK FOR NUMBERS IN THE STRING
                    If InStr(1, CurrentTextBox, i) <> 0 Then
                            MsgBox "Only Characters are allowed in this field. Please re enter", vbInformation
                            ValidateDataType_Advice = False
                    Exit For
                    End If
                Next i
        End Select
Exit Function
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Function

Public Function PFValidateDataType(CurrentText As String, ExpectedDataType_0_Num_1_Text As Integer)
        Select Case ExpectedDataType
            Case 0
            
                'CHECK FOR CHARACTERS IN THE STRING
                ' IF LENGTH IS GREATER THAN ONE AND THERE ARE NO NUMBERS, IT IS A STRING ONLY
                If Len(CurrentText) >= 1 Then
                    If Not IsNumeric(CurrentText) Then
                           MsgBox "Only Numbers are allowed in this field. Please re enter", vbInformation
                           PFValidateDataType = -1
                    End If
                End If
            Case 1
            
                For i = 0 To 9
                    'CHECK FOR NUMBERS IN THE STRING
                    If InStr(1, CurrentText, i) <> 0 Then
                            MsgBox "Only Characters are allowed in this field. Please re enter", vbInformation
                            PFValidateDataType = -1
                    Exit For
                    End If
                Next i
        End Select
End Function

Public Sub ClearText(FrmName As Form)
    Dim CTL As Control
    Dim KARI As Form
    
    Set KARI = FrmName
        For Each CTL In KARI.Controls
            If TypeOf CTL Is TextBox Then
                CTL.Text = ""
            ElseIf TypeOf CTL Is ComboBox Then
                CTL.Text = ""
            ElseIf TypeOf CTL Is CheckBox Then
                CTL.Value = 0
            End If
        Next
End Sub

Public Sub ReverseGreyOut(FrmName As Form)
    Dim CTL As Control
    Dim KARI As Form
    Set KARI = FrmName
        For Each CTL In KARI.Controls
            If TypeOf CTL Is TextBox Then
                CTL.Enabled = True
            ElseIf TypeOf CTL Is ComboBox Then
                CTL.Enabled = True
            ElseIf TypeOf CTL Is CommandButton Then
                CTL.Enabled = True
            End If
        Next
End Sub
Public Sub GreyOut(FrmName As Form)
    Dim CTL As Control
    Dim KARI As Form
    Set KARI = FrmName
        For Each CTL In KARI.Controls
            If TypeOf CTL Is TextBox Then
                CTL.Enabled = False
            ElseIf TypeOf CTL Is ComboBox Then
                CTL.Enabled = False
            End If
        Next
        ClearText KARI
End Sub
Public Sub DisableAllButtons(FrmName As Form)
    Dim CTL As Control
    Dim KARI As Form
    Set KARI = FrmName
        For Each CTL In KARI.Controls
            If TypeOf CTL Is Button Then
                CTL.Enabled = False
            End If
        Next
End Sub
Sub centerform(X As Form)
    Screen.MousePointer = 11
    If X.MDIChild = False Then
        On Error Resume Next: X.Top = (Screen.Height / 2) - X.Height / 2
        X.Left = (Screen.Width / 2) - (X.Width / 2)
    Else
        'MDI.Arrange cascade
    End If
    Screen.MousePointer = 0
End Sub
Public Function ConvertToUppercase(StrText As String)
    ConvertToUppercase = UCase(StrText)
End Function


Sub Main()
    On Error GoTo Err
    Dim FSO As New FileSystemObject
    Dim RsDates As New ADODB.Recordset
    'Get Connection Credentials
        DBUser = "sa"
        DBPassword = "Today123"
        'DBPassword = "CES123"

    '***************
        Conn.ConnectionString = "DSN=NCC;Password='" & DBPassword & "';user id='" & DBUser & "'"
        Conn.Open
        
    '**************
    'LOAD SYSTEM GLOBAL DATE
    RsDates.Open "SELECT CurrentDate FROM DUAL", Conn, adOpenStatic, adLockOptimistic
        With RsDates
            If .EOF = False Then
               GlbSysDate = Format(!CurrentDate, "dd/MMM/YYYY")
            End If
        End With
    RsDates.Close
    
    'CHECK THE LICENCE EXPIRY BEFORE LOGIN.
    CheckLicense
    
    frmSplash.Show
    Exit Sub
Err:
    If Err.Number = -2147467259 Then MsgBox "The Server Machine May be SHUT DOWN or removed from the NETWORK. Resolve these issues and then try again!", vbExclamation, "QUEUE MANAGER MACHINE": Exit Sub
    MsgBox Err.Number & " " & Err.Description, vbCritical
End Sub

Public Function VerifyAccess(StrUserName As String, StrScreenToAccess) As Boolean
On Error GoTo Errorhandler
    Dim RsAccess As New ADODB.Recordset
    RsAccess.Open "SELECT * FROM PROFILES WHERE USERNAME = '" & StrUserName & "' AND RIGHTS = '" & StrScreenToAccess & "'", Conn, adOpenStatic, adLockOptimistic
    If RsAccess.EOF = False Then
        If RsAccess!ACCESS = "True" Then
            VerifyAccess = True
        Else
            VerifyAccess = False
        End If
    End If
Exit Function
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Function
Public Function SwitchScreen(ByVal CurrentForm As Integer, ByVal DestinationForm As Integer) As Boolean
    Dim KARI As Form
    Dim KISH As Form
'I GOT BUGS TRYING TO PASS A FORM AS A FUNCTION PARAMETER SO I DECIDED TO WRITE CODE MINGI WITH THE SAME PURPOSE
        Select Case CurrentForm
            Case 1
                Unload FrmPatients
            Case 2
                Unload FrmObservation
            Case 3
                Unload FrmTreatment
            Case 4
                Unload FrmPharmacy
            Case 5
                Unload FrmCashier
            Case 6
                Unload FrmLaboratory
            Case 7
                Unload FrmSecurity
            Case 9
                Unload FrmReports
            Case 10
                Unload FrmDashBoard
            Case 11
                Unload FrmWard
        End Select

        Select Case DestinationForm
            Case 1
                FrmPatients.Show
            Case 2
                FrmObservation.Show
            Case 3
                FrmWaitingRoom.Show
            Case 4
                FrmPharmacy.Show
            Case 5
                FrmCashier.Show
            Case 6
                FrmLaboratoryWaiting.Show
            Case 7
                FrmSecurity.Show
            Case 9
                FrmReports.Show
            Case 10
                FrmDashBoard.Show
            Case 11
                FrmWard.Show
        End Select
End Function
Public Function DEDUCT_DRUG_FROM_STOCK(ByVal DrugIssued, ByVal Quantity, ByVal UnitsSold) As Boolean
On Error GoTo Errorhandler
Dim DCount As Long
Dim lvShift As Integer
Dim lvAmount As Integer
    Dim ProdID, Pos, DistributionUnit As String
    'IF THERE IS NO INHOUSE PHARAMCY IN THIS HOSPITAL THEN DONT BOTHER CHECKING STOCK.
        If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'ExcludePharmacy'") = 1 Then
            DEDUCT_DRUG_FROM_STOCK = True
            Exit Function
        End If
    '***********************
    GlbUnitQuantity = Quantity
    Pos = InStr(DrugIssued, "-")
    ProdID = GetID_NameFromCombo(DrugIssued, 1) 'Left(DrugIssued, Pos - 2)
    If FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & ProdID & "'") = "" Then
        MsgBox "There is No stock for " & GetID_NameFromCombo(DrugIssued, 2) & " " & "In Pharmacy. Drug Cannot be Dispensed", vbCritical
        Exit Function
    End If
    DCount = FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & ProdID & "'")
    If GlbUnitQuantity = "" Then GlbUnitQuantity = 0
    If GlbUnitQuantity = 0 Then
        DistributionUnit = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & ProdID & "'")
    Else
        DistributionUnit = UnitsSold
        DistributionUnit = GlbUnitQuantity * DistributionUnit
        'GlbUnitQuantity = 0
    End If
    If DCount > 1 Then
        Conn.Execute "UPDATE STOCK_ENTRY SET LASTSTOCKCOUNT = '" & Val(DCount) & "' - " & DistributionUnit & "  WHERE PRODUCTID = '" & ProdID & "'"
        DEDUCT_DRUG_FROM_STOCK = True
        
        'INSERT SALE RECORD TO DRUG SALES REPORT WITH INDICATION IF ITS DAY OR NIGHT SHIFT.
         lvShift = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'DayShift_0_NightShift_1'")
         lvAmount = FindRecord("PRODUCTS", "SALEPRICE", "PRODUCTID = '" & GetID_NameFromCombo(DrugIssued, 1) & "'")
         lvAmount = lvAmount * DistributionUnit
         
        Conn.Execute "INSERT INTO DRUG_SALES_REPORT(PRODUCTID,PRODUCTDESCRIPTION,QUANTITY,SHIFT,SALEDATE,SALEUSER,AMOUNT)" & _
                     "VALUES ('" & GetID_NameFromCombo(DrugIssued, 1) & "','" & GetID_NameFromCombo(DrugIssued, 2) & "', '" & DistributionUnit & "','" & lvShift & "','" & GlbSysDate & "','" & GlbCurrentUser & "','" & lvAmount & "')"
    Else
        MsgBox "This Medicine has run out of stock and will therefore not be added.", vbInformation, "Sorry!"
        DEDUCT_DRUG_FROM_STOCK = False
    End If
Exit Function
Errorhandler:
    'MsgBox "Medicine not yet Maintained in Stock", vbExclamation
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
    'Resume
End Function
Public Function RETURN_DRUG_TO_STOCK(ByVal DrugIssued, ByVal Quantity) As Boolean
On Error GoTo Errorhandler
Dim lvQuantity As Integer
'REMOVE LAB TESTS OR CONSULTATION WITHOUT SEARCHING IN STOCK.
    If InStr(1, DrugIssued, "Lab Test") > 0 Then Exit Function
    If InStr(1, DrugIssued, "Consultation") > 0 Then Exit Function
    'If InStr(1, DrugIssued, "^") = 0 Then Exit Function
    
Dim DCount As Long
    Dim ProdID, Pos, DistributionUnit As String
    Dim QtyArray
    QtyArray = Split(DrugIssued, "^")
    Pos = InStr(DrugIssued, "-")
    ProdID = Left(DrugIssued, Pos - 2)
        DCount = FindRecord("STOCK_ENTRY", "LASTSTOCKCOUNT", "PRODUCTID = '" & ProdID & "'")
    If D = 5 Then
lookup:
        DistributionUnit = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & ProdID & "'")
    Else
        DistributionUnit = FindRecord("PRODUCTS", "PRESCRIPTIONUNIT", "PRODUCTID = '" & ProdID & "'")
        lvQuantity = Quantity
        DistributionUnit = DistributionUnit * Quantity
    End If
        Conn.Execute "UPDATE STOCK_ENTRY SET LASTSTOCKCOUNT = '" & Val(DCount) & "' + " & Val(DistributionUnit) & " WHERE PRODUCTID = '" & ProdID & "'"
        RETURN_DRUG_TO_STOCK = True
        
        'DELETE THE PRE DRUG STOCK AND REFRESH GRID TO AVOID DUPLICATE RETURNS.
Exit Function
Errorhandler:
    If Err.Description = "Subscript out of range" Then
        Err.Clear: GoTo lookup
    Else
        MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
    End If
End Function
Public Function GetID_NameFromCombo(ByVal Combo, ID_or_Name)
    On Error GoTo Errorhandler
    Dim Pos As String
    Pos = InStr(Combo, "-")
    If ID_or_Name = 1 Then
        GetID_NameFromCombo = Left(Combo, Pos - 2)
    Else
        GetID_NameFromCombo = Mid(Combo, Pos + 2, Len(Combo))
    End If
Exit Function
Errorhandler:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Function
Public Sub AuditTrail(ByVal CurrentUser, ByVal CurrentModule, CurrentDate, CurrentTime, ActionTaken)
On Error GoTo Errorhandler
    'INSERT INTO PROAUDIT
    Conn.Execute "INSERT INTO PROAUDIT (SYSTEMUSER,SCREEN,DATE,TIME,ACTION) VALUES('" & CurrentUser & "','" & CurrentModule & "','" & CurrentDate & "', '" & CurrentTime & "','" & ActionTaken & "')"
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Public Function CalculateBMI(lvWeight As Integer, lvHeight As Integer)
    CalculateBMI = CDbl(lvWeight) / CDbl(lvHeight)
End Function
