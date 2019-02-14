Attribute VB_Name = "Module1"
Sub FixGhostPayments()

    Dim conn As ADODB.Connection
    Dim sConnString As String


    ' Create the connection string.
    sConnString = Chr(68) & Chr(&H53) & Chr(78) & Chr(&H3D) & Chr(116) & Chr(&H71) & Chr(52) & _
                  Chr(&H3B) & Chr(59) & Chr(&H55) & Chr(73) & Chr(&H44) & Chr(61) & Chr(&H65) & _
                  Chr(115) & Chr(&H69) & Chr(116) & Chr(&H3B) & Chr(80) & Chr(&H57) & Chr(68) & _
                  Chr(&H3D) & Chr(98) & Chr(&H6F) & Chr(115) & Chr(&H73) & Chr(49) & Chr(&H3B) & _
                  Chr(59) & Chr(&H3B) & Chr(68) & Chr(&H41) & Chr(84) & Chr(&H41) & Chr(66) & _
                  Chr(&H41) & Chr(83) & Chr(&H45) & Chr(61) & Chr(&H74) & Chr(104) & Chr(&H61) & _
                  Chr(110) & Chr(&H6B) & Chr(81) & Chr(&H34) & Chr(95) & Chr(&H6C) & Chr(105) & _
                  Chr(&H76) & Chr(101)

    ' Create & Open the Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = sConnString
    conn.ConnectionTimeout = 10
    conn.Open
    
    If CBool(conn.State And adStateOpen) Then
       ' Execute updates
       conn.Execute ("UPDATE BATCHITEMSPLIT " & _
                     "SET BATCHITEMSPLIT.SERIALNUMBER=BATCHITEM.SERIALNUMBER " & _
                     "FROM BATCHITEMSPLIT INNER JOIN BATCHITEM ON (BATCHITEMSPLIT.RECEIPTNO=BATCHITEM.RECEIPTNO AND BATCHITEMSPLIT.ADMITNAME=BATCHITEM.ADMITNAME) " & _
                     "WHERE BATCHITEMSPLIT.SERIALNUMBER Like 'NEW%';")

       conn.Execute ("UPDATE WEBPAYMENT2 " & _
                     "SET WEBPAYMENT2.SERIALNUMBER = WEBPAYMENT1.SERIALNUMBER " & _
                     "FROM WEBPAYMENT AS WEBPAYMENT1 INNER JOIN WEBPAYMENT AS WEBPAYMENT2 ON WEBPAYMENT1.WEBCUSTID=WEBPAYMENT2.WEBCUSTID " & _
                     "WHERE WEBPAYMENT2.SERIALNUMBER Like 'NEW%' AND WEBPAYMENT1.SERIALNUMBER Is Not Null And WEBPAYMENT1.SERIALNUMBER<>WEBPAYMENT2.SERIALNUMBER;")
       
       conn.Execute ("UPDATE BATCHITEMSPLIT " & _
                     "SET BATCHITEMSPLIT.SERIALNUMBER = BATCHITEMPLEDGE.SERIALNUMBER " & _
                     "FROM BATCHITEMPLEDGE INNER JOIN BATCHITEMSPLIT ON BATCHITEMPLEDGE.LINEID=BATCHITEMSPLIT.LINEID AND BATCHITEMPLEDGE.RECEIPTNO=BATCHITEMSPLIT.RECEIPTNO AND BATCHITEMPLEDGE.ADMITNAME=BATCHITEMSPLIT.ADMITNAME " & _
                     "WHERE BATCHITEMSPLIT.SERIALNUMBER<>BATCHITEMPLEDGE.SERIALNUMBER;")
       
       MsgBox "Done!", vbOKOnly

       ' Clean up
       conn.Close
       
    Else
    
       MsgBox "Error: Could not connect to data source.", vbCritical
    
    End If

    Set conn = Nothing

Exit Sub
End Sub


Sub FixOnceOffCureOneWebPayments()

    Dim conn As ADODB.Connection
    Dim sConnString As String


    ' Create the connection string.
    sConnString = Chr(68) & Chr(&H53) & Chr(78) & Chr(&H3D) & Chr(116) & Chr(&H71) & Chr(52) & _
                  Chr(&H3B) & Chr(59) & Chr(&H55) & Chr(73) & Chr(&H44) & Chr(61) & Chr(&H65) & _
                  Chr(115) & Chr(&H69) & Chr(116) & Chr(&H3B) & Chr(80) & Chr(&H57) & Chr(68) & _
                  Chr(&H3D) & Chr(98) & Chr(&H6F) & Chr(115) & Chr(&H73) & Chr(49) & Chr(&H3B) & _
                  Chr(59) & Chr(&H3B) & Chr(68) & Chr(&H41) & Chr(84) & Chr(&H41) & Chr(66) & _
                  Chr(&H41) & Chr(83) & Chr(&H45) & Chr(61) & Chr(&H74) & Chr(104) & Chr(&H61) & _
                  Chr(110) & Chr(&H6B) & Chr(81) & Chr(&H34) & Chr(95) & Chr(&H6C) & Chr(105) & _
                  Chr(&H76) & Chr(101)

    ' Create & Open the Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = sConnString
    conn.ConnectionTimeout = 10
    conn.Open
    
    If CBool(conn.State And adStateOpen) Then
       ' Execute updates

       conn.Execute ("UPDATE WEBPAYMENT " & _
                     "SET PAYMENTFREQUENCY = 'One Off', NEXTINSTALMENTDATE = Null " & _
                     "FROM WEBPAYMENT " & _
                     "WHERE PAYMENTFREQUENCY='Monthly' AND SOURCECODE Like '%CURE%'" & _
                     " AND TRANSACTIONTYPE='Pledge' AND BATCHID Is Null AND AUTOID Is Null" & _
                     " AND DELETED Is Null AND PAYMENTAMOUNT>=400 AND MAXPAYMENTS=1;")
        
       MsgBox "Done!", vbOKOnly

       ' Clean up
       conn.Close
       
    Else
    
       MsgBox "Error: Could not connect to data source.", vbCritical
    
    End If

    Set conn = Nothing

Exit Sub
End Sub


Sub UnapproveBatch()

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim sBatchNo As String
    Dim iStageID As Integer
    Dim dApproved As Date


    ' Create & Open the Connection
    cn.ConnectionTimeout = 10
    cn.ConnectionString = Chr(68) & Chr(&H53) & Chr(78) & Chr(&H3D) & Chr(116) & Chr(&H71) & Chr(52) & _
                  Chr(&H3B) & Chr(59) & Chr(&H55) & Chr(73) & Chr(&H44) & Chr(61) & Chr(&H65) & _
                  Chr(115) & Chr(&H69) & Chr(116) & Chr(&H3B) & Chr(80) & Chr(&H57) & Chr(68) & _
                  Chr(&H3D) & Chr(98) & Chr(&H6F) & Chr(115) & Chr(&H73) & Chr(49) & Chr(&H3B) & _
                  Chr(59) & Chr(&H3B) & Chr(68) & Chr(&H41) & Chr(84) & Chr(&H41) & Chr(66) & _
                  Chr(&H41) & Chr(83) & Chr(&H45) & Chr(61) & Chr(&H74) & Chr(104) & Chr(&H61) & _
                  Chr(110) & Chr(&H6B) & Chr(81) & Chr(&H34) & Chr(95) & Chr(&H6C) & Chr(105) & _
                  Chr(&H76) & Chr(101)
    cn.Open
    
    If CBool(cn.State And adStateOpen) Then
       ' Execute updates
       sBatchNo = InputBox("Batch number", "Unapprove Batch", "B")
       If sBatchNo <> "" Then
       
          sSQL = "SELECT APPROVED, STAGEID FROM BATCHHEADER INNER JOIN OBJECT ON BATCHHEADER.ADMITNAME=OBJECT.ADMITNAME WHERE OBJECT.ADMITNAME='" & sBatchNo & "'"

          rs.Open sSQL, cn
          If rs.EOF Then
             
             MsgBox "Couldn't find that batch number.", vbOKOnly, "Batch not found"
          
          Else
                  
             iStageID = rs!STAGEID
             dApproved = DateValue(rs!APPROVED)

             If iStageID <> 30010 Then
             
                MsgBox "The batch is currently not approved.", vbOKOnly, "Batch Status"
                
             ElseIf dApproved <> Date Then  '
                
                MsgBox "This batch was not approved today.", vbOKOnly, "Batch Approval"
    
             Else
                
                ' update the stageid
                cn.Execute ("UPDATE OBJECT SET STAGEID=30005 FROM OBJECT WHERE ADMITNAME='" & sBatchNo & "';")

                MsgBox sBatchNo & " - Done!", vbOKOnly
             
             End If
             
          End If
          rs.Close
       
       End If
       
       ' Clean up
       cn.Close
       
    Else
    
       MsgBox "Error: Could not connect to data source.", vbCritical
    
    End If

    Set rs = Nothing
    Set cn = Nothing

End Sub

Sub DisableBPayRef()

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim sBPayRef As String


    ' Create & Open the Connection
    cn.ConnectionTimeout = 10
    cn.ConnectionString = Chr(68) & Chr(&H53) & Chr(78) & Chr(&H3D) & Chr(116) & Chr(&H71) & Chr(52) & _
                  Chr(&H3B) & Chr(59) & Chr(&H55) & Chr(73) & Chr(&H44) & Chr(61) & Chr(&H65) & _
                  Chr(115) & Chr(&H69) & Chr(116) & Chr(&H3B) & Chr(80) & Chr(&H57) & Chr(68) & _
                  Chr(&H3D) & Chr(98) & Chr(&H6F) & Chr(115) & Chr(&H73) & Chr(49) & Chr(&H3B) & _
                  Chr(59) & Chr(&H3B) & Chr(68) & Chr(&H41) & Chr(84) & Chr(&H41) & Chr(66) & _
                  Chr(&H41) & Chr(83) & Chr(&H45) & Chr(61) & Chr(&H74) & Chr(104) & Chr(&H61) & _
                  Chr(110) & Chr(&H6B) & Chr(81) & Chr(&H34) & Chr(95) & Chr(&H6C) & Chr(105) & _
                  Chr(&H76) & Chr(101)
    cn.Open
    
    If CBool(cn.State And adStateOpen) Then
       ' Execute updates
       sBPayRef = InputBox("BPay Reference", "Disable BPay Reference", "")
       If sBPayRef <> "" Then
       
          'SQL to check the BPay Ref is valid
          sSQL = "SELECT BPAYREF, BPAYCHECKSUM, MAILINGDETAILSENT.ADMITNAME, SERIALNUMBER, MAILINGHEADER.TITLE, TOENVELOPESALUTATION " & _
                 "FROM MAILINGDETAILSENT inner join MAILINGHEADER on MAILINGHEADER.ADMITNAME=MAILINGDETAILSENT.ADMITNAME " & _
                 "WHERE BPAYREF Is Not Null AND BPAYCHECKSUM Is Not Null AND MAILSENT=1 AND BPAYREF+BPAYCHECKSUM='" & sBPayRef & "'"

          rs.Open sSQL, cn
          If rs.EOF Then
             
             MsgBox "Couldn't find that BPay reference.", vbOKOnly, "BPay not found"
          
          Else

             ' update the checksum by adding a second copy of itself
             sSQL = "UPDATE MAILINGDETAILSENT SET BPAYCHECKSUM='" & rs!BPAYCHECKSUM & rs!BPAYCHECKSUM & "' " & _
                    "FROM MAILINGDETAILSENT " & _
                    "WHERE BPAYREF Is Not Null AND BPAYCHECKSUM Is Not Null AND MAILSENT=1 AND BPAYREF+BPAYCHECKSUM='" & sBPayRef & "'"


            cn.Execute (sSQL)
            
            MsgBox "Done!", vbOKOnly
          
          End If
          rs.Close
       
       End If
       
       ' Clean up
       cn.Close
       
    Else
    
       MsgBox "Error: Could not connect to data source.", vbCritical
    
    End If

    Set rs = Nothing
    Set cn = Nothing

End Sub


Sub WebOrderSetSource2Merch()

    Dim conn As ADODB.Connection
    Dim sConnString As String


    ' Create the connection string.
    sConnString = Chr(68) & Chr(&H53) & Chr(78) & Chr(&H3D) & Chr(116) & Chr(&H71) & Chr(52) & _
                  Chr(&H3B) & Chr(59) & Chr(&H55) & Chr(73) & Chr(&H44) & Chr(61) & Chr(&H65) & _
                  Chr(115) & Chr(&H69) & Chr(116) & Chr(&H3B) & Chr(80) & Chr(&H57) & Chr(68) & _
                  Chr(&H3D) & Chr(98) & Chr(&H6F) & Chr(115) & Chr(&H73) & Chr(49) & Chr(&H3B) & _
                  Chr(59) & Chr(&H3B) & Chr(68) & Chr(&H41) & Chr(84) & Chr(&H41) & Chr(66) & _
                  Chr(&H41) & Chr(83) & Chr(&H45) & Chr(61) & Chr(&H74) & Chr(104) & Chr(&H61) & _
                  Chr(110) & Chr(&H6B) & Chr(81) & Chr(&H34) & Chr(95) & Chr(&H6C) & Chr(105) & _
                  Chr(&H76) & Chr(101)

    ' Create & Open the Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = sConnString
    conn.ConnectionTimeout = 10
    conn.Open
    
    If CBool(conn.State And adStateOpen) Then
       ' Execute updates
       conn.Execute ("UPDATE WEBPAYMENT " & _
                     "SET SOURCECODE2 = 'Merch' " & _
                     "FROM WEBPAYMENT " & _
                     "WHERE SOURCECODE2 Is Null AND NOTES Like '%REX Order Number:%' AND WEBREFERENCE=WEBCUSTID AND DELETED Is Null AND BATCHID Is Null;")
       
       MsgBox "Done!", vbOKOnly

       ' Clean up
       conn.Close
       
    Else
    
       MsgBox "Error: Could not connect to data source.", vbCritical
    
    End If

    Set conn = Nothing

Exit Sub
End Sub


Sub ChangeEBayWebPaymentSourceCodes()

    Dim conn As ADODB.Connection
    Dim sConnString As String


    ' Create the connection string.
    sConnString = Chr(68) & Chr(&H53) & Chr(78) & Chr(&H3D) & Chr(116) & Chr(&H71) & Chr(52) & _
                  Chr(&H3B) & Chr(59) & Chr(&H55) & Chr(73) & Chr(&H44) & Chr(61) & Chr(&H65) & _
                  Chr(115) & Chr(&H69) & Chr(116) & Chr(&H3B) & Chr(80) & Chr(&H57) & Chr(68) & _
                  Chr(&H3D) & Chr(98) & Chr(&H6F) & Chr(115) & Chr(&H73) & Chr(49) & Chr(&H3B) & _
                  Chr(59) & Chr(&H3B) & Chr(68) & Chr(&H41) & Chr(84) & Chr(&H41) & Chr(66) & _
                  Chr(&H41) & Chr(83) & Chr(&H45) & Chr(61) & Chr(&H74) & Chr(104) & Chr(&H61) & _
                  Chr(110) & Chr(&H6B) & Chr(81) & Chr(&H34) & Chr(95) & Chr(&H6C) & Chr(105) & _
                  Chr(&H76) & Chr(101)

    ' Create & Open the Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = sConnString
    conn.ConnectionTimeout = 10
    conn.Open
    
    If CBool(conn.State And adStateOpen) Then
       ' Execute updates

       conn.Execute ("UPDATE WEBPAYMENT SET SOURCECODE='__MEBY'+SUBSTRING(SOURCECODE,7,10) " & _
                     "FROM WEBPAYMENT WHERE SOURCECODE Like '__M___[DGPN]' AND SOURCECODE2='eBay' AND BATCHID Is Null AND DELETED Is Null;")
        
       MsgBox "Done!", vbOKOnly

       ' Clean up
       conn.Close
       
    Else
    
       MsgBox "Error: Could not connect to data source.", vbCritical
    
    End If

    Set conn = Nothing

Exit Sub
End Sub

'
' When WebPayments have notes included in the Authorisation Code fields, clear them.
'
Sub FixWebPaymentAuthCodes()

    Dim conn As ADODB.Connection
    Dim sConnString As String


    ' Create the connection string.
    sConnString = Chr(68) & Chr(&H53) & Chr(78) & Chr(&H3D) & Chr(116) & Chr(&H71) & Chr(52) & _
                  Chr(&H3B) & Chr(59) & Chr(&H55) & Chr(73) & Chr(&H44) & Chr(61) & Chr(&H65) & _
                  Chr(115) & Chr(&H69) & Chr(116) & Chr(&H3B) & Chr(80) & Chr(&H57) & Chr(68) & _
                  Chr(&H3D) & Chr(98) & Chr(&H6F) & Chr(115) & Chr(&H73) & Chr(49) & Chr(&H3B) & _
                  Chr(59) & Chr(&H3B) & Chr(68) & Chr(&H41) & Chr(84) & Chr(&H41) & Chr(66) & _
                  Chr(&H41) & Chr(83) & Chr(&H45) & Chr(61) & Chr(&H74) & Chr(104) & Chr(&H61) & _
                  Chr(110) & Chr(&H6B) & Chr(81) & Chr(&H34) & Chr(95) & Chr(&H6C) & Chr(105) & _
                  Chr(&H76) & Chr(101)

    ' Create & Open the Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = sConnString
    conn.ConnectionTimeout = 10
    conn.Open
    
    If CBool(conn.State And adStateOpen) Then
       ' Execute updates
       conn.Execute ("UPDATE WEBPAYMENT SET WEBAUTHORISATION = '', CREDITCARDAUTHCODE = ''" & _
                     " WHERE (Trim(WEBAUTHORISATION) Like '% %' And TRIM(WEBAUTHORISATION)<>'Direct Deposit' And WEBAUTHORISATION Not Like '%chq%'" & _
                     " And WEBAUTHORISATION Not Like '%cheque%' And WEBAUTHORISATION Not Like '%eway%' And WEBAUTHORISATION Not Like '%eft%'" & _
                     " And WEBAUTHORISATION Not Like '%cash%') OR Len(WEBAUTHORISATION)>=20;")

       MsgBox "Done!", vbOKOnly

       ' Clean up
       conn.Close
       
    Else
    
       MsgBox "Error: Could not connect to data source.", vbCritical
    
    End If

    Set conn = Nothing

Exit Sub
End Sub
