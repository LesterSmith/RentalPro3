''' Purpose:
''' Split one invoice into two so that one or more
''' pieces of equipment can be checked in without
''' checking in all of the equipment.
''' Author:  Les Smith
''' Date Created: 11/18/2003 at 04:05:00
''' CopyRight:  HHI Software
'''
Public Class CInvoiceSplit
#Region " Public Methods "
   ''' <summary>
   ''' Check to see if the invoice needs to be split.
   ''' Return the invoice id of the invoice to be loaded.
   ''' </summary>
   ''' <param name = "frm"></param>
   ''' <returns>Integer</returns>
   Public Function CheckForSplittingInvoice(ByRef frm As frmSelectCheckInInvoice) As Integer
      Dim i As Integer
      Dim dr As DataRow

      Try
         Dim invID As Integer = frm.dbgEquipment.DataSource.rows(0).item("invoiceid")
         Dim bSplit As Boolean
         Dim cnt As Short = 0

         ' ensure that at least one row is checked
         For i = 0 To frm.dbgEquipment.DataSource.rows.count - 1
            dr = frm.dbgEquipment.DataSource.rows(i)
            If dr("CheckIn") = "true" Then
               cnt += 1
            End If
         Next
         If cnt = 0 Then
            MsgBox("There must be at least one row checked for checkin to work.", MsgBoxStyle.Exclamation)
            Return 0
         End If

         ' now see if any are not checked
         For i = 0 To frm.dbgEquipment.DataSource.rows.count - 1
            dr = frm.dbgEquipment.DataSource.rows(i)
            If dr("CheckIn") = "false" Then
               bSplit = True
               Exit For
            End If
         Next

         If Not bSplit Then Return invID

         Dim sMsg As String
         Dim iRV As Integer
         sMsg = "Checking in only part of the items on the invoice" & Chr(10)
         sMsg &= "will cause the checked in items to be moved to a new" & Chr(10)
         sMsg &= "invoice and the unchecked items to be left on the" & Chr(10)
         sMsg &= "original invoice." & Chr(10)
         sMsg &= "" & Chr(10)
         sMsg &= "Once an invoice is split into two invoices, " & Chr(10)
         sMsg &= "they cannot be put back together.  " & Chr(10) & Chr(10)
         sMsg &= "Are you absolutely sure that you want to check" & Chr(10)
         sMsg &= "in only the items that are checked?" & Chr(10)
         sMsg &= "" & Chr(10)
         sMsg &= "Click Yes to split the invoice or No to cancel." & Chr(10)
         sMsg &= "" & Chr(10)
         iRV = MsgBox(sMsg, CType(292, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Splitting Invoice")

         If iRV = 6 Then
            ' Yes Code goes here
         Else
            ' No code goes here
            Return 0
         End If
         ' we found at least one unchecked invoice,so
         ' we must split the invoice into two
         ' first, we need a new invoice number
         Dim sql As String = "select max(invoiceid) from invoices"
         Dim oDA As New CDataAccess()
         Dim dt As New DataTable("dt")
         If oDA.SendQuery(sql, dt, ConnectString) = 0 Then
            Throw New System.Exception("can't read from invoices")
         End If

         Dim newInvID As Integer = dt.Rows(0).Item(0) + 1

         ' now insert a new invoice with data from the old
         ' invID
         sql = "select * from invoices where invoiceid = " & invID.ToString
         dt = New DataTable("dt")
         If oDA.SendQuery(sql, dt, ConnectString) = 0 Then
            Throw New System.Exception("Can't read invoice: " & invID.ToString)
         End If

         ' tell user what we are doing
         Dim s As String = "Old Invoice is " & invID.ToString & Chr(10)
         s &= "New Invoice is " & newInvID.ToString
         MsgBox(s, MsgBoxStyle.Information)

         sql = "Insert into invoices "
         sql &= " (invoiceid,customerid,status,invoicedate, "
         sql &= "ponumber,ckcardnumber,contactname,paidoption,shiptocustomer,"
         sql &= "shiptoaddress,shiptocity,shiptostate,shiptozip, "
         sql &= "balancedue,notes,exp_month,exp_yr,card_id, "
         sql &= "check_out_employee,check_in_employee,drivers_license,elapsed_time) "
         sql &= " select "
         sql &= newInvID.ToString & ",customerid,status,invoicedate, "
         sql &= "ponumber,ckcardnumber,contactname,paidoption,shiptocustomer,"
         sql &= "shiptoaddress,shiptocity,shiptostate,shiptozip, "
         sql &= "balancedue,notes,exp_month,exp_yr,card_id, "
         sql &= "check_out_employee,check_in_employee,drivers_license,elapsed_time "
         sql &= "from invoices where invoiceid= " & invID.ToString
         Dim serr As String = String.Empty
         If oDA.SendActionSql(sql, ConnectString, serr) <> 1 Then
            Throw New System.Exception("Failed to insert new split invoice: " & newInvID.ToString)
         End If

         ' now move the invoice details to be checked in
         ' to the new invoice number
         For i = 0 To frm.dbgEquipment.DataSource.rows.count - 1
            dr = frm.dbgEquipment.DataSource.rows(i)
            If dr("CheckIn") = "true" Then
               ' equip is to be checked in, move to new invoice
               sql = "update invoice_details set invoiceid = " & newInvID.ToString & " "
               sql &= "where invoiceid = " & invID.ToString & " "
               sql &= "and equip_id = '" & dr("equip_id") & "' "
               If oDA.SendActionSql(sql, ConnectString, serr) = 0 Then
                  Throw New Exception("failed to move invoice detail record for equip: " & dr("eauip_id"))
               End If
            End If
         Next

         ' now we need to put references to the other invoice in the 
         ' two invoice note fields
         ' insert ref to new in old
         sql = "update invoices set notes = notes & " & vbCrLf
         sql &= "'See Invoice: " & newInvID.ToString & "' " & vbCrLf
         sql &= "where invoiceid=" & invID.ToString
         If oDA.SendActionSql(sql, ConnectString, serr) <> 1 Then
            Throw New System.Exception("Failed to insert reference to new invoice in old")
         End If
         ' set ref to old in new
         sql = "update invoices set notes = notes & " & vbCrLf
         sql &= "'See Invoice: " & invID.ToString & "' " & vbCrLf
         sql &= "where invoiceid=" & newInvID.ToString
         If oDA.SendActionSql(sql, ConnectString, serr) <> 1 Then
            Throw New System.Exception("Failed to insert reference to new invoice in old")
         End If
         Return newInvID
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

#End Region



End Class
