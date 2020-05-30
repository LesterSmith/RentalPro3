Public Class CPrintInvoices
   Private m_Preview As Boolean
#Region " Public Methods "
   Public Sub PrintSelectedInvoices(ByRef oFrm As frmSelectInvoicesToPrint)
      ' Print all invoices checked in the passed form's grid.
      ' loop thru the details, total the line items and 
      ' put the totals in vars for later appending to the stringbuilder

      Dim ps As New System.Text.StringBuilder()
      Dim decEP As Decimal
      Dim i As Integer
      Dim dt As New DataTable()
      Dim j As Integer
      Dim SQL As String
      Dim oDA As New CDataAccess()
      Dim dtInv As New DataTable()
      Dim dtDet As New DataTable()
      Dim dtCust As New DataTable()
      Dim oCP As New CPioneerPrint()
      Dim decTotal As Decimal
      Dim sName As String
      Dim deposit As Decimal
      Dim balanceDue As Decimal
      Dim itemTotal As Decimal
      Dim salesTax As Decimal
      Dim amtPaid As Decimal
      Dim laborCost As Decimal
      Dim delivery As Decimal
      Dim pickup As Decimal
      Dim refund As Decimal
      Dim creditMemo As Decimal
      Dim debitMemo As Decimal
      Dim debitNotes As String
      Dim creditNotes As String
      Dim smemo As String
      Dim k As Integer

      Try
         ' get the customer header data.
         Dim CustId As Integer = oFrm.lvInvoices.SelectedItems(0).SubItems(1).Text
         SQL = "select * from customers where customerid = " & CustId
         If oDA.SendQuery(SQL, dtCust, ConnectString) = 0 Then
            Throw New System.Exception("Database error, can't read customer data.")
         End If
         With oCP
            .ComapanyName = IIf(IsDBNull(dtCust.Rows(0).Item("companyname")), "", dtCust.Rows(0).Item("companyname"))
            .BillAddress1 = IIf(IsDBNull(dtCust.Rows(0).Item("billingaddress1")), "", dtCust.Rows(0).Item("billingaddress1"))
            .BillingCity = IIf(IsDBNull(dtCust.Rows(0).Item("city")), "", dtCust.Rows(0).Item("city"))
            .BillingState = IIf(IsDBNull(dtCust.Rows(0).Item("state")), "", dtCust.Rows(0).Item("state"))
            .BillingZip = IIf(IsDBNull(dtCust.Rows(0).Item("postalcode")), "", dtCust.Rows(0).Item("postalcode"))
            .ContactName = IIf(IsDBNull(dtCust.Rows(0).Item("contactname")), "", dtCust.Rows(0).Item("contactname"))
            .CustomerID = IIf(IsDBNull(dtCust.Rows(0).Item("customerid")), "", dtCust.Rows(0).Item("customerid"))
            .TaxId = IIf(IsDBNull(dtCust.Rows(0).Item("tax_id")), "", dtCust.Rows(0).Item("tax_id"))
         End With

         For k = 0 To oFrm.dtInv.Rows.Count - 1
            If oFrm.dtInv.Rows(k).Item("Print") = "true" Then
               ' get the invoice header data
               SQL = "select * from invoices where customerid = " & CustId & " "
               SQL &= "and invoiceid = " & oFrm.dtInv.Rows(k).Item("invoiceid") & " "

               dtInv.Reset()
               If oDA.SendQuery(SQL, dtInv, ConnectString) = 0 Then
                  MsgBox("No invoices for selected customer and options.", MsgBoxStyle.Information)
                  Exit Sub
               End If

               For i = 0 To dtInv.Rows.Count - 1
                  ' loop thru the invoice headers
                  With dtInv.Rows(i)
                     oCP.ShipToName = .Item("shiptocustomer")
                     oCP.ShipToAddress1 = .Item("shiptoaddress")
                     oCP.ShipToCity = .Item("shiptocity")
                     oCP.ShipToZip = .Item("shiptozip")
                     oCP.ShipToState = .Item("shiptostate")
                     oCP.InvoiceType = "Reprint"
                     oCP.PONbr = .Item("ponumber")
                     oCP.CheckNumber = .Item("ckcardnumber")
                     oCP.PaidOption = .Item("paidoption")
                     If oCP.PaidOption = "LC" Or oCP.PaidOption = "CC" Then
                        oCP.CheckNumber = oCP.HideCCNumber(oCP.CheckNumber)
                     End If
                     oCP.InvoiceId = .Item("invoiceid")
                     smemo = IIf(IsDBNull(.Item("notes")), "", .Item("notes"))
                     If Not IsDBNull(.Item("check_in_employee")) Then
                        oCP.CheckOutEmployee = .Item("check_in_employee")
                     Else
                        oCP.CheckOutEmployee = MNS(.Item("check_out_employee"))
                     End If
                  End With

                  ' now get the invoice details
                  SQL = "select * from invoice_details "
                  SQL &= "where invoiceid = " & dtInv.Rows(i).Item("invoiceid") & " "
                  SQL &= "order by record_type "

                  Const DTSP = 70
                  dt.Reset()
                  Dim iRows As Integer = oDA.SendQuery(SQL, dt, ConnectString)
                  If iRows > 0 Then
                     Dim bTotal As Boolean = False
                     For j = 0 To dt.Rows.Count - 1
                        With dt.Rows(j)

                           ' as we loop thru the items, total all the 15's
                           ' total the labor cost so it can be subtracted for tax purposes
                           ' place the other items in the vars
                           ' finally print the total items...
                           Select Case .Item("record_type")
                              Case 15 ' equip item
                                 ' qty
                                 ps.Append(CType(.Item("Quantity"), String).PadLeft(4))
                                 ' skip 3 spaces and print id - name
                                 ps.Append(Space(3) & CType(.Item("Equip_Id"), String).PadRight(10) & " - ")
                                 sName = .Item("equip_name")
                                 If sName.Length > 29 Then
                                    sName = sName.Substring(0, 29)
                                 End If
                                 ' qty
                                 ps.Append(sName.PadRight(30))

                                 ' rental period (Daily...)
                                 ps.Append(Space(9) & CType(.Item("Rental_Period"), String).PadRight(10))
                                 ' price per unit
                                 ps.Append(Format(.Item("priceperunit"), "#,##0.00").PadLeft(10))
                                 decEP = .Item("PriceperUnit") * _
                                       .Item("Quantity") ' + _
                                 '      dt.Rows(i).Item("Deposit")
                                 ps.Append(Space(2) & Format(decEP, "$#,##0.00").PadLeft(10) & vbCrLf)

                                 itemTotal += decEP + dt.Rows(i).Item("deposit")
                                 If .Item("equip_id") = "Labor" Then
                                    laborCost += .Item("PriceperUnit") * .Item("Quantity")
                                 End If
                                 decTotal = itemTotal
                                 ' print the meter reading if applicable
                                 ' if meter_required, print the meter reading at checkout
                                 If MNSng(.Item("meter_out")) > 0 Then
                                    ps.Append(Space(7) & "Meter Out: " & Format(.Item("meter_out"), "0.00"))
                                    If MNSng(.Item("meter_in")) > 0 Then
                                       ps.Append(" - Meter In: " & Format(.Item("meter_in"), "0.00"))
                                    End If
                                    ps.Append(vbCrLf)
                                 End If

                                 ' here, the items are in the report
                                 ' total the debits and credits
                              Case 25 ' deposit
                                 decTotal += .Item("deposit")
                                 deposit += .Item("deposit")
                              Case 35 ' tax
                                 If .Item("salestax") > 0 Then
                                    decTotal += .Item("salestax")
                                    salesTax += .Item("salestax")
                                 End If
                              Case 45 ' delivery
                                 decTotal += .Item("delivery")
                                 delivery += .Item("delivery")
                              Case 46 ' pickup
                                 decTotal += .Item("delivery")
                                 pickup += .Item("delivery")
                              Case 55 ' amtpaid
                                 decTotal -= .Item("amtpaid")
                                 amtPaid += .Item("amtpaid")
                              Case 65 ' refund
                                 Dim valu As Decimal = .Item("amtpaid")
                                 decTotal -= (valu * -1)
                                 refund += .Item("amtpaid")
                              Case 66 ' credit memo
                                 decTotal -= .Item("amtpaid")
                                 creditMemo += .Item("amtpaid")
                              Case 67 ' cash on account
                                 decTotal -= .Item("amtpaid")
                                 amtPaid += .Item("amtpaid")
                              Case 68 ' debit memo
                                 decTotal += .Item("amtpaid")
                                 debitMemo += .Item("amtpaid")
                                 'Case 75 ' bal due
                                 'ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & FormatCurrency(.Item("amtpaid")).PadLeft(10) & vbCrLf)
                           End Select
                        End With
                     Next j

                     If smemo.IndexOf("INVOICE IS VOID") = -1 Then
                        ' now print the appropriate totals
                        If itemTotal > 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "Item Total".PadRight(11) & _
                              FormatCurrency(itemTotal).PadLeft(10) & vbCrLf)
                        End If

                        If deposit > 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "Deposit".PadRight(11) & _
                              FormatCurrency(deposit).PadLeft(10) & vbCrLf)
                        End If

                        If delivery > 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "Delivery".PadRight(11) & _
                              FormatCurrency(delivery).PadLeft(10) & vbCrLf)
                        End If

                        If pickup > 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "Pickup".PadRight(11) & _
                              FormatCurrency(pickup).PadLeft(10) & vbCrLf)
                        End If

                        If salesTax > 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "Sales Tax".PadRight(11) & _
                              FormatCurrency(salesTax).PadLeft(10) & vbCrLf)
                        End If

                        If amtPaid > 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "Pd on Acct".PadRight(11) & _
                              FormatCurrency(amtPaid).PadLeft(10) & vbCrLf)
                        End If

                        If refund > 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "Refund".PadRight(11) & _
                              FormatCurrency(refund).PadLeft(10) & vbCrLf)
                        End If

                        If creditMemo <> 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "CR Memo".PadRight(11) & _
                              FormatCurrency(creditMemo).PadLeft(10) & vbCrLf)

                        End If
                        If debitMemo <> 0 Then
                           ps.Append(vbCrLf & Space(DTSP) & "DB Memo".PadRight(11) & _
                              FormatCurrency(debitMemo).PadLeft(10) & vbCrLf)
                        End If

                        'If decTotal <> 0 Then
                        ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & _
                           FormatCurrency(decTotal).PadLeft(10) & vbCrLf)
                        'End If

                        If smemo.Length > 0 Then
                           Dim outil As New CUtilities()
                           Dim iNL As Integer = outil.MLCount(smemo, 60)
                           Dim m As Integer
                           ps.Append(vbCrLf & vbCrLf & "Notes:" & vbCrLf)

                           For m = 1 To iNL
                              ps.Append(outil.MemoLine(smemo, 60, m) & vbCrLf)
                           Next
                        End If
                     Else
                        ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & _
                           FormatCurrency(0).PadLeft(10) & vbCrLf)
                        ps.Append(vbCrLf & vbCrLf & "Notes: INVOICE IS VOID" & vbCrLf)
                     End If

                     ' clear the totals for the next invoice, if any
                     decTotal = 0
                     debitMemo = 0
                     creditMemo = 0
                     refund = 0
                     amtPaid = 0
                     salesTax = 0
                     delivery = 0
                     deposit = 0
                     itemTotal = 0
                     delivery = 0
                     deposit = 0

                     oCP.InvoiceDate = dtInv.Rows(i).Item("invoicedate")
                     If m_Preview Then
                        oCP.PrintPreview(ps, dtInv.Rows(i).Item("invoiceid"))
                     Else
                        oCP.StartPrint(ps, dtInv.Rows(i).Item("invoiceid"))
                     End If
                     ps = New System.Text.StringBuilder()
                  End If
               Next i
            End If
         Next k
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region




#Region " Public Properties "
   Public Property Preview() As Boolean
      Get
         Return m_Preview
      End Get
      Set(ByVal Value As Boolean)
         m_Preview = Value
      End Set
   End Property

#End Region

End Class
