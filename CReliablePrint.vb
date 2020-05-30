Imports System
Imports System.Windows.Forms
Namespace CReliablePrint
   Public Class CReliablePrint
#Region " Class Level Variables "
      Private fCkIn As frmCheckinNew
      Private fCkOut As frmCustomers
      Private InvID As Integer

      ' following values come from either frmCustomer or frmCheckin
      Private m_ShipToName As String
      Private m_ShipToAddress1 As String
      Private m_ShipToCity As String
      Private m_ShipToState As String
      Private m_ShipToZip As String
      Private m_ComapanyName As String
      Private m_BillAddress1 As String
      Private m_BillingCity As String
      Private m_BillingState As String
      Private m_BillingZip As String
      Private m_InvoiceType As String
      Private m_CustomerID As String
      Private m_PONbr As String
      Private m_ContactName As String
      Private m_CheckNumber As String
      Private m_PaidOption As String
      Private m_InvoiceId As String
      Private m_InvoiceDate As DateTime
      Private m_TaxId As String
      Private _CheckOutEmployee As String
      Private oDA As New CDataAccess()
      Private m_Preview As Boolean


#End Region

#Region " Public Methods "
      Public Function HideCCNumber(ByVal ccn As String) As String
         Dim len As Short = ccn.Length
         Dim i As Integer
         Dim s As String = ccn
         If len > 4 Then
            For i = 1 To len - 4
               Mid(s, i, 1) = "X"
            Next
            Return s
         Else
            Return ccn
         End If
      End Function

      Public Sub PrintCheckOutInvoice(ByVal InvoiceId As Integer)
         ' format the print line
         Dim oP As New CCheckIn()
         Dim equipID As String
         Dim hrPrice As Decimal
         Dim dayPrice As Decimal
         Dim wkPrice As Decimal
         Dim monPrice As Decimal
         Dim ps As New System.Text.StringBuilder()
         Dim ps2 As New System.Text.StringBuilder()
         Dim decEP As Decimal
         Dim SQL As String
         Dim i As Integer
         Dim dt As New DataTable()
         Dim oUtil As New CUtilities()
         Dim decTotal As Decimal
         Dim sName As String
         Dim dueBack As DateTime
         Dim minPrice As Decimal
         Dim rentalPeriods As String = _
            HALF_DAY & DAILY & WEEKLY & MONTHLY & WEEK_END

         ' get customer data and print

         Try
            Dim checkOutEmployee As String
            Dim trash As String = fCkOut.cbEmployees.Text

            If fCkOut.dtFC.Rows.Count > 0 Then

               ' alternate print using dtfc
               For i = 0 To fCkOut.dtFC.Rows.Count - 1
                  If fCkOut.dtFC.Rows(i).RowState <> DataRowState.Deleted Then
                     Dim dr As DataRow = fCkOut.dtFC.Rows(i)
                     ' qty (itemcount)
                     ps.Append(CType(dr("ItemCount"), String).PadLeft(6) & Space(2))
                     ' skip 3 spaces and print id - name(ItemId)
                     ps.Append(CType(dr("ItemId"), String).PadRight(9) & "- ")
                     ' (itemname)
                     sName = dr("Itemname")
                     If sName.Length > 22 Then
                        sName = sName.Substring(0, 22)
                     End If
                     ' qty
                     ps.Append(sName.PadRight(22))

                     ' price per unit(itemprice)
                     ' print minimum price which is half day or day if not hd price
                     If rentalPeriods.IndexOf(dr("itemperiod")) = -1 Or _
                        (dr("itemid") = RERENT And Not MNB(dr("newprices"))) Then
                        ps.Append(Format(dr("ItemPrice"), "#,##0.00").PadLeft(9) & Space(2))
                     Else
                        equipID = dr("itemid")
                        If MNB(dr("newprices")) Then
                           minPrice = oP.GetPriceFromCheckOutData(dr, HALF_DAY)
                        Else
                           minPrice = oP.GetPriceForEquip(equipID, HALF_DAY)
                        End If
                        ps.Append(Format(minPrice, "#,##0.00").PadLeft(9) & Space(2))
                     End If

                     ' get prices to show
                     If dr("rentorsale") = RENT Or (equipID = RERENT And MNB(dr("newprices"))) Then
                        dueBack = CShare.GetDueBackTime(dr("itemperiod"), dr("itemcount"), Me.m_InvoiceDate)
                        If MNB(dr("newprices")) Then
                           hrPrice = oP.GetPriceFromCheckOutData(dr, HOURLY)
                           dayPrice = oP.GetPriceFromCheckOutData(dr, DAILY)
                           wkPrice = oP.GetPriceFromCheckOutData(dr, WEEKLY)
                           monPrice = oP.GetPriceFromCheckOutData(dr, MONTHLY)
                        Else
                           hrPrice = oP.GetPriceForEquip(equipID, HOURLY)
                           dayPrice = oP.GetPriceForEquip(equipID, DAILY)
                           wkPrice = oP.GetPriceForEquip(equipID, WEEKLY)
                           monPrice = oP.GetPriceForEquip(equipID, MONTHLY)
                        End If
                        ' rental price for hour
                        ps.Append(FormatDollars(hrPrice).PadLeft(6) & Space(1))
                        ' day price
                        ps.Append(FormatDollars(dayPrice).PadLeft(6) & Space(1))
                        ' week price
                        ps.Append(FormatDollars(wkPrice).PadLeft(6) & Space(1))
                        ' month price
                        ps.Append(FormatDollars(monPrice).PadLeft(7) & Space(2))
                     Else
                        ps.Append(Space(30))
                     End If

                     '(itemcount*itemprice)
                     decEP = dr("ItemPrice") * _
                           dr("ItemCount")

                     ps.Append(Format(decEP, "#,##0.00").PadLeft(12) & vbCrLf)

                     ' itemdeposit
                     decTotal += decEP + dr("ItemDeposit")

                     ' if meter_required, print the meter reading at checkout
                     If dr("meter_required") Then
                        ps.Append(Space(9) & "Meter Reading: " & Format(dr("hour_meter"), "0.00") & vbCrLf)
                     End If
                  End If
               Next i
               ' add acknowlegement of conditon and fuel
               ps.Append(vbCrLf & vbCrLf & _
                  "          NO VISIBLE DAMAGE TO EQUIPMENT & FULL OF FUEL __________")
               ps.Append(vbCrLf & _
                  "          LOST OR UNRETURNED KEY CHARGE                 __________")

               ' print the totals
               Const totSP = 9
               Const descSP = 22
               ps2.Append(vbCrLf & "Item Total".PadRight(descSP) & fCkOut.txtItemTotal.Text.PadLeft(totSP) & vbCrLf)
               If UnFormat(fCkOut.txtDeposit.Text) > 0 Then
                  ps2.Append(vbCrLf & "Deposit".PadRight(descSP) & fCkOut.txtDeposit.Text.PadLeft(totSP) & vbCrLf)
               End If

               If UnFormat(fCkOut.txtSalesTax.Text) > 0 Then
                  ps2.Append(vbCrLf & "Sales Tax".PadRight(descSP) & fCkOut.txtSalesTax.Text.PadLeft(totSP) & vbCrLf)
               End If

               If UnFormat(fCkOut.txtDelivery.Text) > 0 Then
                  ps2.Append(vbCrLf & "Delivery".PadRight(descSP) & fCkOut.txtDelivery.Text.PadLeft(totSP) & vbCrLf)
               End If

               ps2.Append(vbCrLf & "Total".PadRight(descSP) & fCkOut.txtTotal.Text.PadLeft(totSP) & vbCrLf)

               ps2.Append(vbCrLf & "Amt Paid".PadRight(descSP) & fCkOut.txtAmtPaid.Text.PadLeft(totSP) & vbCrLf)
               ps2.Append(vbCrLf & "Bal Due".PadRight(descSP) & fCkOut.txtBalDue.Text.PadLeft(totSP) & vbCrLf)


               If fCkOut.txtNotes.Text.Trim.Length > 0 Then
                  Dim sMemo As String = fCkOut.txtNotes.Text
                  Dim iNL As Integer = oUtil.MLCount(sMemo, 60)
                  Dim k As Integer
                  ps.Append(vbCrLf & vbCrLf & "     Notes:" & vbCrLf)

                  For k = 1 To iNL
                     ps.Append(oUtil.MemoLine(sMemo, 60, k) & vbCrLf)
                  Next
               End If


               Dim billTo As String = fCkOut.txtCompanyName.Text & vbCrLf
               billTo &= fCkOut.txtBillingAddress1.Text & vbCrLf
               If fCkOut.txtBillingAddress2.Text.Trim.Length > 0 Then
                  billTo &= fCkOut.txtBillingAddress2.Text & vbCrLf
               End If
               billTo &= fCkOut.txtCity.Text & ", " & fCkOut.txtState.Text & " " & fCkOut.txtPostalCode.Text
               Dim shipTo As String = fCkOut.txtShipToCustomer.Text & vbCrLf & _
                                      fCkOut.txtShipAddress1.Text & vbCrLf & _
                                      fCkOut.txtShipCity.Text & vbCrLf
               Dim invDetails As String = ps.ToString
               Dim totalData As String = ps2.ToString
               Dim invHdrData As String = String.Empty
               Dim invNbr As String = InvoiceId.ToString & "/" & m_InvoiceType
               Dim invOut As String = m_InvoiceDate.ToString
               Dim timeIn As String = "N/A"
               Dim elapsedHours As String = "N/A"
               ' if cash customer put phone number in po number
               Dim poNbr As String
               If fCkOut.chkCashCustomer.Checked Then
                  poNbr = fCkOut.txtPhone.Text
               Else
                  poNbr = fCkOut.txtPONbr.Text.Trim
               End If
               Dim agent As String = Me.m_ContactName
               Dim dlNbr As String = fCkOut.textDriversLicence.Text
               Dim jobPhone As String = fCkOut.CustomerPhone
               Dim jobLocation As String = String.Empty
               Dim ckinBy As String = String.Empty
               Dim taxId As String = Me.m_TaxId
               Dim paidOption As String
               Dim writtenby As String = fCkOut.cbEmployees.Text

               Dim txID As String
               Dim pdOption As String
               If m_TaxId.Trim.Length > 0 Then
                  txID = "Tax ID: " & m_TaxId
               End If
               If m_PaidOption = "BC" Then
                  pdOption = "Blank Check#:  " & m_CheckNumber
               ElseIf m_PaidOption = "LC" Then
                  pdOption = "Left Card#:  " & m_CheckNumber
               ElseIf m_PaidOption = "CK" Then
                  pdOption = "Paid Check#:  " & m_CheckNumber
               ElseIf m_PaidOption = "CC" Then
                  pdOption = "Paid Card #:  " & m_CheckNumber
               ElseIf m_PaidOption = "BT" Then
                  pdOption = "Bill To#:  " & m_CustomerID.ToString    'Format(Val(m_CustomerID), "000000")
               ElseIf m_PaidOption = "CA" Then
                  pdOption = "Paid Cash"
               End If

               Dim oPD As New CRelialblePrintObject(tbnameBillTo:=billTo, tbnameShipTo:=shipTo, _
                            tbnameInvoice:=invNbr, tbnameTimeOut:=invOut, tbnameTimeIn:=timeIn, _
                            tbnameElapsed:=elapsedHours, tbnameJobPhone:=jobPhone, tbnameCheckedInBy:=ckinBy, _
                            tbnameDLNumber:=dlNbr, tbnameAgent:=agent, tbnamePONumber:=poNbr, _
                            tbnameDueInTime:=Format(dueBack, "M/d/yy hh:mm tt"), tbnameJobLocation:=jobLocation, _
                            tbnameTotalDesc:=totalData, tbnameTaxId:=txID, _
                            tbnamePaidOption:=pdOption, tbnameDetailLines:=invDetails, tbnameWrittenby:=writtenby)

               If modMain.fMainForm.mnuPreviewBeforePrint.Checked Then
                  oPD.Preview()
               Else
                  oPD.Print()
               End If
            End If
         Catch ex As System.Exception
            StructuredErrorHandler(ex)
         End Try
      End Sub

      Public Sub PrintCheckInInvoice(ByVal InvoiceId As Integer)
         ' format the print line
         Dim ps As New System.Text.StringBuilder()
         Dim ps2 As New System.Text.StringBuilder()

         Dim decEP As Decimal
         Dim SQL As String
         Dim i As Integer
         Dim dt As New DataTable()
         Dim oUtil As New CUtilities()
         Dim decTotal As Decimal
         Dim sName As String
         Dim dueBack As DateTime
         ' get customer data and print
         Try

            'Dim oP As New CCheckIn()
            Dim equipID As String
            Dim hrPrice As Decimal
            Dim dayPrice As Decimal
            Dim wkPrice As Decimal
            Dim monPrice As Decimal

            For i = 0 To fCkIn.dtList.Rows.Count - 1
               If fCkIn.dtList.Rows(i).RowState <> DataRowState.Deleted Then
                  Dim dri As DataRow = fCkIn.dtList.Rows(i)
                  ' qty
                  If "Daily|Half_Day|Weekly|Monthly|Hourly".IndexOf(MNS(dri("rental_period"))) > -1 Then
                     If IsDBNull(dri("rentalduetoreturn")) Then
                        dueBack = Now
                     Else
                        dueBack = dri("rentalduetoreturn")
                     End If
                  End If
                  ps.Append(CType(dri("Quantity"), String).PadLeft(6) & Space(2))
                  ' skip 3 spaces and print id - name

                  equipID = MNS(dri("Equip_Id"))
                  ps.Append(equipID.PadRight(10) & " - ")
                  sName = dri("equip_name")
                  If sName.Length > 22 Then
                     sName = sName.Substring(0, 22)
                  End If
                  ps.Append(sName.PadRight(22))

                  ' rental price for desired period
                  ' price per unit
                  ps.Append(Format(dri("priceperunit"), "#,##0.00").PadLeft(9) & Space(2))

                  ' on checkin don't list the whatif prices
                  ps.Append(Space(30))

                  decEP = dri("PriceperUnit") * dri("Quantity")
                  ps.Append(Format(decEP, "#,##0.00").PadLeft(10) & vbCrLf)

                  ' if meter_required, print the meter reading at checkout
                  If dri("meterin") > 0 Then
                     ps.Append(Space(9) & "Meter Out: " & Format(dri("meterout"), "0.00") & _
                                          " In: " & Format(dri("meterin"), "0.00") & vbCrLf)
                  End If
               End If
            Next i

            'Const DTSP = 64
            Const totSp = 9
            Const descSP = 22
            If Not fCkIn.voidInvoice Then
               ' print the totals
               ps2.Append(vbCrLf & "Item Total".PadRight(descSP) & fCkIn.txtItemTotal.Text.PadLeft(totSp) & vbCrLf)
               If UnFormat(fCkIn.txtDeposit.Text) <> 0 Then
                  ps2.Append(vbCrLf & "Deposit".PadRight(descSP) & fCkIn.txtDeposit.Text.PadLeft(totSp) & vbCrLf)
               End If

               If UnFormat(fCkIn.txtSalesTax.Text) > 0 Then
                  ps2.Append(vbCrLf & "Sales Tax".PadRight(descSP) & fCkIn.txtSalesTax.Text.PadLeft(totSp) & vbCrLf)
               End If

               If UnFormat(fCkIn.txtDelivery.Text) <> 0 Then
                  ps2.Append(vbCrLf & "Delivery".PadRight(descSP) & fCkIn.txtDelivery.Text.PadLeft(totSp) & vbCrLf)
               End If

               If UnFormat(fCkIn.textManualPickup.Text) <> 0 Then
                  ps2.Append(vbCrLf & "Pickup".PadRight(descSP) & fCkIn.textManualPickup.Text.PadLeft(totSp) & vbCrLf)
               End If
               ps2.Append(vbCrLf & "Total".PadRight(descSP) & fCkIn.txtTotal.Text.PadLeft(totSp) & vbCrLf)
               If UnFormat(fCkIn.txtAmtPaid.Text) <> 0 Then
                  ps2.Append(vbCrLf & "Paid/CkOut".PadRight(descSP) & fCkIn.txtAmtPaid.Text.PadLeft(totSp) & vbCrLf)
               End If

               If UnFormat(fCkIn.txtAmtPaidAtCkIn.Text) > 0 Then
                  ps2.Append(vbCrLf & "Paid/CkIn".PadRight(descSP) & fCkIn.txtAmtPaidAtCkIn.Text.PadLeft(totSp) & vbCrLf)
               End If
               If UnFormat(fCkIn.txtBalDue.Text) < 0 Then
                  Dim valu As Decimal = UnFormat(fCkIn.txtBalDue.Text)
                  ps2.Append(vbCrLf & "Refund Due".PadRight(descSP) & _
                     FormatCurrency(valu * -1).PadLeft(totSp) & vbCrLf)
               Else
                  ps2.Append(vbCrLf & "Bal Due".PadRight(descSP) & fCkIn.txtBalDue.Text.PadLeft(totSp) & vbCrLf)
               End If


               If fCkIn.txtNotes.Text.Trim.Length > 0 Then
                  Dim sMemo As String = fCkIn.txtNotes.Text
                  Dim iNL As Integer = oUtil.MLCount(sMemo, 60)
                  Dim k As Integer
                  ps.Append(vbCrLf & vbCrLf & "Notes:" & vbCrLf)

                  For k = 1 To iNL
                     ps.Append(oUtil.MemoLine(sMemo, 60, k) & vbCrLf)
                  Next
               End If
            Else
               ps.Append(vbCrLf & "Bal Due".PadRight(22) & FormatCurrency(0).PadLeft(12) & vbCrLf)
               ps.Append(vbCrLf & vbCrLf & "  Notes:" & " INVOICE IS VOID" & vbCrLf)
            End If

            ' get invoice header data from checkout
            Dim dti As New DataTable()
            SQL = "select * from invoices where invoiceid = " & fCkIn.txtInvoiceID.Text & ""
            If oDA.SendQuery(SQL, dti, ConnectString) = 0 Then
               MsgBox("Can't read invoice header record.", MsgBoxStyle.Critical)
               Exit Sub
            End If
            Dim dr As DataRow = dti.Rows(0)

            Dim billTo As String = fCkIn.txtCompanyName.Text & vbCrLf
            billTo &= fCkIn.txtBillingAddress1.Text & vbCrLf
            If fCkIn.txtBillingAddress2.Text.Trim.Length > 0 Then
               billTo &= fCkIn.txtBillingAddress2.Text & vbCrLf
            End If
            billTo &= fCkIn.txtCity.Text & ", " & fCkIn.txtState.Text & " " & fCkIn.txtPostalCode.Text
            Dim shipTo As String = fCkIn.txtShipToCustomer.Text & vbCrLf & _
                                   fCkIn.txtShipAddress1.Text & vbCrLf & _
                                   fCkIn.txtShipCity.Text & vbCrLf

            Dim invoice As String = fCkIn.txtInvoiceID.Text & "/Rental Check In"
            Dim timeOut As String = fCkIn.lblCkOutDate.Text
            Dim timeIn As String = fCkIn.lblCkInDate.Text
            'Dim elapsedHours As Integer = _
            '    DateDiff(DateInterval.Hour, _
            '    CType(fCkIn.lblCkOutDate.Text, DateTime), _
            '    CType(fCkIn.lblCkInDate.Text, DateTime))
            Dim elapsedhours As String = modAutoCalc.ElapsedTime
            Dim jobPhone As String = "None"
            Dim dlnbr As String = MNS(fCkIn.textDriversLicence.Text)
            Dim agent As String = MNS(dr("contactname"))
            Dim poNbr As String = MNS(fCkIn.txtPONbr.Text)
            Dim writtenBy As String = MNS(dr("check_out_employee"))
            Dim ckInBy As String
            'Dim trash As String = fCkIn.cbEmployees.Text
            'If PrintInitialsOnly Then
            '   writtenBy = oUtil.GetToken(trash)
            'Else
            '   writtenBy = oUtil.GetToken(trash, " ")
            '   writtenBy = oUtil.GetToken(trash, " ")
            'End If
            'trash = fCkIn.cbEmployees.Text
            'If PrintInitialsOnly Then
            '   ckInBy = oUtil.GetToken(trash)
            'Else
            '   ckInBy = oUtil.GetToken(trash, " ")
            '   ckInBy = oUtil.GetToken(trash, " ")
            'End If
            writtenBy = fCkIn.lblWrittenBy.Text
            ckInBy = fCkIn.cbEmployees.Text
            Dim detailLines As String = ps.ToString
            Dim dueIn As String = ""
            Dim jobLocation As String = ""
            Dim totalDesc As String = ps2.ToString()
            Dim txID As String
            If m_TaxId.Length > 0 Then
               txID = "Tax ID: " & Me.TaxId
            End If
            Dim pdOption As String
            If m_PaidOption = "BC" Then
               pdOption = "Blank Check#:  " & m_CheckNumber & vbCrLf
            ElseIf m_PaidOption = "LC" Then
               pdOption = "Left Card#:  " & m_CheckNumber & vbCrLf
            ElseIf m_PaidOption = "CK" Then
               pdOption = "Paid Check#:  " & m_CheckNumber & vbCrLf
            ElseIf m_PaidOption = "CC" Then
               pdOption = "Paid Card #:  " & m_CheckNumber & vbCrLf
            ElseIf m_PaidOption = "BT" Then
               pdOption = "Bill To#:  " & m_CustomerID.ToString & vbCrLf
            ElseIf m_PaidOption = "CA" Then
               pdOption = "Paid Cash" & vbCrLf
            End If

            Dim oPD As New CRelialblePrintObject(tbnamebillto:=billTo, _
                tbnameshipTo:=shipTo, tbnameInvoice:=invoice, tbnameTimeOut:=timeOut, _
                tbnameTimeIn:=timeIn, tbnameElapsed:=elapsedhours, _
                tbnameJobPhone:=jobPhone, tbnameCheckedInBy:=ckInBy, tbnameDlNumber:=dlnbr, _
                tbnameAgent:=agent, tbnamePONumber:=poNbr, tbnameWrittenBy:=writtenBy, _
                tbnameDetailLines:=detailLines, tbnameDueInTime:=dueBack, tbnameJobLocation:=jobLocation, _
                tbnameTotalDesc:=totalDesc, tbnameTaxId:=txID, tbnamePaidOption:=pdOption)


            If modMain.fMainForm.mnuPreviewBeforePrint.Checked Then
               oPD.Preview()
            Else
               oPD.Print()
            End If
         Catch ex As System.Exception
            StructuredErrorHandler(ex)
         End Try
      End Sub

      Public Sub PrintSelectedInvoices(ByRef oFrm As frmSelectInvoicesToPrint)
         ' Print all invoices checked in the passed form's grid.
         ' loop thru the details, total the line items and 
         ' put the totals in vars for later appending to the stringbuilder

         Dim ps As New System.Text.StringBuilder()
         Dim ps2 As New System.Text.StringBuilder()
         Dim decEP As Decimal
         Dim i As Integer
         Dim dt As New DataTable()
         Dim j As Integer
         Dim SQL As String
         Dim oDA As New CDataAccess()
         Dim dtInv As New DataTable()
         Dim dtDet As New DataTable()
         Dim dtCust As New DataTable()
         Dim oCP As CRelialblePrintObject
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
            Dim CompanyName As String = IIf(IsDBNull(dtCust.Rows(0).Item("companyname")), "", dtCust.Rows(0).Item("companyname"))
            Dim BillAddress1 As String = IIf(IsDBNull(dtCust.Rows(0).Item("billingaddress1")), "", dtCust.Rows(0).Item("billingaddress1"))
            Dim BillingCity As String = IIf(IsDBNull(dtCust.Rows(0).Item("city")), "", dtCust.Rows(0).Item("city"))
            Dim BillingState As String = IIf(IsDBNull(dtCust.Rows(0).Item("state")), "", dtCust.Rows(0).Item("state"))
            Dim BillingZip As String = IIf(IsDBNull(dtCust.Rows(0).Item("postalcode")), "", dtCust.Rows(0).Item("postalcode"))
            Dim ContactName As String = IIf(IsDBNull(dtCust.Rows(0).Item("contactname")), "", dtCust.Rows(0).Item("contactname"))
            Dim CustomerID As Integer = IIf(IsDBNull(dtCust.Rows(0).Item("customerid")), "", dtCust.Rows(0).Item("customerid"))
            Dim TaxId As String = IIf(IsDBNull(dtCust.Rows(0).Item("tax_id")), "", dtCust.Rows(0).Item("tax_id"))
            Dim billTo As String = CompanyName & vbCrLf & _
                                   BillAddress1 & vbCrLf & _
                                   BillingCity & ", " & BillingState & " " & BillingZip & vbCrLf
            Dim ShipToName As String
            Dim ShipToAddress1 As String
            Dim ShipToCity As String
            Dim ShipToZip As String
            Dim ShipToState As String
            Dim InvoiceType As String
            Dim PONbr As String
            Dim CheckNumber As String
            Dim PaidOption As String
            Dim shipTo As String
            Dim agent As String
            Dim timeIN As String
            Dim dr As DataRow
            Dim timeOut As String
            Dim elapsedHours As Single
            Dim dl As String
            Dim writtenBy As String
            Dim checkInBy As String
            Dim detailLines As String
            Dim totalDesc As String
            Dim elapsedTime As String

            For k = 0 To oFrm.dtInv.Rows.Count - 1
               ps.Length = 0
               ps2.Length = 0
               elapsedHours = 0
               elapsedTime = String.Empty
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
                        dr = dtInv.Rows(i)
                        elapsedTime = MNS(dr("elapsed_time"))
                        PaidOption = MNS(dr("paidoption"))
                        ShipToName = dr("shiptocustomer")
                        ShipToAddress1 = dr("shiptoaddress")
                        ShipToCity = dr("shiptocity")
                        ShipToZip = dr("shiptozip")
                        ShipToState = dr("shiptostate")
                        shipTo = ShipToName & vbCrLf & _
                                 ShipToAddress1 & vbCrLf & _
                                 ShipToCity & ", " & ShipToState & " " & ShipToZip & vbCrLf

                        InvoiceType = "Reprint"
                        PONbr = dr("ponumber")
                        writtenBy = dr("check_out_employee")
                        checkInBy = MNS(dr("check_in_employee"))
                        CheckNumber = dr("ckcardnumber")
                        PaidOption = dr("paidoption")
                        If PaidOption = "LC" Or PaidOption = "CC" Then
                           CheckNumber = HideCCNumber(CheckNumber)
                        End If
                        InvoiceId = dr("invoiceid") & "/Reprint"
                        agent = dr("contactname")
                        smemo = IIf(IsDBNull(dr("notes")), "", dr("notes"))
                        If Not IsDBNull(.Item("check_in_employee")) Then
                           CheckOutEmployee = dr("check_in_employee")
                        Else
                           CheckOutEmployee = MNS(dr("check_out_employee"))
                        End If
                        dl = MNS(dr("drivers_license"))
                     End With

                     ' now get the invoice details
                     SQL = "select * from invoice_details "
                     SQL &= "where invoiceid = " & dtInv.Rows(i).Item("invoiceid") & " "
                     SQL &= "order by record_type "

                     Const DTSP = 0 '70
                     dt.Reset()
                     timeIN = FMDNS(Now)

                     Dim iRows As Integer = oDA.SendQuery(SQL, dt, ConnectString)
                     If iRows > 0 Then
                        Dim bTotal As Boolean = False
                        For j = 0 To dt.Rows.Count - 1
                           With dt.Rows(j)
                              dr = dt.Rows(j)

                              ' as we loop thru the items, total all the 15's
                              ' total the labor cost so it can be subtracted for tax purposes
                              ' place the other items in the vars
                              ' finally print the total items...
                              Select Case dr("record_type")
                                 Case 15 ' equip item
                                    ' qty
                                    timeOut = FMDNS(dr("returned_date"))
                                    If elapsedHours = 0 And Not IsDBNull(dr("returned_date")) Then
                                       elapsedHours = DateDiff(DateInterval.Hour, dr("rented_date"), dr("returned_date"))
                                    End If

                                    ps.Append(CType(dr("Quantity"), String).PadLeft(6) & Space(2))
                                    ' skip 3 spaces and print id - name
                                    ps.Append(CType(dr("Equip_Id"), String).PadRight(10) & " - ")
                                    sName = .Item("equip_name")
                                    If sName.Length > 22 Then
                                       sName = sName.Substring(0, 22)
                                    End If
                                    ' qty
                                    ps.Append(sName.PadRight(22))

                                    ' rental price for desired period
                                    ' price per unit
                                    ps.Append(Format(dr("priceperunit"), "#,##0.00").PadLeft(9) & Space(2))

                                    ' on checkin don't list the whatif prices
                                    ps.Append(Space(30))

                                    decEP = dr("PriceperUnit") * dr("Quantity")
                                    ps.Append(Format(decEP, "#,##0.00").PadLeft(10) & vbCrLf)


                                    itemTotal += decEP + dr("deposit")
                                    If dr("equip_id") = "Labor" Then
                                       laborCost += dr("PriceperUnit") * dr("Quantity")
                                    End If
                                    decTotal = itemTotal
                                    ' print the meter reading if applicable
                                    ' if meter_required, print the meter reading at checkout
                                    If dr("meter_in") > 0 Then
                                       ps.Append(Space(9) & "Meter Out: " & Format(dr("meter_out"), "0.00") & _
                                                            " In: " & Format(dr("meter_in"), "0.00") & vbCrLf)
                                       ps.Append(vbCrLf)
                                    End If

                                    ' here, the items are in the report
                                    ' total the debits and credits
                                 Case 25 ' deposit
                                    decTotal += dr("deposit")
                                    deposit += dr("deposit")
                                 Case 35 ' tax
                                    If dr("salestax") > 0 Then
                                       decTotal += dr("salestax")
                                       salesTax += dr("salestax")
                                    End If
                                 Case 45 ' delivery
                                    decTotal += dr("delivery")
                                    delivery += dr("delivery")
                                 Case 46 ' pickup
                                    decTotal += dr("delivery")
                                    pickup += dr("delivery")
                                 Case 55 ' amtpaid
                                    decTotal -= dr("amtpaid")
                                    amtPaid += dr("amtpaid")
                                 Case 65 ' refund
                                    Dim valu As Decimal = dr("amtpaid")
                                    decTotal -= (valu * -1)
                                    refund += dr("amtpaid")
                                 Case 66 ' credit memo
                                    decTotal -= dr("amtpaid")
                                    creditMemo += dr("amtpaid")
                                 Case 67 ' cash on account
                                    decTotal -= dr("amtpaid")
                                    amtPaid += dr("amtpaid")
                                 Case 68 ' debit memo
                                    decTotal += dr("amtpaid")
                                    debitMemo += dr("amtpaid")
                                    'Case 75 ' bal due
                                    'ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & FormatCurrency(.Item("amtpaid")).PadLeft(10) & vbCrLf)
                              End Select
                           End With
                        Next j
                        Const descSP = 22
                        Const totSP = 9
                        If smemo.IndexOf("INVOICE IS VOID") = -1 Then
                           ' now print the appropriate totals
                           If itemTotal > 0 Then
                              ps2.Append(vbCrLf & "Item Total".PadRight(descSP) & _
                                 FormatCurrency(itemTotal).PadLeft(totSP) & vbCrLf)
                           End If

                           If deposit > 0 Then
                              ps2.Append(vbCrLf & Space(DTSP) & "Deposit".PadRight(descSP) & _
                                 FormatCurrency(deposit).PadLeft(totSP) & vbCrLf)
                           End If

                           If delivery > 0 Then
                              ps2.Append(vbCrLf & Space(DTSP) & "Delivery".PadRight(descSP) & _
                                 FormatCurrency(delivery).PadLeft(totSP) & vbCrLf)
                           End If

                           If pickup > 0 Then
                              ps2.Append(vbCrLf & Space(DTSP) & "Pickup".PadRight(descSP) & _
                                 FormatCurrency(pickup).PadLeft(totSP) & vbCrLf)
                           End If

                           If salesTax > 0 Then
                              ps2.Append(vbCrLf & Space(DTSP) & "Sales Tax".PadRight(descSP) & _
                                 FormatCurrency(salesTax).PadLeft(totSP) & vbCrLf)
                           End If

                           If amtPaid > 0 Then
                              ps2.Append(vbCrLf & Space(DTSP) & "Pd on Acct".PadRight(descSP) & _
                                 FormatCurrency(amtPaid).PadLeft(totSP) & vbCrLf)
                           End If

                           If refund > 0 Then
                              ps2.Append(vbCrLf & Space(DTSP) & "Refund".PadRight(descSP) & _
                                 FormatCurrency(refund).PadLeft(totSP) & vbCrLf)
                           End If

                           If creditMemo <> 0 Then
                              ps2.Append(vbCrLf & Space(DTSP) & "CR Memo".PadRight(descSP) & _
                                 FormatCurrency(creditMemo).PadLeft(totSP) & vbCrLf)
                           End If

                           If debitMemo <> 0 Then
                              ps2.Append(vbCrLf & Space(DTSP) & "DB Memo".PadRight(descSP) & _
                                 FormatCurrency(debitMemo).PadLeft(totSP) & vbCrLf)
                           End If

                           ps2.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(descSP) & _
                              FormatCurrency(decTotal).PadLeft(totSP) & vbCrLf)

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
                           ps2.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & _
                              FormatCurrency(0).PadLeft(10) & vbCrLf)
                           ps2.Append(vbCrLf & vbCrLf & "Notes: INVOICE IS VOID" & vbCrLf)
                        End If
                        detailLines = ps.ToString
                        totalDesc = ps2.ToString

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
                        If elapsedTime.Length = 0 Then
                           elapsedTime = Format(elapsedHours, "#,##0") & " Hrs"
                        End If
                        InvoiceDate = dtInv.Rows(i).Item("invoicedate")
                        Dim oPD As New CRelialblePrintObject(tbnamebillto:=billTo, _
                            tbnameshipTo:=shipTo, tbnameInvoice:=InvoiceId, tbnameTimeOut:=timeOut, _
                            tbnameTimeIn:=timeIN, tbnameElapsed:=elapsedTime, _
                            tbnameJobPhone:="N/A", tbnameCheckedInBy:=checkInBy, tbnameDlNumber:=dl, _
                            tbnameAgent:=agent, tbnamePONumber:=PONbr, tbnameWrittenBy:=writtenBy, _
                            tbnameDetailLines:=detailLines, tbnameDueInTime:="N/A", tbnameJobLocation:="N/A", _
                            tbnameTotalDesc:=totalDesc, tbnameTaxId:=String.Empty, tbnamePaidOption:=PaidOption)


                        If modMain.fMainForm.mnuPreviewBeforePrint.Checked Then
                           oPD.Preview()
                        Else
                           oPD.Print()
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




#Region " Constructor "
      Public Sub New()
         ' Caller will set properties prior to calling PrintSelectedInvoices
      End Sub

      Public Sub New(ByRef f As frmCheckinNew)
         fCkIn = f
         m_ShipToName = f.txtShipToCustomer.Text.Trim
         m_ShipToAddress1 = f.txtShipAddress1.Text.Trim
         m_ShipToCity = f.txtShipCity.Text
         m_ShipToState = f.txtShipState.Text
         m_ShipToZip = f.txtShipZip.Text
         m_ComapanyName = f.txtCompanyName.Text.Trim
         m_BillAddress1 = f.txtBillingAddress1.Text.Trim
         m_BillingCity = f.txtCity.Text
         m_BillingState = f.txtState.Text
         m_BillingZip = f.txtPostalCode.Text
         m_InvoiceType = "Rental Check In"
         m_CustomerID = f.txtCustomerID.Text
         m_CustomerID = f.txtCustomerID.Text
         m_PONbr = f.txtPONbr.Text
         m_ContactName = f.txtContactName.Text
         m_CheckNumber = f.txtCheckNumber.Text
         m_TaxId = f.txtTaxID.Text
         If f.optBillTo.Checked Then
            m_PaidOption = "BT"
         ElseIf f.optCash.Checked Then
            m_PaidOption = "CA"
         ElseIf f.optLeftBlankCheck.Checked Then
            m_PaidOption = "BC"
         ElseIf f.optLeftCardNumber.Checked Then
            m_PaidOption = "LC"
            m_CheckNumber = HideCCNumber(m_CheckNumber)
         ElseIf f.optPaidByCheck.Checked Then
            m_PaidOption = "CK"
         ElseIf f.optPaidByCreditCard.Checked Then
            m_PaidOption = "CC"
            m_CheckNumber = HideCCNumber(m_CheckNumber)
         End If
         m_InvoiceDate = Now
      End Sub

      Public Sub New(ByRef f As frmCustomers)

         Try
            fCkOut = f
            m_ShipToName = f.txtShipToCustomer.Text.Trim
            m_ShipToAddress1 = f.txtShipAddress1.Text.Trim
            m_ShipToCity = f.txtShipCity.Text
            m_ShipToState = f.txtShipState.Text
            m_ShipToZip = f.txtShipZip.Text
            m_ComapanyName = f.txtCompanyName.Text.Trim
            m_BillAddress1 = f.txtBillingAddress1.Text.Trim
            m_BillingCity = f.txtCity.Text
            m_BillingState = f.txtState.Text
            m_BillingZip = f.txtPostalCode.Text
            If f.chkCkOutAndIN.Checked Then
               m_InvoiceType = "Check Out & In"
            Else
               m_InvoiceType = "Rental Check Out"
            End If
            m_CustomerID = f.txtCustomerID.Text
            m_CustomerID = f.txtCustomerID.Text
            m_PONbr = f.txtPONbr.Text
            m_ContactName = f.txtContactName.Text
            m_CheckNumber = f.txtCheckNumber.Text
            If f.optBillTo.Checked Then
               m_PaidOption = "BT"
            ElseIf f.optCash.Checked Then
               m_PaidOption = "CA"
            ElseIf f.optLeftBlankCheck.Checked Then
               m_PaidOption = "BC"
            ElseIf f.optLeftCardNumber.Checked Then
               m_PaidOption = "LC"
               m_CheckNumber = HideCCNumber(m_CheckNumber)
            ElseIf f.optPaidByCheck.Checked Then
               m_PaidOption = "CK"
            ElseIf f.optPaidByCreditCard.Checked Then
               m_PaidOption = "CC"
               m_CheckNumber = HideCCNumber(m_CheckNumber)
            End If
            m_InvoiceDate = f.dtpCkOutDateReset.Value.ToString
            m_TaxId = f.txtTaxID.Text
            Me._CheckOutEmployee = f.CheckOutEmployee
         Catch ex As System.Exception
            StructuredErrorHandler(ex)
         End Try
      End Sub

#End Region





      ''' This is a standalone print class.

#Region "Property Methods"
      Public Property ShipToName() As String
         Get
            Return m_ShipToName
         End Get
         Set(ByVal Value As String)
            m_ShipToName = Value
         End Set
      End Property

      Public Property ShipToAddress1() As String
         Get
            Return m_ShipToAddress1
         End Get
         Set(ByVal Value As String)
            m_ShipToAddress1 = Value
         End Set
      End Property

      Public Property ShipToCity() As String
         Get
            Return m_ShipToCity
         End Get
         Set(ByVal Value As String)
            m_ShipToCity = Value
         End Set
      End Property

      Public Property ShipToState() As String
         Get
            Return m_ShipToState
         End Get
         Set(ByVal Value As String)
            m_ShipToState = Value
         End Set
      End Property

      Public Property ShipToZip() As String
         Get
            Return m_ShipToZip
         End Get
         Set(ByVal Value As String)
            m_ShipToZip = Value
         End Set
      End Property

      Public Property ComapanyName() As String
         Get
            Return m_ComapanyName
         End Get
         Set(ByVal Value As String)
            m_ComapanyName = Value
         End Set
      End Property

      Public Property BillAddress1() As String
         Get
            Return m_BillAddress1
         End Get
         Set(ByVal Value As String)
            m_BillAddress1 = Value
         End Set
      End Property

      Public Property BillingCity() As String
         Get
            Return m_BillingCity
         End Get
         Set(ByVal Value As String)
            m_BillingCity = Value
         End Set
      End Property

      Public Property BillingState() As String
         Get
            Return m_BillingState
         End Get
         Set(ByVal Value As String)
            m_BillingState = Value
         End Set
      End Property

      Public Property BillingZip() As String
         Get
            Return m_BillingZip
         End Get
         Set(ByVal Value As String)
            m_BillingZip = Value
         End Set
      End Property

      Public Property InvoiceType() As String
         Get
            Return m_InvoiceType
         End Get
         Set(ByVal Value As String)
            m_InvoiceType = Value
         End Set
      End Property

      Public Property CustomerID() As String
         Get
            Return m_CustomerID
         End Get
         Set(ByVal Value As String)
            m_CustomerID = Value
         End Set
      End Property

      Public Property PONbr() As String
         Get
            Return m_PONbr
         End Get
         Set(ByVal Value As String)
            m_PONbr = Value
         End Set
      End Property

      Public Property ContactName() As String
         Get
            Return m_ContactName
         End Get
         Set(ByVal Value As String)
            m_ContactName = Value
         End Set
      End Property

      Public Property CheckNumber() As String
         Get
            Return m_CheckNumber
         End Get
         Set(ByVal Value As String)
            m_CheckNumber = Value
         End Set
      End Property

      Public Property PaidOption() As String
         Get
            Return m_PaidOption
         End Get
         Set(ByVal Value As String)
            m_PaidOption = Value
         End Set
      End Property

      Public Property InvoiceId() As String
         Get
            Return m_InvoiceId
         End Get
         Set(ByVal Value As String)
            m_InvoiceId = Value
         End Set
      End Property

      Public Property InvoiceDate() As DateTime
         Get
            Return m_InvoiceDate
         End Get
         Set(ByVal Value As DateTime)
            m_InvoiceDate = Value
         End Set
      End Property

      Public Property TaxId() As String
         Get
            Return m_TaxId
         End Get
         Set(ByVal Value As String)
            m_TaxId = Value
         End Set
      End Property

      Public Property CheckOutEmployee() As String
         Get
            Return _CheckOutEmployee
         End Get
         Set(ByVal Value As String)
            _CheckOutEmployee = Value
         End Set
      End Property
#End Region

   End Class
   ''' The constructor accepts the text for the print objects,
   ''' Once instantiated,
   ''' Simply call the Print or Preview method.
   Public Class CRelialblePrintObject
      Dim previewDialog As New PrintPreviewDialog()
      WithEvents PrintDoc As Printing.PrintDocument

#Region " Constructor "

      Public Sub New( _
         ByVal tbnameBillTo As String, _
         ByVal tbnameShipTo As String, _
         ByVal tbnameInvoice As String, _
         ByVal tbnameTimeOut As String, _
         ByVal tbnameTimeIn As String, _
         ByVal tbnameElapsed As String, _
         ByVal tbnameJobPhone As String, _
         ByVal tbnameCheckedInBy As String, _
         ByVal tbnameDLNumber As String, _
         ByVal tbnameAgent As String, _
         ByVal tbnamePONumber As String, _
         ByVal tbnameDueInTime As String, _
         ByVal tbnameJobLocation As String, _
         ByVal tbnameTotalDesc As String, _
         ByVal tbnameTaxid As String, _
         ByVal tbnamePaidOption As String, _
         ByVal tbnameDetailLines As String, _
         ByVal tbnameWrittenBy As String)
         AddObjectToList("Courier New", _
            11.25, _
            106.462472088623, _
            125.249989306641, _
            100.199973730469, _
            100.199991445312, _
            tbnameBillTo, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            106.462472088623, _
            237.974979682617, _
            100.199973730469, _
            100.199991445312, _
            tbnameShipTo, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            588.674845666504, _
            87.6749925146484, _
            100.199973730469, _
            100.199991445312, _
            tbnameInvoice, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            588.674845666504, _
            162.824986098633, _
            100.199973730469, _
            100.199991445312, _
            tbnameTimeOut, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            588.674845666504, _
            125.249989306641, _
            100.199973730469, _
            100.199991445312, _
            tbnameTimeIn, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            588.674845666504, _
            187.874983959961, _
            100.199973730469, _
            100.199991445312, _
            tbnameElapsed, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            12.5249967163086, _
            356.962469523926, _
            100.199973730469, _
            100.199991445312, _
            tbnameJobPhone, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            200.399947460937, _
            325.649972197266, _
            100.199973730469, _
            100.199991445312, _
            tbnameCheckedInBy, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            200.399947460937, _
            356.962469523926, _
            100.199973730469, _
            100.199991445312, _
            tbnameDLNumber, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            407.062393280029, _
            325.649972197266, _
            100.199973730469, _
            100.199991445312, _
            tbnameAgent, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            407.062393280029, _
            356.962469523926, _
            100.199973730469, _
            100.199991445312, _
            tbnamePONumber, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            594.937344024658, _
            356.962469523926, _
            100.199973730469, _
            100.199991445312, _
            tbnameDueInTime, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            607.462340740967, _
            325.649972197266, _
            100.199973730469, _
            100.199991445312, _
            tbnameJobLocation, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            9.75, _
            538.574767181923, _
            814.124899833984, _
            100.199956685009, _
            100.199987671875, _
            tbnameTotalDesc, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            576.149848950195, _
            244.237479147949, _
            100.199973730469, _
            100.199991445312, _
            tbnameTaxid, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            576.149848950195, _
            275.549976474609, _
            100.199973730469, _
            100.199991445312, _
            tbnamePaidOption, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            9.75, _
            12.5249945856261, _
            413.324949146484, _
            100.199956685009, _
            100.199987671875, _
            tbnameDetailLines, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")
         AddObjectToList("Courier New", _
            11.25, _
            12.5249967163086, _
            325.649972197266, _
            100.199973730469, _
            100.199991445312, _
            tbnameWrittenBy, _
            False, _
            False, _
            False, _
            "Text", _
            1, _
            "Black")


         PrintDoc = New Printing.PrintDocument()
      End Sub

#End Region


#Region " Public Methods "
      Public Sub Print()
         'Dim PD As New Printing.PrintDocument()
         'PD.
         'PD.Print()
         PrintDoc.DocumentName = "Reliable Rental"
         PrintDoc.Print()
      End Sub

      Public Sub Preview()
         PrintDoc.DocumentName = "Reliable Rental"
         previewDialog.Document = PrintDoc
         previewDialog.ShowDialog()
         PrintDoc.Dispose()
         previewDialog.Dispose()
      End Sub


      Public Sub AddObjectToList(ByVal FontName As String, _
                                 ByVal FontSize As Single, _
                                 ByVal Left As Single, _
                                 ByVal Top As Single, _
                                 ByVal Height As Single, _
                                 ByVal Width As Single, _
                                 ByVal TextString As String, _
                                 ByVal FontBold As Boolean, _
                                 ByVal FontItalic As Boolean, _
                                 ByVal FontUnderline As Boolean, _
                                 ByVal Box As String, _
                                 ByVal PenWidth As Single, _
                                 ByVal Color As String)
         With printObject
            .FontName = FontName
            .FontSize = FontSize
            .XPos = Left
            .YPos = Top
            .xWidth = Width
            .yHeight = Height
            .Text = TextString
            .FontBold = FontBold
            .FontItalic = FontItalic
            .FontUnderlne = FontUnderline
            .Box = Box
            .PenWidth = PenWidth
            .Color = Color
         End With
         printObjectArray.Add(printObject)
      End Sub


      Public Structure PrintObjects
         Dim FontName As String
         Dim FontSize As Single
         Dim XPos As Single
         Dim YPos As Single
         Dim yHeight As Single
         Dim xWidth As Single
         Dim FontBold As Boolean
         Dim FontUnderlne As Boolean
         Dim FontItalic As Boolean
         Dim Text As String
         Dim Name As String
         Dim Left As Single
         Dim Top As Single
         Dim Height As Single
         Dim Width As Single
         Dim HCI As Single
         Dim VCI As Single
         Dim CharWidth As Single
         Dim LineHeight As Single
         Dim Points As Single
         Dim Box As String
         Dim PenWidth As Single
         Dim Color As String
      End Structure

#End Region
      Public printObjectArray As New ArrayList()
      Private printObject As PrintObjects



#Region " Private Methods "
      Private Sub PrintDoc_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDoc.PrintPage
         Dim i As Integer

         For i = 0 To printObjectArray.Count - 1
            printObject = printObjectArray.Item(i)
            With printObject
               Dim fontAttributes As System.Drawing.FontStyle = _
                 (IIf(printObject.FontBold, 1, 0) Or _
                 IIf(printObject.FontItalic, 2, 0) Or _
                 IIf(printObject.FontUnderlne, 4, 0))
               Dim printFont As Font
               If fontAttributes <> 0 And fontAttributes <> FontStyle.Regular Then
                  printFont = New Font(.FontName, .FontSize, fontAttributes)
               Else
                  printFont = New Font(.FontName, .FontSize)
               End If
               If printObject.Box = "Box" Then
                  Dim c As Color
                  c = Color.FromName(.Color)
                  Dim myPen As New Pen(c, .PenWidth)
                  e.Graphics.DrawRectangle(myPen, printObject.XPos, printObject.YPos, printObject.xWidth, printObject.yHeight)
               ElseIf .Box = "Line" Then
                  Dim c As Color
                  c = Color.FromName(.Color)
                  Dim myPen As New Pen(c, .PenWidth)
                  e.Graphics.DrawLine(myPen, printObject.XPos, printObject.YPos, printObject.XPos + printObject.xWidth, printObject.YPos + printObject.yHeight)
               ElseIf .Box = "Text" Then
                  Dim b As New SolidBrush(Color.FromName(printObject.Color))
                  e.Graphics.DrawString(printObject.Text, _
                                          printFont, _
                                          b, _
                                          printObject.XPos, _
                                          printObject.YPos, _
                                          New StringFormat())
               End If
            End With
         Next i
         e.HasMorePages = False
      End Sub

#End Region


   End Class

End Namespace