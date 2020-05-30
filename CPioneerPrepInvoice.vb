Imports System
Imports System.Windows.Forms
Imports System.IO
Namespace CPioneerInvoice
    Public Class CPioneerPrepInvoice
#Region " Class Level Variables "
        Private fCkIn As frmCheckinNew
        Private fCkOut As frmCustomers
        Private InvID As Integer
        Private oDA As New CDataAccess()
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

        Private Function GetCompanyAddress() As String
            Dim dtAddress As New DataTable
            Dim sql2 As String = "select top 1 * from Configuration"
            Dim cda As New CDataAccess()
            cda.SendQuery(sql2, dtAddress, ConnectString)

            Dim dr As DataRow = dtAddress.Rows(0)
            With dr
                Dim address As String = _
                    dr.Item("Corporate_Name").ToString & vbCrLf & _
                    dr.Item("Address1").ToString & vbCrLf
                Dim a2 As String = IIf(String.IsNullOrEmpty(dr.Item("Address2").ToString), String.Empty, dr.Item("Address2").ToString)
                If a2 <> String.Empty Then
                    address &= a2 & vbCrLf
                End If
                address &= dr.Item("City").ToString & ", " & dr.Item("State").ToString & " " & dr.Item("Zip").ToString & vbCrLf
                address &= dr.Item("Phone").ToString & " * " & dr.Item("Fax").ToString
                Return address
            End With

        End Function
        Public Sub PrintCheckOutInvoice(ByVal InvoiceId As Integer)
            ' format the print line
            Dim ps As New System.Text.StringBuilder()
            Dim decEP As Decimal
            Dim SQL As String
            Dim i As Integer
            Dim dt As New DataTable()
            Dim oUtil As New CUtilities()
            Dim decTotal As Decimal
            Dim sName As String
            ' get customer data and print

            Try
                Dim trash As String = fCkOut.cbEmployees.Text
                If PrintInitialsOnly Then
                    CheckOutEmployee = oUtil.GetToken(trash)
                Else
                    CheckOutEmployee = oUtil.GetToken(trash, " ")
                    CheckOutEmployee = oUtil.GetToken(trash, " ")
                End If

                If fCkOut.dtFC.Rows.Count > 0 Then

                    ' alternate print using dtfc
                    For i = 0 To fCkOut.dtFC.Rows.Count - 1
                        With fCkOut.dtFC.Rows(i)
                            ' qty (itemcount)
                            ps.Append(CType(.Item("ItemCount"), String).PadLeft(4) & Space(2))
                            ' skip 3 spaces and print id - name(ItemId)
                            ps.Append(CType(.Item("ItemId"), String).PadRight(10) & " - ")
                            ' (itemname)
                            sName = .Item("Itemname")
                            If sName.Length > 27 Then
                                sName = sName.Substring(0, 27)
                            End If
                            ' qty
                            ps.Append(sName.PadRight(29))

                            ' rental period (Daily...)(Itemperiod)
                            ps.Append(Space(3) & CType(.Item("ItemPeriod"), String).PadRight(9))
                            ' price per unit(itemprice)
                            ps.Append(Format(.Item("ItemPrice"), "#,##0.00").PadLeft(10))
                            '(itemcount*itemprice)
                            decEP = .Item("ItemPrice") * _
                                  .Item("ItemCount") ' + _
                            '      dt.Rows(i).Item("Deposit")
                            ps.Append(Format(decEP, "#,##0.00").PadLeft(10) & vbCrLf)
                            ' itemdeposit
                            decTotal += decEP + .Item("ItemDeposit")

                            ' if meter_required, print the meter reading at checkout
                            If .Item("meter_required") Then
                                ps.Append(Space(9) & "Meter Reading: " & Format(.Item("hour_meter"), "0.00") & vbCrLf)
                            End If
                        End With
                    Next i

                    ' print the totals
                    Const DTSP = 59
                    ps.Append(vbCrLf & Space(DTSP) & "Item Total".PadRight(11) & fCkOut.txtItemTotal.Text.PadLeft(10) & vbCrLf)
                    If UnFormat(fCkOut.txtDeposit.Text) > 0 Then
                        ps.Append(vbCrLf & Space(DTSP) & "Deposit".PadRight(11) & fCkOut.txtDeposit.Text.PadLeft(10) & vbCrLf)
                    End If

                    If UnFormat(fCkOut.txtSalesTax.Text) > 0 Then
                        ps.Append(vbCrLf & Space(DTSP) & "Sales Tax".PadRight(11) & fCkOut.txtSalesTax.Text.PadLeft(10) & vbCrLf)
                    End If

                    If UnFormat(fCkOut.txtDelivery.Text) > 0 Then
                        ps.Append(vbCrLf & Space(DTSP) & "Delivery".PadRight(11) & fCkOut.txtDelivery.Text.PadLeft(10) & vbCrLf)
                    End If

                    ps.Append(vbCrLf & Space(DTSP) & "Total".PadRight(11) & fCkOut.txtTotal.Text.PadLeft(10) & vbCrLf)

                    ps.Append(vbCrLf & Space(DTSP) & "Amt Paid".PadRight(11) & fCkOut.txtAmtPaid.Text.PadLeft(10) & vbCrLf)
                    ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & fCkOut.txtBalDue.Text.PadLeft(10) & vbCrLf)

                    If fCkOut.txtNotes.Text.Trim.Length > 0 Then
                        Dim sMemo As String = fCkOut.txtNotes.Text
                        Dim iNL As Integer = oUtil.MLCount(sMemo, 60)
                        Dim k As Integer
                        ps.Append(vbCrLf & vbCrLf & "Notes:" & vbCrLf)

                        For k = 1 To iNL
                            ps.Append(oUtil.MemoLine(sMemo, 60, k) & vbCrLf)
                        Next
                    End If


                    Dim billTo As String = fCkOut.txtCompanyName.Text & vbCrLf & _
                                      fCkOut.txtBillingAddress1.Text & vbCrLf & _
                                      fCkOut.txtCity.Text & ", " & fCkOut.txtState.Text & " " & fCkOut.txtPostalCode.Text
                    Dim shipTo As String = fCkOut.txtShipToCustomer.Text & vbCrLf & _
                                      fCkOut.txtShipAddress1.Text & vbCrLf & _
                                      fCkOut.txtShipCity.Text & vbCrLf
                    Dim invDetails As String = ps.ToString
                    Dim invHdrData As String = _
                        "Inv Type: " & m_InvoiceType & vbCrLf & _
                        "Inv Date: " & m_InvoiceDate.ToString & vbCrLf
                    If m_InvoiceType = "Reprint" Then
                        invHdrData &= "Printed: " & Now.ToString & vbCrLf
                    End If
                    invHdrData &= "Customer: " & Format(Val(m_CustomerID), "0") & vbCrLf
                    invHdrData &= "Invoice #: " & Format(InvoiceId, "0") & vbCrLf 'Val(fCkOut.txtInvoiceID.Text)
                    invHdrData &= "P.O. #: " & m_PONbr & vbCrLf
                    invHdrData &= "Contact: " & m_ContactName & vbCrLf
                    If m_TaxId.Trim.Length > 0 Then
                        If m_TaxId.Length > 4 Then
                            invHdrData &= "Tax ID: Ending in " & m_TaxId.Substring(m_TaxId.Length - 4, 4) & vbCrLf
                        Else
                            invHdrData &= "Tax ID: " & m_TaxId & vbCrLf
                        End If
                    End If
                    If m_PaidOption = "BC" Then
                        invHdrData &= "Blank Check #:  " & m_CheckNumber & vbCrLf
                    ElseIf m_PaidOption = "LC" Then
                        If m_CheckNumber.Length > 4 Then
                            invHdrData &= "Left Card #:  Ending in" & m_CheckNumber.Substring(m_CheckNumber.Length - 4, 4) & vbCrLf
                        Else
                            invHdrData &= "Left Card #:  " & m_CheckNumber & vbCrLf
                        End If
                    ElseIf m_PaidOption = "CK" Then
                        invHdrData &= "Paid by Check #:  " & m_CheckNumber & vbCrLf
                    ElseIf m_PaidOption = "CC" Then
                        invHdrData &= "Paid by Card #:  " & m_CheckNumber & vbCrLf
                    ElseIf m_PaidOption = "BT" Then
                        invHdrData &= "Bill To #:  " & m_CustomerID.ToString & vbCrLf  'Format(Val(m_CustomerID), "000000")
                    ElseIf m_PaidOption = "CA" Then
                        invHdrData &= "Paid by Cash" & vbCrLf
                    End If
                    If m_InvoiceType = "Rental Check Out" Then
                        invHdrData &= "CkOut Emp: " & Me._CheckOutEmployee & vbCrLf
                    ElseIf m_InvoiceType = "Rental Check In" Then
                        invHdrData &= "CkIn Emp: " & Me._CheckOutEmployee & vbCrLf
                    End If

                    Dim companyLogo As String = ""
                    Dim companyAddress As String = ""
                    Dim title As String = ""

                    If fCkOut.chkPrintToFile.Checked Then 'OrElse modMain.fMainForm.mnuPreviewBeforePrint.Checked Then
                        companyLogo = "Pioneer" & vbCrLf + "Rental"
                        'companyAddress = "1514 Edgefield Rd, Hwy 25" & vbCrLf & _
                        '                 "North Augusta, SC 29860" & vbCrLf & _
                        '                 "803-613-0850 * 803-278-6866 FAX"
                        companyAddress = GetCompanyAddress()
                        title = "SALES CONTRACT/INVOICE"
                    End If

                    Dim printerName As String = String.Empty

                    If fCkOut.chkPrintToFile.Checked Then
                        printerName = "CutePDF Writer"
                    End If

                    'Dim printFileName As String = String.Empty
                    'If fCkOut.chkPrintToFile.Checked AndAlso Not String.IsNullOrEmpty(CutePDFFilePath) AndAlso Not String.IsNullOrEmpty(m_ComapanyName) Then
                    '    printFileName = Path.Combine(CutePDFFilePath, m_ComapanyName + "_" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".pdf")
                    'End If

                    Dim oPD As New PioneerNewPrintObject(companyLogo, companyAddress, title, shipTo, billTo, invHdrData, invDetails, printerName)

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
            Dim checkOutEm
            ' get customer data and print
            Try
                If fCkIn.dtList.Rows.Count > 0 Then
                    Dim trash As String = fCkIn.cbEmployees.Text
                    If PrintInitialsOnly Then
                        CheckOutEmployee = oUtil.GetToken(trash)
                    Else
                        CheckOutEmployee = oUtil.GetToken(trash, " ")
                        CheckOutEmployee = oUtil.GetToken(trash, " ")
                    End If

                    For i = 0 To fCkIn.dtList.Rows.Count - 1
                        With fCkIn.dtList.Rows(i)
                            ' qty
                            ps.Append(CType(.Item("Quantity"), String).PadLeft(4) & Space(2))
                            ' skip 3 spaces and print id - name
                            ps.Append(CType(.Item("Equip_Id"), String).PadRight(10) & " - ")
                            sName = .Item("equip_name")
                            If sName.Length > 27 Then
                                sName = sName.Substring(0, 27)
                            End If
                            ' qty
                            ps.Append(sName.PadRight(29))

                            ' rental period (Daily...)
                            ps.Append(Space(3) & CType(.Item("Rental_Period"), String).PadRight(9))
                            ' price per unit
                            ps.Append(Format(.Item("priceperunit"), "#,##0.00").PadLeft(10))
                            decEP = .Item("PriceperUnit") * _
                                   .Item("Quantity") ' + _
                            ps.Append(Format(decEP, "#,##0.00").PadLeft(10) & vbCrLf)

                            ' if meter_required, print the meter reading at checkout
                            If .Item("meterin") > 0 Then
                                ps.Append(Space(9) & "Meter Out: " & Format(.Item("meterout"), "0.00") & _
                                                     " In: " & Format(.Item("meterin"), "0.00") & vbCrLf)
                            End If

                        End With
                    Next i

                    Const DTSP = 59
                    If Not fCkIn.voidInvoice Then
                        ' print the totals
                        ps.Append(vbCrLf & Space(DTSP) & "Item Total".PadRight(11) & fCkIn.txtItemTotal.Text.PadLeft(10) & vbCrLf)
                        If UnFormat(fCkIn.txtDeposit.Text) <> 0 Then
                            ps.Append(vbCrLf & Space(DTSP) & "Deposit".PadRight(11) & fCkIn.txtDeposit.Text.PadLeft(10) & vbCrLf)
                        End If

                        If UnFormat(fCkIn.txtSalesTax.Text) > 0 Then
                            ps.Append(vbCrLf & Space(DTSP) & "Sales Tax".PadRight(11) & fCkIn.txtSalesTax.Text.PadLeft(10) & vbCrLf)
                        End If

                        If UnFormat(fCkIn.txtDelivery.Text) <> 0 Then
                            ps.Append(vbCrLf & Space(DTSP) & "Delivery".PadRight(11) & fCkIn.txtDelivery.Text.PadLeft(10) & vbCrLf)
                        End If

                        If UnFormat(fCkIn.textManualPickup.Text) <> 0 Then
                            ps.Append(vbCrLf & Space(DTSP) & "Pickup".PadRight(11) & fCkIn.textManualPickup.Text.PadLeft(10) & vbCrLf)
                        End If
                        ps.Append(vbCrLf & Space(DTSP) & "Total".PadRight(11) & fCkIn.txtTotal.Text.PadLeft(10) & vbCrLf)
                        If UnFormat(fCkIn.txtAmtPaid.Text) <> 0 Then
                            ps.Append(vbCrLf & Space(DTSP) & "Paid/CkOut".PadRight(11) & fCkIn.txtAmtPaid.Text.PadLeft(10) & vbCrLf)
                        End If

                        If UnFormat(fCkIn.txtAmtPaidAtCkIn.Text) > 0 Then
                            ps.Append(vbCrLf & Space(DTSP) & "Paid/CkIn".PadRight(11) & fCkIn.txtAmtPaidAtCkIn.Text.PadLeft(10) & vbCrLf)
                        End If
                        If UnFormat(fCkIn.txtBalDue.Text) < 0 Then
                            Dim valu As Decimal = UnFormat(fCkIn.txtBalDue.Text)
                            ps.Append(vbCrLf & Space(DTSP) & "Refund Due".PadRight(11) & _
                               FormatCurrency(valu * -1).PadLeft(10) & vbCrLf)
                        Else
                            ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & fCkIn.txtBalDue.Text.PadLeft(10) & vbCrLf)
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
                        ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & FormatCurrency(0).PadLeft(10) & vbCrLf)
                        ps.Append(vbCrLf & vbCrLf & "Notes:" & " INVOICE IS VOID" & vbCrLf)
                    End If

                    ' get invoice header data from checkout
                    Dim dti As New DataTable()
                    SQL = "select * from invoices where invoiceid = " & InvoiceId & ""
                    If oDA.SendQuery(SQL, dti, ConnectString) = 0 Then
                        MsgBox("Can't read invoice header record.", MsgBoxStyle.Critical)
                        Exit Sub
                    End If
                    Dim dr As DataRow = dti.Rows(0)

                    Dim billTo As String = fCkIn.txtCompanyName.Text & vbCrLf & _
                                           fCkIn.txtBillingAddress1.Text & vbCrLf & _
                                           fCkIn.txtCity.Text & ", " & fCkIn.txtState.Text & " " & fCkIn.txtPostalCode.Text
                    Dim shipTo As String = fCkIn.txtShipToCustomer.Text & vbCrLf & _
                                           fCkIn.txtShipAddress1.Text & vbCrLf & _
                                           fCkIn.txtShipCity.Text & vbCrLf
                    Dim invDetails As String = ps.ToString
                    Dim invHdrData As String = _
                        "Inv Type: " & m_InvoiceType & vbCrLf & _
                        "Inv Date: " & m_InvoiceDate.ToString & vbCrLf
                    If m_InvoiceType = "Reprint" Then
                        invHdrData &= "Printed: " & Now.ToString & vbCrLf
                    End If
                    invHdrData &= "Customer: " & Format(Val(m_CustomerID), "0") & vbCrLf
                    invHdrData &= "Invoice #: " & Format(Val(fCkIn.txtInvoiceID.Text), "0") & vbCrLf
                    invHdrData &= "P.O. #: " & m_PONbr & vbCrLf
                    invHdrData &= "Contact: " & m_ContactName & vbCrLf
                    If m_TaxId.Trim.Length > 4 Then
                        invHdrData &= "Tax ID: Ending in " & m_TaxId.Substring(m_TaxId.Length - 4, 4) & vbCrLf
                    End If
                    If m_PaidOption = "BC" Then
                        invHdrData &= "Blank Check #:  " & m_CheckNumber & vbCrLf
                    ElseIf m_PaidOption = "LC" Then
                        If m_CheckNumber.Length > 4 Then
                            invHdrData &= "Left Card #:  Ending in" & m_CheckNumber.Substring(m_CheckNumber.Length - 4, 4) & vbCrLf
                        End If
                    ElseIf m_PaidOption = "CK" Then
                        invHdrData &= "Paid by Check #:  " & m_CheckNumber & vbCrLf
                    ElseIf m_PaidOption = "CC" Then
                        invHdrData &= "Paid by Card #:  " & m_CheckNumber & vbCrLf
                    ElseIf m_PaidOption = "BT" Then
                        invHdrData &= "Bill To #:  " & m_CustomerID.ToString & vbCrLf  'Format(Val(m_CustomerID), "000000")
                    ElseIf m_PaidOption = "CA" Then
                        invHdrData &= "Paid by Cash" & vbCrLf
                    End If
                    If m_InvoiceType = "Rental Check Out" Then
                        invHdrData &= "CkOut Emp: " & Me._CheckOutEmployee & vbCrLf
                    ElseIf m_InvoiceType = "Rental Check In" Then
                        invHdrData &= "CkIn Emp: " & Me._CheckOutEmployee & vbCrLf
                    End If

                    Dim companyLogo As String = ""
                    Dim companyAddress As String = ""
                    Dim title As String = ""
                    If fCkIn.chkPrintToFile.Checked Then 'OrElse modMain.fMainForm.mnuPreviewBeforePrint.Checked Then
                        companyLogo = "Pioneer" & vbCrLf + "Rental"
                        'companyAddress = "1514 Edgefield Rd, Hwy 25" & vbCrLf & _
                        '                 "North Augusta, SC 29860" & vbCrLf & _
                        '                 "803-613-0850 * 803-278-6866 FAX"
                        companyAddress = GetCompanyAddress()
                        title = "SALES CONTRACT/INVOICE"
                    End If

                    Dim printerName As String = String.Empty
                    If fCkIn.chkPrintToFile.Checked Then
                        printerName = "CutePDF Writer"
                    End If

                    'Dim printFileName As String = String.Empty
                    'If fCkIn.chkPrintToFile.Checked AndAlso Not String.IsNullOrEmpty(CutePDFFilePath) AndAlso Not String.IsNullOrEmpty(m_ComapanyName) Then
                    '    printFileName = Path.Combine(CutePDFFilePath, m_ComapanyName + "_" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".pdf")
                    'End If

                    Dim oPD As New PioneerNewPrintObject(companyLogo, companyAddress, title, shipTo, billTo, invHdrData, invDetails, printerName)
                    oPD.PrinterName = printerName

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
            CustomerNameForInvoiceFile = CleanFileName(m_ComapanyName)
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
                CustomerNameForInvoiceFile = CleanFileName(m_ComapanyName)
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
                m_InvoiceDate = Now
                m_TaxId = f.txtTaxID.Text
                Me._CheckOutEmployee = f.CheckOutEmployee
            Catch ex As System.Exception
                StructuredErrorHandler(ex)
            End Try
        End Sub

#End Region

#Region " Propertys "
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
#Region " Private Methods "
        Private Function CleanFileName(ByVal filename As String) As String
            Return String.Join("_", filename.Split(Path.GetInvalidFileNameChars()))
        End Function

#End Region
    End Class
    ' ''' This is a standalone print class.
    ' ''' The constructor accepts the text for the print objects,
    ' ''' Once instantiated,
    ' ''' Simply call the Print or Preview method.
    'Public Class PioneerNewPrintObject
    '    Dim previewDialog As New PrintPreviewDialog
    '    WithEvents PrintDoc As Printing.PrintDocument
    '    Public Sub New( _
    '       ByVal tbnameLogo As String, _
    '       ByVal tbnameCompanyName As String, _
    '       ByVal tbnameTitle As String, _
    '       ByVal tbnameShipTo As String, _
    '       ByVal tbnameBillTo As String, _
    '       ByVal tbnameInvoiceHdrData As String, _
    '       ByVal tbnameDetails As String)
    '        AddObjectToList("Courier New", _
    '           18, _
    '           25.0499989270663, _
    '           25.0499978298926, _
    '           150.299993562398, _
    '           62.6249945747316, _
    '           tbnameLogo, _
    '           True, _
    '           True, _
    '           False, _
    '           "Text", _
    '           3, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           12, _
    '           212.925008448929, _
    '           25.0500003455172, _
    '           300.600011927899, _
    '           62.6250008637929, _
    '           tbnameCompanyName, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           12, _
    '           576.150022861807, _
    '           25.0500003455172, _
    '           200.400007951933, _
    '           50.1000006910343, _
    '           tbnameTitle, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           11.25, _
    '           37.5749828151941, _
    '           100.199990053375, _
    '           400.799816695404, _
    '           100.199990053375, _
    '           tbnameShipTo, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           11.25, _
    '           37.5749828151941, _
    '           212.924978863422, _
    '           400.799816695404, _
    '           100.199990053375, _
    '           tbnameBillTo, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           11.25, _
    '           488.474776597524, _
    '           112.724988810047, _
    '           325.649851065016, _
    '           112.724988810047, _
    '           tbnameInvoiceHdrData, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           11.25, _
    '           37.5749828151941, _
    '           350.699965186813, _
    '           776.549644847345, _
    '           488.474951510204, _
    '           tbnameDetails, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")

    '        PrintDoc = New Printing.PrintDocument
    '    End Sub


    '    Public Structure PrintObjects
    '        Dim FontName As String
    '        Dim FontSize As Single
    '        Dim XPos As Single
    '        Dim YPos As Single
    '        Dim yHeight As Single
    '        Dim xWidth As Single
    '        Dim FontBold As Boolean
    '        Dim FontUnderlne As Boolean
    '        Dim FontItalic As Boolean
    '        Dim Text As String
    '        Dim Name As String
    '        Dim Left As Single
    '        Dim Top As Single
    '        Dim Height As Single
    '        Dim Width As Single
    '        Dim HCI As Single
    '        Dim VCI As Single
    '        Dim CharWidth As Single
    '        Dim LineHeight As Single
    '        Dim Points As Single
    '        Dim Box As String
    '        Dim PenWidth As Single
    '        Dim Color As String
    '    End Structure
    '    Public printObjectArray As New ArrayList
    '    Private printObject As PrintObjects

    '    Public Sub AddObjectToList(ByVal FontName As String, _
    '                               ByVal FontSize As Single, _
    '                               ByVal Left As Single, _
    '                               ByVal Top As Single, _
    '                               ByVal Height As Single, _
    '                               ByVal Width As Single, _
    '                               ByVal TextString As String, _
    '                               ByVal FontBold As Boolean, _
    '                               ByVal FontItalic As Boolean, _
    '                               ByVal FontUnderline As Boolean, _
    '                               ByVal Box As String, _
    '                               ByVal PenWidth As Single, _
    '                               ByVal Color As String)
    '        With printObject
    '            .FontName = FontName
    '            .FontSize = FontSize
    '            .XPos = Left
    '            .YPos = Top
    '            .xWidth = Width
    '            .yHeight = Height
    '            .Text = TextString
    '            .FontBold = FontBold
    '            .FontItalic = FontItalic
    '            .FontUnderlne = FontUnderline
    '            .Box = Box
    '            .PenWidth = PenWidth
    '            .Color = Color
    '        End With
    '        printObjectArray.Add(printObject)
    '    End Sub
    '    Public Sub Print()
    '        Dim PD As New Printing.PrintDocument
    '        PD.Print()
    '    End Sub
    '    Public Sub Preview()
    '        PrintDoc.DocumentName = "Print Object Test"
    '        previewDialog.Document = PrintDoc
    '        previewDialog.ShowDialog()
    '        PrintDoc.Dispose()
    '        previewDialog.Dispose()
    '    End Sub
    '    Private Sub PrintDoc_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDoc.PrintPage
    '        Dim i As Integer

    '        For i = 0 To printObjectArray.Count - 1
    '            printObject = printObjectArray.Item(i)
    '            With printObject
    '                Dim fontAttributes As System.Drawing.FontStyle = _
    '                  (IIf(printObject.FontBold, 1, 0) Or _
    '                  IIf(printObject.FontItalic, 2, 0) Or _
    '                  IIf(printObject.FontUnderlne, 4, 0))
    '                Dim printFont As Font
    '                If fontAttributes <> 0 And fontAttributes <> FontStyle.Regular Then
    '                    printFont = New Font(.FontName, .FontSize, fontAttributes)
    '                Else
    '                    printFont = New Font(.FontName, .FontSize)
    '                End If
    '                If .Box = "Box" Then
    '                    Dim c As Color
    '                    c = Color.FromName(.Color)
    '                    Dim myPen As New Pen(c, .PenWidth)
    '                    e.Graphics.DrawRectangle(myPen, printObject.XPos, printObject.YPos, printObject.xWidth, printObject.yHeight)
    '                ElseIf .Box = "Line" Then
    '                    Dim c As Color
    '                    c = Color.FromName(.Color)
    '                    Dim myPen As New Pen(c, .PenWidth)
    '                    e.Graphics.DrawLine(myPen, printObject.XPos, printObject.YPos, printObject.XPos + printObject.xWidth, printObject.YPos + printObject.yHeight)
    '                ElseIf .Box = "Text" Then
    '                    Dim b As New SolidBrush(Color.FromName(.Color))
    '                    e.Graphics.DrawString(printObject.Text, _
    '                                          printFont, _
    '                                          b, _
    '                                          .XPos, _
    '                                          .YPos, _
    '                                          New StringFormat)
    '                End If
    '            End With
    '        Next i
    '        e.HasMorePages = False
    '    End Sub
    'End Class

    ' old class
    '    ''' This is a standalone print class.
    '    ''' The constructor accepts the text for the print objects,
    '    ''' Once instantiated,
    '    ''' Simply call the Print or Preview method.
    '    Public Class CPioneerPrintObject
    '        'Private _Left As Single
    '        'Private _Top As Single
    '        'Private _Height As Single
    '        'Private _Width As Single
    '        'Private _FontName As String
    '        'Private _FontSize As Single
    '        'Private _FontBold As Boolean
    '        'Private _FontItalic As Boolean
    '        'Private _FontUnderline As Boolean
    '        'Private _Text As String
    '        'Private _Box As Boolean
    '        'Private _PenWidth As Single
    '        Dim previewDialog As New PrintPreviewDialog()
    '        WithEvents PrintDoc As Printing.PrintDocument
    '        Public Sub New( _
    '            ByVal tbnameLogo As String, _
    '            ByVal tbnameCompanyName As String, _
    '            ByVal tbnameTitle As String, _
    '            ByVal tbnameShipTo As String, _
    '            ByVal tbnameBillTo As String, _
    '            ByVal tbnameInvoiceHdrData As String, _
    '            ByVal tbnameDetails_Total As String)
    '            AddObjectToList("Courier New", _
    '               18, _
    '               25.0499989270663, _
    '               25.0499978298926, _
    '               150.299993562398, _
    '               62.6249945747316, _
    '               tbnameLogo, _
    '               True, _
    '               True, _
    '               False, _
    '               "Text", _
    '               3, _
    '               "Black")
    '            AddObjectToList("Courier New", _
    '               12, _
    '               212.925008448929, _
    '               25.0500003455172, _
    '               300.600011927899, _
    '               62.6250008637929, _
    '               tbnameCompanyName, _
    '               True, _
    '               False, _
    '               False, _
    '               "Text", _
    '               1, _
    '               "Black")
    '            AddObjectToList("Courier New", _
    '               12, _
    '               576.150022861807, _
    '               25.0500003455172, _
    '               200.400007951933, _
    '               50.1000006910343, _
    '               tbnameTitle, _
    '               True, _
    '               False, _
    '               False, _
    '               "Text", _
    '               1, _
    '               "Black")
    '            AddObjectToList("Courier New", _
    '               11.25, _
    '               37.5749828151941, _
    '               100.199990053375, _
    '               400.799816695404, _
    '               100.199990053375, _
    '               tbnameShipTo, _
    '               True, _
    '               False, _
    '               False, _
    '               "Text", _
    '               1, _
    '               "Black")
    '            AddObjectToList("Courier New", _
    '               11.25, _
    '               37.5749828151941, _
    '               212.924978863422, _
    '               400.799816695404, _
    '               100.199990053375, _
    '               tbnameBillTo, _
    '               True, _
    '               False, _
    '               False, _
    '               "Text", _
    '               1, _
    '               "Black")
    '            AddObjectToList("Courier New", _
    '               11.25, _
    '               488.474776597524, _
    '               112.724988810047, _
    '               325.649851065016, _
    '               112.724988810047, _
    '               tbnameInvoiceHdrData, _
    '               True, _
    '               False, _
    '               False, _
    '               "Text", _
    '               1, _
    '               "Black")
    '            AddObjectToList("Courier New", _
    '               11.25, _
    '               37.5749828151941, _
    '               350.699965186813, _
    '               776.549644847345, _
    '               488.474951510204, _
    '               tbnameDetails, _
    '               True, _
    '               False, _
    '               False, _
    '               "Text", _
    '               1, _
    '               "Black")

    '            PrintDoc = New Printing.PrintDocument()
    '        End Sub
    '        Public Structure PrintObjects
    '            Dim FontName As String
    '            Dim FontSize As Single
    '            Dim XPos As Single
    '            Dim YPos As Single
    '            Dim yHeight As Single
    '            Dim xWidth As Single
    '            Dim FontBold As Boolean
    '            Dim FontUnderlne As Boolean
    '            Dim FontItalic As Boolean
    '            Dim Text As String
    '            Dim Name As String
    '            Dim Left As Single
    '            Dim Top As Single
    '            Dim Height As Single
    '            Dim Width As Single
    '            Dim HCI As Single
    '            Dim VCI As Single
    '            Dim CharWidth As Single
    '            Dim LineHeight As Single
    '            Dim Points As Single
    '            Dim Box As Boolean
    '            Dim PenWidth As Single
    '        End Structure
    '        Public printObjectArray As New ArrayList()
    '        Private printObject As PrintObjects

    '        Public Sub AddObjectToList(ByVal FontName As String, _
    '                                   ByVal FontSize As Single, _
    '                                   ByVal Left As Single, _
    '                                   ByVal Top As Single, _
    '                                   ByVal Height As Single, _
    '                                   ByVal Width As Single, _
    '                                   ByVal TextString As String, _
    '                                   ByVal FontBold As Boolean, _
    '                                   ByVal FontItalic As Boolean, _
    '                                   ByVal FontUnderline As Boolean, _
    '                                   ByVal Box As Boolean, _
    '                                   ByVal PenWidth As Single)
    '            With printObject
    '                .FontName = FontName
    '                .FontSize = FontSize
    '                .XPos = Left
    '                .YPos = Top
    '                .xWidth = Width
    '                .yHeight = Height
    '                .Text = TextString
    '                .FontBold = FontBold
    '                .FontItalic = FontItalic
    '                .FontUnderlne = FontUnderline
    '                .Box = Box
    '                .PenWidth = PenWidth
    '            End With
    '            printObjectArray.Add(printObject)
    '        End Sub
    '        Public Sub Print()
    '            Dim PD As New Printing.PrintDocument()
    '            PD.Print()
    '        End Sub
    '        Public Sub Preview()
    '            PrintDoc.DocumentName = "Print Object Test"
    '            previewDialog.Document = PrintDoc
    '            previewDialog.ShowDialog()
    '            PrintDoc.Dispose()
    '            previewDialog.Dispose()
    '        End Sub
    '        Private Sub PrintDoc_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDoc.PrintPage
    '            Dim i As Integer

    '            For i = 0 To printObjectArray.Count - 1
    '                printObject = printObjectArray.Item(i)
    '                With printObject
    '                    Dim fontAttributes As System.Drawing.FontStyle = _
    '                      (IIf(printObject.FontBold, 1, 0) Or _
    '                      IIf(printObject.FontItalic, 2, 0) Or _
    '                      IIf(printObject.FontUnderlne, 4, 0))
    '                    Dim printFont As Font
    '                    If fontAttributes <> 0 And fontAttributes <> FontStyle.Regular Then
    '                        printFont = New Font(.FontName, .FontSize, fontAttributes)
    '                    Else
    '                        printFont = New Font(.FontName, .FontSize)
    '                    End If
    '                    If .Box Then
    '                        Dim myPen As New Pen(Color.Black, .PenWidth)
    '                        e.Graphics.DrawRectangle(myPen, printObject.XPos, printObject.YPos, printObject.xWidth, printObject.yHeight)
    '                    Else
    '                        e.Graphics.DrawString(printObject.Text, _
    '                                              printFont, _
    '                                              Brushes.Black, _
    '                                              .XPos, _
    '                                              .YPos, _
    '                                              New StringFormat())
    '                    End If
    '                End With
    '            Next i
    '            e.HasMorePages = False
    '        End Sub

    '#Region "Object Properties"
    '        'Public Property Left() As Single
    '        '   Get
    '        '      Return _Left
    '        '   End Get
    '        '   Set(ByVal Value As Single)
    '        '      _Left = Value
    '        '   End Set
    '        'End Property
    '        'Public Property Top() As Single
    '        '   Get
    '        '      Return _Top
    '        '   End Get
    '        '   Set(ByVal Value As Single)
    '        '      _Top = Value
    '        '   End Set
    '        'End Property

    '        'Public Property Height() As Single
    '        '   Get
    '        '      Return _Height
    '        '   End Get
    '        '   Set(ByVal Value As Single)
    '        '      _Height = Value
    '        '   End Set
    '        'End Property
    '        'Public Property Width() As Single
    '        '   Get
    '        '      Return _Width
    '        '   End Get
    '        '   Set(ByVal Value As Single)
    '        '      _Width = Value
    '        '   End Set
    '        'End Property

    '        'Public Property FontName() As String
    '        '   Get
    '        '      Return _FontName
    '        '   End Get
    '        '   Set(ByVal Value As String)
    '        '      _FontName = Value
    '        '   End Set
    '        'End Property
    '        'Public Property FontSize() As Single
    '        '   Get
    '        '      Return _FontSize
    '        '   End Get
    '        '   Set(ByVal Value As Single)
    '        '      _FontSize = Value
    '        '   End Set
    '        'End Property
    '        'Public Property FontBold() As Boolean
    '        '   Get
    '        '      Return _FontBold
    '        '   End Get
    '        '   Set(ByVal Value As Boolean)
    '        '      _FontBold = Value
    '        '   End Set
    '        'End Property
    '        'Public Property FontItalic() As Boolean
    '        '   Get
    '        '      Return _FontItalic
    '        '   End Get
    '        '   Set(ByVal Value As Boolean)
    '        '      _FontItalic = Value
    '        '   End Set
    '        'End Property
    '        'Public Property FontUnderline() As Boolean
    '        '   Get
    '        '      Return _FontUnderline
    '        '   End Get
    '        '   Set(ByVal Value As Boolean)
    '        '      _FontUnderline = Value
    '        '   End Set
    '        'End Property

    '        'Public Property Text() As String
    '        '   Get
    '        '      Return _Text
    '        '   End Get
    '        '   Set(ByVal Value As String)
    '        '      _Text = Value
    '        '   End Set
    '        'End Property
    '        'Public Property PenWidth() As Single
    '        '   Get
    '        '      Return _PenWidth
    '        '   End Get
    '        '   Set(ByVal Value As Single)
    '        '      _PenWidth = Value
    '        '   End Set
    '        'End Property
    '#End Region

    '    End Class

    ''' This is a standalone print class.
    ''' The constructor accepts the text for the print objects,
    ''' Once instantiated,
    ''' Simply call the Print or Preview method.
    ''' 
    '*****************************************************************************
    'Public Class PioneerNewPrintObject
    '    Dim previewDialog As New PrintPreviewDialog
    '    WithEvents PrintDoc As Printing.PrintDocument
    '    Public Sub New( _
    '       ByVal tbnameLogo As String, _
    '       ByVal tbnameCompanyName As String, _
    '       ByVal tbnameTitle As String, _
    '       ByVal tbnameShipTo As String, _
    '       ByVal tbnameBillTo As String, _
    '       ByVal tbnameInvoiceHdrData As String, _
    '       ByVal tbnameDetails As String)
    '        AddObjectToList("Courier New", _
    '           18, _
    '           40.6859337567782, _
    '           55.2374921078467, _
    '           165.93592839211, _
    '           92.8124888526857, _
    '           tbnameLogo, _
    '           True, _
    '           True, _
    '           False, _
    '           "Text", _
    '           3, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           12, _
    '           223.348965319961, _
    '           45.1749965308199, _
    '           311.023968798932, _
    '           82.7499970490956, _
    '           tbnameCompanyName, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           12, _
    '           586.57397973284, _
    '           45.1749965308199, _
    '           210.823964822966, _
    '           70.2249968763371, _
    '           tbnameTitle, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           11.25, _
    '           47.3474418453455, _
    '           119.067173738678, _
    '           410.572275725555, _
    '           119.067173738678, _
    '           tbnameShipTo, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           11.25, _
    '           47.3474418453455, _
    '           231.792162548725, _
    '           410.572275725555, _
    '           119.067173738678, _
    '           tbnameBillTo, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           11.25, _
    '           498.247235627675, _
    '           131.59217249535, _
    '           335.422310095167, _
    '           131.59217249535, _
    '           tbnameInvoiceHdrData, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")
    '        AddObjectToList("Courier New", _
    '           11.25, _
    '           47.3474418453455, _
    '           369.567148872116, _
    '           786.322103877497, _
    '           507.342135195507, _
    '           tbnameDetails, _
    '           True, _
    '           False, _
    '           False, _
    '           "Text", _
    '           1, _
    '           "Black")

    '        PrintDoc = New Printing.PrintDocument
    '    End Sub
    '    Public Structure PrintObjects
    '        Dim FontName As String
    '        Dim FontSize As Single
    '        Dim XPos As Single
    '        Dim YPos As Single
    '        Dim yHeight As Single
    '        Dim xWidth As Single
    '        Dim FontBold As Boolean
    '        Dim FontUnderlne As Boolean
    '        Dim FontItalic As Boolean
    '        Dim Text As String
    '        Dim Name As String
    '        Dim Left As Single
    '        Dim Top As Single
    '        Dim Height As Single
    '        Dim Width As Single
    '        Dim HCI As Single
    '        Dim VCI As Single
    '        Dim CharWidth As Single
    '        Dim LineHeight As Single
    '        Dim Points As Single
    '        Dim Box As String
    '        Dim PenWidth As Single
    '        Dim Color As String
    '    End Structure
    '    Public printObjectArray As New ArrayList
    '    Private printObject As PrintObjects

    '    Public Sub AddObjectToList(ByVal FontName As String, _
    '                               ByVal FontSize As Single, _
    '                               ByVal Left As Single, _
    '                               ByVal Top As Single, _
    '                               ByVal Height As Single, _
    '                               ByVal Width As Single, _
    '                               ByVal TextString As String, _
    '                               ByVal FontBold As Boolean, _
    '                               ByVal FontItalic As Boolean, _
    '                               ByVal FontUnderline As Boolean, _
    '                               ByVal Box As String, _
    '                               ByVal PenWidth As Single, _
    '                               ByVal Color As String)
    '        With printObject
    '            .FontName = FontName
    '            .FontSize = FontSize
    '            .XPos = Left
    '            .YPos = Top
    '            .xWidth = Width
    '            .yHeight = Height
    '            .Text = TextString
    '            .FontBold = FontBold
    '            .FontItalic = FontItalic
    '            .FontUnderlne = FontUnderline
    '            .Box = Box
    '            .PenWidth = PenWidth
    '            .Color = Color
    '        End With
    '        printObjectArray.Add(printObject)
    '    End Sub
    '    Public Sub Print()
    '        'Dim PD As New Printing.PrintDocument
    '        PrintDoc.Print()
    '        'PD.Print()
    '    End Sub
    '    Public Sub Preview()
    '        PrintDoc.DocumentName = "Print Object Test"
    '        previewDialog.Document = PrintDoc
    '        previewDialog.ShowDialog()
    '        PrintDoc.Dispose()
    '        previewDialog.Dispose()
    '    End Sub
    '    Private Sub PrintDoc_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDoc.PrintPage
    '        Dim i As Integer

    '        For i = 0 To printObjectArray.Count - 1
    '            printObject = printObjectArray.Item(i)
    '            With printObject
    '                Dim fontAttributes As System.Drawing.FontStyle = _
    '                  (IIf(printObject.FontBold, 1, 0) Or _
    '                  IIf(printObject.FontItalic, 2, 0) Or _
    '                  IIf(printObject.FontUnderlne, 4, 0))
    '                Dim printFont As Font
    '                If fontAttributes <> 0 And fontAttributes <> FontStyle.Regular Then
    '                    printFont = New Font(.FontName, .FontSize, fontAttributes)
    '                Else
    '                    printFont = New Font(.FontName, .FontSize)
    '                End If
    '                If .Box = "Box" Then
    '                    Dim c As Color
    '                    c = Color.FromName(.Color)
    '                    Dim myPen As New Pen(c, .PenWidth)
    '                    e.Graphics.DrawRectangle(myPen, printObject.XPos, printObject.YPos, printObject.xWidth, printObject.yHeight)
    '                ElseIf .Box = "Line" Then
    '                    Dim c As Color
    '                    c = Color.FromName(.Color)
    '                    Dim myPen As New Pen(c, .PenWidth)
    '                    e.Graphics.DrawLine(myPen, printObject.XPos, printObject.YPos, printObject.XPos + printObject.xWidth, printObject.YPos + printObject.yHeight)
    '                ElseIf .Box = "Text" Then
    '                    Dim b As New SolidBrush(Color.FromName(.Color))
    '                    e.Graphics.DrawString(printObject.Text, _
    '                                          printFont, _
    '                                          b, _
    '                                          .XPos, _
    '                                          .YPos, _
    '                                          New StringFormat)
    '                End If
    '            End With
    '        Next i
    '        e.HasMorePages = False
    '    End Sub
    'End Class
    '*****************************************************************************
    Public Class CPrintInvoices
        Private m_Preview As Boolean
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



        Public Property Preview() As Boolean
            Get
                Return m_Preview
            End Get
            Set(ByVal Value As Boolean)
                m_Preview = Value
            End Set
        End Property
    End Class

End Namespace
