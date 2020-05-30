Public Class CPrintSaleTaxReport
#Region " Class Level Variables "
   Private oDA As New CDataAccess()
   Private SQL As String
   Private m_Preview As Boolean
   Private m_StDate As Date
   Private m_EndDate As Date
   Private m_AcctBasis As String


#End Region


#Region " Public Methods "
   Public Sub PrintTaxReport()
      ' print sales tax

      Dim ps As New System.Text.StringBuilder()
      Dim i As Integer
      Dim dt As New DataTable()
      Dim oCP As New CPrintStringNew()
      Dim invTotal As Decimal
      Dim taxTotal As Decimal
      Dim sName As String
      Dim title As String = ReportName
      Dim subTitle As String = _
         "Sales & Tax Report for Invoices Dated: " & _
         Me.StDate.ToShortDateString & " To " & Me.EndDate.ToShortDateString

      Try
         Dim colHdr As String = _
             "Customer Name".PadRight(27) & _
             "Invoice ID".PadRight(11) & _
             "Invoice Date".PadRight(12) & _
             "Tax ID".PadRight(12) & _
             "Invoice Amt".PadLeft(10) & _
             "Sales Tax".PadLeft(10)

         ' if cash basis, only report that portion of the original sale that has
         ' been paid for... 
         ' AmtPaid (sales tax item) is the original amt owed
         ' AmtPaid - BalanceDue (invoice hdr) = amount of orig paid for and reportable
         ' but is only reportable once, so the financing deals will cause a real problem
         ' and should not be done...  No real company finances sales as part of it's business
         ' model.  Any financing would be separate.
         ' if accrual basis, all of amtpaid is reportable
         SQL = "Select  d.invoiceid, d.customer_id, i.invoicedate,d.amtpaid, "
         SQL &= "d.salestax, tax_id,companyname "
         SQL &= "from invoice_details d,customers c, invoices i "
         SQL &= "where i.customerid = c.customerid and d.record_type=35 "
         SQL &= "and d.invoiceid=i.invoiceid "
         SQL &= "and invoicedate >= #" & Me.StDate & "# and invoicedate < #" & DateAdd(DateInterval.Day, 1, Me.EndDate) & "# "
         If m_AcctBasis = "CASH" Then
            SQL &= "and i.status = 'CLOSED' "
         End If
         SQL &= "order by companyname,d.invoiceid"
         If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
            MsgBox("No sales tax records to print for selected range.", MsgBoxStyle.Information)
            Exit Sub
         End If

         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               ps.Append(LS(CType(dt.Rows(i).Item("companyname"), String), 26).PadRight(27))
               ps.Append(CType(.Item("invoiceid"), String).PadRight(12))
               ps.Append(CType(DateValue(.Item("invoicedate")), String).PadRight(12))
               If IsDBNull(.Item("tax_id")) Then
                  ps.Append(Space(12))
               Else
                  ps.Append(CType(.Item("tax_id"), String).PadRight(12))
               End If
               ps.Append(FormatCurrency(.Item("amtpaid")).PadLeft(10))
               ps.Append(FormatCurrency(.Item("salestax")).PadLeft(10) & vbCrLf)
               invTotal += .Item("amtpaid")
               taxTotal += .Item("salestax")
            End With
         Next i

         ps.Append(vbCrLf & Space(27 + 12 + 12) & _
            "Total".PadRight(12) & _
            FormatCurrency(invTotal).PadLeft(10) & _
            FormatCurrency(taxTotal).PadLeft(10) & vbCrLf)

         oCP.TitleFontStyle = "BI"
         oCP.TitleFontSize = REPORT_TITLE_FONT_SIZE
         If Me.Preview Then
            oCP.PrintPreview(80, _
               ps.ToString, _
               ReportName, _
               subTitle, _
               colhdr1:=colHdr)
         Else
            oCP.StartPrint(80, _
               ps.ToString, _
               ReportName, _
               "Invoices Dated: " & Me.StDate.ToString & " To " & Me.EndDate.ToString, _
               colhdr1:=colHdr)
         End If
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

   Public Property StDate() As Date
      Get
         Return m_StDate
      End Get
      Set(ByVal Value As Date)
         m_StDate = Value
      End Set
   End Property

   Public Property EndDate() As Date
      Get
         Return m_EndDate
      End Get
      Set(ByVal Value As Date)
         m_EndDate = Value
      End Set
   End Property

    Public Property AcctBasis As String
        Get
            Return m_AcctBasis
        End Get
        Set(ByVal Value As String)
           m_AcctBasis = Value
        End Set
    End Property


#End Region

End Class
