<%
Class Payment
  Private m_Id
  Private m_Status
  Private m_Curency
  Private m_Amount
  Private m_PurchaseOrder
  Private m_InvoicePdfUrl
  Private m_InvoiceReceiptUrl

  Public Property Let Id(p_Data)
      m_Id = p_Data
  End Property

  Public Property Get Id()
     Id = m_Id
  End Property

  Public Property Let Status(p_Data)
      m_Status = p_Data
  End Property

  Public Property Get Status()
     Status = m_Status
  End Property

  Public Property Let Curency(p_Data)
      m_Curency = p_Data
  End Property

  Public Property Get Curency()
     Curency = m_Curency
  End Property

  Public Property Let Amount(p_Data)
      m_Amount = p_Data
  End Property

  Public Property Get Amount()
     Amount = m_Amount
  End Property

  Public Property Let PurchaseOrder(p_Data)
      m_PurchaseOrder = p_Data
  End Property

  Public Property Get PurchaseOrder()
     PurchaseOrder = m_PurchaseOrder
  End Property

  Public Property Let InvoicePdfUrl(p_Data)
      m_InvoicePdfUrl = p_Data
  End Property

  Public Property Get InvoicePdfUrl()
     InvoicePdfUrl = m_InvoicePdfUrl
  End Property

  Public Property Let InvoiceReceiptUrl(p_Data)
      m_InvoiceReceiptUrl = p_Data
  End Property

  Public Property Get InvoiceReceiptUrl()
     InvoiceReceiptUrl = m_InvoiceReceiptUrl
  End Property

End Class
%>