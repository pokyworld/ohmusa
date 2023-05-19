<%
Class Order
  Public OrderItems
  Public Customer
  Public Payment
  Private m_OrderId
  Private m_NetSubTotal
  Private m_SubTotal
  Private m_Discount
  Private m_Shipping
  Private m_Tax
  Private m_Total

  Public Property Let OrderId(p_Data)
      m_OrderId = p_Data
  End Property

  Public Property Get OrderId()
     OrderId = m_OrderId
  End Property

  Public Property Let NetSubTotal(p_Data)
      m_NetSubTotal = p_Data
  End Property

  Public Property Get NetSubTotal()
     NetSubTotal = m_NetSubTotal
  End Property

  Public Property Let SubTotal(p_Data)
      m_SubTotal = p_Data
  End Property

  Public Property Get SubTotal()
     SubTotal = m_SubTotal
  End Property

  Public Property Let Discount(p_Data)
      m_Discount = p_Data
  End Property

  Public Property Get Discount()
     Discount = m_Discount
  End Property

  Public Property Let Shipping(p_Data)
      m_Shipping = p_Data
  End Property

  Public Property Get Shipping()
     Shipping = m_Shipping
  End Property

  Public Property Let Tax(p_Data)
      m_Tax = p_Data
  End Property

  Public Property Get Tax()
     Tax = m_Tax
  End Property

  Public Property Let Total(p_Data)
      m_Total = p_Data
  End Property

  Public Property Get Total()
     Total = m_Total
  End Property

  ' Private Sub Class_Initialize()

	' End Sub
  ' Private Sub Class_Terminate()

	' End Sub

End Class
%>