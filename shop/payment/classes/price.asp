<%
Class Price
  Public Product
  Private m_Curency
  Private m_Amount
  Private m_Tax
  Private m_SubTotal

  Public Property Let Curency(p_Data)
      m_Curency = UCase(p_Data)
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

  Public Property Let Tax(p_Data)
      m_Tax = p_Data
  End Property

  Public Property Get Tax()
     Tax = m_Tax
  End Property

  Public Property Let SubTotal(p_Data)
      m_SubTotal = p_Data
  End Property

  Public Property Get SubTotal()
     SubTotal = m_SubTotal
  End Property

  ' Private Sub Class_Initialize()

	' End Sub
  ' Private Sub Class_Terminate()

	' End Sub

End Class
%>