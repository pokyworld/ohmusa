<%
Class Payment
  Private m_Method
  Private m_Status
  Private m_Curency
  Private m_Amount

  Public Property Let Method(p_Data)
      m_Method = p_Data
  End Property

  Public Property Get Method()
     Method = m_Method
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

End Class
%>