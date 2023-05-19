<%
Class Payment
  Private m_Id
  Private m_Status
  Private m_Curency
  Private m_Amount

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

End Class
%>