<%
Class Customer
  Public BillingAddress
  Public InvoiceAddress
  Public ShippingAddress
  Private m_FullName
  Private m_Login
  Private m_UserId
  Private m_StripeCustomerId
  Private m_Email
  Private m_Phone

  Public Property Let FullName(p_Data)
      m_FullName = p_Data
  End Property

  Public Property Get FullName()
     FullName = m_FullName
  End Property

  Public Property Let Login(p_Data)
      m_Login = p_Data
  End Property

  Public Property Get Login()
     Login = m_Login
  End Property

  Public Property Let UserId(p_Data)
      m_UserId = p_Data
  End Property

  Public Property Get UserId()
     UserId = m_UserId
  End Property

  Public Property Let StripeCustomerId(p_Data)
      m_StripeCustomerId = p_Data
  End Property

  Public Property Get StripeCustomerId()
     StripeCustomerId = m_StripeCustomerId
  End Property

  Public Property Let Email(p_Data)
      m_Email = p_Data
  End Property

  Public Property Get Email()
     Email = m_Email
  End Property

  Public Property Let Phone(p_Data)
      m_Phone = p_Data
  End Property

  Public Property Get Phone()
     Phone = m_Phone
  End Property

End Class
%>