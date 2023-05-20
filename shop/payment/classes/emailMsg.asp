<%
Class EmailMsg
  Private m_MailId
  Private m_MailRef
  Private m_MailFrom
  Private m_MailDisplayName
  Private m_Subject
  Private m_TextBody
  Private m_HtmlBody
  Public MailTo
  Public MailCC

  Public Property Let MailId(p_Data)
    m_MailId = p_Data
  End Property

  Public Property Get MailId()
    MailId = m_MailId
  End Property

  Public Property Let MailRef(p_Data)
    m_MailRef = p_Data
  End Property

  Public Property Get MailRef()
    MailRef = m_MailRef
  End Property

  Public Property Let Subject(p_Data)
    m_Subject = p_Data
  End Property

  Public Property Get Subject()
    Subject = m_Subject
  End Property

  Public Property Let TextBody(p_Data)
    m_TextBody = p_Data
  End Property

  Public Property Get TextBody()
    TextBody = m_TextBody
  End Property

  Public Property Let HtmlBody(p_Data)
    m_HtmlBody = p_Data
  End Property

  Public Property Get HtmlBody()
    HtmlBody = m_HtmlBody
  End Property

  Public Property Let MailFrom(p_Data)
    m_MailFrom = p_Data
  End Property

  Public Property Get MailFrom()
    MailFrom = m_MailFrom
  End Property

  Public Property Let MailDisplayName(p_Data)
    m_MailDisplayName = p_Data
  End Property

  Public Property Get MailDisplayName()
    MailDisplayName = m_MailDisplayName
  End Property

End Class
%>