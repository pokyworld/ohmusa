<%
Class Address
  Private m_AddrName
  Private m_Line1
  Private m_Line2
  Private m_Locality
  Private m_City
  Private m_State
  Private m_Zip
  Private m_Country

  Public Property Let AddrName(p_Data)
      m_AddrName = p_Data
  End Property

  Public Property Get AddrName()
     AddrName = m_AddrName
  End Property

  Public Property Let Line1(p_Data)
      m_Line1 = p_Data
  End Property

  Public Property Get Line1()
     Line1 = m_Line1
  End Property

  Public Property Let Line2(p_Data)
      m_Line2 = p_Data
  End Property

  Public Property Get Line2()
      Line2 = m_Line2
  End Property

  Public Property Let City(p_Data)
      m_City = p_Data
  End Property

  Public Property Get City()
     City = m_City
  End Property

  Public Property Let State(p_Data)
      m_State = p_Data
  End Property

  Public Property Get State()
     State = m_State
  End Property

  Public Property Let Zip(p_Data)
      m_Zip = p_Data
  End Property

  Public Property Get Zip()
     Zip = m_Zip
  End Property

  Public Property Let Country(p_Data)
      m_Country = p_Data
  End Property

  Public Property Get Country()
    Country = m_Country
  End Property

End Class

%>