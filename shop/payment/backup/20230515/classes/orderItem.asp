<%
Class OrderItem
  Public Price
  Private m_Line
  Private m_Quantity
  Private m_SubTotal

  Public Property Let Line(p_Data)
      m_Line = p_Data
  End Property

  Public Property Get Line()
    Line = m_Line
  End Property

  Public Property Let Quantity(p_Data)
    m_Quantity = p_Data
  End Property

  Public Property Get Quantity()
    Quantity = m_Quantity
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