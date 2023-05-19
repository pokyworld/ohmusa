<%
Class Product
  Private m_ID
  Private m_SKU
  Private m_Name
  Private m_Slug
  Private m_Size
  Private m_Color
  Private m_Description
  Private m_ProductUrl
  Private m_ImageUrl

  Public Property Let ID(p_Data)
      m_ID = p_Data
  End Property

  Public Property Get ID()
     ID = m_ID
  End Property

  Public Property Let SKU(p_Data)
      m_SKU = p_Data
  End Property

  Public Property Get SKU()
     SKU = m_SKU
  End Property

  Public Property Let Name(p_Data)
      m_Name = p_Data
      m_Slug = GetSlug(m_Name) & "-" & UCase(m_SKU)
      m_ProductUrl = "/product/" & m_Slug
  End Property

  Public Property Get Name()
     Name = m_Name
  End Property

  Public Property Let Slug(p_Data)
      m_Slug = p_Data
  End Property

  Public Property Get Slug()
     Slug = m_Slug
  End Property

  Public Property Let Size(p_Data)
      m_Size = p_Data
  End Property

  Public Property Get Size()
     Size = m_Size
  End Property

  Public Property Let Color(p_Data)
      m_Color = p_Data
  End Property

  Public Property Get Color()
     Color = m_Color
  End Property

  Public Property Let Description(p_Data)
      m_Description = p_Data
  End Property

  Public Property Get Description()
    If m_Description = "" then
         Description = "No Description"
    Else
         Description = m_Description
    End If
  End Property

  Public Property Let ProductUrl(p_Data)
      m_ProductUrl = p_Data
  End Property

  Public Property Get ProductUrl()
     ProductUrl = m_ProductUrl
  End Property

  Public Property Let ImageUrl(p_Data)
      m_ImageUrl = p_Data
  End Property

  Public Property Get ImageUrl()
     ImageUrl = m_ImageUrl
  End Property

  ' Private Sub Class_Initialize()

	' End Sub
  ' Private Sub Class_Terminate()

	' End Sub

End Class
%>