<% 
  Class Product
    Private p_seq
    Private p_sku
    Private p_title
    Private p_Description
    Private p_Price
    Private p_population
    Private p_capital

    ' getter and setter
    Public Property Get SKU()
      SKU = p_sku
    End Property
    Public Property Let SKU(value)
      p_sku = value
    End Property

    Public Property Get Title()
      Title = p_title
    End Property
    Public Property Let Title(value)
      p_title = value
    End Property

    Public Property Get Description()
      Description = p_Description
    End Property
    Public Property Let Description(value)
      p_Description = value
    End Property
    
    Public Property Get Price()
      Price = p_Price
    End Property
    Public Property Let Price(value)
      p_Price = value
    End Property    
  End Class
%>
