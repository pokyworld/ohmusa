<%
  
  ' Set newOrderItems = Server.CreateObject("Scripting.Dictionary")

  TotalSoFar = 0
  For Each Item In newOrderItems.Items
      Set newOrderItem = Item
      TotalSoFar = TotalSoFar + newOrderItem.Price.Amount
  Next
  
  PromoDiscount = FormatNumber(PromoDiscount, 2) * 100
  PromoCode = PromoCode
  Discount = FormatNumber(Discount, 2) * 100
  Shipping = Shipping
  ShippingTax = Shipping * TaxRate
  Total = TotalSoFar - Discount - PromoDiscount + Shipping + ShippingTax

  Set newOrder = SetOrder(OrderId, OrderDate, newCustomer, newOrderItems, Curency, PromoCode, PromoDiscount, Discount, Shipping, newPayment)

  Set newOrderItems = Nothing

  ' PrintOrder(newOrder)  
%>