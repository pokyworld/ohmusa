<%
  
  ' Set newOrderItems = Server.CreateObject("Scripting.Dictionary")

  TotalSoFar = 0
  For Each Item In newOrderItems.Items
      Set newOrderItem = Item
      TotalSoFar = TotalSoFar + newOrderItem.Price.Amount
  Next
  
  Discount = Round(TotalSoFar * DiscountRate, 2)
  If Discount <> dbDiscount * 100 Then : Discount = dbDiscount * 100
  If Shipping <> dbShipping * 100 Then : Shipping = dbShipping
  ShippingTax = Shipping * TaxRate
  Total = TotalSoFar - Discount + Shipping + ShippingTax

  Set newOrder = SetOrder(OrderId, newCustomer, newOrderItems, Discount, Shipping, newPayment)

  Set newOrderItems = Nothing

  ' PrintOrder(newOrder)  
%>