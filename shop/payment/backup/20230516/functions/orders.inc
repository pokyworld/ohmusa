
<%

Function SetCustomer(FullName, Login, UserId, StripeCustomerId, Email, Phone, Addresses)
  Set billingAddr = Addresses.Item("BillingAddress")
  Set invoiceAddr = Addresses.Item("InvoiceAddress")
  Set shippingAddr = Addresses.Item("ShippingAddress")
  Set newCustomer = new Customer
  Set newCustomer.BillingAddress = billingAddr
  Set newCustomer.InvoiceAddress = invoiceAddr
  Set newCustomer.ShippingAddress = shippingAddr
  newCustomer.FullName = FullName
  newCustomer.Login = Login
  newCustomer.UserId = UserId
  newCustomer.StripeCustomerId = StripeCustomerId
  newCustomer.Email = Email
  newCustomer.Phone = Phone
  Set SetCustomer = newCustomer
End Function

Function SetOrder(OrderId, Customer, OrderItems, Curency, PromoCode, PromoDiscount, Discount, Shipping, Payment)
  Shipping = Shipping * 100 '// Stored as cente
  Set newOrder = new Order
  Set newOrder.OrderItems = OrderItems
  Set newOrder.Customer = Customer
  newOrder.OrderId = OrderId
  newOrder.Curency = Curency
  newOrder.PromoCode = PromoCode
  newOrder.PromoDiscount = PromoDiscount
  newOrder.Discount = Discount
  newOrder.Shipping = Shipping
  
  '// Sum Amount, Tax And SubTotal fields
  NetTotalSoFar = 0
  TaxSoFar = 0
  TotalSoFar = 0

  For Each Item In OrderItems
    Set newOrderItem = OrderItems(Item)
      NetTotalSoFar = NetTotalSoFar + (newOrderItem.Price.Amount * newOrderItem.Quantity)
      TaxSoFar = TaxSoFar + (newOrderItem.Price.Tax * newOrderItem.Quantity)
      TotalSoFar = NetTotalSoFar + TaxSoFar '// Amount including Tax
  Next

  SubTotal = NetTotalSoFar - Discount - PromoDiscount + Shipping
  TotalTax = SubTotal * TaxRate
  Total = SubTotal + TotalTax

  Set newOrder.OrderItems = OrderItems
  newOrder.NetSubTotal = NetTotalSoFar
  newOrder.SubTotal = SubTotal
  newOrder.Tax = TotalTax
  newOrder.Total = Total

  Set SetOrder = newOrder
End Function

Function SetOrderItem(Line, ProductId, SKU, Name, Curency, Amount, Tax, Quantity)
  Set newProduct = new Product
  newProduct.ID = ProductId
  newProduct.SKU = SKU
  newProduct.Name = Name

  Set newPrice = new Price
  Set newPrice.Product = newProduct
  newPrice.Curency = Curency
  newPrice.Amount = Amount
  newPrice.Tax = Tax
  newPrice.SubTotal = newPrice.Amount + newPrice.Tax

  Set newOrderItem = new OrderItem
  Set newOrderItem.Price = NewPrice
  newOrderItem.Line = Line
  newOrderItem.Quantity = Quantity
  newOrderItem.SubTotal = newOrderItem.Price.SubTotal * newOrderItem.Quantity
  Set SetOrderItem = newOrderItem
End Function

Sub PrintCustomer(Customer)
  out = "FullName: " & Customer.FullName & "<br/>"
  If Not Customer.UserId = "" Then : out = out & "UserId: " & Customer.UserId & "<br/>"
  If Not Customer.StripeCustomerId = "" Then : out = out & "StripeCustomerId: " & Customer.StripeCustomerId & "<br/>"
  If Not Customer.Login = "" Then : out = out & "Login: " & Customer.Login & "<br/>"
  If Not Customer.Email = "" Then : out = out & "Email: " & Customer.Email & "<br/>"
  If Not Customer.Phone = "" Then : out = out & "Phone: " & Customer.Phone & "<br/>"
  out = out & "Billing Address: " & "<br/>"
  If Not Customer.BillingAddress.Line1 = "" Then : out = out & vbTab & "Addr: " & Customer.BillingAddress.Line1 & "<br/>"
  If Not Customer.BillingAddress.Line2 = "" And Customer.BillingAddress.Line2 <> Customer.BillingAddress.Line1 Then : out = out & vbTab & "Addr: " & Customer.BillingAddress.Line2 & "<br/>"
  If Not Customer.BillingAddress.City = "" Then : out = out & vbTab & "City: " & Customer.BillingAddress.City & "<br/>"
  If Not Customer.BillingAddress.State = "" Then : out = out & vbTab & "State: " & Customer.BillingAddress.State & "<br/>"
  If Not Customer.BillingAddress.Zip = "" Then : out = out & vbTab & "Zip: " & Customer.BillingAddress.Zip & "<br/>"
  If Not Customer.BillingAddress.Country = "" Then : out = out & vbTab & "Country: " & Customer.BillingAddress.Country & "<br/>"
  out = out & "Shipping Address: " & "<br/>"
  If Not Customer.ShippingAddress.Line1 = "" Then : out = out & vbTab & "Addr: " & Customer.ShippingAddress.Line1 & "<br/>"
  If Not Customer.ShippingAddress.Line2 = "" And Customer.ShippingAddress.Line2 <> Customer.ShippingAddress.Line1 Then : out = out & vbTab & "Addr: " & Customer.ShippingAddress.Line2 & "<br/>"
  If Not Customer.ShippingAddress.City = "" Then : out = out & vbTab & "City: " & Customer.ShippingAddress.City & "<br/>"
  If Not Customer.ShippingAddress.State = "" Then : out = out & vbTab & "State: " & Customer.ShippingAddress.State & "<br/>"
  If Not Customer.ShippingAddress.Zip = "" Then : out = out & vbTab & "Zip: " & Customer.ShippingAddress.Zip & "<br/>"
  If Not Customer.ShippingAddress.Country = "" Then : out = out & vbTab & "Country: " & Customer.ShippingAddress.Country & "<br/>"
  Rw(out)
End Sub

Sub PrintOrder(Order)
  RW("OrderId: " & Order.OrderId)
  For Each Item in Order.OrderItems
      Set newOrderItem = Order.OrderItems(Item)
      out = "#" & newOrderItem.Line & "<br/>"
      out = out & "SKU: " & newOrderItem.Price.Product.SKU & "<br/>"
      out = out & "Product ID: " & newOrderItem.Price.Product.ID & "<br/>"
      out = out & "ProductName: " & newOrderItem.Price.Product.Name & "<br/>"
      ' out = out & "Slug: " & newOrderItem.Price.Product.Slug & "<br/>"
      ' out = out & "Color: " & newOrderItem.Price.Product.Color & "<br/>" 
      ' out = out & "Size: " & newOrderItem.Price.Product.Size & "<br/>" 
      ' out = out & "ProductUrl: " & Application("https_host") & newOrderItem.Price.Product.ProductUrl & "<br/>"
      out = out & "Currency: " & UCase(newOrderItem.Price.Curency) & "<br/>" 
      out = out & "Amount: " & FormatNumber(newOrderItem.Price.Amount/100,2) & "<br/>" 
      out = out & "Tax: " & FormatNumber(newOrderItem.Price.Tax/100,2) & "<br/>" 
      out = out & "Quantity: " & FormatNumber(newOrderItem.Quantity,0) & "<br/>" 
      out = out & "SubTotal: " & FormatNumber(newOrderItem.SubTotal/100,2)
      RW(out)
  Next
  out = "Currency: " & UCase(newOrder.Curency) & "<br/>" 
  out = out & "NetSubTotal: " & FormatNumber(Order.NetSubTotal/100,2) & "<br/>" 
  If Order.Discount >= 1 Then : out = out & "Discount: (" & FormatNumber(Order.Discount/100,2) & ")<br/>" 
  If Order.PromoDiscount >= 1 Then : out = out & "Promo " & UCase(Order.PromoCode) & ": (" & FormatNumber(Order.PromoDiscount/100,2) & ")<br/>" 
  out = out & "Shipping: " & FormatNumber(Order.Shipping/100,2) & "<br/>" 
  out = out & "SubTotal: " & FormatNumber(Order.SubTotal/100,2) & "<br/>" 
  out = out & "Tax: " & FormatNumber(Order.Tax/100,2) & "<br/>" 
  out = out & "Total: " & FormatNumber(Order.Total/100,2)
  RW(out)
End Sub

Function GetStripe(apiKey, url)
  Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
  http.open "GET", url, False
  http.setRequestHeader "Authorization", "Bearer " & apiKey
  http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  http.send
  result = http.responseText
  Set http = Nothing
  GetStripe = result
End Function

Function PostStripe(apiKey, url, reqBody)
  Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
  http.open "POST", url, False
  http.setRequestHeader "Authorization", "Bearer " & apiKey
  http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  http.send reqBody
  result = http.responseText
  Set http = Nothing
  PostStripe = result
End Function

Function CheckForCoupon(apiKey, reqUrl)
  ' RW("Checking: " & reqUrl)
  result = GetStripe(apiKey, reqUrl)  
  If Len(result) >= 1 Then
    Set ojCoupon = new aspJSON
    ojCoupon.loadJSON(result)
    If Len(ojCoupon.data("id")) >= 1 Then 
      CouponId = ojCoupon.data("id")
      ' RW("Found: " & CouponId)
    Else
      CouponId = ""
    End If
    Set ojCoupon = Nothing
    CheckForCoupon = CouponId
  Else
    CheckForCoupon = ""
  End If
End Function

Function CreateCoupon(apiKey, CouponId, CouponName, Curency, Amount)
  '// Remember Coupon 
  CouponMemo = CouponId
  reqUrl = "https://api.stripe.com/v1/coupons/" & CouponId
  result = CheckForCoupon(apiKey, reqUrl)
  If Len(result) >= 1 Then 
    CreateCoupon = result
  Else
    CouponId = CouponMemo
    ' RW("Creating coupon: " & CouponId)
    reqUrl = "https://api.stripe.com/v1/coupons"
    reqCoupon = "id=" & CouponId
    reqCoupon = reqCoupon & "&name=" & UCase(CouponName)
    reqCoupon = reqCoupon & "&amount_off=" & Amount
    reqCoupon = reqCoupon & "&duration=once"
    reqCoupon = reqCoupon & "&currency=" & LCase(Curency)
    ' RW(ReqCoupon)
    result = PostStripe(apiKey, reqUrl, reqCoupon)
    ' RW(result)
    If Len(result) >= 1 Then
      Set ojCoupon = new aspJSON
      ojCoupon.loadJSON(result)
      CouponId = ojCoupon.data("id")
      Set ojCoupon = Nothing
      reqUrl = "https://api.stripe.com/v1/coupons/" & CouponId
      result = CheckForCoupon(apiKey, reqUrl)
      If Len(result) >= 1 Then
        CreateCoupon = CouponId
      Else
        CreateCoupon = ""
      End If
    End If
  End If
End Function

Function GetStripeSession(apiKey, StripeInvoiceId)
  reqUrl = "https://api.stripe.com/v1/checkout/sessions"
  reqUrl = reqUrl & "/" & StripeSessionId
  result = GetStripe(apiKey, reqUrl)
  If Len(result) >= 1 Then
    GetStripeSession = result
  Else
    GetStripeSession = ""
  End If
End Function

Function GetStripeInvoice(apiKey, StripeInvoiceId)
  reqUrl = "https://api.stripe.com/v1/invoices"
  reqUrl = reqUrl & "/" & StripeInvoiceId
  result = GetStripe(apiKey, reqUrl)
  If Len(result) >= 1 Then
    GetStripeInvoice = result
  Else
    GetStripeInvoice = ""
  End If
End Function


%>