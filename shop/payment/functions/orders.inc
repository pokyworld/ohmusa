
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

Function SetPayment(StripeChargeId, Curency, PurchaseOrder)
  Set newPayment = new Payment
  newPayment.Id = StripeChargeId
  newPayment.Curency = Curency
  newPayment.PurchaseOrder = PurchaseOrder
  result = GetStripeCharge(secret_key, StripeChargeId)
  Set ojPayment = new aspJSON
  ojPayment.loadJSON(result)
  newPayment.Amount = ojPayment.data("amount")
  newPayment.Status = ojPayment.data("status")
  Set ojPayment = Nothing
  Set SetPayment = newPayment
End Function

Function SetOrder(OrderId, OrderDate, Customer, OrderItems, Curency, PromoCode, PromoDiscount, Discount, Shipping, Payment)
  Shipping = Shipping * 100 '// Stored as cents
  Set newOrder = new Order
  Set newOrder.OrderItems = OrderItems
  Set newOrder.Customer = Customer
  Set newOrder.Payment = Payment
  newOrder.OrderId = OrderId
  newOrder.OrderDate = OrderDate
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

Function SetOrderItem(Line, ProductId, SKU, Name, Curency, Amount, Tax, Quantity, ImageUrl)
  Set newProduct = new Product
  newProduct.ID = ProductId
  newProduct.SKU = SKU
  newProduct.Name = Name
  newProduct.ImageUrl = ImageUrl

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
  RW("OrderId: " & Order.OrderId & "<br/>" & _
    "OrderDate: " & FormatDateTime(Order.OrderDate, 0) & "<br/>" & _
    "")
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
      out = out & "ImageUrl: " & Application("https_host") & newOrderItem.Price.Product.ImageUrl & "<br/>"
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
  out = out & "Total: " & FormatNumber(Order.Total/100,2) & "<br/>" 
  out = out & "<br/>Payment Ref: " & Order.Payment.Id & "<br/>"
  out = out & "Payment Status: " & PCase(Order.Payment.Status) & "<br/>"
  out = out & "PurchaseOrder: " & Order.Payment.PurchaseOrder & "<br/>"
  out = out & "Payment Amount: " & FormatNumber(Order.Payment.Amount/100,2) & "<br/>"
  BalanceDue = Round(Order.Total/100,2) - Round(Order.Payment.Amount/100,2)
  out = out & "Balance Due: " & FormatNumber(BalanceDue,2)

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

Function GetStripePayment(apiKey, StripePaymentId)
  reqUrl = "https://api.stripe.com/v1/payment_intents"
  reqUrl = reqUrl & "/" & StripePaymentId
  result = GetStripe(apiKey, reqUrl)
  If Len(result) >= 1 Then
    GetStripePayment = result
  Else
    GetStripePayment = ""
  End If
End Function

Function GetStripeCharge(apiKey, StripeChargeId)
  reqUrl = "https://api.stripe.com/v1/charges"
  reqUrl = reqUrl & "/" & StripeChargeId
  result = GetStripe(apiKey, reqUrl)
  If Len(result) >= 1 Then
    GetStripeCharge = result
  Else
    GetStripeCharge = ""
  End If
End Function

Function UpdateStripeId(consumer_id, stripe_mode, StripeCustomerId)
  consumer_id = CInt(Trim(consumer_id&""))
  stripe_mode = UCase(Trim(stripe_mode&""))
  mode_test = "TEST|LIVE"
  StripeCustomerId = Trim(StripeCustomerId&"")
  If Len(stripe_mode) < 4 Then : UpdateStripeId = "" : Exit Function
  If InStr(mode_test, stripe_mode) = 0 Then : UpdateStripeId = "" : Exit Function
  If Len(StripeCustomerId) < 10 Then : UpdateStripeId = "" : Exit Function
  
  StripeId = stripe_mode & "|" & StripeCustomerId

  sql = "UPDATE dbo.consumer SET stripeId = '" & StripeId &"' "
  sql = sql & "WHERE consumer_id = " & consumer_id & ";"

  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  cnn.Open
  cnn.Execute(sql)
  Set cnn = Nothing
  UpdateStripeId = StripeId
End Function

Function UpdateStripePaymentId(orderid, StripeChargeId)
  orderid = CLng(Trim(orderid&""))
  StripeChargeId = Trim(StripeChargeId&"")
  If Len(StripeChargeId) < 10 Then : UpdateStripePaymentId = "" : Exit Function

  sql = "UPDATE dbo.orders SET stripePaymentId = '" & StripeChargeId &"' "
  sql = sql & "WHERE orderid = " & orderid & ";"

  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  cnn.Open
  cnn.Execute(sql)
  Set cnn = Nothing
  UpdateStripePaymentId = StripeChargeId
End Function

Function InsertStripeKeys(orderid, stripeMode, stripeSessionId, stripeCustomerId, stripePayIntentId, stripeInvoiceId, stripeChargeId, stripeInvoicePDF, stripeInvoiceReceipt)
  If Len(Trim(orderid)) = 0 Then : InsertStripeKeys = "" : Exit Function
  If Len(Trim(stripeSessionId)) = 0 Then : InsertStripeKeys = "" : Exit Function

  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  cnn.Open

  '// 1st check if already exists: update vs insert
  sql = "SELECT ISNULL(MAX(id), 0) AS TransId FROM dbo.stripeTransactions "
  sql = sql & "WHERE orderid = " & orderid & " "
  sql = sql & "AND stripeSessionId = '" & stripeSessionId & "' "
  sql = sql & "AND stripeCustomerId = '" & stripeCustomerId & "' "
  sql = sql & "AND stripePayIntentId = '" & stripePayIntentId & "' "
  sql = sql & "AND stripeInvoiceId = '" & stripeInvoiceId & "' "
  sql = sql & "AND stripeChargeId = '" & stripeChargeId & "';"
  ' RW(sql)

  Set rsTrans = cnn.Execute(sql)
  If Not rsTrans.BOF And Not rsTrans.EOF Then
    TransId = CLng(rsTrans("TransId"&""))
  End If
  Set rsTrans = Nothing

  If TransId >= 1 Then
    '// Update keys
    sql = "UPDATE dbo.stripeTransactions SET "
    sql = sql & "stripeMode = '" & stripeMode & "', "
    sql = sql & "stripeCustomerId = '" & stripeCustomerId & "', "
    sql = sql & "stripePayIntentId = '" & stripePayIntentId & "', "
    sql = sql & "stripeInvoiceId = '" & stripeInvoiceId & "', "
    sql = sql & "stripeChargeId = '" & stripeChargeId & "', "
    sql = sql & "stripeInvoicePDF = '" & stripeInvoicePDF & "', "
    sql = sql & "stripeInvoiceReceipt = '" & stripeInvoiceReceipt & "' "
    sql = sql & "WHERE id = " & TransId
    ' RW(sql)
    cnn.Execute(sql)
    InsertStripeKeys = TransId
  Else
    '// Insert keys
    sql = "INSERT INTO dbo.stripeTransactions (orderid, stripeMode, "
    sql = sql & "stripeSessionId, stripeCustomerId, stripePayIntentId, "
    sql = sql & "stripeInvoiceId, stripeChargeId, stripeInvoicePDF, "
    sql = sql & "stripeInvoiceReceipt) VALUES ("
    sql = sql & "" & orderid & ", "
    If Len(Trim(UCase(stripeMode))) >= 1 Then : sql = sql & "'" & UCase(stripeMode) & "', " : Else : sql = sql & "null, "
    If Len(Trim(stripeSessionId)) >= 1 Then : sql = sql & "'" & stripeSessionId & "', " : Else : sql = sql & "null, "
    If Len(Trim(stripeCustomerId)) >= 1 Then : sql = sql & "'" & stripeCustomerId & "', " : Else : sql = sql & "null, "
    If Len(Trim(stripePayIntentId)) >= 1 Then : sql = sql & "'" & stripePayIntentId & "', " : Else : sql = sql & "null, "
    If Len(Trim(stripeInvoiceId)) >= 1 Then : sql = sql & "'" & stripeInvoiceId & "', " : Else : sql = sql & "null, "
    If Len(Trim(stripeChargeId)) >= 1 Then : sql = sql & "'" & stripeChargeId & "', " : Else : sql = sql & "null, "
    If Len(Trim(stripeInvoicePDF)) >= 1 Then : sql = sql & "'" & stripeInvoicePDF & "', " : Else : sql = sql & "null, "
    If Len(Trim(stripeInvoiceReceipt)) >= 1 Then : sql = sql & "'" & stripeInvoiceReceipt & "');" : Else : sql = sql & "null);"
    ' RW(sql)
  
    cnn.Execute(sql)
  
    '// Get TransId
    sql = "SELECT ISNULL(MAX(id), 0) AS TransId FROM dbo.stripeTransactions "
    sql = sql & "WHERE orderid = " & orderid & " "
    sql = sql & "AND stripeSessionId = '" & stripeSessionId & "' "
    sql = sql & "AND stripeCustomerId = '" & stripeCustomerId & "' "
    sql = sql & "AND stripePayIntentId = '" & stripePayIntentId & "' "
    sql = sql & "AND stripeInvoiceId = '" & stripeInvoiceId & "' "
    sql = sql & "AND stripeChargeId = '" & stripeChargeId & "';"
    ' RW(sql)
  
    Set rsTrans = cnn.Execute(sql)
    If Not rsTrans.BOF And Not rsTrans.EOF Then
      TransId = CLng(rsTrans("TransId"&""))
      InsertStripeKeys = TransId
    Else
      InsertStripeKeys = ""
    End If
    Set rsTrans = Nothing
  End If  '// TransId >= 1
  Set cnn = Nothing
End Function

Function GetLastOrderPayment(orderid)
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  cnn.Open

  sql = "SELECT Max(Id) AS TransID, StripeSessionId "
  sql = sql & "FROM dbo.stripeTransactions "
  sql = sql & "WHERE orderid = " & orderid & " "
  sql = sql & "AND stripeChargeId Is Not Null "
  sql = sql & "GROUP BY StripeSessionId;"

  Set rsTrans = cnn.Execute(sql)
  If Not rsTrans.BOF And Not rsTrans.EOF Then
    TransId = CLng(Trim(rsTrans("TransId"&"")))
    StripeSessionId = Trim(rsTrans("StripeSessionId"&""))
  End If
  Set rsTrans = Nothing

  GetLastOrderPayment = StripeSessionId
End Function

Function CreatePaymentEmail(Order, StripeCustomerId)
  Set newEmailMsg = new EmailMsg

  MAIL_FROM = Application("MAIL_FROM")
  MAIL_DISPLAY_NAME = Application("MAIL_DISPLAY_NAME")
  
  Set newMailToItems = Server.CreateObject("Scripting.Dictionary")
  ' counter = 0
  ' For Each Item In Order.Customer.Email
  '   With newMailToItems
  '     .Add counter, """" & Order.Customer.FullName & """ " & Order.Customer.Email
  '     counter = counter + 1
  '   End With
  ' Next
  ' For Each key In newMailToItems.Keys
  '   RW(key & "," & newMailToItems.Item(key))
  ' Next

  With newMailToItems
    .Add 0, """" & Order.Customer.FullName & """ " & Order.Customer.Email
    ' .Add 1, """Developer"" butlerscripts@gmail.com"
    counter = counter + 1
  End With

  Set newEmailMsg.MailTo = newMailToItems
  Set newMailToItems = Nothing

  Set newMailCCItems = Server.CreateObject("Scripting.Dictionary")
  With newMailCCItems
    .Add 0, Order.Customer.Email
    .Add 1, """Developer"" butlerscripts@gmail.com"
    .Add 2, """Silverado Cruz"" silverado.cruiser@gmail.com"
    .Add 3, """Eric Vuong OMH"" evuong2000@gmail.com"
  End With
  Set newEmailMsg.MailCC = newMailCCItems
  Set newMailCCItems = Nothing

  newEmailMsg.MailFrom = MAIL_FROM
  newEmailMsg.MailDisplayName = MAIL_DISPLAY_NAME
  newEmailMsg.MailRef = Order.OrderId
  newEmailMsg.Subject = "OMH Order: " & Order.OrderId & " (Payment confirmation)"

  newEmailMsg.HtmlBody = FetchHtmlTemplate(OrderId, StripeCustomerId)
  newEmailMsg.TextBody = "Order: " & Order.OrderId & vbCrlf & "Stripe payment successful."
  
  result = SendNow(newEmailMsg)
  
  CreatePaymentEmail = result

End Function

Function SendNow(EmailMsg)

  Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
  Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
  Const cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
  Const cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
  Const cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
  Const cdoSendUsername = "http://schemas.microsoft.com/cdo/configuration/sendusername"
  Const cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
  Const cdoSendTLS = "http://schemas.microsoft.com/cdo/configuration/sendtls"
  Const cdoSMTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
  Const cdoSendUsingPort = 2
  Const cdoDispositionNotificationTo = "urn:schemas:mailheader:disposition-notification-to"
  Const cdoReturnReceiptTo = "urn:schemas:mailheader:return-receipt-to"
  Const cdoReturnPath = "urn:schemas:mailheader:return-path"

  MAIL_MAILER = Application("MAIL_MAILER")
  MAIL_PORT = Application("MAIL_PORT")
  MAIL_USERNAME = Application("MAIL_USERNAME")
  MAIL_PASSWORD = Application("MAIL_PASSWORD")
  MAIL_HOST = Application("MAIL_HOST")
  MAIL_ENCRYPTION = Application("MAIL_ENCRYPTION")

  Set MailToItems = EmailMsg.MailTo
  If MailToItems.Count = 0 Then SendNow = "Nothing to Send" : Exit Function
  MailTo = Join(MailToItems.Items, ",")

  Set MailCCItems = EmailMsg.MailCC

  MailId = EmailMsg.MailId
  MailRef = EmailMsg.MailRef
  Subject = EmailMsg.Subject
  TextBody = EmailMsg.TextBody
  HtmlBody = EmailMsg.HtmlBody

  MailFrom = EmailMsg.MailFrom
  MailDisplayName = EmailMsg.MailDisplayName

  Dim cdoConfig, cdoMsg

  Set cdoConfig = CreateObject("CDO.Configuration") 

  With cdoConfig.Fields 
    .Item(cdoSendUsingMethod) = 2 ' cdoSendUsingPort 
    .Item(cdoSMTPServer) = MAIL_HOST
    .Item(cdoSMTPServerPort) = MAIL_PORT
    .Item(cdoSMTPConnectionTimeout) = 180
  	.Item(cdoSMTPAuthenticate) = 1
	  .Item(cdoSendUsername) = MAIL_USERNAME
	  .Item(cdoSendPassword) = MAIL_PASSWORD
    .Item(cdoSendTLS) = True
    ' .Item(cdoSMTPUseSSL) = True

    ' RW("cdoSendUsingMethod: " & .Item(cdoSendUsingMethod))
    ' RW("cdoSMTPServer: " & .Item(cdoSMTPServer))
    ' RW("cdoSMTPServerPort: " & .Item(cdoSMTPServerPort))
    ' RW("cdoSMTPAuthenticate: " & .Item(cdoSMTPAuthenticate))
    ' RW("cdoSendUsername: " & .Item(cdoSendUsername))
    ' RW("cdoSendTLS: " & .Item(cdoSendTLS))
    ' RW("cdoSMTPUseSSL: " & .Item(cdoSMTPUseSSL))

    .Update 
  End With 
  MailTo = """Eric Vuong OMH"" evuong2000@gmail.com, ""Des Butler"" butlerscripts@gmail.com,""Silverado Cruz"" silverado.cruiser@gmail.com"
  Set cdoMsg = CreateObject("CDO.Message") 
  With cdoMsg 
    Set .Configuration = cdoConfig 
    .From = """" & MailDisplayName & """" & MailFrom
    .To = MailTo
    .cc = MailCC
    .Subject = Subject
    .TextBody = TextBody
    .HTMLBody = HtmlBody
    ' .CreateMHTMLBody "https://asus-aio.windows/shop/payment/templates/payment_conf.asp?orderid=" & MailRef

    ' .addAttachment (path)

    ' RW(.From)
    ' RW(.To)
    ' RW(.Cc)
    ' RW(Subject)
    ' RW(.TextBody)
    ' RW(.HtmlBody)

    ' RW("Sending...")

    .Send 

  End With 

  Set cdoMsg = Nothing
  Set cdoConfig = Nothing

  SendNow = Subject

End Function

Function FetchHtmlTemplate(OrderId, StripeCustomerId)
  host = "https://" & Request.ServerVariables("HTTP_HOST")
  url = host & "/shop/payment/templates/payment_conf.asp?orderid=" & OrderId & "&stripe=" & StripeCustomerId
  Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
  http.open "GET", url, False
  http.send
  result = http.responseText
  Set http = Nothing
  FetchHtmlTemplate = result
End Function
%>