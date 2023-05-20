<%
'// Variables defined here: TaxRate, Shipping (cost) //
  If Len(Session("mapMarkup")) > 0 Then
    MapMargin = CDbl(Session("mapMarkup"))
  Else
    MapMargin = CDbl(Trim(Application("defaultMarginFactor")&"")) 
  End If

  TaxRate = CDbl(Trim(Application("defaultTaxRate")&""))  '// Percent as decimal
  TaxOnTotalOnly = True

  Set newOrderItems = Server.CreateObject("Scripting.Dictionary")

  sql = "SELECT o.orderid, "
  sql = sql & "CASE WHEN c.consumer_id = NULL THEN 0 ELSE c.consumer_id END AS consumer_id, "

  ' sql = sql & "CASE WHEN c.stripeCustomerId = NULL THEN "" ELSE c.stripeCustomerId END AS stripeCustomerId, "
  ' sql = sql & "'TEST|cus_NWDRWQVYGMvPYp' AS stripeCustomerId, "
  sql = sql & "'' AS stripeCustomerId, "

  sql = sql & "o.[login], o.orderdate, p.Product_ID AS productId, p.Thumb_Img AS image, d.Product_Code AS productCode, "
  sql = sql & "c.stripeId, d.Product_Name AS productName, od.quantity, ROUND(od.price,2) AS price, "
  sql = sql & "o.shippingCost, o.promoCode, o.promoDiscount, o.discount, o.total AS subTotal, "
  sql = sql & "0 AS salesTax, o.total, o.Contactname, o.ShiptoCompanyName, o.stripePaymentId, "
  sql = sql & "o.address1, o.address2, o.city, o.state, o.zip, o.country, o.email, o.phone, py.po "
  sql = sql & "FROM dbo.orders o "
  sql = sql & "LEFT JOIN dbo.consumer c ON c.[login] = o.[login] "
  sql = sql & "INNER JOIN dbo.orderdetail od ON od.orderid = o.orderid "
  sql = sql & "INNER JOIN dbo.dropshiptemplate2013 d ON d.itemid = od.productid "
  sql = sql & "INNER JOIN dbo.Products p ON p.item = d.product_code "
  sql = sql & "INNER JOIN dbo.payment py ON py.paymentId = o.paymentId "
  sql = sql & "WHERE o.orderid = " & orderId & " "

  'Create connection and query category data.
  strSQLCmd1 = sql
  Set cnn1 = Server.CreateObject("ADODB.Connection")
  cnn1.ConnectionString = Application.Contents("dbConnStr")
  cnn1.Open
  Set rsOrder = Server.CreateObject("ADODB.Recordset")

  rsOrder.Open strSQLCmd1, cnn1, 3

  If Not rsOrder.BOF And Not rsOrder.EOF Then
    Counter = 1
    Do while Not rsOrder.EOF
      Line = Counter
      ProductId = CInt(Trim(rsOrder("productId")&""))
      ' ImageUrl = "../../ThumbImages/" & Trim(rsOrder("image")&"")
      ImageUrl = "https://omhusa.com/ThumbImages/" & Trim(rsOrder("image")&"")
      SKU = UCase(Trim(rsOrder("productCode")&""))
      ProductName = Trim(rsOrder("productName")&"")
      Curency = "USD"
      Amount = CDbl(Trim(rsOrder("price")&"")) * 100
      Quantity = CInt(Trim(rsOrder("quantity")&""))
      PromoCode = Trim(rsOrder("promoCode")&"")
      PromoDiscount = CDbl(Trim(rsOrder("promoDiscount")&""))
      Discount = CDbl(Trim(rsOrder("discount")&""))

      dbSubTotal = CDbl(Trim(rsOrder("subTotal")&"")) 
      dbShipping = CDbl(Trim(rsOrder("shippingCost")&"")) 
      dbTax = CDbl(Trim(rsOrder("salesTax")&""))
      dbTotal = CDbl(Trim(rsOrder("total")&"")) 

      OrderDate = FormatDtTime(Trim(rsOrder("orderdate")&""))
      FullName = Trim(rsOrder("contactName")&"")
      Login = Trim(rsOrder("login")&"")
      UserId = Trim(rsOrder("consumer_id")&"")
      StripeCustomerId = ""
      StripeId = Trim(rsOrder("stripeId")&"")
      If Len(StripeId) >= 6 Then 
        aryStripe = Split(StripeId, "|")
        If IsArray(aryStripe) = True And (aryStripe(0) = "TEST" Or aryStripe(0) = "LIVE") Then
        '// Check that any ID is valid for the current stripe_mode (TEST/LIVE)
          If UCase(aryStripe(0)) = UCase(stripe_mode) Then
            StripeCustomerId = Trim(aryStripe(1))
          End If
        End If
      End If
      StripePaymentId = Trim(rsOrder("stripePaymentId")&"")
      PurchaseOrder = Trim(rsOrder("po")&"")
      If Len(PurchaseOrder) = 0 Then : PurchaseOrder = "None"
      Email = LCase(Trim(rsOrder("email")&""))
      Phone = Trim(rsOrder("phone")&"")

      ShipAddrContact = Trim(rsOrder("ContactName")&"")
      ShipAddrCompany = Trim(rsOrder("ShiptoCompanyName")&"")
      ShipAddrLine1 = Trim(rsOrder("address1")&"")
      ShipAddrLine2 = Trim(rsOrder("address2")&"")
      ShipAddrCity = Trim(rsOrder("city")&"")
      ShipAddrState = Trim(rsOrder("state")&"")
      ShipAddrZip = Trim(rsOrder("zip")&"")
      ShipAddrCountry = Trim(rsOrder("country")&"")
      UseShippingAddrForBillingAddr = True

      Set newOrderItem = SetOrderItem(Line, ProductId, SKU, ProductName, Curency, Amount, Tax, Quantity, ImageUrl)
      newOrderItems.Add Counter, newOrderItem

      Counter = Counter + 1
      rsOrder.MoveNext
    Loop
  Else
    RW("No Order Found")
    Response.End
  End If
  Set rsOrder = Nothing
%>