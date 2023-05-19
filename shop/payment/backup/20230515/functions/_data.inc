<%
'// Variables defined here: TaxRate, Shipping (cost) //
  If Len(Session("mapMarkup")) > 0 Then
    MapMargin = CDbl(Session("mapMarkup"))
  Else
    ' MapMargin = CDbl(Trim(Application("defaultMarginFactor")&"")) 
    MapMargin = 1.15
  End if
  ' Shipping = CDbl(Trim(Application("defaultShippingRate")&""))  '// In cents, no decimal
  Shipping = 0
  ' TaxRate = CDbl(Trim(Application("defaultTaxRate")&""))  '// Percent as decimal
  TaxRate = 0
  DiscountRate = 0
  TaxOnTotalOnly = True
  ' PromoCode = "OMHSPECIAL10"

  Set newOrderItems = Server.CreateObject("Scripting.Dictionary")

  sql = "SELECT o.orderid, "
  sql = sql & "CASE WHEN c.consumer_id = NULL THEN 0 ELSE c.consumer_id END AS consumer_id, "

  ' sql = sql & "CASE WHEN c.stripeCustomerId = NULL THEN "" ELSE c.stripeCustomerId END AS stripeCustomerId, "
  ' sql = sql & "'TEST|cus_NWDRWQVYGMvPYp' AS stripeCustomerId, "
  sql = sql & "'' AS stripeCustomerId, "

  sql = sql & "o.[login], o.orderdate, p.Product_ID AS productId, d.Product_Code AS productCode, "
  sql = sql & "d.Product_Name AS productName, od.quantity, ROUND(od.price,2) AS price, "
  sql = sql & "o.shippingCost, o.promoDiscount AS discount, o.total AS subTotal, "
  sql = sql & "0 AS salesTax, o.total, o.Contactname, o.ShiptoCompanyName, "
  sql = sql & "o.address1, o.address2, o.city, o.state, o.zip, o.country, o.email, o.phone "
  sql = sql & "FROM dbo.orders o "
  sql = sql & "LEFT JOIN dbo.consumer c ON c.[login] = o.[login] "
  sql = sql & "INNER JOIN dbo.orderdetail od ON od.orderid = o.orderid "
  sql = sql & "INNER JOIN dbo.dropshiptemplate2013 d ON d.itemid = od.productid "
  sql = sql & "INNER JOIN dbo.Products p ON p.item = d.product_code "
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
      SKU = UCase(Trim(rsOrder("productCode")&""))
      ProductName = Trim(rsOrder("productName")&"")
      Curency = "USD"
      Amount = CDbl(Trim(rsOrder("price")&"")) * 100
      Quantity = CInt(Trim(rsOrder("quantity")&""))

      dbSubTotal = CDbl(Trim(rsOrder("subTotal")&"")) 
      dbShipping = CDbl(Trim(rsOrder("shippingCost")&"")) 
      dbDiscount = CDbl(Trim(rsOrder("discount")&""))
      dbTax = CDbl(Trim(rsOrder("salesTax")&""))
      dbTotal = CDbl(Trim(rsOrder("total")&"")) 

      FullName = Trim(rsOrder("contactName")&"")
      Login = Trim(rsOrder("login")&"")
      UserId = Trim(rsOrder("consumer_id")&"")
      StripeCustomerId = ""
      StripeId = Trim(rsOrder("stripeCustomerId")&"")
      If Len(StripeId) >= 6 Then 
        aryStripe = Split(StripeId, "|")
        If IsArray(aryStripe) = True And (aryStripe(0) = "TEST" Or aryStripe(0) = "LIVE") Then
        '// Check that any ID is valid for the current stripe_mode (TEST/LIVE)
          If UCase(aryStripe(0)) = UCase(stripe_mode) Then
            StripeCustomerId = Trim(aryStripe(1))
          End If
        End If
      End If
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

      ' RW("Line: "& Line & "<br/>" & _
      '   "ProductId: "& ProductId & "<br/>" & _
      '   "SKU: "& SKU & "<br/>" & _
      '   "ProductName: "& ProductName & "<br/>" & _
      '   "Curency: "& Curency & "<br/>" & _
      '   "Amount: "& Amount & "<br/>" & _
      '   "Tax: "& Tax & "<br/>" & _
      '   "Quantity: "& Quantity & "<br/>" & _
      '   "")

      Set newOrderItem = SetOrderItem(Line, ProductId, SKU, ProductName, Curency, Amount, Tax, Quantity)
      newOrderItems.Add Counter, newOrderItem

      Counter = Counter + 1
      rsOrder.MoveNext
    Loop
    ' Response.Write "<h1>Database Values</h1>"

    ' RW("SubTotal: "& FormatNumber(dbSubTotal, 2) & "<br/>" & _
    '   "Shipping: "& FormatNumber(dbShipping, 2) & "<br/>" & _
    '   "Discount: "& FormatNumber(dbDiscount, 2) & "<br/>" & _
    '   "Tax: "& FormatNumber(dbTax, 2) & "<br/>" & _
    '   "Total: "& FormatNumber(dbTotal, 2) & "<br/>" & _
    '   "")

    ' RW("FullName: "& FullName & "<br/>" & _
    '   "Login: "& Login & "<br/>" & _
    '   "UserId: "& UserId & "<br/>" & _
    '   "StripeCustomerId: "& StripeCustomerId & "<br/>" & _
    '   "Email: "& Email & "<br/>" & _
    '   "Phone: "& Phone & "<br/>" & _
    '   "")

    ' RW("ShipAddrContact: "& ShipAddrContact & "<br/>" & _
    '   "ShipAddrCompany: "& ShipAddrCompany & "<br/>" & _
    '   "ShipAddrLine1: "& ShipAddrLine1 & "<br/>" & _
    '   "ShipAddrLine2: "& ShipAddrLine2 & "<br/>" & _
    '   "ShipAddrCity: "& ShipAddrCity & "<br/>" & _
    '   "ShipAddrState: "& ShipAddrState & "<br/>" & _
    '   "ShipAddrZip: "& ShipAddrZip & "<br/>" & _
    '   "ShipAddrCountry: "& ShipAddrCountry & "<br/>" & _
    '   "UseShippingAddrForBillingAddr: "& UseShippingAddrForBillingAddr & "<br/>" & _
    '   "")
  Else
    RW("No Order Found")
    Response.End
  End If
  Set rsOrder = Nothing
%>