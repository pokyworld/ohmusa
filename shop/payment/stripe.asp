<%@ Language=VBScript %>
<!-- #include file="../../include/asp_lib.inc.asp" -->
<!-- #include file="../../include/sqlCheckInclude.asp" -->
<!--#include virtual="/shop/payment/classes/aspJSON1.17.asp"-->
<!--#include virtual="/shop/payment/functions/helpers.inc"-->
<!--#include virtual="/shop/payment/classes/customer.asp"-->
<!--#include virtual="/shop/payment/classes/payment.asp"-->
<!--#include virtual="/shop/payment/classes/product.asp"-->
<!--#include virtual="/shop/payment/classes/price.asp"-->
<!--#include virtual="/shop/payment/classes/orderItem.asp"-->
<!--#include virtual="/shop/payment/classes/order.asp"-->
<!--#include virtual="/shop/payment/classes/address.asp"-->
<%
  Response.Expires=0
  Response.CacheControl= "no-cache"
  Response.AddHeader "Pragma", "no-cache"

  OrderId = Trim(Request("orderid")&"")
  If Len(OrderId) = 0 Then : RW("No Order ID") : Response.End

  stripe_mode = UCase(Trim(Application("stripe_mode")&""))
  If stripe_mode = "TEST" Then
    secret_key = Trim(Application("stripe_test_sk")&"")
    public_key = Trim(Application("stripe_test_sk")&"")
  Else
    secret_key = Trim(Application("stripe_live_sk")&"")
    public_key = Trim(Application("stripe_live_sk")&"")
  End If
  
%>
<!--// order functions -->
<!--#include virtual="/shop/payment/functions/orders.inc"-->

<!--// fetch data //-->
<!--#include virtual="/shop/payment/functions/_data.inc"-->
<!--// build customer object //-->
<!--#include virtual="/shop/payment/functions/_customer.inc"-->
<!--// build payment object //-->
<!--#include virtual="/shop/payment/functions/_payment.inc"-->
<!--// build order object //-->
<!--#include virtual="/shop/payment/functions/_order.inc"-->

<!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Old Modern Handicrafts | Payment Success</title>
    <style>
      * { font-family: sans-serif;}
      a { text-decoration: none; font-weight:600;font-size:0.9em;}
      a:hover{ text-decoration: underline;}
      html, body, #image-wrapper { display:flex; flex-direction:column; flex: 1; min-height:100vh; width: 100%;}
      #image-wrapper { display:flex;padding:1rem;justify-content:center;align-items:center;}
      #image-wrapper img { height:5rem; width:auto; }
    </style>
  </head>
  <body>
    <div id="image-wrapper">
      <img src="/shop/payment/assets/images/loading.gif" alt="Loading.." />

<%
  '// Print out Customer/Order details from newOrder object  
    ' PrintCustomer(newOrder.Customer) 
    ' PrintOrder(newOrder) 

  '// Request Body starts here

  reqBody = "mode=payment"
  reqBody = reqBody & "&client_reference_id=" & newOrder.OrderId
  reqBody = reqBody & "&invoice_creation[enabled]=true"

  '// Customer
  If Len(newOrder.Customer.StripeCustomerId) >= 1 Then
    reqBody = reqBody & "&customer=" & newOrder.Customer.StripeCustomerId 
  Else
    If Len(newOrder.Customer.Email) >= 1 Then : reqBody = reqBody & "&customer_email=" & newOrder.Customer.Email
  End If
  ' If Len(newOrder.Customer.FullName) >= 1 Then : reqBody = reqBody & "&customer_details[name]=" & newOrder.Customer.FullName
  ' If Len(newOrder.Customer.Phone) >= 1 Then : reqBody = reqBody & "&customer_details[phone]=" & newOrder.Customer.Phone

  '// Order Lines
  For Each Item In newOrder.OrderItems
    Set newOrderItem = newOrder.OrderItems(Item)
    counter = Item - 1
    reqBody = reqBody & "&line_items[" & counter & "][quantity]=" & newOrderItem.Quantity
    reqBody = reqBody & "&line_items[" & counter & "][price_data][currency]=usd"
    reqBody = reqBody & "&line_items[" & counter & "][price_data][unit_amount]=" & newOrderItem.Price.SubTotal
    reqBody = reqBody & "&line_items[" & counter & "][price_data][product_data][name]=" & newOrderItem.Price.Product.Name
    ' reqBody = reqBody & "&line_items[" & counter & "][id]=" & newOrderItem.Price.Product.ID
    ' reqBody = reqBody & "&line_items[" & counter & "][price_data][product_data][sku]=TEES0001REDXL"
  Next

  '// Discounts
  If newOrder.PromoDiscount >= 1 Or newOrder.Discount >= 1 Then
    CouponId = "DISC_" & OrderId
    If Len(newOrder.PromoCode) >= 1 Then : CouponName = newOrder.PromoCode : Else : CouponName = "Discount"
    TotalDiscount = newOrder.PromoDiscount + newOrder.Discount
    Coupon = CreateCoupon(secret_key, CouponId, CouponName, newOrder.Curency,TotalDiscount)
    If Len(Coupon) = 0 Then : Coupon = CouponId
    reqBody = reqBody & "&discounts[0]coupon=" & Coupon
  End If

  '// Shipping
  ' reqBody = reqBody & "&shipping_address_collection[allowed_countries][0]=US"
  ' reqBody = reqBody & "&shipping_address_collection[allowed_countries][1]=CA"

  ' reqBody = reqBody & "&shipping_address_collection[shipping_address][recipient]=" & newOrder.Customer.FullName
  ' reqBody = reqBody & "&shipping_address_collection[shipping_address][addressLine][0]=" & newOrder.Customer.BillingAddress.Line1
  ' If Len(newOrder.Customer.BillingAddress.Line2) >= 1 Then : reqBody = reqBody & "&shipping_address_collection[shipping_address][addressLine][1]=" & newOrder.Customer.BillingAddress.Line2
  ' reqBody = reqBody & "&shipping_address_collection[shipping_address][city]=" & newOrder.Customer.BillingAddress.City
  ' reqBody = reqBody & "&shipping_address_collection[shipping_address][region]=" & newOrder.Customer.BillingAddress.State
  ' reqBody = reqBody & "&shipping_address_collection[shipping_address][postalCode]=" & newOrder.Customer.BillingAddress.Zip
  ' reqBody = reqBody & "&shipping_address_collection[shipping_address][country]=" & newOrder.Customer.BillingAddress.Country

  reqBody = reqBody & "&shipping_options[0][shipping_rate_data][type]=fixed_amount"
  reqBody = reqBody & "&shipping_options[0][shipping_rate_data][fixed_amount][currency]=usd"
  reqBody = reqBody & "&shipping_options[0][shipping_rate_data][fixed_amount][amount]=" & newOrder.Shipping
  If newOrder.Shipping = 0 Then
    reqBody = reqBody & "&shipping_options[0][shipping_rate_data][display_name]=Free shipping"
  Else
    reqBody = reqBody & "&shipping_options[0][shipping_rate_data][display_name]=Shipping Cost"
  End If

  '// Optional Metadata Fields
  reqBody = reqBody & "&metadata[order_id]=" & newOrder.OrderId
  If Len(newOrder.Customer.Email) >= 1 Then : reqBody = reqBody & "&metadata[email]=" & newOrder.Customer.Email

  '// Gonna be a nightmare: Tax
  ' reqBody = reqBody & "&tax_rates[0][display_name]=Sales Tax"
  ' reqBody = reqBody & "&tax_rates[0][inclusive]=false"
  ' reqBody = reqBody & "&tax_rates[0][percentage]=10"

  '// Payment Types
  reqBody = reqBody & "&payment_method_types[0]=card"
  ' reqBody = reqBody & "&payment_method_types[1]=link"
  ' reqBody = reqBody & "&payment_method_types[2]=google"
  ' reqBody = reqBody & "&payment_method_types[3]=apple"

  '// Redirection Urls
  reqBody = reqBody & "&success_url=" & Server.URLEncode("http://" & Request.ServerVariables("SERVER_NAME") & "/shop/payment/success.asp?id={CHECKOUT_SESSION_ID}&orderid=" & OrderId & "&userid=" & newOrder.Customer.UserId)
  reqBody = reqBody & "&cancel_url=" & Server.URLEncode("http://" & Request.ServerVariables("SERVER_NAME") & "/shop/payment/aborted.asp?id={CHECKOUT_SESSION_ID}&orderid=" & OrderId & "&userid=" & newOrder.Customer.UserId )

  ' RW(reqBody)
  
  reqUrl = "https://api.stripe.com/v1/checkout/sessions"

  result = PostStripe(secret_key, reqUrl, reqBody)

  Set ojStripeSession = new aspJSON
  ojStripeSession.loadJSON(result)
  ' RW(ojStripeSession.JSONoutput)
  redirUrl = ojStripeSession.data("url")
  Set ojStripeSession = Nothing

  BalanceDue = Round(newOrder.Total/100,2) - Round(newOrder.Payment.Amount/100,2) 
  If BalanceDue = 0 Then
    StripeSessionId = GetLastOrderPayment(orderid)
    params = "id=" & StripeSessionId & "&orderid=" & OrderId & "&userid=" & UserId
    paidRedirUrl = "/shop/payment/success.asp?" & params
  End If
  '// doesn't work; don't know why; using JS instead
  ' Response.Redirect redirUrl
%>
    </div>
  <script>
  document.addEventListener("DOMContentLoaded", () =>{
    const redir = "<%=redirUrl%>";
    const balance = "<%=BalanceDue%>";
    const paidRedir = "<%=paidRedirUrl%>";
    if(balance > 0) {
      window.location = redir; 
    } else {
      // window.location = "/shop/orderconfirmRetail.asp";
      window.location = "<%=paidRedir%>";
    }
  });
  </script>
</body>
</html>
