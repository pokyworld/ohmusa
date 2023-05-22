<%@ Language=VBScript %>
<!--#include file="../../include/asp_lib.inc.asp" -->
<!--#include file="../../include/sqlCheckInclude.asp" -->
<!--#include virtual="/shop/payment/classes/aspJSON1.17.asp"-->
<!--#include virtual="/shop/payment/functions/helpers.inc"-->
<!--#include virtual="/shop/payment/classes/customer.asp"-->
<!--#include virtual="/shop/payment/classes/payment.asp"-->
<!--#include virtual="/shop/payment/classes/product.asp"-->
<!--#include virtual="/shop/payment/classes/price.asp"-->
<!--#include virtual="/shop/payment/classes/orderItem.asp"-->
<!--#include virtual="/shop/payment/classes/order.asp"-->
<!--#include virtual="/shop/payment/classes/address.asp"-->
<!--#include virtual="/shop/payment/classes/emailMsg.asp"-->

<%
  Response.Expires=0
  Response.CacheControl= "no-cache"
  Response.AddHeader "Pragma", "no-cache"

  Application("MAIL_FROM") = "orders@omhusa.com"
  Application("MAIL_CC") = "eric@omhusa.com"
  Application("MAIL_DISPLAY_NAME") = "OMH USA Orders"
  Application("MAIL_MAILER") = "smtp"
  Application("MAIL_HOST") = "mail.omhusa.com"
  Application("MAIL_PORT") = "587"
  Application("MAIL_USERNAME") = "service@omhusa.com"
  Application("MAIL_PASSWORD") = "OMH3750SG@$"
  Application("MAIL_ENCRYPTION") = "tls"

  ' RW("MAIL_FROM: " & Application("MAIL_FROM"))
  ' RW("MAIL_DISPLAY_NAME: " & Application("MAIL_DISPLAY_NAME"))
  ' RW("MAIL_MAILER: " & Application("MAIL_MAILER"))
  ' RW("MAIL_HOST: " & Application("MAIL_HOST"))
  ' RW("MAIL_PORT: " & Application("MAIL_PORT"))
  ' RW("MAIL_USERNAME: " & Application("MAIL_USERNAME"))
  ' RW("MAIL_ENCRYPTION: " & Application("MAIL_ENCRYPTION"))

  If Len(Trim(Request.QueryString("id")&"")) >= 1 Then : StripeSessionId = Trim(Request.QueryString("id")&"")
  If Len(Trim(Request.QueryString("orderid")&"")) >= 1 Then : OrderId = Trim(Request.QueryString("orderid")&"")
  If Len(Trim(Request.QueryString("userid")&"")) >= 1 Then : UserId = Trim(Request.QueryString("userid")&"")

  If Len(OrderId) = 0 Then : RW("No Order ID") : Response.End

  stripe_mode = UCase(Trim(Application("stripe_mode")&""))
  If stripe_mode = "TEST" Then
    secret_key = Trim(Application("stripe_test_sk")&"")
    public_key = Trim(Application("stripe_test_sk")&"")
  Else
    secret_key = Trim(Application("stripe_live_sk")&"")
    public_key = Trim(Application("stripe_live_sk")&"")
  End If

  ' If Len(OrderId) >= 1 And Len(StripeSessionId) >= 1 Then
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

<%
  '// Get Session data from stripe to validate cancelled
    StripeSessionResult = GetStripeSession(secret_key, StripeSessionId)
    If Len(StripeSessionResult) >= 1 Then
      ' RW(StripeSessionResult)
      Set ojSession = new aspJSON
      ojSession.loadJSON(StripeSessionResult)
      redirUrl = ojSession.data("url")
      paymentStatus = ojSession.data("payment_status")
      status = ojSession.data("status")
      ' RW("Payment Status: " & paymentStatus)
      Set ojSession = Nothing
    End If
%>

<%
    result = CreatePaymentEmail(newOrder, StripeCustomerId, null)
    RW(result)
    
    CCList = Application("MAIL_CC")
    result = CreatePaymentEmail(newOrder, StripeCustomerId, CCList)


  '// Print out Customer/Order details from newOrder object  
    ' PrintCustomer(newOrder.Customer) 
    ' PrintOrder(newOrder)

  ' End If  '// Len(OrderId) >= 1 And Len(StripeSessionId) >= 1

  ' RW(StripeSessionResult)

%>
