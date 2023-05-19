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

<%
  Response.Expires=0
  Response.CacheControl= "no-cache"
  Response.AddHeader "Pragma", "no-cache"

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

  MAIL_MAILER = Application("MAIL_MAILER")
  MAIL_HOST = Application("MAIL_HOST")
  MAIL_PORT = Application("MAIL_PORT")
  MAIL_USERNAME = Application("MAIL_USERNAME")
  MAIL_PASSWORD = Application("MAIL_PASSWORD")
  MAIL_ENCRYPTION = Application("MAIL_ENCRYPTION")
  
  If Len(OrderId) >= 1 And Len(StripeSessionId) >= 1 Then
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
    <title>Old Modern Handicrafts | Payment Aborted</title>
    <style>
      * { font-family: sans-serif;}
      a { text-decoration: none; font-weight:600;font-size:0.9em;}
      a:hover{ text-decoration: underline;}
    </style>
  </head>
  <body>
    <div style="display:flex;padding:1rem;justify-content:center;">
      <div style="display:flex;flex:1;flex-direction:column;align-items:flex-start;max-width: 800px;">
        <h1>Order: <%=OrderId%></h1>
        <h2>Payment: UNPAID/ABORTED</h2>
        <div style="display: flex;justify-content:center;gap:1rem;">
          <button id="redir">Retry Payment</button>
        </div>
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
      ' RW(redirUrl)
      ' RW("Session Status: " & status)
      RW("Payment Status: " & paymentStatus)
      Set ojSession = Nothing
    ' Else
      
    End If
  '// Print out Customer/Order details from newOrder object  
    PrintCustomer(newOrder.Customer) 
    PrintOrder(newOrder)

  End If  '// Len(OrderId) >= 1 And Len(StripeSessionId) >= 1

%>
      </div>
    </div>
    <script>
    document.addEventListener("DOMContentLoaded", () =>{
      const payBtn = document.querySelector("#redir");
      payBtn.addEventListener("click", (e) => {
        e.preventDefault();
        window.location = redir;
      });
      const redir = "<%=redirUrl%>";
      console.log(redir);
      setTimeout(() => {
      //window.location = redir;
      },5000);
      // window.location = redir;
      // window.open(redir, "_blank");
    });
    </script>
  </body>
</html>
<%
  ' RW(StripeSessionResult)



%>
