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

  ' stripe_mode = Application("stripe_mode")
  ' If stripe_mode = "TEST" Then
  '   secret_key = Application("stripe_test_sk")
  '   public_key = Application("stripe_test_sk")
  ' Else
  '   secret_key = Application("stripe_live_sk")
  '   public_key = Application("stripe_live_sk")
  ' End If
  
  stripe_mode = "TEST"
  secret_key = "sk_test_51N4RDzAefTrv2FbjdmD3EvTuTVA2iZM3lo5s1Hb4uYXQnj5p6o19tKp7DqtF4n4QQrchjkTDw9xlnClhaLPw8zsf0068pHpZAA"

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
    <title>Old Modern Handicrafts | Payment Success</title>
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
        <h2>Payment: SUCCESS</h2>
        <div style="display: flex;justify-content:center;gap:1rem;">
          <button id="invoice_pdf">Download Invoice</button>
          <button id="invoice_online">View Invoice/Receipt</button>
        </div>
<%
  '// Get Session data from stripe to validate cancelled
    StripeSessionResult = GetStripeSession(secret_key, StripeSessionId)
    If Len(StripeSessionResult) >= 1 Then
      Set ojSession = new aspJSON
      ojSession.loadJSON(StripeSessionResult)
      redirUrl = ojSession.data("url")
      PaymentStatus = ojSession.data("payment_status")
      StripeInvoiceId = ojSession.data("invoice")
      status = ojSession.data("status")
      Set ojSession = Nothing

      StripeInvoiceResult = GetStripeInvoice(secret_key, StripeInvoiceId)
      If Len(StripeInvoiceResult) >= 1 Then
        Set ojInvoice = new aspJSON
        ojInvoice.loadJSON(StripeInvoiceResult)
        StripeOnlineInvoice = ojInvoice.data("hosted_invoice_url")
        StripePDFInvoice = ojInvoice.data("invoice_pdf")
        StripeChargeId = ojInvoice.data("charge")
        RW("Payment Status: " & PaymentStatus & "<br/>" & _
        "Payment Reference: " & StripeChargeId & "<br/>" & _
        "")
        Set ojInvoice = Nothing
      ' Else
      End If
      
    ' Else
      
    End If

  '// Print out Customer/Order details from newOrder object  
    PrintCustomer(newOrder.Customer) 
    PrintOrder(newOrder)

  End If  '// Len(OrderId) >= 1 And Len(StripeSessionId) >= 1
%>
        <div style="display: flex;justify-content:center;gap:1rem;">
          <a href="<%=StripePDFInvoice%>" target="_blank">Download Invoice</a>
          <a href="<%=StripeOnlineInvoice%>" target="_blank">View Invoice/Receipt</a>
        </div>
      </div>
    </div>
    <script>
    document.addEventListener("DOMContentLoaded", () =>{
      const invoice_pdf_url = "<%=StripePDFInvoice%>";
      const invoice_online_url = "<%=StripeOnlineInvoice%>";
      const downlownBtn = document.querySelector("#invoice_pdf");
      const viewBtn = document.querySelector("#invoice_online");
      downlownBtn.addEventListener("click", (e) => {
        e.preventDefault();
        window.open(invoice_pdf_url, "_blank");
      });
      viewBtn.addEventListener("click", (e) => {
        e.preventDefault();
        window.open(invoice_online_url, "_blank");
      });
    });
    </script>

</body>
</html>
<%
  ' RW(StripeSessionResult)
  ' RW(StripeInvoiceResult)
%>