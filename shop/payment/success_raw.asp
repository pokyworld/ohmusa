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
      SessionPaymentStatus = ojSession.data("payment_status")
      StripeCustomerId = ojSession.data("customer")
      StripeInvoiceId = ojSession.data("invoice")
      StripePaymentId = ojSession.data("payment_intent")
      StripeSuccessUrl = Replace(ojSession.data("success_url"), "{CHECKOUT_SESSION_ID}", StripeSessionId)
      StripeCancelUrl = Replace(ojSession.data("success_url"), "{CHECKOUT_SESSION_ID}", StripeSessionId)
      RW("StripeSuccessUrl: " & StripeSuccessUrl)
      RW("StripeCancelUrl: " & StripeCancelUrl)
      status = ojSession.data("status")
      Set ojSession = Nothing

      StripeInvoiceResult = GetStripeInvoice(secret_key, StripeInvoiceId)
      If Len(StripeInvoiceResult) >= 1 Then
        Set ojInvoice = new aspJSON
        ojInvoice.loadJSON(StripeInvoiceResult)
        StripeOnlineInvoice = ojInvoice.data("hosted_invoice_url")
        StripePDFInvoice = ojInvoice.data("invoice_pdf")
        StripeChargeId = ojInvoice.data("charge")
        RW("Payment Status: " & SessionPaymentStatus & "<br/>" & _
        "Payment Reference: " & StripeChargeId & "<br/>" & _
        "")
        Set ojInvoice = Nothing
      End If

      StripePaymentResult = GetStripePayment(secret_key, StripePaymentId)
      If Len(StripePaymentResult) >= 1 Then
        Set ojPayment = new aspJSON
        ojPayment.loadJSON(StripePaymentResult)
        ' RW(ojPayment.JSONoutput)
        StripeChargeId = ojPayment.data("latest_charge")
        PaymentStatus = ojPayment.data("status")
        ' RW("SessionPaymentStatus: " & SessionPaymentStatus & "<br/>" & _
        '   "PaymentStatus: " & PaymentStatus & "<br/>" & _
        '   "StripeSessionId: " & StripeSessionId & "<br/>" & _
        '   "StripeCustomerId: " & StripeCustomerId & "<br/>" & _
        '   "StripePaymentId: " & StripePaymentId & "<br/>" & _
        '   "StripeInvoiceId: " & StripeInvoiceId & "<br/>" & _
        '   "StripeChargeId: " & StripeChargeId & "<br/>" & _
        '   "StripePDFInvoice: " & StripePDFInvoice & "<br/>" & _
        '   "StripeOnlineInvoice: " & StripeOnlineInvoice & "<br/>" & _
        '   "")
        Set ojPayment = Nothing
      End If
      
      ' StripeChargeResult = GetStripeCharge(secret_key, StripeChargeId)
      ' If Len(StripeChargeResult) >= 1 Then
      '   Set ojCharge = new aspJSON
      '   ojCharge.loadJSON(StripeChargeResult)
      '   StripeChargeId = ojCharge.data("latest_charge")
      '   Set ojCharge = Nothing
      ' End If
      
      If Len(Trim(userid)) > 0 Then
        StripeId = UpdateStripeId(UserId, stripe_mode, StripeCustomerId)
      End If
      ' RW("StripeId: " & StripeId)
      TransId = InsertStripeKeys(orderid, stripe_mode, StripeSessionId, StripeCustomerId, StripePaymentId, StripeInvoiceId, stripeChargeId, StripePDFInvoice, StripeOnlineInvoice)
      ' RW("TransId: " & TransId)
      PaymentId = UpdateStripePaymentId(OrderId, StripeChargeId)
      ' RW("PaymentId: " & PaymentId)

    End If

  '// Print out Customer/Order details from newOrder object  
    ' PrintCustomer(newOrder.Customer) 
    ' PrintOrder(newOrder)

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
  RW(StripeSessionResult)
  RW(StripeInvoiceResult)
  ' RW(StripePaymentResult)
  ' RW(StripeChargeResult)
  
%>