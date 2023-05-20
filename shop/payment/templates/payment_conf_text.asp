<%@ Language=VBScript %>
<!--#include virtual="/include/asp_lib.inc.asp" -->
<!--#include virtual="/include/sqlCheckInclude.asp" -->
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

  If Len(Trim(Request.QueryString("orderid")&"")) >= 1 Then : OrderId = Trim(Request.QueryString("orderid")&"")

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
  <!--#include virtual="/shop/payment/functions/orders.inc"-->
  <!--#include virtual="/shop/payment/functions/_data.inc"-->
  <!--#include virtual="/shop/payment/functions/_customer.inc"-->
  <!--#include virtual="/shop/payment/functions/_payment.inc"-->
  <!--#include virtual="/shop/payment/functions/_order.inc"-->
<%
  sql = "SELECT TOP 1 Max(id) AS id, orderid, stripeInvoicePDF, stripeInvoiceReceipt "
  sql = sql & "FROM dbo.stripeTransactions "
  sql = sql & "WHERE orderid = " & OrderId & " "
  sql = sql & "AND stripeInvoicePDF IS NOT NULL "
  sql = sql & "AND stripeInvoiceReceipt IS NOT NULL "
  sql = sql & "GROUP BY orderid, stripeInvoicePDF, stripeInvoiceReceipt "
  sql = sql & "ORDER BY Max(id) DESC;"
    
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  cnn.Open
  Set rsStripe = Server.CreateObject("ADODB.Recordset")
  rsStripe.Open sql, cnn, 3

  If Not rsStripe.BOF And Not rsStripe.EOF Then
    StripeInvoicePDF = Trim(rsStripe("stripeInvoicePDF")&"")
    StripeInvoiceReceipt = Trim(rsStripe("stripeInvoiceReceipt")&"")
  End If
  Set rsStripe = Nothing

  '// Print out Customer/Order details from newOrder object  
    ' PrintCustomer(newOrder.Customer) 
    ' PrintOrder(newOrder)

%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>

<body>
  <style>
    html,
    body {
      display: flex;
      flex-direction: column;
      flex: 1;
      width: 100%;
      min-height: 100vh;
      margin: 0;
      padding: 0;
    }

    * {
      font-family: sans-serif;
    }

    @media screen and (max-width: 600px) {
      .mainTable .category {
        display: none;
      }
    }

    .button-group {
      display: flex;
      justify-content: center;
      gap: 1rem;
      margin-bottom: 1rem;
    }

    .stripe-result {
      padding: 0 2rem;
    }

    @media screen and (max-width: 800px) {

      .stripe-result {
        padding: 0.25rem 0 0.5rem;
      }
    }

    .stripe-result .table-wrapper {
      display: flex;
      flex-grow: 1;
      flex: 1;
      width: 100%;
      flex-wrap: wrap;
      justify-content: space-between;
      gap: 2rem;
      padding-bottom: 1rem;
    }

    @media screen and (max-width: 800px) {

      .stripe-result .table-wrapper {
        gap: 1rem;
        margin-left: -0.25rem;
      }
    }

    .stripe-result .table-wrapper hr {
      border-top: solid 1px lightgray;
      margin-left: -3px;
      margin-right: -3px;
      border-collapse: collapse;
    }

    .stripe-result .table-wrapper .tbl-address,
    .stripe-result .table-wrapper .tbl-contact,
    .stripe-result .table-wrapper .tbl-totals {
      display: flex;
      flex: 1;
    }

    .stripe-result .table-wrapper .tbl-address th,
    .stripe-result .table-wrapper .tbl-contact th,
    .stripe-result .table-wrapper .tbl-totals th {
      text-align: left;
    }

    .stripe-result .table-wrapper .tbl-address td,
    .stripe-result .table-wrapper .tbl-contact td,
    .stripe-result .table-wrapper .tbl-totals td {
      padding: 0.1px 0.5rem;
    }

    @media screen and (max-width: 450px) {

      .stripe-result .table-wrapper .tbl-address td,
      .stripe-result .table-wrapper .tbl-contact td,
      .stripe-result .table-wrapper .tbl-totals td {
        font-size: 0.9em;
        width: 90vw;
      }
    }

    .stripe-result .table-wrapper .tbl-address td span,
    .stripe-result .table-wrapper .tbl-contact td span,
    .stripe-result .table-wrapper .tbl-totals td span {
      font-weight: 600;
      margin-right: 1rem;
    }

    .stripe-result .table-wrapper .tbl-totals {
      justify-content: flex-end;
      flex: 1;
      width: 50%;
    }

    @media screen and (max-width: 401px) {

      .stripe-result .table-wrapper .tbl-totals {
        width: 100%;
      }
    }

    .stripe-result .table-wrapper .tbl-totals tr.border td,
    .stripe-result .table-wrapper .tbl-totals tr.border th {
      border-top: solid 2px lightgray;
      border-bottom: solid 2px lightgray;
    }

    .stripe-result .table-wrapper .order-lines,
    .stripe-result .table-wrapper .tbl-totals,
    .stripe-result .table-wrapper .tbl-payment {
      width: 100%;
      padding: 0 0.5rem;
    }

    .stripe-result .table-wrapper .order-lines th,
    .stripe-result .table-wrapper .order-lines td,
    .stripe-result .table-wrapper .tbl-totals th,
    .stripe-result .table-wrapper .tbl-totals td,
    .stripe-result .table-wrapper .tbl-payment th,
    .stripe-result .table-wrapper .tbl-payment td {
      padding: 0.2rem 0.25rem;
      text-align: left;
      vertical-align: middle;
    }

    .stripe-result .table-wrapper .order-lines th.right,
    .stripe-result .table-wrapper .order-lines td.right,
    .stripe-result .table-wrapper .tbl-totals th.right,
    .stripe-result .table-wrapper .tbl-totals td.right,
    .stripe-result .table-wrapper .tbl-payment th.right,
    .stripe-result .table-wrapper .tbl-payment td.right {
      text-align: right;
    }

    .stripe-result .table-wrapper .order-lines th.center,
    .stripe-result .table-wrapper .order-lines td.center,
    .stripe-result .table-wrapper .tbl-totals th.center,
    .stripe-result .table-wrapper .tbl-totals td.center,
    .stripe-result .table-wrapper .tbl-payment th.center,
    .stripe-result .table-wrapper .tbl-payment td.center {
      text-align: center;
    }

    .stripe-result .table-wrapper .order-lines th img,
    .stripe-result .table-wrapper .order-lines td img,
    .stripe-result .table-wrapper .tbl-totals th img,
    .stripe-result .table-wrapper .tbl-totals td img,
    .stripe-result .table-wrapper .tbl-payment th img,
    .stripe-result .table-wrapper .tbl-payment td img {
      height: 40px;
      width: auto;
    }
    
    /* .banner { 
      display: flex;
      flex: 1;
      flex-grow: 1;
      background-image: url(https://omhusa.com/amazingslider/jessica/1.jpg); 
      background-repeat:no-repeat;
      background-size: contain;
      background-position: center;
      min-height: 100px;
      margin-bottom: 0px;
    } */

    .banner img {
      width: 100%;
      height: auto;
    }

    @media screen and (min-width: 601px) {
      .stripe-result {
        padding: 2rem;
        margin-bottom: 2rem;
      }
      .banner-row {
        /* display: none; */
      }
      table, tr, td, th {
        max-width: 600px;
      }

    }

    @media (max-width: 600px) {
      .stripe-result {
        padding: 1rem 0;
        margin-bottom: 2rem;
      }
      .banner-row {
        /* display: table-row; */
      }
      table, tr, td, th {
        width: 100vw;
      }
    }
    .thfeatured {
      background-color: darkslateblue;
      color: white;
      padding: 0.5rem;
    }
    .link-group {
      display:flex;justify-content:center;gap:1rem;
    }
    .footer-message {
      padding: 1rem;
      font-size: 0.9em;
    }
  </style>
  <div class="stripe-result">
    <table>
      <tbody>
        <!-- <tr class="banner-row">
          <td>
            <div class="banner"></div>
          </td>
        </tr> -->
        <tr class="banner-row">
          <td>
            <div class="banner">
              <img src="https://omhusa.com/amazingslider/jessica/1.jpg" />
            </div>
          </td>
        </tr>
        <tr>
          <th class="thfeatured">
            Payment SUCCESS: (Order <%=OrderId%>)
          </th>
        </tr>
        <tr>
          <td style="padding: 0;">
            <div class="table-wrapper">
              <table class="tbl-contact">
                <tbody>
                  <tr>
                    <th>Contact Details</th>
                  </tr>
                  <tr>
                    <td><span>Name:</span><%=newOrder.Customer.FullName%></td>
                  </tr>
                  <tr>
                    <td><span>Email:</span><%=newOrder.Customer.Email%></td>
                  </tr>
                  <tr>
                    <td><span>Phone:</span><%=newOrder.Customer.Phone%></td>
                  </tr>
                </tbody>
              </table>
              <table class="tbl-address">
                <tbody>
                  <tr>
                    <th>Address</th>
                  </tr>
                  <tr>
                    <td><%=newOrder.Customer.BillingAddress.Line1%></td>
                  </tr>
<%                  
  If Len(Trim(newOrder.Customer.BillingAddress.Line2)) = 0 Or newOrder.Customer.BillingAddress.Line2 = newOrder.Customer.BillingAddress.Line1 Then
  Else
%>
                  <tr>
                    <td nowrap><%=newOrder.Customer.BillingAddress.Line2%></td>
                  </tr>
<%
  End If
%>                  
                  <tr>
                    <td><%=newOrder.Customer.BillingAddress.City%></td>
                  </tr>
                  <tr>
                    <td><%=newOrder.Customer.BillingAddress.State%>, <%=newOrder.Customer.BillingAddress.Zip%></td>
                  </tr>
                  <tr>
                    <td nowrap><%=newOrder.Customer.BillingAddress.Country%></td>
                  </tr>
                </tbody>
              </table>
            </div>
            <hr>
            <div class="table-wrapper">
              <table class="order-lines">
                <div>
                  <tr>
                    <!-- <th>&nbsp;</th> -->
                    <th>SKU</th>
                    <th>Product</th>
                    <th class="center">Quantity</th>
                    <th class="right">Amount</th>
                  </tr>
<%
  For Each Item in newOrder.OrderItems
    Set newOrderItem = newOrder.OrderItems(Item)
%>
                  <tr>
                    <!--<td><img src="<%=newOrderItem.Price.Product.ImageUrl%>" alt=""></td>-->
                    <td><%=newOrderItem.Price.Product.SKU%></td>
                    <td><%=newOrderItem.Price.Product.Name%></td>
                    <td class="center"><%=newOrderItem.Quantity%></td>
                    <td class="right">$&nbsp;<%=FormatNumber(newOrderItem.SubTotal/100,2)%></td>
                  </tr>
<%
  Next
%>              
                </div>
              </table>
            </div>
            <hr>
            <div class="table-wrapper">
              <table class="tbl-totals">
                <tr>
                  <td><span>Net SubTotal:</span></td>
                  <td class="right">$&nbsp;<%=FormatNumber(newOrder.NetSubTotal/100,2)%></td>
                </tr>
<%
    If newOrder.PromoDiscount >= 1 Then
%>              
                <tr>
                  <td><span><%=newOrder.PromoCode%>:</span></td>
                  <td class="right">$&nbsp;(<%=FormatNumber(newOrder.PromoDiscount/100,2)%>)</td>
                </tr>
<%
    End If

    If newOrder.Discount >= 1 Then
%>              
                <tr>
                  <td><span>Discount:</span></td>
                  <td class="right">$&nbsp;(<%=FormatNumber(newOrder.Discount/100,2)%>)</td>
                </tr>
<%
    End If
%>
                <tr>
                  <td><span>Shipping:</span></td>
                  <td class="right">$&nbsp;<%=FormatNumber(newOrder.Shipping/100,2)%></td>
                </tr>
                <tr>
                  <td><span>SubTotal:</span></td>
                  <td class="right">$&nbsp;<%=FormatNumber(newOrder.SubTotal/100,2)%></td>
                </tr>
                <!--<tr>
                  <td><span>Tax(es):</span></td>
                  <td class="right">$&nbsp;<%=FormatNumber(newOrder.Tax/100,2)%></td>
                </tr>-->
                <tr class="border">
                  <td><span>Total:</span></td>
                  <td class="right bold">$&nbsp;<%=FormatNumber(newOrder.Total/100,2)%></td>
                </tr>
                <tr><td colspan="2">&nbsp;</td></tr>
                <tr>
                  <td><span>Purchase Order:</span>&nbsp;</td>
                  <td id="payment-ref" class="right"><%=newOrder.Payment.PurchaseOrder%></td>
                </tr>
                <tr>
                  <td><span>Payment Ref:</span>&nbsp;</td>
                  <td id="payment-ref" class="right"><%=newOrder.Payment.Id%></td>
                </tr>
                <tr>
                  <td><span>Payment Status:</span></td>
                  <td class="right"><%=UCase(newOrder.Payment.Status)%></td>
                </tr>
                <tr>
                  <td><span>Amount Paid:</span></td>
                  <td class="right">$&nbsp;<%=FormatNumber(newOrder.Payment.Amount/100,2)%></td>
                </tr>
<%
    BalanceDue = Round((newOrder.Total / 100), 2) - Round((newOrder.Payment.Amount / 100), 2)
%>  
                <tr class="border">
                  <td><span>Balance Due:</span></td>
                  <td class="right bold">$&nbsp;<%=FormatNumber(BalanceDue, 2)%></td>
                </tr>
              </table>
            </div>
            <div class="link-group">
<%
        If Len(StripeInvoicePDF) >= 1 Then
%>
              &nbsp;<a href="<%=StripeInvoicePDF%>" target="_blank">Download Invoice</a>
<%
        End If
            
        If Len(StripeInvoiceReceipt) >= 1 Then
%>
              &nbsp;<a href="<%=StripeInvoiceReceipt%>" target="_blank">View Invoice/Receipt</a>
<%
        End If
%>  
            </div>
            <div class="footer-message">Thank you for your order with OMH. <br/>If you have any questions please contact us at (909) 598 2525<p>
          </td>
        </tr>
      </tbody>
    </table>
  </div>
</body>

</html>
