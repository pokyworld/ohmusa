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

    Dim strSQLCateCombo, cnn1, strSQLCmd1
    Dim rstCategory

  '********************************************************************************************************************************************************************************************************

  ' need these ASP section for category menu

  ' SQL statement for creating combo box. If name has more than 13 char then add ... as a tail.

  sql = "select Left(Category_Name, 23)+Left('...', Len(Category_Name) - Len(Left(Category_Name, 23))), "
  sql = sql & "Category_ID from Category where status <>'inactive' "
  sql = sql & "order by Category_Name asc "
  strSQLCateCombo = sql

  ' Create connection and query category data.

  sql = "select Category_ID, Category_Name from Category where status <>'inactive' "
  sql = sql & "order by upper(Category_Name) asc "
  strSQLCmd1 = sql

  Set cnn1 = Server.CreateObject("ADODB.Connection")
  cnn1.ConnectionString = Application.Contents("dbConnStr")
  cnn1.Open

  Set rstCategory = Server.CreateObject("ADODB.Recordset")
  rstCategory.Open strSQLCmd1, cnn1, 3

  ' end category menu ASP

  '********************************************************************************************************************************************************************************************************
  '************************************************************************************************************************
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
    <link rel="stylesheet" type="text/css" href="include/product_stylesheet.css">

    <script language="JavaScript1.2" src="include/javascript.js"></script>
  </head>

  <body>
    <table class="fixedTable">
      <tr>
        <td class="mainleft">&nbsp;</td>
        <td class="maincenter">
          <!--#Include virtual="/shop/payment/include/headerRetail.asp"  -->
          <table class="mainTable">
            <tr>
<% 
          if not isnull(Request.Cookies("screenSize")) and len(trim(Request.Cookies("screenSize")))>0 then
            if (cint((Request.Cookies("screenSize"))) <600) then 
%>
              <td class="category" hidden="true">
<% 
            else 
%>
              <td class="category">
<% 
            end if '// (cint((Request.Cookies("screenSize"))) <600)
          else 
%>
              <td class="category">
<% 
          end if  '// not isnull(Request.Cookies("screenSize")) and len(trim(Request.Cookies
%>


<% 
              If rstCategory.RecordCount > 0 Then
%>
                <table class="table_outer_border">
                  <tr><th class="thcategoryBGcolor">CATEGORIES</th></tr>
                  <tr><td width="100%" align="left">&nbsp;</td></tr>
                  <tr>
                    <td align="left" class="tdmargin10">
                      <span class="cssLink"><a href="../productsearchRetail.asp?pCategoryID=-1" title="Ship Model - New Products "> <strong>New Products!!!</strong> </a></span>
                    </td>
                  </tr>
<% 
                While Not rstCategory.EOF
                  CategoryID = rstCategory("Category_ID")
                  CategoryName = rstCategory("Category_Name")
%>
                  <tr><td width="100%" align="left">&nbsp;</td></tr>
                  <tr>
                    <td align="left" class="tdmargin10">
                      <span class="cssLink">
                        <a href="../productsearchRetail.asp?pCategoryID=<%=CategoryID%>" title="Ship Model - <%=CategoryName%>"> <%=CategoryName%></a>
                      </span>
                    </td>
                  </tr>
<% 
                  rstCategory.MoveNext 
                Wend 
                rstCategory.Close 
                cnn1.Close 
                Set rstCategory=Nothing 
                Set cnn1=Nothing
%>
                </table>
<%
              End If  '// rstCategory.RecordCount> 0

%>
                <br />
                <table class="table_outer_border">
                  <tr>
                    <th class="thcategoryBGcolor">
                      LINKS</th>
                  </tr>

                  <tr>
                    <td width="100%" align="left">&nbsp;</td>
                  </tr>

                  <tr>

                    <td class="tdmargin10">


                      <p align="center">
                        <a href="productsearchRetail.asp?pCategoryID=-3" title="Items on sale">
                          <img border="0" src="../../images/SALE.jpg"><br />

                        </a>
                      </p>

                      <p align="center">
                        <a href="catalog_r.asp" title="catalog">
                          <img border="0" src="../../images/catalog.JPG"><br />
                        </a>
                      </p>

                    </td>
                  </tr>

                  <tr>
                    <td width="100%" align="left">&nbsp;</td>
                  </tr>

                </table>
              </td>

              <!--end  <td class="category"> -->

              <td class="pageContent">
                <!--start content about us -->
                
<%
  ' ********************************************************************************************
%>
<%
  '----------------Middle content start--------------

  '// Get Session data from stripe to validate cancelled
    StripeSessionResult = GetStripeSession(secret_key, StripeSessionId)
    If Len(StripeSessionResult) >= 1 Then
      Set ojSession = new aspJSON
      ojSession.loadJSON(StripeSessionResult)
      redirUrl = ojSession.data("url")
      SessionPaymentStatus = ojSession.data("payment_status")
      StripeCustomerId = ojSession.data("customer")
      ' StripeInvoiceId = ojSession.data("invoice")
      ' StripePaymentId = ojSession.data("payment_intent")
      status = ojSession.data("status")
      Set ojSession = Nothing

    End If

  '// Print out Customer/Order details from newOrder object  
    ' PrintCustomer(newOrder.Customer) 
    ' PrintOrder(newOrder)
%>
<link rel="stylesheet" href="template.css" />
  <div class="stripe-result">
    <table class="table98border_aboutus">
      <tr>
        <th class="thfeatured">
          Payment ABORTED: (Order <%=newOrder.OrderId%>)
        </th>
      </tr>
      <tr>
        <td style="padding: 0;">
          <div class="table-wrapper">
            <table class="tbl-contact">
              <tr>
                <th>Contact Details</th>
              </tr>
              <tr>
                <td><span>Contact Name:</span><%=newOrder.Customer.FullName %></td>
              </tr>
              <tr>
                <td><span>Email:</span><%=newOrder.Customer.Email%></td>
              </tr>
              <tr>
                <td><span>Phone:</span><%=newOrder.Customer.Phone%></td>
              </tr>
            </table>
            <table class="tbl-address">
              <tr>
                <th>Address</th>
              </tr>
              <tr>
                <td nowrap><%=newOrder.Customer.BillingAddress.Line1 %></td>
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
            </table>
          </div>
          <hr/>
          <div class="table-wrapper">
            <table class="order-lines">
              <tr>
                <th>&nbsp;</th>
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
                <td><img src="<%=newOrderItem.Price.Product.ImageUrl%>" alt=""></td>
                <td><%=newOrderItem.Price.Product.SKU%></td>
                <td><%=newOrderItem.Price.Product.Name%></td>
                <td class="center"><%=newOrderItem.Quantity%></td>
                <td class="right">&dollar;&nbsp;<%=FormatNumber(newOrderItem.SubTotal/100,2)%></td>
              </tr>
<%
  Next
%>              
            </table>
          </div>
          <hr />
          <div class="table-wrapper">
            <table class="tbl-totals">
              <tr>
                <td><span>Net SubTotal:</span></td>
                <td class="right">&dollar;&nbsp;<%=FormatNumber(newOrder.NetSubTotal/100,2)%></td>
              </tr>
<%
  If newOrder.PromoDiscount >= 1 Then
%>              
              <tr>
                <td><span><%=newOrder.PromoCode%>:</span></td>
                <td class="right">&dollar;&nbsp;(<%=FormatNumber(newOrder.PromoDiscount/100,2)%>)</td>
              </tr>
<%
  End If

  If newOrder.Discount >= 1 Then
%>              
              <tr>
                <td><span>Discount:</span></td>
                <td class="right">&dollar;&nbsp;(<%=FormatNumber(newOrder.Discount/100,2)%>)</td>
              </tr>
<%
  End If
%>
              <tr>
                <td><span>Shipping:</span></td>
                <td class="right">&dollar;&nbsp;<%=FormatNumber(newOrder.Shipping/100,2)%></td>
              </tr>
              <tr>
                <td><span>SubTotal:</span></td>
                <td class="right">&dollar;&nbsp;<%=FormatNumber(newOrder.SubTotal/100,2)%></td>
              </tr>
              <!--<tr>
                <td><span>Tax(es):</span></td>
                <td class="right">&dollar;&nbsp;<%=FormatNumber(newOrder.Tax/100,2)%></td>
              </tr>-->
              <tr class="border">
                <td><span>Total:</span></td>
                <td class="right bold">&dollar;&nbsp;<%=FormatNumber(newOrder.Total/100,2)%></td>
              </tr>
              <tr><td colspan="2">&nbsp;</td></tr>
              <!-- <tr>
                <td><span>Payment Ref:</span></td>
                <td class="right"><%=newOrder.Payment.Id%></td>
              </tr> -->
              <tr>
                <td><span>Payment Status:</span></td>
                <td class="right"><%=UCase(newOrder.Payment.Status)%></td>
              </tr>
              <!-- <tr>
                <td><span>Amount Paid:</span></td>
                <td class="right">&dollar;&nbsp;<%=FormatNumber(newOrder.Payment.Amount/100,2)%></td>
              </tr> -->
<%
  BalanceDue = Round((newOrder.Total / 100), 2) - Round((newOrder.Payment.Amount / 100), 2)
  If BalanceDue = 0 Then
    StripeSessionId = GetLastOrderPayment(orderid)
    params = "id=" & StripeSessionId & "&orderid=" & OrderId & "&userid=" & UserId
    Response.Redirect "/shop/payment/success.asp?" & params
  End If
%>  
              <tr class="border">
                <td><span>Balance Due:</span></td>
                <td class="right bold">&dollar;&nbsp;<%=FormatNumber(BalanceDue, 2)%></td>
              </tr>
            </table>
          </div>
          <div class="button-group">
            <button id="retry">Retry Payment</button>
          </div>

        </td>
      </tr>
    </table><!--end table98 -->
  </div>
  <div class="order-message">
    <h2>Thank you for your Order.</h2>
      <h4>Unfortunately the order has been placed on hold</h4>
      <h4>It remains UNPAID, Please <a href="<%=redirUrl%>">try again</a></h4>
  </div>
  <script>
    document.addEventListener("DOMContentLoaded", () =>{
      const redir = "<%=redirUrl%>";
      const payBtn = document.querySelector("#retry");
      const portrait = window.matchMedia("(orientation: portrait)");
      portrait.addEventListener("change", function(e) {
        if(e.matches) {
          location.reload();
        }
      });
      payBtn.addEventListener("click", (e) => {
        e.preventDefault();
        window.location = redir;
      });
    });
  </script>
<%
  '----------------Middle content end--------------- 
%>
<%
  ' ********************************************************************************************
%>
                </table><!--end table98 -->
              </td>
            </tr>
          </table><!--end mainTable-->
          <!--Include virtual="/shop/payment/include/FooterRetail.asp"  -->
        </td>
        <!--end mainCenter-->
        <td class="mainright"> </td>
      </tr>
    </table>
  </body>
</html>
<%
Else '// no order id
  If Len(OrderId) < 1 Then : RW("Missing OrderId") : Response.End
  If Len(StripeSessionId) < 1 Then : RW("Missing StripeSessionId") : Response.End
End If
%>
