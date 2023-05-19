
<%@ Language=VBScript %>
<%option explicit%>
<%
response.Expires=0
response.CacheControl= "no-cache"
response.AddHeader "Pragma", "no-cache"


'--------------------------------------------------------------
'      Coded By: Eric
'       Purpose: Display all category and search product form.

'   Used Tables: products
'  Invoked From: productsearch
'       Invokes: order.asp
'Included Files: header.htm, footerRetail.asp, animate.js, StyleSheet.css, asp_lib.inc.asp.
'--------------------------------------------------------------
'Updated By   Eric    Date 01/04/2011   Comments
'updated 2/22/2023 promo code is added, promocode can be enable in server setting, shipping cost discount were revised
'Display products details
'Updated by Eric 03/31/2020 changed to dropshiptemplate2013 instead of dstemplate table
' updated checkoutprocess.asp also 'Updated by Eric 03/31/2020 changed to dropshiptemplate2013 instead of dstemplate table
'--------------------------------------------------------------
%>
<!-- #include file="../include/asp_lib.inc.asp" -->
<!-- #include file="../include/sqlCheckInclude.asp" -->

<%

Dim strSQLCateCombo, cnn1, strSQLCmd1

Dim rstCategory
dim volumeDiscount1000

dim volumeDiscount2000
dim volumeDiscount5000

volumeDiscount1000=0.05
volumeDiscount2000=0.075
volumeDiscount5000=0.1

dim enablePromo

enablePromo=0
dim mapMarkup
if len(session("mapMarkup")) > 0 then
	mapMarkup=Cdbl (session("mapMarkup"))
else
	mapMarkup=1
end if








'********************************************************************************************************************************************************************************************************
'need these ASP section for category menu
'SQL statement for creating combo box. If name has more than 13 char then add ... as a tail.
strSQLCateCombo = "select Left(Category_Name, 23)+Left('...', Len(Category_Name) - Len(Left(Category_Name, 23))), Category_ID from Category where status <>'inactive' order by Category_Name asc "

'Create connection and query category data.
strSQLCmd1 = "select Category_ID, Category_Name from Category where status <>'inactive' order by upper(Category_Name) asc"
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.ConnectionString = Application.Contents("dbConnStr")
cnn1.Open
Set rstCategory = Server.CreateObject("ADODB.Recordset")
rstCategory.Open strSQLCmd1, cnn1, 3
'end category menu ASP
'********************************************************************************************************************************************************************************************************
   
   
   
'************************************************************************************************************************   
'start cart ASP   




%>
<%

Dim cnn, rst, rstCount, objMail, strSQLCmd, strMailContent
Dim strSuccess
dim straddAction, strUpdateAction
dim strLogin
dim intAddProduct_id, price, total, carttotal
dim redirectUrl
dim intquantity
dim strsqlcmdupdate
dim removeitemcode
dim consumerSaleEnable
consumerSaleEnable=1
dim consumerUser
dim logged_in

Dim UserIPAddress
UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If UserIPAddress = "" Then
  UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
End If


consumerUser=1
enablePromo=1


removeitemcode=request.querystring("removeitemcode")
strLogin = session("login")
If len(strLogin)=0 then
	strlogin ="temp" & userIPaddress & "_"&Session.SessionID
	session("templogin")=strlogin
	logged_in=0
else 
	logged_in=1

end if






straddAction = request.form("additemtocart")
strUpdateAction=request.form("pUpdate")
dim dsproId
dsproID=trim(request.querystring("dsProid"))

if len(request.querystring("addproduct_id")) > 0 then
	intAddProduct_id=cint(request.querystring("addproduct_id"))
	straddAction = "add"
else
	intAddProduct_id= Cint(request.form("addproduct_id"))
	
end if

dim strPromoCode
strPromoCode=request.form("promoCode")


dim strSQLgetItemID, rstTemp, rstPromo
strSQLgetItemID="select itemId from dropshiptemplate2013 inner join products on products.item=dropshiptemplate2013.product_code where product_id="&intAddproduct_id


 'Create connection to execute insert command
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  set rst = Server.CreateObject("adodb.RecordSet")
  set rstTemp=Server.CreateObject("adodb.RecordSet")
  set rstPromo=Server.CreateObject("adodb.RecordSet")
  
  
  set rstCount = Server.CreateObject("adodb.RecordSet")
  cnn.Open
  

'if add to cart from drop ship
if len(dsproId)>0 then
	intAddProduct_id=Cint(dsproId)
	straddAction="add"
	session("addDropship")=1
	
	

' add product from form
' Need to convert product_id to itemId in dropshiptemplate2013
elseif (intAddProduct_id>0) then
	session("addDropship")=0
   rstTemp.open strSQLgetItemID, cnn, 3
   if (rstTemp.eof) then
 	 	intAddProduct_id=-1
 	
 	 else	
   			intAddProduct_id=rstTemp("itemID")
   end if
   
 end if
 
  

%>
<%
'update shopping cart


if len(strUpdateAction) > 0 then
	dim i, cartSize
 	i=1
 	cartSize=Cint(request.form("pCartSize"))
 	while i<cartSize+1
 		intQuantity =Cint(Request.Form("quantity" & i))
 	'	response.write("quantity " & i& ":" & intquantity& "<br>")

 	dim cart_product_id
 	cart_product_id=request.form("product" & i)
 	if intQuantity >0 then
		 strSQLCmdupdate = "update shoppingcart set quantity = " & intquantity& " where product_id= " & cart_product_id
	else
		strSQLcmdupdate= "delete from shoppingcart where product_id= " & cart_product_id & "and login = '"&strlogin&"'"
	end if
	   rst.open strSqlcmdUpdate, cnn,3
		i=i+1
 	Wend    
end if

'add to cart


if len(straddaction)>0 and intAddProduct_id>0 Then	
 strSqlcmd="select count(product_id) as countproduct from shoppingcart where product_id= '" & intAddproduct_id & "' and login = '"&strlogin &"'"
 	 rstCount.open strSqlcmd, cnn,3
 	 'if same product is already in the cart, just add 1 to the quantity
 	 if Cint(rstCount("countproduct")) > 0 then
 	 	strSqlCmd = "update shoppingcart set quantity = quantity+1 where product_id= '" & intAddproduct_id&"'"
 	 else
		 strSQLcmd="insert into shoppingcart (login, product_id, quantity, shopdate, ip_address, logged_in) VALUES ('"& strlogin & "','" & intAddproduct_id & "', '1',getdate(),'" & userIPaddress & "' , '" & logged_in & "')"
		 
	end if
	 rst.open strSqlcmd, cnn,3
	 'response.write(strSQLcmd)
 	
end if
'remove item in cart
if len(removeitemcode)>0 then
	  strSQLCmd= "delete from shoppingcart where product_id='" & removeitemcode & "' and login = '"&strlogin&"'"
	    rst.open strSqlcmd, cnn,3
end if
	    


'VIEW CART	
dim strSQLcmd2
if len(straddaction)>-1 Then		


	  strSQLCmd= "SELECT distinct (item), shoppingcart.*, dropshiptemplate2013.shipping_cost,dropshiptemplate2013.available as available, dropshiptemplate2013.quantityonhand as quantityonhand, dropshiptemplate2013.Product_Name AS product_name, dropshiptemplate2013.sale_map as sale, dropshiptemplate2013.map as price, dropshiptemplate2013.product_code as itemcode, products.thumb_img as thumb_image, products.product_id as mainProduct_id FROM shoppingcart INNER JOIN " &_ 
                      " dropshiptemplate2013 ON shoppingcart.product_id = dropshiptemplate2013.itemID inner join products on item=product_code where login = '"&strlogin&"'"
  
	  'strSQLCmd2= "SELECT shoppingcart.*, products.Product_Name AS product_name, products.sale as sale, products.price as price, products.item as itemcode, products.thumb_img as thumb_image FROM shoppingcart INNER JOIN " &_ 
       '               " products ON shoppingcart.product_id = products.product_id where login = '"&strlogin&"'"

  rst.open strSqlcmd, cnn,3

  
  
  '-------------------------------------------------------------
  'Check promo code
  '---------------------------------------------------------------
  
  if len(strPromoCode)>0 and enablePromo=1 then
  dim discount, maxDiscount
  discount=0
  maxDiscount=0
  dim promoDesc
  
    
	  
  
	  strSQLCmd= "select *, isnull(max_amount, 0) as maxDiscount from promo where promo_code like '%" & strPromoCode &"'"
	  
	    rstPromo.open strSqlcmd, cnn,3
		if not rstPromo.eof then
				
					
			If dateDiff("d", now(), rstpromo("end_date"))>=0 and DateDiff("d", rstPromo("start_date"), now()) >=0 then
			
		' date is valid
				discount=rstpromo("discount_percent")
				
				
				maxDiscount= rstpromo("maxDiscount")
				if maxDiscount  > 0 then
					
					promoDesc =rstpromo("promo_description") & " discount (Max " & maxDiscount & "$ )"
				else
					promoDesc =rstpromo("promo_description") & " discount " 
				end if
				
				
			
			else
		'date is not valid
				promoDesc = "Promo Code Date is not valid"
				'response.write(promoDesc & "<br>")
			end if
			
		
		else
				promoDesc ="Promo Code is not valid"
	
		end if
	end if
  '-------------------------------------------------------------
  'end Check promo code
  '---------------------------------------------------------------


 	
%>



  

 
                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
 
 <script language="JavaScript1.2" src="../include/javascript.js"></script>
 
 <script>
   function isInt(value) {
   return !isNaN(value) && parseInt(Number(value)) == value;
}
     function validateTextBoxes() {
         
             var elLength = document.cart.elements.length;

             for (i = 0; i < elLength; i++) {
                 var type = cart.elements[i].type;
                
                 


                 if (type == "text") {


                        if (cart.elements[i].value.trim().length == 0) {
                            // alert("quantity can't be blank");
                             cart.elements[i].value=0;
                             continue;
                        }

                     
                     
                     if (isNaN(cart.elements[i].value.trim() ) and cart.elements[i].name<>"promoCode") {
               
                        alert("Quantity  is not valid ");
                        cart.elements[i].focus();
                        return false;}
            
                     
                 }

             }
             return true;
         }

         
     
</script>

<meta name="viewport" content="width=device-width, initial-scale=0.75">
  </head>
  <body>
 
 
 
 
 
 
 
 <table class="fixedTable" >



<tr>
	<td class= "mainleft" >  </td>
	<td class = "maincenter" >   
	
<!--#Include file="headerRetail.asp"  -->
    <table class="searchTable">
        <tr>
                          
                        <td class="cssTextCENTER" height="28" width="100%">
                            <form action="productsearchRetail.asp" method="POST" name="SearchForm">
                         
                               
                          <%Call SQLCombo("pCategoryID", "1", "", strSQLCateCombo, "All categories", "- - - - - - - - -", "0", "0")%>

                            Name / SKU
                                <input name="formSearch" type="hidden" value="yes" />
                                <input name="pProductName" size="15" type="text" />
                                <input name="pAction" type="submit" value="Search" />
                            </form>
                        </td>
          
        </tr>
</table>
<table class="mainTable">
  <tr>
    
         <%
	
	if not isnull(Request.Cookies("screenSize")) and len(trim(Request.Cookies("screenSize")))>0 then
	
	if (cint((Request.Cookies("screenSize"))) <600) then
			
	%>
	
    <td class="category" hidden="true" >
    <%
	else
	%>
	<td class="category"  >
	<% end if
	else
	%>
	<td class="category"  >
	<%
	end if
	%>
	
    
    
    
   
                 
                  <%
                  If rstCategory.RecordCount > 0 Then
                  %>
                  <table  class="table_outer_border" >                  
                  <tr >
                   <th  class ="thcategoryBGcolor"  >
                   CATEGORIES</th>
                   </tr>
                   
				   
				    <tr>
                      <td width="100%" align="left">&nbsp;</td>
                    </tr>
				   <tr>
                      <td align="left" class = "tdmargin10">
					 
					 
                      <span class="cssLink"><a href="productsearchRetail.asp?pCategoryID=-1" title="Ship Model - New Products "> <strong>New Products!!!</strong>  </a></span></td>
                    </tr>
					
					
                    <%
                    While Not rstCategory.EOF
                    %>
                   
                    
                   
                    <tr>
                      <td width="100%" align="left">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="left" class = "tdmargin10">
					 
					 
                      <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=<%=rstCategory("Category_ID")%>" title="Ship Model - <%=rstCategory("Category_Name")%>"><%=rstCategory("Category_Name")%> </a>
                          </span></td>
                    </tr>
                    <%
                      rstCategory.MoveNext
                    Wend
                    rstCategory.Close
                    cnn1.Close
                    Set rstCategory = Nothing
                    Set cnn1 = Nothing
                    %>
                   
            
                
                 
                    </table>
                  <%
                  End If
                  %>
                  
				  
				     <br />
				  <table  class="table_outer_border" > 
                  <tr >
                   <th  class ="thcategoryBGcolor"  >
                   LINKS</th>
                   </tr>
                   
                    <tr>
                      <td width="100%" align="left">&nbsp;</td>
                    </tr>
                    
                      <tr  >
      
                      <td  class="tdmargin10">
                      
                     
					   <p align="center">                   
                        <a href="productsearchRetail.asp?pCategoryID=-3" title="Items on sale"> 
                        <img border="0" src="../images/SALE.jpg" ><br />
					  
					    </a>
					  </p>
					  
                       <p align="center">                   
                        <a href="catalog_r.asp" title="catalog"> 
                        <img border="0" src="../images/catalog.JPG" ><br />
					   	</a>
					   </p>
					  
					  </td>
                    </tr>
               
                   <tr>
                      <td width="100%" align="left">&nbsp;</td>
                    </tr>
                  
                  </table>
				  
				  
				  
                </td>
                
             <!--end   <td class="category"> -->
    
    
  
    
    
    <td class="pageContent">
    
    
    
  
      
      
               <!--start content about us -->
      <table class="table98border_aboutus">
	<tr>
      <th class="thfeatured" colspan = "3" >
          SHOPPING CART
     </th>
     </tr>
     
     
     
     <%'***************************************************************************************************%>
     <%'----------------cart content start--------------%>
  	
  	
  	
  	
  	
  	<tr>
		<td align="left" valign="top" >&nbsp;
	    	</td>
   
										<td >
                                            <div align="center">
                                              <center>
                                             
										


									<p align="center">&nbsp; 
                                            <%if intAddproduct_id=-1 then%>
                                           Requested item not added, it is not currently available
								        	<%end if%>
								        	<%
								        	dim checkoutlink
								        	
								        	checkoutlink=1
								        	if rst.eof then%>
                                           Please add items to your cart
                                            	
								        	<%
								        	'set checkoutLink inactive
								        		checkoutlink=0
								        	end if%> 
								        	<br /><br />
								        	
								        	
								        	
<form method="POST" name="cart" action="cartRetail.asp" id="theForm" >
<table class="table98border_aboutus">

 
 



    <th class="thShoppingCart">&nbsp;</th>
  

  <th class="thShoppingCart">Img</th>

    <th class="thShoppingCart">SKU</th>

	 <th class="thShoppingCart">Name</th>
     <th class="thShoppingCart">Avail</th>
    <th class="thShoppingCart">Quan</th>
     <th class="thShoppingCart">Price</th>
     <th class="thShoppingCart">Total</th>
                                              
												<%
												dim count, quantityCount
												quantityCount=0
												count=1
												cartTotal=0
												dim shippingCost
												shippingCost=0
												dim shippingmethod
												'default shipping method: ground
												shippingmethod=0
												
												WHILE  (not rst.eof)  %>
												
												<tr>
                                                  <td width="10%">
                                                

														   <p align="center">
                                                

														   <a href="cartRetail.asp?removeitemcode=<%=rst("product_id")%>" onClick="">   <img border="0" src="../images/delete.gif" alt="Remove this item"></a>
	
													</td>
                                                 <td width="20%">
                                                 <p align="center">
                                                 <a href="productsdetailsRetail.asp?ProductID=<%=rst("mainProduct_id")%>">
                                                 <img src="../thumbimages/<%=rst("thumb_image")%>" width="50" border="0" ></a></td>

                                                  <td width="15%"><%=rst("itemcode")%>&nbsp;</td>
                                                  <td width="35%"><%=rst("product_name")%>&nbsp;</td>
                                                   <td width="10%"><%=left(rst("available"), 5)%>&nbsp;</td>
                                                  <td width="10%">
                                                                                               
												 <input type="text" name="quantity<%=count%>" value="<%=rst("quantity")%>" size="4">
												 <input type="hidden" name="product<%=count%>" value="<%=rst("product_id")%>">
                                                  </td>
                                                  <%
                                                 
                                                  	if not isnull(rst("sale")) then
                                                  		price=cdbl(rst("sale"))* mapMarkup
                                                  	elseif not isnull(rst("price")) then
                                                  		price=Cdbl(rst("price"))*mapMarkup
                                                  	else 
                                                  		price=0
                                                  	end if
                                                  	price=round (price, 2)
                                                  	
                                              
                                                  	

                                                  	
                                                 %> 	                                                  
                                                  <td width="10%"><%=price%>&nbsp;</td>
                                                  <%
                                                  if not isNull(rst("quantity")) then
                                                  	total=Cdbl(rst("quantity"))*price
                                                  else
                                                  	total=0
                                                  end if
                                                  total=round (total, 2)
                                                  
                                                  %>
       		        							  <td width="20%"><%=total%>&nbsp;</td>
                                                 <%
                                                 cartTotal=cartTotal+total
                                                 
                                                 
                                                 if rst("shipping_cost")>=216 then
                                                 	'via truck
                                                 	shippingmethod=1
                                                 	
                                                 elseif rst("shipping_cost")>=215 then
                                                 	'via fedex ground oversize 3
	                                               	shippingmethod=2
                                                 end if
                                                 
                                                 quantityCount=quantityCount+rst("quantity")
                                                 if (not isNull(rst("shipping_cost"))) and rst("shipping_cost")>0 then
                                                 	shippingCost=shippingCost+Cdbl(rst("shipping_cost")*rst("quantity"))
                                                 else
                                                 	'response.write("Shipping cost is not available")
                                                 	shippingCost=0
                                                 		
                                                 end if
												 ' free shipping to consumer
												 ' 
												 shippingCost=0 
                                               
                                                 
                                                 rst.movenext()
                                                 
                                                 count=count+1
                                                 
                                             
                                                %> 
                                                </tr>
                                                <% WEND
                                                'shippingCost combination discount
                                                dim note
                                                note=""
                                                ' multiple items shipping discount 10% off shipping Co
                                                if quantityCount>1 and quantityCount<5 then
                                                			shippingCost=ShippingCost*.9                                                			                                     
                                                elseif quantityCount>=5 and quantityCount<10 then
                                                			shippingCost=ShippingCost*.85   
												elseif quantityCount>=10 and quantityCount<20 then
															shippingCost=shippingCost*.8
												elseif quantityCount>=20 and quantityCount<50 then
															shippingCost=shippingCost*.75
															
												elseif quantityCount>=50 then
															shippingCost=shippingCost*.7															
                                                end if
                                                
                                                ' high volume shipping Discount for multiple items
                                                
												if quantityCount>1 then
												
													if cartTotal>=1000 and cartTotal<2000 then
														shippingCost=shippingCost*0.9
													 elseif cartTotal>=2000 and cartTotal<5000  then
														shippingCost=shippingCost*.85
													 elseif cartTotal>=5000 then
														 shippingCost=shippingCost*.8
													 end if                                                
																									 
													if shippingCost>500 then
														'shippingCost=0
													end if
                                                end if
												
                                                
                                                
                                                                                                
                                                dim volumeDiscount
                                                volumeDiscount=0
                                                if carttotal>=1000 and cartTotal<2000 and QuantityCount>=3 then
                                                	volumediscount=CartTotal*volumeDiscount1000
                                                elseif carttotal>=2000 and cartTotal<5000 and QuantityCount>=3 then
                                                	volumediscount=CartTotal*volumeDiscount2000
                                                elseif carttotal>=5000 and QuantityCount>=3 then
                                                	volumediscount=CartTotal*volumeDiscount5000
                                                end if
                                                
                                                
                                                if shippingCost=0 then
                                                	Note="* Shipping cost will be provided upon order processing. <br> "
                                                end if
                                                If cartTotal>=500 and volumediscount=0 then
                                                	Note=Note & "Discount is available when you order 1000$+ with a minimum of 3 in cart total quantity<br>"
                                                end if
                                                shippingCost=round(shippingCost,2)
                                                cartTotal=round(cartTotal,2)
                                                volumeDiscount=round(volumeDiscount, 2)
                                                
                                                session("shippingCost")=shippingCost
                                                session("cartTotal")=cartTotal
                                                session("volumeDiscount")=volumeDiscount
                                                	
                                                %>  
                                              </table>
                                              </center>
                                            </div>
                                            
                                      
                                               <div align="center">
                                                 <center>
<table class="table98border_aboutus">
                                                 <tr>
												<td>
												</td>
												 
                                                  
                                                   <td align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                                                   Total&nbsp; &nbsp;
                                                   </td>
                                                   <td align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                                                   $ <%=cartTotal%>&nbsp;</td>
                                                 </tr>
                                                 <% if volumeDiscount>0 then%>
                                                 
                                                 <tr>
												  <td>
												 
												 
												 </td>
												 
                                                   <td  align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                                                  Volume Discount </td>
                                                   <td  align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                                                 
                                                   - $ <%=volumeDiscount%>&nbsp;</td>
                                                 </tr>
                                                 <%end if%>
												 
												 
												 												 
												<% 
												 
												 '--------------------------------------------------------------------------------------------
												 'if enable promo =1 

												dim promoDiscount
												promoDiscount=0
												if enablePromo=1 then%>
                                                 
                                                 <tr>
												
												 
												<td>
												 Promo Code <input type="text" name="promoCode" size="20"> (Click update to apply)
												 </td>
												 <td  align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
												<%if len(promoDesc) > 0 then
														response.write (promoDesc)
													end if
													
																										
												 %>										 
												 </td>
												<%
											
                                                if discount> 0 then
													promoDiscount = round( cartTotal * discount/100, 2)
													'only allow promo discount at Max discount
													
													if promoDiscount> maxDiscount then
														promodiscount=MaxDiscount
													end if
														
													session("promoDiscount")=promoDiscount
													session("promoName")=promoDesc
													session("promoCode")=strpromoCode
													
												else
													promodiscount=0
													session("promoDiscount")=""
													session("promoName")=""
													
												end if

												session("promoDiscount")=promoDiscount
												
												%>
												
												 <td  align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
													<%
													if discount> 0 then
														response.write ("- $ " & round(promoDiscount,2))

													end if %>
												</td>
                                               </tr>
											 
												 
											   <%
											     '--------------------------------------------------------------------------------------------
												 
												 else
												 'if enable promo =0 
												 
												 end if
												 if consumerSaleEnable=1 and consumerUser=1 then
													shippingCost=0
													Note = "* Free domestic ground shipping except orders to Hawaii and Alaska<br>"
												 end if
												 
												 %>
												     
											   
											   
											   
												
												<tr>
												 <td>
												 </td>
												 

												 <td  align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                                                
                                                   Domestic Ground Shipping 
                                                   Cost&nbsp;&nbsp;&nbsp; </td>
                                                   <td  align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                                                  &nbsp;<%=ShippingCost%>&nbsp; 
                                                   <font color="#FF0000">*</font></td>
                                                 </tr>
                                                 <tr>
												 <td></td>
                                                   <td  align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                                                  Grand Total&nbsp;&nbsp;&nbsp; </td>
                                                   <td  align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                                                   $ <%=Carttotal+ShippingCost-volumeDiscount - promoDiscount%>&nbsp;</td>
                                                 </tr>
                                                 



 </table>
                                            
                                      
                                      
                                         
											     </center>
                                            </div>
                                      
                                         
											 <p align="center">
                                      
                                      
                                             <input type="hidden" value="<%=count-1%>" name="pCartSize">
                                             <input type="hidden" name="pUpdate" value="update" >
                                             <input type= "submit" name="update" border="0" src="../images/update.gif"  value=" Update " onClick="return validateTextBoxes();">
                                            
                                            
                                              
                                              <%if session("addDropship")=1 then %>
                                            
                                              <input type = "button" value=" Continue Shopping " onclick = "location.href='dropshiplist.asp'" />
                                              
                                              <%else%>
                                               <input type = "button" value=" Continue Shopping " onclick = "location.href='productsretail.asp'" />
                                              <%end if%>
                                              
                                              

											 <!--<input type="submit" value="Update" name="pUpdate">-->
											 
                                      
                                      
                                      
                                         
											
                                            
                                      
                                      
                                         
											 <p align="center">
                                      
                                             
                                     
											 
                                            </p>
                                      
                                      
                                         
											 <p align="center">
                                      
                                      
                                      
										<% if checkoutlink<>0 then %>
                                                <input type = "button" value=" Checkout " onclick = "location.href='checkoutRetail.asp'" />
                                          </p>
                                      <%end if%>
                                      
                                         
											 <p align ="left">
                                      
                                             <font class="attentionText">
                                             
                                             Note: <br /></font>
                                             <font>
                                             
                                       
                                             <%response.write(note)%>
                                             * For items with limited quantity, please contact us to verify availability <br>
                                             * Soldout items will be automatically added to back order 
                                             unless you notify us otherwise.
                                             </font>
                                             </p>
                                             
										
                                      
                                         
											 <p align="center">
                                       <a href="https://www.positivessl.com" target ="_blank"  style="font-family: arial; font-size: 10px; color: #212121; text-decoration: none;"><img src="https://www.positivessl.com/images-new/PositiveSSL_tl_white.png" alt="SSL Certificate" title="SSL Certificate" border="0" /></a>
                                        </p>
                                      
                                         
												 

                                                                               
                                            
                                            </form>
                                            
                                            
                                            
                                              </font>
                                              </font>
										</td>
										
									 <td align="left" valign="top" >&nbsp;</td>
   
									</tr>
									
								
									
				
								
  	
     
     <%  
    end if
    cnn.Close
    Set cnn = Nothing
    '----------------cart content end--------------
   
   %>
   
			
		</table><!--end table98 --> 
    
       </td>
       </tr>
       </table>
        <!--end mainTable-->
        
        <!--#Include file="footerRetail.asp"  --> 
        
        </td>
        <!--end mainCenter-->
        



<td class = "mainright" >    </td>
</tr>

</table>
 
 
 

  
  
  
  
  
  
  </body>
  
  
  
  
  

  
  
  
  
  
  
  
