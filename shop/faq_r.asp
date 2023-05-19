
<%
'--------------------------------------------------------------
'      Coded By: Eric
'       Purpose: Display all category and search product form.
'   Used Tables: products
'  Invoked From: productsearch
'       Invokes: order.asp
'Included Files: header.asp, footer.asp, animate.js, StyleSheet.css, asp_lib.inc.asp.
'--------------------------------------------------------------
'Updated By   Eric    Date 01/04/2011   Comments
'Display products details
'--------------------------------------------------------------
%>
<!-- #include file="../include/asp_lib.inc.asp" -->
<!-- #include file="../include/sqlCheckInclude.asp" -->
<%

Dim strSQLCateCombo, cnn, strSQLCmd,strSQL

Dim rstCategory

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
   
  
 
  %>
 
                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
  </head>
  <body>
 
 
 
 
 
 <table class="fixedTable" >



<tr>
	<td class= "mainleft" >  </td>
	<td class = "maincenter" >   
	
<!--#Include file="headerRetail.asp"  -->
   
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
                        <img border="0" src="images/SALE.jpg" ><br />
					  
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
      <th class="thfeatured" >
          TERMS &amp; CONDITIONS</th>
     </tr>
  	<tr>
				<td >
				
			
					<b>1. Construction</b>
                    <br />
                    <br />

                    OMH models are all handcrafted by skillful and experienced craftsmen. They are 
                    built to scale according to the original plans or pictures of the actual ship 
                    using &quot;plank on frame&quot; or &quot;plank on bulk head&quot; construction method.

                    OMH models are not KIT; they are completely built from scratch. Each model comes with a 
                    nice wooden stand for display.

                    Please click <a href = "stepstobuild_r.asp"> here </a> to learn more.<br />
                    <br />
                                     <b>   2. Ship Measurement<br /></b>
                    <br />

All ships in the USA are measured in maximum display dimension (inches) and illustrated in the below picture.<br />
                    <br />
                    <img border="0" src="../images/measurementInches.jpg" width="488" height="338"><br />
                    <br />
                    <b>3. Material<br /></b>
                    <br />

                    At OMH, we put the beauty of each model first, so we use the highest quality of 
                    wood that are either locally available or from imported source such as rose wood, ebony, mahogany, teak, western red cedar… All 
                    timbers are air dry or kiln dry before production to ensure OMH products withstand 
                    different climate across the globe.<br />
                    <br />
             <b>4. How long it takes to build a model?<br /></b>
                    <br />

                    It takes about two hundred hours to complete a model such as 36&quot; HMS Victory. Custom built 
                    models require longer time. Our certificate of authenticity will state the 
                    number of hours required to finish each stage in contruction the model.<br />
                    <br />
                    <b>5. How to order?<br /></b>
                    <br />

                    Now you can order online, give us a call or email us if you have further questions. <br />
                    <br />

<b>6. Payment <br /></b>
                    <br />
                    You can pay online using any major credit cards. We also accept other payment methods such as checks, wire transfer, ACH...We are working to add other payment methods such as Apple Pay, Google Pay, Stripe for your convenient.<br />
                    <br />

<b>7. Delivery<br /></b>
                    <br />

                    Most items will be shipped within two business days via Fedex/UPS. Shipments 
                    take approximately 2-5 business days within the 48 states. A tracking number is 
                    provided upon shipment of each order. Oversized items may be shipped via common 
                    carriers. <br />
                    <br />

<b>8. Shipping & Freight<br /></b>
                    <br />

                    Shipping cost is Free within the 48 continental states for retail orders. For Express, Hawaii, Alaska or international orders, please contact 
                    us for exact shipping quote.<br />
                    <br />

<b>9. Minimum order<br /></b>
                    <br />

                    There is no minimum order for consumers.
                    <br />
                    <br />
                 <b>   10. Packaging<br /></b>
                    <br />

Each ship model is individually packed in a carton box. On many models, a 
                    
                               wooden frame
						is added for maximum protection. Most models are packed fully assembled.

For simple sail boats, the masts are folded down to save shipping space. Setting up the masts is 
                    simple with included instruction.<br />
                    <br />
                    <b>   11. How can I order from overseas?</b>

                    <br />
                    <br />
                    Please contact us to get a shipping quote before placing an order.
                    
                    <br />
					
					<b>
					
					12. Natural Materials </b>
					</br></br>
					
					
					Wood is a natural material that has different colors, grain patterns and overall texture and appearance. All wood will change color under the effect of UV light and oxidation. 
					Our products are made of natural wood, thus the final products' color may be different from the pictures on the web site or catalog or any pictures
					in email communication.
					</br></br>
					
					
                                    <b>    13. Returns<br />
                    </b>
                    <br />
                    Returns are expected to be made within 7 days of shipment receipt. In the event 
                    of processing a return, please forward the following information to
                    <a href="mailto:service@omhusa.com">service@omhusa.com</a>:<br />
                    <p>1. Invoice number<br />
                    2. Recipient Name (if drop shipped)<br />
                    3. Reasons for return<br />
                    4. Photos or damages/defects (If applicable)<br />
                    5. Description of damages/defects (If applicable)</p>
                    <p>Returns are subject to a 15% restocking fee if items have no damages or defects. 
                        Please do not deduct from your payment until you have received our credit memo.</p>
                    <b>     14. Damages or Defective Good</b><br />
                                        <br />
                    Any goods received damaged or defective can be returned following the 
                    instructions in our return policy statement. Once we assess the damages, two 
                    options are open:<br />
                    <br />
                    1. When deemed fixable, OMH will send parts for customer to repair the model<br />
                    2. When damage is unable to be fixed, OMH will ship out a new replacement at no 
                    cost. If replacement is not available in a reasonable time, OMH will issue 100% 
                    refund for the cost of the item.<br />
                    <br />
                    Any goods damaged in shipment need to be reported to us within 7 days so we can 
                    file a claim. Shipping cost is non refundable in any case.<br />
                    <br />
                    <b>    15. Backorders<br />
                    <br />
                    </b>
                    
                    Our
                    standard policy is to ship all back orders. Lead time of order may vary, please 
                    contact OMH for more information.<br />
                    <br />
                    <b>    16. Price<br />
                    </b>
                    <br />
                    Retail prices include ground shipping cost to 48 states. All prices are subject to change without 
                    notice. <br />
                    <br />
                    <b>    17. Catalog request</b>

                    <br />
                    <br />
                    Please click 
                    
                    <a href="catalog_r.asp">here</a> to see an online catalog and request one mailed to you.
                    <br />
                    <br />
                    <b>    18. Other Conditions</b><br />
                    <br />
                    Typographical and stenographic errors subject to correction. Purchaser terms 
                    inconsistent with those stated herein which may appear on purchaser&#39;s formal 
                    order will not be binding to OMH.
                    <br />
                    <br />
                    For further information, please feel free to contact us. We will answer your question 
                    in a timely manner. 	
</td>
</tr>
                    
			
  	<tr>
				<td >
				
			
					&nbsp;</td>
</tr>
                    
			
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
  
  
  
  
  
  
  
  
  
  
  
  
  
  
