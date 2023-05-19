<%@ Language=VBScript %>
<% Option Explicit %>
<%
'--------------------------------------------------------------
'      Coded By: Tan Pham on 01/01/2001.
'       Purpose: Display all category and search product form.
'   Used Tables: Category.
'  Invoked From: index.asp.
'       Invokes: search_product.asp
'Included Files: headerRetail.asp, footerRetail.asp, animate.js, StyleSheet.css, asp_lib.inc.asp.
'--------------------------------------------------------------
'Updated By   Trang Truong   Date 21/03/2001    Comments
'Display products for category
'--------------------------------------------------------------
%>
<!-- #include file = "../include/asp_lib.inc.asp" -->
<%
Session("SqlPage") = "productSearchRetail.asp"
%>
<!-- #include file="../include/sqlCheckInclude.asp" -->
<%
Dim intCategoryID, strProductName, intWhichPage, intRecNum, intMaxProduct, intNumOfPage, intProductFound
Dim strSQLCateCombo, cnn,cnn1, rstProduct, strSQLCmd,strSQL,strSQLCmd1
Dim strCategoryName
Dim ProductPrice
Dim ProductQuan
Dim intNewproduct
dim retailMarkup
retailMarkup=2.0
session("retailMarkup")=retailMarkup
dim mapMarkup
mapMarkup=1.15
session("mapMarkup")=mapMarkup


dim displaySalePrice_retail

displaySalePrice_retail = 0
 dim totalColumn
 totalColumn=4
			  
if not isnull(Request.Cookies("screenSize")) and len(trim(Request.Cookies("screenSize")))>0 then
	
	if (cint((Request.Cookies("screenSize"))) <600) then
			totalColumn=2
	end if
	
end if


Dim rstCategory

'SQL statement for creating combo box. If name has more than 13 char then add ... as a tail.
strSQLCateCombo = "select Left(Category_Name, 23)+Left('...', Len(Category_Name) - Len(Left(Category_Name, 23))), Category_ID from Category where status <>'inactive' order by Category_Name asc "

'Create connection and query category data.
strSQLCmd1 = "select Category_ID, Category_Name from Category where status <>'inactive' order by upper(Category_Name) asc"
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.ConnectionString = Application.Contents("dbConnStr")
cnn1.Open
Set rstCategory = Server.CreateObject("ADODB.Recordset")
rstCategory.Open strSQLCmd1, cnn1, 3





'Product number displayed on one page (4, 8, 12, 16 . . .)
intMaxProduct = 20

'Initial intRecNum
intRecNum = 1

'Search from form

'If Len(Request.Form("pAction")) > 0 Then
If Len(Request.Form("formSearch")) > 0 Then
'if Request.Form("hide")=1 then
  'Always display page 1
  intWhichPage = 1


  intCategoryID = Cint(Request.Form("pCategoryID"))
  strProductName = fixstring(Trim(Request.Form("pProductName")))

  'Check if search in all category
  If intCategoryID > 0 Then
    If Len(strProductName) > 0 Then
      'strSQLCmd = "select * from Products where active > 0 and Category_ID = " & intCategoryID &_
       ' " and Product_Name like '%" & strProductName & "%' order by item"
       strSQLCmd="select Product_Id, Product_Name, Specification, Long_Desc," &_
		         " Short_Desc, Thumb_Img, Large_Img, Price, item, sale, Quantity, IsNew," &_
		         " Category_Name , rating, total_review, map, sale_map " &_
		         " from Products_review a, Category b where inUSA=1 and price is not null and a.Category_Id = " & intCategoryId  &_
		         " and a.Category_Id = b.Category_Id " &_
		         " and (Product_Name like '%" & strProductName & "%' or Specification like '%" & strProductName & "%' or item  like '%" & strProductName & "%') and price is not null order by item"

    Else
      'strSQLCmd = "select * from Products_review where Category_ID = " & intCategoryID &_
      '  " order by Product_Name asc"
       strSQLCmd="select Product_Id, Product_Name, Specification, Long_Desc," &_
		         " Short_Desc, Thumb_Img, Large_Img, Price, item,  sale, Quantity, IsNew," &_
		         " Category_Name , rating, total_review, map, sale_map " &_
		         " from Products_review a, Category b where  inUSA=1 and price is not null and a.Category_Id = " & intCategoryId  &_
		         " and a.Category_Id = b.Category_Id " &_
		         " order by item"

    End If
  Else
    If Len(strProductName) > 0 Then
      strSQLCmd = "select * from Products_review where inUSA=1 and price is not null and Product_Name like '%" & strProductName & "%' or inusa=1 and item like '%" & strProductName & "%' or inusa=1 and price is not null and specification like '%" & strProductName & "%' order by item"

    Else
      strSQLCmd = "select * from Products_review where inUSA=1 and price is not null order by item"
    End If
  End If
Else
  intWhichPage = cint(Request.QueryString("pWhichPage"))
  intCategoryId= cint(Request.QueryString("pCategoryID"))
  'Ckeck if go to next page
  If intWhichPage > 0 Then
    if len(Session("strSQLCmd"))> 0 then
        strSQLCmd = Session("strSQLCmd")
     elseif intCategoryId >=0 then
     
      strSQLCmd="select Product_Id, Product_Name, Specification, Long_Desc," &_
		      " Short_Desc, Thumb_Img, Large_Img, Price, sale, item,  Quantity, IsNew," &_
		      " Category_Name , rating, total_review, map, sale_map " &_
		      " from Products_review a, Category b where inUSA=1 and price is not null and a.Category_Id = " & intCategoryId  &_
		      " and a.Category_Id = b.Category_Id order by item"
      else
             strSQLCmd="select Product_Id, Product_Name, Specification, Long_Desc," &_
		      " Short_Desc, Thumb_Img, Large_Img, Price, sale, item,  Quantity, IsNew," &_
		      " Category_Name, rating, total_review, map, sale_map " &_
		      " from Products_review a, Category b where inUSA=1 and price is not null "   &_
		      " and a.Category_Id = b.Category_Id order by item"
 
     end if
     
    
    
    
    intWhichPage = CInt(intWhichPage)
  Else
    intWhichPage = 1
    intCategoryID = Cint(Trim(Request.QueryString("pCategoryID")))
    'display only NEW item, categoryid=-1, when customer click on NEW products.
    If intCategoryID=-1 then
    	 strSQLCmd = "select * from Products_review where inUsa=1 and isNew=1 and price is not null order by product_Id desc "
    'if category=-2 then display all products that overstock from the factory: (quantity=5)
    elseif intCategoryID=-2 then
     	strSQLCmd = "select * from Products_review where inUSA=1 and price is not null and overstock=1 order by item "
    'display products on sale
	elseif intCategoryID=-3 then
     	strSQLCmd = "select * from Products_review where inUSA=1 and sale < price and price is not null order by item "
		
		
		' favorite items BEST SELLER
	elseif intCategoryID=-4 then
     	strSQLCmd = "select p.* from Products_review p join bestseller2022 s on p.item=s.item  order by product_name "
		' halloween
	elseif intCategoryID=-5 then
     	strSQLCmd = "select * from Products_review where inUSA=1 and halloween=1 and price is not null order by product_name "
	'XMAS 
elseif intCategoryID=-6 then
    	strSQLCmd = "select * from Products_review where inUSA=1 and xmas =1 and price is not null order by product_name "
		
		
	else
    ' display products when customer click on each category
    strSQLCmd="select Product_Id, Product_Name, Specification, Long_Desc," &_
		      " Short_Desc, Thumb_Img, Large_Img, Price, sale, item, Quantity, IsNew," &_
		      " Category_Name, rating, total_review, map, sale_map " &_
		      " from Products_review a, Category b where inUSA=1 and price is not null and a.Category_Id = " & intCategoryId  &_
		      " and a.Category_Id = b.Category_Id order by item"
		      
		      
	End if
  End If
 
End If


			  
	
'Create connection and query category data.
Set cnn = Server.CreateObject("ADODB.Connection")
cnn.ConnectionString = Application.Contents("dbConnStr")

cnn.Open
Set rstProduct = Server.CreateObject("ADODB.Recordset")
 

rstProduct.Open strSQLCmd, cnn, 3

'Write down session if there is more one page
If rstProduct.RecordCount > intMaxProduct Then
  Session("strSQLCmd") = strSQLCmd
  intNumOfPage = rstProduct.RecordCount\intMaxProduct
  If (rstProduct.RecordCount mod intMaxProduct) > 0 Then
    intNumOfPage = intNumOfPage + 1
  End If
Else
  intNumOfPage = 1
End If

intProductFound = rstProduct.RecordCount

'if form is used
'If Len(Request.Form("pAction")) > 0 Then
If Len(Request.Form("formSearch")) > 0 Then

    intCategoryID = Cint(Trim(Request.Form("pCategoryID")))
else
    intCategoryID = Cint(Trim(Request.QueryString("pCategoryID")))
end if

If intProductFound > 0 and intCategoryID >0  Then
   strCategoryName = rstProduct("Category_Name")
End If
%>
<html>
<head>

<title>Model ship - Battleship, Boats, Canoes, Cruiseship, Tallship, Yacht, Speed boat, Souvenir </title>

<meta name="keywords" content="model ships, ship model, sailing ship model, wooden boat, tallship, tallship model, historic ship, wooden ship model, Handmade ship model, speed boat, tall ship, modern yacht, old style yacht, historic ship model, museum ship model, Queen Mary, Riva, HMS Victory, Sovereign of the Seas, San Felipe, Esmeralda, USS Constituion, USS Constellation, Wasa, Vasa, Soleil Royal, Friesland, Titanic, Shrimp boat, Normandie, Queen Mary 2, Zeven Provincien, Lady Washington, Mikasa, Australia, Batavia, Shamrock, Endeavour, Bounty, Handcraft ship model, handicraft, handicrafts, manufacturer, exporter, ship leading builder">
<meta name="DESCRIPTION" content="Fine, handcrafted model ships and sail boat models are available from Handicraftscan.com.  See a large number of photos of our beautiful tall ships, speed boats, fishing boats, canoes, and yachts from ole to modern style online now, tallship, tallship model, historic ship, wooden ship model, Handmade ship model, speed boat, tall ship, modern yacht, old style yacht, historic ship model, museum ship model, Queen Mary, Riva, HMS Victory, Sovereign of the Seas, San Felipe, Esmeralda, USS Constituion, USS Constellation, Wasa, Vasa, Soleil Royal, Friesland, Titanic, Shrimp boat, Normandie, Queen Mary 2, Zeven Provincien, Lady Washington, Mikasa, Australia, Batavia, Shamrock, Endeavour, Bounty, Handcraft ship model, handicraft, handicrafts, manufacturer, exporter, ship leading builder">
<meta name="Destination" content="Model ship - Battleship, Boats, Canoes, Cruiseship, Tallship, Yacht, Speed boat, Souvenir">

<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico" />
<link rel="stylesheet" type="text/css" href="../product_stylesheet.css">

<script language="JavaScript1.2" src="kill-mouse.js" type="text/javascript"></script>
<script type='text/javascript' src='https://platform-api.sharethis.com/js/sharethis.js#property=6202fc68049246001a151155&product=inline-share-buttons' async='async'></script>
 
 

<script type="text/javascript">
	document.cookie = "screenSize=" + screen.width;
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
                            <form action="productSearchRetail.asp" method="POST" name="SearchForm">
                         
                               
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
    
    <td class="category">
    
    
   
                 
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
					 
					 
                      <span class="cssLink"><a href="productSearchRetail.asp?pCategoryID=-1" title="Ship Model - New Products "> <i class="AttentionText">New Products!!!</i>  </a></span></td>
                    </tr>
					
                    <%
                    While Not rstCategory.EOF
                    %>
                   
                    
                   
                    <tr>
                      <td width="100%" align="left">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="left" class = "tdmargin10">
					 
					 
                     
                      <a href="productSearchRetail.asp?pCategoryID=<%=rstCategory("Category_ID")%>" title="Ship Model - <%=rstCategory("Category_Name")%>"><%=rstCategory("Category_Name")%> </a>
                         </td>
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
                        <a href="productSearchRetail.asp?pCategoryID=-3" title="Items on sale"> 
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
    
    
  
    
    
    <td class="productSearchlist">
    
    
  
  
    
      <table class="table98border" >
          <tr>
      <th class="thfeatured"  >
      PRODUCTS
      </th>
      </tr>
      
      
        <tr>
          <td width="100%">
        

            <table class="table98noborder">
            <%
           
            
            
              if intCategoryID >0 then
		            Response.Write ("<tr><td colspan=""4"" width=""100%"" height=""25"">" &_
                        "Product(s) for <b>" & (strCategoryName) & "</b>." )
        		    Response.Write ("</td></tr>")
            elseif Cint(intCategoryID) =-2 then
            	 Response.Write ("<tr><td colspan=""4"" width=""100%""  height=""15"">" &_
                        "<b>Overstock from the factory, contact us for special discount</b><br>")
                        
	            Response.Write ("</td></tr>")

				Response.Write ("<tr><td colspan=""4"" width=""100%""  height=""15"">" &_
                        "<b>Click on the # for airplanes, vehicles Close out <a href='productSearchRetail.asp?pcategoryid=26'>1</a><a href='productSearchRetail.asp?pcategoryid=30'> 2 </a> <a href='productSearchRetail.asp?pcategoryid=2'> 3 </a> <a href='productSearchRetail.asp?pcategoryid=15'> 4 </a>  <a href='productSearchRetail.asp?pcategoryid=5'> 5 </a>  <a href='productSearchRetail.asp?pcategoryid=14'> 6 </a>     </b><br>")
                        
	            Response.Write ("</td></tr>")
				
				'NEW PRODUCTS
	         elseif Cint(intCategoryID) =-1 then
            	 Response.Write ("<tr><td colspan=""4"" width=""100%""  height=""25"">" &_
                        "<b>New Products </b>")
	            Response.Write ("</td></tr>")
				' ON SALE
	         elseif Cint(intCategoryID) =-3 then
            	 Response.Write ("<tr><td colspan=""4"" width=""100%""  height=""25"">" &_
                        "<b>Products On Sale </b>")
	            Response.Write ("</td></tr>")
				'TRICK OR TREAT
			elseif Cint(intCategoryID) =-5 then
            	 Response.Write ("<tr><td colspan=""4"" width=""100%""  height=""25"">" &_
                        " <img src=""../images/pirateIcon1.jpg"" width=""50""> <b>Trick or Treat??!! </b> <img width=""50"" src=""../images/pirateIcon1.jpg""> ")
	            Response.Write ("</td></tr>")
				
				'XMAS
				elseif Cint(intCategoryID) =-6 then
            	 Response.Write ("<tr><td colspan=""4"" width=""100%""  height=""25"">" &_
                        " <b>Happy Holidays </b> <img src=""../images/xmas/gift.png"" width=""50""> <img width=""50"" src=""../images/xmas/candycane.png""> ")
	            Response.Write ("</td></tr>")
				
				'BEST SELLERS 
				elseif Cint(intCategoryID) =-4 then
            	 Response.Write ("<tr><td colspan=""4"" width=""100%""  height=""25"">" &_
                        " <img src=""../images/bestseller.jpg"" width=""50""> <b>OMH Top 100 Best Sellers </b> <img width=""50"" src=""../images/bestseller.jpg""> ")
	            Response.Write ("</td></tr>")
				
	            
			else	

            Response.Write ("<tr><td colspan=""4"" width=""100%"" height=""25"">" &_
                        "Product(s) for all Categories." )
            Response.Write ("</td></tr>")
            end if
            
            
            
            
            
            
            If rstProduct.RecordCount > 0 Then
            %>

              <%
              Dim strThumbImage(4)
              Dim strLargeImage(4)
              Dim strProductNames(4)
              Dim strPrices(4)
              Dim strSale(4)
			  
			  Dim strmapPrices(4)
              Dim strmapSale(4)
              
			  
              Dim discount(4)
              Dim strProductIDs(4)
              Dim strLongDes(4)
              Dim strSpecifications(4)
              Dim strQuantities(4)
			  dim intRating(4)
			  dim intTotalRating(4)
			   	
					
			  
              
              Dim IsNew(4)
              Dim intProductOnRow
			
			   intProductOnRow = 0
			
			  
			  
              While Not rstProduct.EOF and intRecNum <= intWhichPage*intMaxProduct 
                'Display if this product in right page
				If (intRecNum > (intWhichPage - 1) * intMaxProduct) and (intRecNum <= intWhichPage * intMaxProduct) Then
				intProductOnRow = 0
				while Not rstProduct.EOF and intProductOnRow<totalColumn
                

					strThumbImage(intProductOnRow) = Trim(rstProduct("Thumb_Img"))
                    strLargeImage(intProductOnRow) = Trim(rstProduct("Large_Img"))
                    strProductNames(intProductOnRow) = (HTML_Encode(Trim(rstProduct("Product_Name"))))
                    strProductIDs(intProductOnRow) = rstProduct("Product_ID")
                    strPrices(intProductOnRow) = rstProduct("Price")
                    strSale(intProductOnRow)=rstProduct("sale")
					
                    strmapPrices(intProductOnRow) = rstProduct("map")
                    strmapSale(intProductOnRow)=rstProduct("sale_map")
					
					
                    strQuantities(intProductOnRow) = rstProduct("Quantity")
                    IsNew(intProductOnRow) = rstProduct("IsNew")
                 
                    strSpecifications(intProductOnRow) = rstProduct("Specification")
                    strLongDes(intProductOnRow) = rstProduct("Long_Desc")
					
						  intRating(intProductOnRow)=rstProduct("rating")
				  intTotalRating(intProductOnRow)=rstProduct("total_review")
				  
				  
                    rstProduct.MoveNext
                    intRecNum = intRecNum + 1
                    intProductOnRow = intProductOnRow+1
					
                wend
				
  
             
             
              for i = 0 to totalColumn-1
                      
                if isNull(strSale(i)) then
                    strSale(i) = strPrices(i)
                end if
                
                
                
                  if cdbl(strPrices(i))> 0 then
                   discount(i) = round(100*(1.0-cdbl(strSale(i))/cdbl(strPrices(i))), 0)
                   
                    else
                    discount(i)=0
                   end if
				   
				   
				   '----------------------------------------------------------------------
				   'calculate price for wholesaler and consumer
				   if len(Session("wholesaler")) > 1 then
                       
                     
                     else
					 ' change to MAP MARK UP
                   strprices(i)=round(cdbl(strmapPrices(i))*mapMarkup, 2)
				 
				   strSale(i)=round(cdbl(strmapSale(i))*mapMarkup, 2)
                     
                   end if
                       
                                          
				    'end display price before for wholesaler and consumer
				   '----------------------------------------------------------------------------
				   
				   
				   
				   
				   
				   
              next
                  'end calculate discount
           
                 
                  '***************************************************************** 
                  
                  
                %>
                  <tr>
                  <%
                  Select case intProductOnRow
                    Case 4
                  %>
                      <td width="25%" class="tdImgBox " height="79">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(0)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(0)%>"  width="100"  >
                        </a>
                      </td>
                      <td class="tdImgBox " height="79">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(1)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(1)%>" width="100" >
                        </a>
                      </td>
                      <td  height="79" class="tdImgBox ">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(2)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(2)%>"  width="100">
                        </a>
                      </td>
                      <td  height="79" class="tdImgBox ">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(3)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(3)%>"      width="100"               </a>
                      </td>
                  <%
                  Case 3
                  %>
                      <td  width = "33%" height="79" class="tdImgBox ">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(0)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(0)%>"  width="100">
                        </a>
                      </td>
                      <td  height="79" class="tdImgBox ">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(1)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(1)%>"  width="100">
                        </a>
                      </td>
                      <td  height="79" class="tdImgBox">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(2)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(2)%>"  width="100">
                        </a>
                      </td>
					  <td  height="79" class="tdImgBoxEmpty" >&nbsp;</td>
                    
                  <%
                  Case 2
                  %>
                      <td width="50%" height="79" class="tdImgBox ">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(0)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(0)%>"  width="100">
                        </a>
                      </td>
                      <td  height="79" class="tdImgBox ">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(1)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(1)%>" width="100">
                        </a>
                      </td>
                      <td  height="79" class="tdImgBoxEmpty" >&nbsp;</td>
                      <td  height="79" class="tdImgBoxEmpty">&nbsp;</td>
                  <%
                  Case 1
                  %>
                      <td  height="79" class="tdImgBox ">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(0)%>">
                          <img alt="click here to view detail" border="0" src="../ThumbImages/<%=strThumbImage(0)%>" width="100">
                        </a>
                      </td>
                      <td  height="79" class="tdImgBoxEmpty" >&nbsp;</td>
                      <td height="79" class="tdImgBoxEmpty" >&nbsp;</td>
                      <td  height="79" class="tdImgBoxEmpty" >&nbsp;</td>
                  <%
                  End Select
                  %>
                  </tr>
                  <tr>
                  <%
                  Select case intProductOnRow
                    Case 4
                  %>
                      
                      
                      <td  class= "cssTextCenterBottom " height="44">
                     
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(0)%>"><%=left(strProductNames(0), 30)%></a>
						
						
						
						
						
						<br><%For i = 1 To intRating(0)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(0)> 0 then
  %>
  <%=intTotalRating(0)%> Ratings<br><%end if%>
  
  
  
                        
                        <%
                        if IsNew(0)=1 then
                       	%>
                        <font class="attentionTextNew">
									NEW!!!</font>                   	
                       <%	
                       end if
                        %>
                        
                        
                        <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(0)%>
                       
                       <% if cdbl(strSale(0)) < cdbl(strPrices(0)) then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(0)%> <%end if %> </font> 
                       
                       <%else
                       
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                   
                      %> <br>
					  
					  
	
	

  
					  Price: $ <%=strPrices(0)%>
                       
                       <% if cdbl(strSale(0)) < cdbl(strPrices(0)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(0)%> <%end if %> </font> 
                       
              
                       
                       
                       <%
                       
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                        
                        
                        
                       <br />
              
  		
                       
                      </td>
                      
                      
                      
                      
                      
                      
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(1)%>"><%=left(strProductNames(1), 30)%></a>
						
						<br><%For i = 1 To intRating(1)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(1)> 0 then
  %>
  <%=intTotalRating(1)%> Ratings<br>
  <%end if%>
  
  
                        
                    
                        <%
                        if IsNew(1)= 1 then
                       	%>
                       <font class="attentionTextNew">
									NEW!!!</font>                         	
                       <%	
                       end if
                        %>
                        
                        
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(1)%>                                              
                          <% if cdbl(strSale(1)) < cdbl(strPrices(1)) then %>                                              
                        <font class="salePriceText">  Sale $ <%=strSale(1)%> <%end if %> </font> 
                    
                       
                       <%else
                       
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                   
                      %> <br> Price: $ <%=strPrices(1)%>
                       
                       <% if cdbl(strSale(1)) < cdbl(strPrices(1)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(1)%> <%end if %> </font> 
                       
                       
                       
                       <%
                       
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                       
                        
                        <br>

                      </td>
                      
                      
                      
                      
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(2)%>"><%=left(strProductNames(2), 30)%></a>
              
		<br><%For i = 1 To intRating(2)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(2)> 0 then
  %>
  <%=intTotalRating(2)%> Ratings<br>
  <%end if%>
  
  
                        <%
                        if IsNew(2)= 1 then
                       	%>
                      <font class="attentionTextNew">
									NEW!!!    </font>                    	
                       <%	
                       end if
                        %>
                        
                        
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(2)%>                                              
                         <% if cdbl(strSale(2)) < cdbl(strPrices(2)) then %>                                              
                      <font class="salePriceText">  Sale $ <%=strSale(2)%> <%end if %> </font> 
                    
                       
                       <%else
                       
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                
                     %> <br> Price: $ <%=strPrices(2)%>
                       
                       <% if cdbl(strSale(2)) < cdbl(strPrices(2)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(2)%> <%end if %> </font> 
                       
                       
                       
                       <%
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                       
                        <br>
                     
                      </td>
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(3)%>"><%=left(strProductNames(3), 30)%></a>
						
		<br><%For i = 1 To intRating(3)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(3)> 0 then
  %>
  <%=intTotalRating(3)%> Ratings<br>
  <%end if%>
  
  
                      
                        <%
                        if IsNew(3)= 1 then
                       	%>
                      <font class="attentionTextNew">
									NEW!!!    </font>                     	
                       <%	
                       end if
                        %>
                        
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(3)%>                                              
                          <% if cdbl(strSale(3)) < cdbl(strPrices(3)) then %>                                              
                        <font class="salePriceText">  Sale $ <%=strSale(3)%> <%end if %> </font> 
                    
                       
                       <%else
                       
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                  
                     %> <br> Price: $ <%=strPrices(3)%>
                       
                       <% if cdbl(strSale(3)) < cdbl(strPrices(3)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(3)%> <%end if %> </font> 
                       
                       
                       
                       <%
                       
                       
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                         <br>
                      
                      </td>
                  <%
                  Case 3
                  %>
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(0)%>"><%=strProductNames(0)%></a>
						
							<br><%For i = 1 To intRating(0)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(0)> 0 then
  %>
  <%=intTotalRating(0)%> Ratings<br>
  <%end if%>
  
  
						
						
                       
                        <%
                        if IsNew(0)=1 then
                       	%>
                        <font class="attentionTextNew">
									NEW!!!      </font>                  	
                       <%	
                       end if
                        %>
                        
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(0)%>                                              
                     <% if cdbl(strSale(0)) < cdbl(strPrices(0)) then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(0)%> <%end if %> </font> 
                    
                       
                       <%else
                       
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                  
                      %> <br>Price: $ <%=strPrices(0)%>
                       
                       <% if cdbl(strSale(0)) < cdbl(strPrices(0)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(0)%> <%end if %> </font> 
                       
                       
                       
                       <%
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                        <br>
              
                      </td>
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(1)%>"><%=strProductNames(1)%></a>
						
						
						
							<br><%For i = 1 To intRating(1)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(1)> 0 then
  %>
  <%=intTotalRating(1)%> Ratings<br>
  <%end if%>
  
  
  
                        <%
                        if IsNew(1)= 1 then
                       	%>
                        <font class="attentionTextNew">
									NEW!!!        </font>                	
                       <%	
                       end if
                        %>
                        
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(1)%>                                              
                     <% if cdbl(strSale(1)) < cdbl(strPrices(1)) then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(1)%> <%end if %> </font> 
                    
                       
                       <%else
                       
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                   
                      %> <br>Price: $ <%=strPrices(1)%>
                       
                       <% if cdbl(strSale(1)) < cdbl(strPrices(1)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(1)%> <%end if %> </font> 
                       
                       
                       <%
                       
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                        <br>
                 
                      </td>
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(2)%>"><%=strProductNames(2)%></a>
						
							<br><%For i = 1 To intRating(2)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(2)> 0 then
  %>
  <%=intTotalRating(2)%> Ratings<br>
  <%end if%>
  
  
                        <%
                        if IsNew(2)= 1 then
                       	%>
                        <font class="attentionTextNew">
									NEW!!!    </font>                    	
                       <%	
                       end if
                        %>
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(2)%>                                              
                      <% if cdbl(strSale(2)) < cdbl(strPrices(2)) then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(2)%> <%end if %> </font> 
                    
                       
                       <%else
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                
                     %> <br> Price: $ <%=strPrices(2)%>
                       
                       <% if cdbl(strSale(2)) < cdbl(strPrices(2))and displaySalePrice_retail=1  then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(2)%> <%end if %> </font> 
                       
                       
                       
                       <%
                       
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                        <br>
                    
                      </td>
                      <td  class="cssTextCenterBottom" height="44">&nbsp;</td>
                  <%
                  Case 2
                  %>
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(0)%>"><%=strProductNames(0)%></a>
						
							<br><%For i = 1 To intRating(0)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(0)> 0 then
  %>
  <%=intTotalRating(0)%> Ratings<br>
  <%end if%>
  
  
  
  
                        <%
                        if IsNew(0)=1 then
                       	%>
                        <font class="attentionTextNew">
									NEW!!!   </font>                     	
                       <%	
                       end if
                        %>
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(0)%>                                              
                       <% if cdbl(strSale(0)) < cdbl(strPrices(0)) then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(0)%> <%end if %> </font> 
                       
                       <%else
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                 
                      %> <br>Price: $ <%=strPrices(0)%>
                       
                       <% if cdbl(strSale(0)) < cdbl(strPrices(0)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(0)%> <%end if %> </font> 
                       
                       
                       
                       <%
                       
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                         <br>
                 
                      </td>
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(1)%>"><%=strProductNames(1)%></a>
						
							<br><%For i = 1 To intRating(1)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(1)> 0 then
  %>
  <%=intTotalRating(1)%> Ratings<br>
  <%end if%>
  
  
  
                        <%
                        if IsNew(1)=1 then
                       	%>
                        <font class="attentionTextNew">
									NEW!!!          </font>              	
                       <%	
                       end if
                        %>
                        
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(1)%>                                              
                        <% if cdbl(strSale(1)) < cdbl(strPrices(1)) then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(1)%> <%end if %> </font> 
                    
                       
                       <%else
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                   
                     %> <br> Price: $ <%=strPrices(1)%>
                       
                       <% if cdbl(strSale(1)) < cdbl(strPrices(1)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(1)%> <%end if %> </font> 
                       
                       
                       
                       <%
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                        <br>
                   
                      </td>
                      <td  class="cssTextCenterBottom" height="44">&nbsp;</td>
                      <td  class="cssTextCenterBottom" height="44">&nbsp;</td>
                  <%
                  Case 1
                  
                        
                        
                  %>
                      <td  class="cssTextCenterBottom" height="44">
                        <a href="productsdetailsRetail.asp?productId=<%=strProductIDs(0)%>"><%=strProductNames(0)%></a>
						
							<br><%For i = 1 To intRating(0)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  if intTotalRating(0)> 0 then
  %>
  <%=intTotalRating(0)%> Ratings<br>
  <%end if%>
  
  
  
                        <%
                        if IsNew(0)=1 then
                       	%>
                        <font class="attentionTextNew">
									NEW!!!          </font>              	
                       <%	
                       end if
                        %>
                         <%
                        '******************************************************************************
                        ' display price if user log in
                       if len(Session("wholesaler")) > 1 then
                       
                       %>
                       <br />
                       $ <%=strPrices(0)%>                                              
                       <% 
                     
                        
                       
                       
                       if cdbl(strSale(0)) < cdbl(strPrices(0)) then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(0)%> <%end if %> </font> 
                       
                       <%else
                       
                       ' NOT LOG IN
                       ' IF ON SALE, DISPLAY ON SALE %
                   
                     %> <br> Price: $ <%=strPrices(0)%>
                       
                       <% if cdbl(strSale(0)) < cdbl(strPrices(0)) and displaySalePrice_retail=1 then %>                                              
                     <font class="salePriceText">  Sale $ <%=strSale(0)%> <%end if %> </font> 
                       
                       
                       
                       <%
                       
                       end if
                       ' end display price
                        '******************************************************************************
                       %>
                       
                        <br>
                       
                      </td>
                      <td  class="cssTextCenterBottom" height="44">&nbsp;</td>
                      <td  class="cssTextCenterBottom" height="44">&nbsp;</td>
                      <td  class="cssTextCenterBottom" height="44">&nbsp;</td>
                  <%
                  End Select
                  %>
                  </tr>
                  <tr>
                    <td width="100%" colspan="4" height="21">&nbsp;</td>
                  </tr>
              <%
                Else
                  rstProduct.MoveNext
                  intRecNum = intRecNum + 1
                End If
                
              Wend

              rstProduct.Close
              cnn.Close
              Set rstProduct = Nothing
              Set cnn = Nothing
              %>
              <tr>
                <td width="100%" colspan="4" align="center" height="44">
                  <table border="0" cellpadding="0" cellspacing="0" >
                    <tr>
                    
                    

            <%
              'Print out page links
              If intNumOfPage > 1 Then
               
 

            
              Response.Write("<TD width=""80%"" colspan=""3"" class=""cssTextCenter"">")
              %>
              <div class="paginationlinks">
              <%
               if intWhichPage > 1 then
                      Response.Write("<a  href=""productSearchRetail.asp?pCategoryID="& intcategoryID&""&"&pWhichPage=" & 1 & """>" & "|<&nbsp;</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  " )
                      Response.Write("<a  href=""productSearchRetail.asp?pCategoryID="& intcategoryID&""&"&pWhichPage=" & intWhichPage-1 & """>" & "<&nbsp;</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
              end if
                
                
              Dim i 
              For i = 1 To intNumOfPage
               If i <> intWhichPage Then



                'if search form is used
					'If Len(Request.Form("pAction")) > 0 Then
					If Len(Request.Form("formSearch")) > 0 Then

                   intCategoryID = Cint(Trim(Request.Form("pCategoryID")))
                else
                    intCategoryID = Cint(Trim(Request.QueryString("pCategoryID")))
                end if
                'intCategoryID= Request.QueryString("pCategoryID")

                'Response.Write("<a href=""search_product.asp?pWhichPage=" & i & """>" & i & "&nbsp;</a>")

                Response.Write("<a  href=""productSearchRetail.asp?pCategoryID="& intcategoryID&""&"&pWhichPage=" & i  & """>" & i  & "&nbsp;</a>")







               Else
               
                Response.Write("<b>" & i & "&nbsp;</b>")
               End If
               
               
               
               
               if i mod 20 = 0 then
               		'response.write("<br>")
               end if
              Next
            
              if intwhichpage <intNumOfPage then
                Response.Write(" &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a  href=""productSearchRetail.asp?pCategoryID="& intcategoryID&""&"&pWhichPage=" & intWhichPage+1 & """>" & "></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
                Response.Write("<a  href=""productSearchRetail.asp?pCategoryID="& intcategoryID&""&"&pWhichPage=" & intNumofPage & """>" & ">|</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>")
                    
               

                
              
                
      
       
                  
              end if
               
              Response.Write("<br>page&nbsp;" & intWhichPage & "&nbsp;of&nbsp;" & intNumOfPage & "</div></TD>")
              
            
                
            
              
              Else
              %>
              <td width="46%" align="center">
              
              </td>
            <%
              End If
              %>
              

              
                <td width="27%" align="center" valign="bottom">
                
                
                  
            </td>
                </tr>
                    <tr>

              <td width="46%" align="center">&nbsp;
              
                  </td>
              

              
                <td width="27%" align="left" valign="bottom">&nbsp;
                
                                  </td>
                </tr>
                    <tr>

              <td width="46%" align="center">
              
             
                  
                  </td>
              

              
                <td width="27%" align="left" valign="bottom">
               
</td>
                </tr>
                </table>
				 <br>
				  
				
				
				  
				  <!-- ShareThis BEGIN --><div class="sharethis-inline-share-buttons"></div><!-- ShareThis END -->
				  <br>
				  
				  <br>
				  
				  <br>
				  <br>
				  
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>	  <br>
				  
				  <br>	  <br>
				  <br>	  <br>

				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  <br>
				  
				  
                </td>
              </tr>
			  
            </table>
            <%
            End If
            %>
          </td>
        </tr>
      </table>
    </td>
	<td>
	</td>
	
  
  </tr>
</table>

<!--#Include file="footerRetail.asp"  -->
</td>

<td class = "mainright" >    </td>
</tr>



</table>
</body>
</html>
