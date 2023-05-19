
<%
'--------------------------------------------------------------
'      Coded By: Eric
'       Purpose: Display all category and search product form.
'   Used Tables: products
'  Invoked From: productsearch
'       Invokes: order.asp
'Included Files: headerRetail.asp, footerRetail.asp, animate.js, StyleSheet.css, asp_lib.inc.asp.
'--------------------------------------------------------------
'Updated By   Eric    Date 01/04/2011   Comments
'Display products details
'--------------------------------------------------------------
%>
<!-- #include file="../include/asp_lib.inc.asp" -->
<!-- #include file="../include/sqlCheckInclude.asp" -->
<%

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   	'if not login then redirect to login page
   	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  'check if not login
						if len(Session("consumer")) <1 then
  							
						  'response.redirect("login.asp")
						else
						 ' response.write ("user " & Session("consumer"))


						
				      end if 
					  

Dim strSQLCateCombo, cnn, rstProduct, strSQLCmd,strSQL
Dim ProductPrice(1)
Dim ProductQuan(1)

dim intProductId, intReview
dim strThumbImage(1)
dim strLargeImage(1) 
dim strProductName(1)
dim strProductID(1)
dim strItem(1)
dim strPrice(1) 
dim strRegPrice(1)
Dim strQuantity(1) 
dim strSpecification(1) 
dim strLongDes(1) 
dim strHistory(1)
dim rel_prodID
Dim rstCategory
dim strMSRP(1)
dim quantityonhand(1)
dim strRetailSalePrice(1)
dim msrpMarkup, mapMarkup
msrpMarkup=session("retailMarkup")


if len(session("mapMarkup")) > 0 then
	mapMarkup=Cdbl (session("mapMarkup"))
else
	mapMarkup=1
end if

dim intstar(1), strComment(1), strReviewby(1), strReviewDate(1), strSource(1)
dim intTotalReview(1)
dim rating(20), reviewby(20), reviewdate(20), reviewsource(20), reviewComment(20), title(20)

dim strName, strEmail, 	strTitle, strRating, Source, Comment, ipaddress


dim folder
dim item, itemID, itemCode


'check if login
'if len(Session("consumer")) > 1 then


Set objFSO = Server.CreateObject("Scripting.FileSystemObject")






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
intproductid=0
   if len(Request.QueryString("ProductId"))=0 then
   else
	intProductId = Cint(TRIM(Request.QueryString("ProductId")))
   end if
   
   itemCode=(TRIM(Request.QueryString("item")))
   
   if isnull(Request.QueryString("review")) or len(Request.QueryString("review")) = 0 then 
		intReview=0
	else
		intReview = Cint(TRIM(Request.QueryString("review")))
		
   end if
  'Create connection and query category data.
  
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.ConnectionString = Application.Contents("dbConnStr")
	cnn.Open
	Set rstProduct = Server.CreateObject("ADODB.Recordset")
	Set rstReview = Server.CreateObject("ADODB.Recordset")
	
 
  dim counter, reviewActive
  counter=0
  'to diplay product together with its related product
  
 
 
 
 '----------------------------------------------------------------------------------------------------------
'Insert Reviews into database 
 
'Get data
strName = fixstring(Trim(Request.Form("FullName")))
dim insertReviewSuccess 
insertReviewSuccess  = 0
dim reviewUser
reviewUser=Session("consumer")

if intReview=1 and len(strName)>0  then 
	'strName = fixstring(Trim(Request.Form("FullName")))
	strEmail = fixstring(Trim(Request.Form("Email")))
	strTitle = fixstring(Trim(Request.Form("Title")))
	strRating = fixstring(Trim(Request.Form("Rating")))
	Source = fixstring(Trim(Request.Form("Source")))
	Comment = fixstring(Trim(Request.Form("Comment")))
	ipaddress= fixstring(trim(Request.ServerVariables("remote_addr")))
	
	if strRating>2 then
		reviewActive=1
	else	
		reviewActive=0
	end if
	

if len(strName)> 0 and len(strRating)>0 then
 strSQLcmd="insert into reviews (reviewid ,comment,reviewby, reviewdate,stars,source,productId,review_user, active,title,email ) select max(reviewid) + 1, '" &  Comment & "' , '" &  strName & "', getdate(), " &  strRating & ", '" & Source & "' , " & intProductId & ",'" & reviewUser & "'," & reviewActive & ",'" &  strTitle & "', '" & strEmail & "' from reviews "
 
 'response.write (strsqlcmd)
	rstReview.open strSqlcmd, cnn,3
	insertReviewSuccess =1
	
	
	dim rstMsg
	set rstMsg=  Server.CreateObject("adodb.RecordSet")
	rstMsg.open "select * from screenmessage", cnn, 3
  
  
	
	
	  'Send mail when someone write a review
  '---------------------------------------------------------------------------------------------------
  strMailContent = "Dear OMHUSA," & "<br><br>"
  strMailContent = strMailContent & "You have received a review from "
  strMailContent = strMailContent & strName & "<br>"
  
   strMailContent = strMailContent & "Email: " & strEmail & "<br>"
   strMailContent = strMailContent & "User Name: " & reviewUser & "<br>"
   
   strMailContent = strMailContent & "Product link: " & "https://omhusa.com/productsdetailsRetail.asp?productId=" & intProductId & "<br>"
      
   strMailContent = strMailContent & "Feedback Short Description: " & strTitle & "<br>"
   strMailContent = strMailContent & "Rating: " & strRating & " stars <br>"
   
   strMailContent = strMailContent & "Source: " & Source & "<br>"
   
   
   'strMailContent = strMailContent & "ip address: " & ipaddress & "<br>"
   strMailContent = strMailContent & "Full Review: " & "<br>"
   strMailContent = strMailContent & comment & "<br><br>"
  
       
 
 dim sch, cdoconfig, cdomessage
   sch = "http://schemas.microsoft.com/cdo/configuration/" 
 
    Set cdoConfig = CreateObject("CDO.Configuration") 
 
    With cdoConfig.Fields 
        .Item(sch & "sendusing") = 2 ' cdoSendUsingPort 
        .Item(sch & "smtpserver") = rstMsg("ms13")
	.Item(sch & "smtpauthenticate") =1
	.Item(sch & "sendusername") =rstMsg("ms12")
	.Item(sch & "sendpassword") =rstMsg("ms11")
        .update 
    End With 
 
    Set cdoMessage = CreateObject("CDO.Message") 
    dim strSubject, feedbackcontainlink
	

    With cdoMessage 
        Set .Configuration = cdoConfig 
        .From = strEmail

	 .To = rstMsg("ms14")
	


        .Subject = " Product Review from: " & strName
        .HTMLBody = strMailContent
		.cc=rstMsg("ms17")
	
	'if InStr(.htmlbody, "http")=0 then 
        .Send 
		on error resume next
		
	
    End With 
    
' Error Handler
If Err.Number <> 0  Then
   ' Error Occurred / Trap it
   response.write ("Send email failed")

   On Error Goto 0  ' But don't let other errors hide!
   ' Code to cope with the error here
End If
'On Error Goto 0 ' Reset error handling.




 
    Set cdoMessage = Nothing 

    Set cdoConfig = Nothing 
	
	'----------------------------------------------------------------------------------------------------------------------------------------
	'end send email
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
end if

'rstReview.close()
	
		


end if
		
'----------------------------------------------------------------------------------------------------------
		
  %>
 
                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">

 
<script language="JavaScript1.2" src="../include/javascript.js"></script>
<script src="../include/spotlight_master/dist/spotlight.bundle.js"></script>

 <script language="JavaScript1.2">
 
 function validateData()
 {
 var StrName=document.getElementsByName("FullName");
 var StrEmail=document.getElementsByName("Email");
 var StrSource=document.getElementsByName("Source");
 var StrRating=document.getElementsByName("Rating");
 var StrComment=document.getElementsByName("Comment");
 var StrTitle=document.getElementsByName("Title");
 
	


//Check if Name is empty
  if (isBlank(StrName[0].value)){
    alert("Please input your full name!");
    StrName[0].focus(); 
	return false;
  }
  
 
	   
  
  //Check if Email is valid
  if (! isEmail(StrEmail[0].value)){
   alert("Invalid email address!");
   StrEmail[0].focus();
    return false;
  }
  
  
  
  //Check if title is empty
  if (isBlank(StrTitle[0].value)){
    alert("Please input your short description!");
    StrTitle[0].focus();
    return false;
  }
  
  //Check if title is empty
  if  (StrComment[0].value.toUpperCase().indexOf("HTTP") >= 0) {
  
  
   alert("Please not include hyperlink in the comment!");
    StrComment[0].focus();
   return false;
  }
  
  
  
  
  return true;
}


 
 
 
 </script>
 
 
  
 <script type='text/javascript' src='https://platform-api.sharethis.com/js/sharethis.js#property=6202fc68049246001a151155&product=inline-share-buttons' async='async'>
  </script>
  <style>
  
   input.input10 {
	font-family:   "Bookman Old Style", "helvetica neue unltralight", "Verdana" , "arial";
	font-size: small;
	color: #757170;
	/*color: #666666;*/
	background-color: #EDE9E8;
}



</style>



<script type="text/javascript">
	document.cookie = "screenSize=" + screen.width;
	
</script>
<%
if isnull(Request.Cookies("screenSize")) or len(trim(Request.Cookies("screenSize")))=0 then

	
	%>
	
	
<script type="text/javascript">
window.onload = function() {
    if(!window.location.hash) {
        window.location = window.location + '#loaded';
        window.location.reload();
    }
}
</script>
<%
		
	
end if
%>

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
                                <input class="input10" name="pAction" type="submit" value="Search" />
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
					 
					 
                      <span class="cssLink"><a href="productsearchRetail.asp?pCategoryID=-1" title="Ship Model - New Products "> <i class="AttentionText">New Products!!!</i>  </a></span></td>
                    </tr>
					
                   
                    <%
                    While Not rstCategory.EOF
                    %>
                   
                    
                   
                    <tr>
                      <td width="100%" align="left">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="left" class = "tdmargin10">
					 
					 
                      <span class="leftlink">
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
    
    
  
     
                  
				  
    
    <td class="productSearchlist">
    
    
    
  
    
 <%

 


	  
  	  
	  
dim reviewCount



  Do 
reviewCount=0
'For main product counter =0, counter =1 is for related product
    if counter =0 then
		if intProductId>0 then
			strSQLCmd="select products.*, review_summary.*, isnull(dropshiptemplate2013.quantityonhand, 0) as quantityonhand, isnull(dropshiptemplate2013.Description, '') as description, isnull(dropshiptemplate2013.History, '') as history, isnull(dropshiptemplate2013.map, 0) as map, isnull(dropshiptemplate2013.sale_Map, 0) as saleMap from products left join dropshiptemplate2013 on products.item=dropshiptemplate2013.product_code left  join review_summary on products.product_id=review_summary.productid where products.Product_Id = " & intProductId 
			strSQlcmd_review = "select top 10 * from reviews join products on reviews.productId=products.product_id where reviews.active =1 and reviews.productid= " & intProductId &  " order by reviewdate desc"
		elseif len(itemcode)>0 then
			strSQLCmd="select products.*, review_summary.*, isnull(dropshiptemplate2013.quantityonhand, 0) as quantityonhand, isnull(dropshiptemplate2013.Description, '') as description, isnull(dropshiptemplate2013.History, '') as history , isnull(dropshiptemplate2013.map, 0) as map, isnull(dropshiptemplate2013.sale_Map, 0) as saleMap from products join dropshiptemplate2013 on products.item=dropshiptemplate2013.product_code left  join review_summary on products.product_id=review_summary.productid where products.item = '" & itemCode &"'"
			strSQlcmd_review = "select top 10 * from reviews join products on reviews.productId=products.product_id where reviews.active =1 and products.item = '" & itemCode & "' order by reviewdate desc"
		end if
		
		
    elseif isNull(rel_prodID)=false or len(rel_ProdId)>0 then
	'for related product; this one happen when counter=1
    	intProductId=Cint(rel_ProdId)
	   	strSQLCmd="select products.*, review_summary.*, isnull(dropshiptemplate2013.quantityonhand, 0) as quantityonhand, isnull(dropshiptemplate2013.Description, '') as description, isnull(dropshiptemplate2013.History, '') as history , isnull(dropshiptemplate2013.map, 0) as map, isnull(dropshiptemplate2013.sale_Map, 0) as saleMap from products left join dropshiptemplate2013 on products.item=dropshiptemplate2013.product_code left join review_summary on products.product_id=review_summary.productid where products.Product_Id = " & intProductId
		strSQlcmd_review = "select top 10 * from reviews where active=1 and  productid= " & intProductId  & " order by reviewdate desc"
		
		
    end if
    


	rstProduct.Open strSQLCmd, cnn, 3
	rstReview.open strSQlcmd_review, cnn, 3
	

'Write down session if there is more one page
 If Not rstProduct.EOF Then

                  'Get data
                  
             
				    strThumbImage(counter) = rstProduct("Thumb_Img")
					strItem(counter)=rstProduct("item")
					
                  strLargeImage(counter) = rstProduct("Large_Img")
                  strProductName(counter) = Trim(rstProduct("Product_Name"))
                  strPrice(counter) = rstProduct("sale")
				  if not isnull(rstProduct("price")) then
				    strRegprice(counter)=cdbl(rstProduct("price"))
					else
					strRegprice(counter)=0
				end if
					
                  strQuantity(counter) = rstProduct("Quantity")
   				    strSpecification(counter) = trim(rstProduct("USASpec"))
					
					
					strHistory(counter)=trim(rstProduct("history"))
					
					if len(trim(rstProduct("Long_Desc"))) > 0 then
						strLongDes(counter) = trim(rstProduct("Long_Desc"))						
					else
						strLongDes(counter) = trim(rstProduct("description"))
				  end if
				  
				   strLongDes(counter)=replace(strLongDes(counter), chr(13)+ chr(10), "<br>")
				   strHistory(counter)=replace(strHistory(counter), chr(13)+ chr(10), "<br>")
				  
                  rel_prodID=rstProduct("rel_prodId")
				
				  intStar(counter)=(rstProduct("rating"))
				  intTotalReview(counter)=rstProduct("total_review")
				  if not isnull(rstProduct("quantityonhand")) then
					quantityonhand(counter)=cint(rstProduct("quantityonhand"))
				else 
					quantityonhand(counter)=0
				end if
				
				  
				  
				  
				  ' change MSRP to MAP price
				  if not isnull(rstProduct("map")) then
				  
					'strMSRP(counter)=round(cdbl(rstProduct("price"))*msrpMarkup, 2)
					' change MSRP to MAP price
					strMSRP(counter)=round(cdbl(rstProduct("map"))*mapMarkup, 2)
					
				  else
				  strMSRP(counter)=0
				  end if
				  
				  if not isnull(rstProduct("saleMap")) then
					strRetailSalePrice(counter)=round(cdbl(rstProduct("saleMap"))*mapMarkup, 2)
				  else
				  
				  strRetailSalePrice(counter)=strMSRP(counter)
				  end if
				 
				  
				  
				  
				  
				'if item is not on sale then price is regular price otherwise price is sale price
				if isNull(strPrice(counter)) = true and isNull(strRegPrice(counter)) =false then
					strPrice(counter)=cdbl(strRegPrice(counter))
				end if


 %>
 
 <%
 
 
  If Not rstReview.eof Then
  
  do 


                  'Get data
				  title (reviewcount) = rstReview("title")
				  
				  rating (reviewCount) = rstReview("stars")
				  reviewBy(reviewcount)=  rstReview ("reviewby") 
				  reviewdate(reviewcount) = rstReview("reviewdate")
				reviewSource(reviewCount)=	rstReview("source")
				reviewComment(reviewCount)= rstReview("Comment")
				rstReview.movenext
				reviewcount=reviewcount+1
				
		loop until rstReview.eof	

  	
  end if
  
  'item=ItemCode
  item=strItem(counter)  
  
             

 %>
 
 
  <!--<table class="productDetail" name="maintable1" > -->
  
 
<table class="tableShowImages">
 
 
   <tr>
      <th class="thfeatured"  colspan="2"  >
      PRODUCT DETAILS
      </th>
      </tr>
      
     <tr>
      
      <td  >
	  </td>
       
	
    </tr>
    
    <tr>
      
      <td ><b>
	<b>Item: </b> <%=item%>  
	 |  Name: </b> <%=strProductName(counter)%> 
	  <%
	  if isnull(intstar(counter)) then
	  
	 else
	
	%>
	</break></break>
		  
	<%
	response.write ( "    ")
	
For i = 1 To intstar(counter)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  %>
   <a href="#productReview"> ( <%=intTotalReview(counter)%>  ratings )</a>
  <%  
	 end if
	
	  %>
	
	  
	  
	  </td>
	  
        
     
      
    </tr>
<tr>

<td>
	<!--display NK spotlight images here
	-----------------------------------------------------------------------------------------------------------
	-->
	<div class="spotlight-group">
	 <table class="tableShowImages">

<%



dim mainimage
item=trim(item)
 folder = "/nk/" &item& "/thumbnails"

 
	  	  ' if item exist in drop ship list AND image folder exist
If objFSO.FolderExists (Server.MapPath(folder))then


 
      'Product and images are found
      'Display detail 
	  
    	 Set objFolder = objFSO.GetFolder(Server.MapPath(folder))
		if objFSO.FolderExists (Server.MapPath("/nk")) then
		    Set objFolder2 = objFSO.GetFolder(Server.MapPath("/nk"))
		end if 
	
		
dim i, j, column, emptyColumn
i=0
column=6

if not isnull(Request.Cookies("screenSize")) and len(trim(Request.Cookies("screenSize")))>0 then

	if (cint((Request.Cookies("screenSize"))) <600) then
			column=2
	elseif (cint((Request.Cookies("screenSize"))) <1000) then
			column=4
	end if

	
end if


For Each objFile in objFolder.files
if ucase(right(objfile.name, 3))="JPG" then
	largeImage=left(objFile.name, len(objFile.name)-7) + replace(right(objFile.name, 7), "S", "L")
	if i=0 then
		mainimage=largeImage
	end if
		
	
'RESPONSE.WRITE (i & ", " & column & "<br>")
	
	if i mod column = 0 then
	
	
%>
 <tr>
 <% end if%>

<td style="vertical-align:middle;">   

<p align="center">

<% if ucase(right(largeImage, 3))="JPG" then %>
    
	<a  class="spotlight" href="/nk/<%=item%>/images/<%=largeImage%>"  title= "<%=Ucase(item) & "  " & strProductName(counter)%>" >   
<img   class="my_img" width="100"  src="/nk/<%=item%>/thumbnails/<%=objfile.name%>" border="0" ></a>


<%	'response.write (i & " : " & "t:" & objfile.name & " L : " & largeImage  )
end if
%>




</td>

<%
i=i+1
if i mod column =0 then
	%>
	</tr>

	
	<%
end if


end if
Next

if i mod column<>0 then

    
    emptyColumn= column- (i mod column)
    for j = 1 to emptyColumn%>
    <td class="tdImgBoxEmpty" >
    </td>
    <%
    next%>
    </tr>
	
	
    <%
end if
END IF

' IF FOLDER OF ITEM EXIST

	
%>

</table>
</div>

<!--END display NK images here
	-----------------------------------------------------------------------------------------------------------
	-->
	

</td>


</tr>


   
    
    <tr>
    <td> 
	
	

	
	
	


 
    <%if len(mainImage)=0 then %>
		<img src ="../largeimages/<%=strLargeImage(counter)%>"  style = "float: right; margin: 10px; max-width: 75%;" />
	<% else%>
   
   <img src ="/nk/<%=item%>/images/<%=mainImage%>" style = "float: right; margin: 10px; max-width: 70%;" />
   <%end if%>
   
   
	
    <b>Specification:</b> <%=strSpecification(counter)%>
	
    <br />   <br />
   
  
				    
     
                        <%
                        '******************************************************************************
                        ' always display price for consumers
                       if 1>2 then
                       
                       %>
					    <b>Price: </b>
                       <br />
                       $ <%=  strRegprice(counter)%>
                       
                       <% if cdbl(strPrice(counter)) < cdbl(strRegprice(counter)) then %>                                              
                     <font class="salePriceText">  Sale $<%=  strPrice(counter)%> <%end if %> </font> 
					 <form method="post" action="cartRetail.asp">
    
	<input type="hidden" name="addproduct_ID" value="<%=intProductid%>">
	<input type="hidden" name="additemtocart" value="add">
    <div align="center"><input class="orderButton"  type="submit" name="addtocart" border="0" value="Order" align="center"></div>
   	</form>
		
							
                       
                       <%else '	for consumers				   
					   %>
					    <b>Price: </b> $
					   <%=strMSRP(counter)%> 
					    <% if cdbl(strMSRP(counter)) > cdbl(strRetailSalePrice(counter)) then %>                                              
                     <font class="salePriceText">     NOW:  $ <%=  strRetailSalePrice(counter)%> 
					 <% end if 
						response.write ("    ")
					 %>
					 </font> 
					 <% if quantityonhand(counter) <= 0 then
							response.write (" ( <font class=""salePriceText"">  This item is sold out at the moment</font> )")
								
					 elseif quantityonhand(counter)>0 and quantityonhand(counter)<20 then
							response.write ("( " & quantityonhand(counter) & " in stock )" )
					else 
						response.write ("(20+ in stock)" )
					end if
					%>
					
						
						
					<br><br>
					 
					 
					<form method="post" action="cartRetail.asp">
    
	<input type="hidden" name="addproduct_ID" value="<%=intProductid%>">
	<input type="hidden" name="additemtocart" value="add">
    <div><input type="submit" class="orderButton" name="addtocart" border="0"  value="Order" align="center"></div>
   	</form>
		
 
					   
                       
					   
                       <% 
					   'end consumer
										   
					   end if
                       ' end display price
                        '******************************************************************************
                       %>
                       <br>
					   
                       
                       
   
    
    
    
   <br />   
 
         <b>Description: </b> </br>
		 <%=strLongDes(counter)%>
		 </br></br>
		 
		 <% if len(strHistory(counter))>0 then%>
			<b>History: </b> </br>
			<%=strHistory(counter)%>
		<%end if%>
		
		 
	
         
		 
		 
    <form method="post" action="cartRetail.asp" align="center" name="cartForm">
    
	<input type="hidden" name="addproduct_ID" value="<%=intProductid%>">
	<input type="hidden" name="additemtocart" value="add">
	

	
	
  		</form>
         



	  
	  
<%	
		if reviewcount <=0 then

%>


<br>

<form align="left">
<input  type="hidden" value="Be the first one to rate this item now" name="button3" onClick="javascript:document.location.href='productsdetailsRetail.asp?review=1&productId=<%=intProductId%>'">

	<input  type="button" value="Be the first one to rate this item now" name="button3" onClick="javascript:document.location.href='productsdetailsRetail.asp?review=1&productId=<%=intProductId%>'">
</form>

<%
else

%>

<form align="left">
	<input  type="button" value="Rate this item now" name="button3" onClick="javascript:document.location.href='productsdetailsRetail.asp?review=1&productId=<%=intProductId%>'">
</form>

<%end if	

	  	 
		 
		 if intReview=1 and insertReviewSuccess = 0 and  counter=0 then
	%>
	
	 <script language="JavaScript1.2">
	
		var button1=document.getElementsByName("button3");
		button1[0].style.visibility="hidden";
	
		button1[1].style.visibility="hidden";
				
		</script>
	


	
	
			
			
		<table width="80%">
		
	
    	<form name="ReviewForm" method="post" action="productsdetailsRetail.asp?review=1&productId=<%=intProductId%>">			
		
		
	<tr>
	<th colspan="2">
    Please let us know what you think about this item:</th>
	
	<tr>
	
		<td width="20%"  >Your Name</td><td width="60%" > <input type="text" name="FullName" > <font style="color:#FF0000" size = "4"> *</font></td></tr>
	
	<tr>		<td>Your Email</td><td> <input name="Email" type = "text"> <font style="color:#FF0000" size = "4"> *</font> </td></tr>
	
		
	<tr><td> One Line Title</td><td> <input name="Title" type = "text" > <font style="color:#FF0000" size = "4"> *</font> </td></tr>
	
	<tr><td>Star Rating</td><td> 
	<select name="Rating" id="Rating">
  <option value="5"><font style="color:#FFA500" size = "4"> 5 stars</font></option>
  <option value="4"><font style="color:#FFA500" size = "4"> 4 stars</font></option>
  <option value="3"><font style="color:#FFA500" size = "4"> 3 stars</font></option>
  <option value="2"><font style="color:#FFA500" size = "4"> 2 stars</font></option>
  <option value="1"><font style="color:#FFA500" size = "4"> 1 star</font></option>
  
	</select></td></tr>
	

	
	
	<tr><td>
	You purchased this item from </td><td>
	<select name="Source" id="Source">
  <option value="OMHUSA">Old Modern Handicrafts</option>
  <option value="Amazon">Amazon</option>
  <option value="Captain Jim">Captain Jim</option>
  
  <option value="Wayfair">Wayfair</option>
  <option value="SpyglassShop3">Spyglass Shop</option>
  <option value="123 Stores">123 Stores</option>
  
  <option value="Others">Others</option>
  
	</select>
	</td></tr>
  
	
	
	
	<tr><td>
	Comment </td>
	<td> <textarea name="Comment" 
                    style="height: 100px; width: 300px"></textarea> (No hyperlink please)
					</td></tr>
					
					
					
		
	
	
  <tr><td>  </td>
			  <td>
			  
			  <input class="input10" type="submit" value="Submit Review" onClick="return validateData()" name="button10">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;			<input  type="button" value="Cancel" name="button12" onClick="javascript:document.location.href='productsdetailsRetail.asp?productId=<%=intProductId%>'"></td>
							
							
							</tr>
			  
  	</form>

</table>
	<% 
	elseif insertReviewSuccess =1 and counter = 0 then
	
	response.write  ("Thank you for your review!")
	%>
	
	 <script language="JavaScript1.2">
	
		var button1=document.getElementsByName("button3");
		button1[0].style.visibility="hidden";
	
		button1[1].style.visibility="hidden";
				
		</script>
	
	<%
	
	
	end if
%>
</td>
</tr>


      
      <tr>
      <td align="center">
    

<br>
  
    <form method="post" action="cartRetail.asp">
    
	<input type="hidden" name="addproduct_ID" value="<%=intProductid%>">
	<input type="hidden" name="additemtocart" value="add">
    <div align="center"><input class="orderButton" style="background-color:#F9E65C;" type="submit" name="addtocart" border="0" value="Order" align="center"></div>
   <br>
   <br>
   <br>
   <br>
   

   
	
	
	
  		</form>
   <br>
  
   
   <%if reviewcount> 0 then %>
   <a name="productReview"></a>
   <B>PRODUCT REVIEWS  </B><br>
   



   
   <% END IF%>
   <p align="left">

   <%
   
For i = 0 To reviewcount-1
if reviewcount> 0 then
  

	For j=1 To rating(i)

%>
	 <font style="color:#FFA500" size = "4"> *</font>
<%
  Next
  
  Response.write (" <b> " & title(i) & " </b> <br>")
  Response.write ("Reviewed by: " & Reviewby(i) & " , Date: " & reviewdate(i) & " , Source: " & Reviewsource(i) & "</br>")
  
				
  %>
  </br>
  <%
	response.write (" Comment: "  & ReviewComment(i) & "<br><br>")
	  
	  else
end if
	  
Next

%>

	  
<%	
if intReview<>1 then
		if reviewcount <=0 then

%>


<br>

<form align="left">
<input  type="hidden" value="Be the first one to rate this item now" name="button3" onClick="javascript:document.location.href='productsdetailsRetail.asp?review=1&productId=<%=intProductId%>'">

	<input  type="button" value="Be the first one to rate this item now" name="button3" onClick="javascript:document.location.href='productsdetailsRetail.asp?review=1&productId=<%=intProductId%>'">
</form>

<%
else

%>

<form align="left">
	<input  type="button" value="Rate this item now" name="button3" onClick="javascript:document.location.href='productsdetailsRetail.asp?review=1&productId=<%=intProductId%>'">
</form>

<%end if
end if	%>

 
  </p>   
   <br>
   <br><br>
   
<!-- ShareThis BEGIN --><div class="sharethis-inline-share-buttons"></div><!-- ShareThis END -->
   

  
 
 	

</td>
      </tr>
      
</table>

	 
<!--end productDetail table  --> 
	


  <%
   if counter=0 and isNUll(rel_prodID) = false then
  %>
<br />
  <b>Related products: </b>
  <%  
   end if
   
  %>
  
  
  
  
  <%
end if
rstproduct.close
rstReview.close 
counter=counter+1

loop until counter =2 or isNull(rel_prodID)

%>
  
  
  
    
    
    
    </td>
     
<!--end "productSearchlist"> td--> 
  
  </tr>
</table><!--end maintable--> 

<!--#Include file="footerRetail.asp"  --> 

</td><!--end mainCenter--> 




<td class = "mainright" >    </td>
</tr>



</table>
 
 

 
 
 
  
  
  
  
  
  
  </body>
  
  
  
  
  
  
  
  
  
  
  
  
  
  
