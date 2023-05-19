
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
<link rel="shortcut icon" type="image/x-icon" href="http://www.omhvn.com/favicon.ico" />


  </head>
  <body>
 
 
 
 
 
 <table class="fixedTable" >



<tr>
	<td class= "mainleft" >  </td>
	<td class = "maincenter" >   
	
<!--#Include file="headerretail.asp"  -->
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
      <th class="thfeatured"  >
      ABOUT US
      </th>
	  
	<tr>
										<td  >
										 <img border="0" src="../images/t034l01.jpg" width="290" style = "float: right; margin: 10px;">
										
                                        <p></p>
										<br>

	  
                                   		 <p>In such a fast-paced world, classic beauty seems gone too soon but the memories 
                                             live on. At Old Modern Handicrafts, we work hard to bring those memories back to 
                                             life. <b>Old Modern Handicrafts</b> (OMH) uses the <b>old</b> way of building ship 
                                             models while integrating <b>modern</b> technology. The traditional plank on 
                                             frame or plank on bulk head method is done by hand and followed by our high-tech 
                                             tools such as laser machines to create beautiful, accurate details that are 
                                             scaled down in precise proportions from the original model. The <b>old</b> 
                                             and <b>modern</b> technology work seamlessly together to create a product that 
                                             will become the center of attention for any home or office.</p>									  
								
								<p >
											 Besides scale models, OMH also makes wooden canoes and kayaks that looks beautiful and also performs well on water. These boats are strip built from Canadian western red cedar and encapsulated in fiber glass. This combination is strong yet completely transparent so that the beauty of the wood will be cherished and protected for generations. Here are 
                                            some products highlights that lead 
                                            us to success.                                           </p>
										
										<ul type="disc" >
                                            <li > Extensive research through original plans and 
                                            pictures make our models authentic.</li>
                                            <li > 100% hand built from scratch using “plank on 
                                            frame” or &quot;plank on bulkhead&quot; construction method</li>
                                            <li > Each model was built by skillful master 
                                            craftsmen with many intricate details.</li>
                                            <li  >                                            Hundreds of hours required to finish a model 
                                            at museum-quality level such as HMS 
                                            Victory or USS Constitution.</li>
                                            <li  >                                            Made of finest wood like rosewood, western red cedar, mahogany, 
                                            teak and other exotic wood.</li>
                                            <li  >                                            Chrome,  brass or cast metal fittings and ornaments 
                                            constitute the excellence of our 
                                            models</li>
                                            <li  >                                            Each model goes through a demanding quality 
                                            control process before leaving the 
                                            workshop</li>
                                          </ul>
                                             
                                             </td>
									</tr>
														
									<tr>
										<td  >
										    
											
										  
								            <p > We also custom build models for our 
                                            customers. We need as many 
                                            pictures and if possible the 
                                            original plan for the best results.</p>       

<p>
For a detailed timeline of our company, please click <a href="omhtimelinedesign.pdf" target="blank" > here </a>. 
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
<br>
<br>
<br>
<br>

</p>
											
                                            
                                            
                                          </td>
									</tr>
									
								</table>
								 
                            <!--end table98 --> 
						 
      </td>
      </tr>
     </table> <!--end maintable --> 

      
<!--#Include file="footerRetail.asp"  --> 

</td> 

<!--end  mainCenter-->



<td class = "mainright" >    </td>
</tr>



</table>
 

 
 
 
  
  
  
  
  
  
  </body>
  
  
  
  
  
  
  
  
  
  
  
  
  
  
