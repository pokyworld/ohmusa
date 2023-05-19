
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

Dim strSQLCateCombo, cnn1, strSQLCmd,strSQL

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
 
 <%
Dim objFSO, objFile, objfile2, objFolder, objfolder2, subfolder
dim largeImage

dim folder
dim item, itemID


'check if login
'if len(Session("wholesaler")) > 1 then


Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

item=request.querystring("size")
if strcomp(item, "l")=0 then
	item="CatalogL"
elseif strcomp(item, "s")=0 then
	item="CatalogS"
else
	item="Catalog2022"
end if


folder = "/" & item& "/thumbnails"
dim serverpath
serverpath = server.mappath("/")
Set objFolder = objFSO.GetFolder(serverpath & folder)



%>

                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
 <script type="text/javascript" src="../include/lightbox204/js/prototype.js"></script>
<script type="text/javascript" src="../include/lightbox204/js/scriptaculous.js?load=effects,builder"></script>
<script type="text/javascript" src="../include/lightbox204/js/lightbox.js"></script>
<link rel="stylesheet" href="../include/lightbox204/css/lightbox.css" type="text/css" media="screen" />
<script src="../include/spotlight_master/dist/spotlight.bundle.js"></script>

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
                   CATALOG</th>
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
                   
                   
                    <tr>
                    <td>
                    </td>
      
                  
                    </tr>
                 
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
  <th class="thfeatured" colspan="6">
CATALOG
  </th>
  </tr>
  
     
     
 

<%
dim i, column
i=0
column=6


For Each objFile in objFolder.files
	
'	largeImage=left(objFile.name, len(objFile.name)-7) + replace(right(objFile.name, 7), "S", "L")
largeImage=objFile.name
	
	if i mod column = 0 then
	
	
%>
 <tr>
 <% end if%>

<td  >

<% if ucase(right(largeImage, 3))="JPG" then %>
    
	<a class="spotlight"  href="../<%=item%>/images/<%=largeImage%>" rel="lightbox[1]" title= "<%=largeImage%>" >   
<img class="my_img" src="../<%=item%>/thumbnails/<%=objfile.name%>" border="0" ></a>
<%end if%>

</td>

<%
i=i+1
if i mod column =0 then

	%>
	
	
	</tr>
	<%
end if


Next

%>

<tr><td colspan="6">
<a href="omh2022catalog.pdf" class="cssTextCenter"><u>Click here to download</u></a> the full catalog PDF file (HiRes) and <a href="Contactus_r.asp"> <u>contact us</u> </a> to request a hard copy

     </td>  </tr>
     
  
	
						


  

  
  
  

  </table>
					
      
		
		 
      
      		
</td>

<!--end td class pagecontent -->



</tr>
	 
      
</table>
      
        <!--end content contactus -->
    
        
      
      
      
      


<!--#Include file="footerretail.asp"  --> 
 
</td>

 <!--end mainCenter -->



<td class = "mainright" >    </td>
</tr>
</table>
 
 
 

  
  
  
  
  
  
  </body>
  
  
  
  
  </html>
  