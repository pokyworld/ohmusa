<%@ Language=VBScript %>
<% Option Explicit %>

<%
'response.Expires=0
'response.CacheControl= "no-cache"
'response.AddHeader "Pragma", "no-cache"

%>


<!-- #include file="../include/asp_lib.inc.asp" -->
<!-- #include file="../include/sqlCheckInclude.asp" -->


<%
'--------------------------------------------------------------
'      Coded By: Tan Pham on 01/01/2001.
'       Purpose: Display all category and search product form.
'   Used Tables: Category.
'  Invoked From: index.asp.
'       Invokes: search_product.asp
'Included Files: headerRetail.asp, footerRetail.asp, animate.js, StyleSheet.css, asp_lib.inc.asp.
'--------------------------------------------------------------
'Updated By TrangTruong     Date28/03/00   Comments
'updated by Eric V 07/26/2019
'--------------------------------------------------------------


   
Dim strSQLCateCombo, cnn, rstCategory, strSQLCmd

'SQL statement for creating combo box. If name has more than 13 char then add ... as a tail.
strSQLCateCombo = "select Left(Category_Name, 23)+Left('...', Len(Category_Name) - Len(Left(Category_Name, 23))), Category_ID from Category where status <>'inactive' order by Category_Name asc "

'Create connection and query category data.
strSQLCmd = "select Category_ID, Category_Name from Category where status <>'inactive' order by upper(Category_Name) asc"
Set cnn = Server.CreateObject("ADODB.Connection")
cnn.ConnectionString = Application.Contents("dbConnStr")
cnn.Open
Set rstCategory = Server.CreateObject("ADODB.Recordset")
rstCategory.Open strSQLCmd, cnn, 3
%>


<html>
<head>

<title>Ship Model Wholesale - Vintage Nautical Gifts</title>

<meta name="keywords" content="model ship, ship model, model boat, wooden boat, tall ship model, wooden ship model, Handmade ship model, speed boat, wholesale">
<meta name="DESCRIPTION" content="Ship model manufacturer and wholesale, sail boat models, wooden boat, wooden canoe, kayak, tall ship model, model speed boat, drop ship direct">
<meta name="Destination" content="tallship, tall ship model, historic ship, wooden ship model, Handmade ship model, speed boat, tall ship, modern yacht, old style yacht, historic ship model, museum ship model, Queen Mary, Riva, HMS Victory, Sovereign of the Seas, San Felipe, Esmeralda, USS Constituion, USS Constellation, Wasa, Vasa, Soleil Royal, Friesland, Titanic, Shrimp boat, Normandie, Queen Mary 2, Zeven Provincien, Lady Washington, Mikasa, Australia, Batavia, Shamrock, Endeavour, Bounty, Handcraft ship model, handicraft, handicrafts, manufacturer, exporter, ship leading builder">

<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
<link rel="shortcut icon" type="image/x-icon" href="http://www.omhvn.com/favicon.ico" />


<script type='text/javascript' src='https://platform-api.sharethis.com/js/sharethis.js#property=6202fc68049246001a151155&product=inline-share-buttons' async='async'></script>
 


<script language="JavaScript1.2" src="../Jan2005/kill-mouse.js" type="text/javascript"></script>
    <style type="text/css">
     
  
       
     
  
    </style>
    
    
     
    <!-- Insert to your webpage before the </head> -->
    <script src="../amazingslider/sliderengine/jquery.js"></script>
    <script src="../amazingslider/sliderengine/amazingslider.js"></script>
    <script src="../amazingslider/sliderengine/initslider-1.js"></script>
   
    <!-- End of head section HTML codes -->
	
	<script type="text/javascript">
	document.cookie = "screenSize=" + screen.width;
</script>
<meta name="viewport" content="width=device-width, initial-scale=.75">

    
    </head>

<body  >
<table class="fixedTable" >



<tr>
		<td class = "mainleft"  >  </td>
	<td class = "maincenter"   >
	

<!--#include file="headerRetail.asp"-->
<table class="searchTable" >
 <tr>
                <td width="100%" align = "center" valign="top" colspan="3" height="25" class="cssTextCenter">
                  <form method="POST" name="SearchForm" action="productsearchRetail.asp">
                    
                    <%Call SQLCombo("pCategoryID", "1", "", strSQLCateCombo, "All categories", "- - - - - - - - -", "0", "0")%>
                    Name / SKU
                    
                    <input type="hidden" name="formSearch" value="yes">
                    <input type="text" name="pProductName" size="15">
                    <input type="submit" value="Search" name="pAction">
                  </form>
                </td>
				
				
				
              </tr>

  </table>
  
  <%
  'if Hour(Now()) mod 2 =0 then
  'change slider every hour
  
  if Hour(Now()) mod 2 >=0 then
  ' condition always wrong, so always use slider2
  'updated 7.26.2019 Eric
  
  
  '
   %>
<!-- begin slider 1-->
  <table class="table98slider" >
  <tr>
  <td>



    <div id="amazingslider-1" style="display:block;position:relative;margin:  0px 0px 0px; ">
        <ul class="amazingslider-slides" style="display:none;">
            <li><img src="../amazingslider/images/1.jpg" /></a></li>
            <li><a href="registration.asp" target="_self"><img src="../amazingslider/images/2.jpg"  /></a></li>
            <li><a href="aboutus.asp" target="_self"><img src="../amazingslider/images/3.jpg"  /></a></li>
            <li><a href="" target="_self"><img src="../amazingslider/images/4.jpg"  /></a></li>
            <li><a href="https://www.facebook.com/OmhUsa/" target="_blank"><img src="../amazingslider/images/5.jpg" /></a></li>
            <li><a href="" target="_self"><img src="../amazingslider/images/6.jpg" /></a></li>
        </ul>
        <div class="amazingslider-engine" style="display:none;"><a href="http://amazingslider.com" title="Responsive jQuery Image Slideshow">Responsive jQuery Image Slideshow</a></div>
    </div>
    

        </td>
     </tr>
  </table>
  <!-- End of slider 1-->
  
  <% else  %>
   
   <!-- begin slider 2-->
  <table class="table98slider" >
  <tr>
  <td>



    <div id="amazingslider-1" style="display:block;position:relative;margin:  0px 0px 0px; ">
        <ul class="amazingslider-slides" style="display:none;">
            <li><img src="../amazingslider/images2/1.jpg" /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=1" target="_self"><img src="../amazingslider/images2/2.jpg"  /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=10" target="_self"><img src="../amazingslider/images2/3.jpg"  /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=41" target="_self"><img src="../amazingslider/images2/4.jpg"  /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=27" target="_self"><img src="../amazingslider/images2/5.jpg"  /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=40" target="_self"><img src="../amazingslider/images2/6.jpg"/></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=42" target="_self"><img src="../amazingslider/images2/7.jpg" /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=12" target="_self"><img src="../amazingslider/images2/8.jpg" /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=38" target="_self"><img src="../amazingslider/images2/9.jpg" /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=22" target="_self"><img src="../amazingslider/images2/10.jpg"  /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=24" target="_self"><img src="../amazingslider/images2/11.jpg" /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=31" target="_self"><img src="../amazingslider/images2/12.jpg" /></a></li>
            <li><a href="productsearchRetail.asp?pCategoryID=7" target="_self"><img src="../amazingslider/images2/13.jpg" /></a></li>
            
            
        </ul>
        <div class="amazingslider-engine" style="display:none;"><a href="http://amazingslider.com" title="Responsive jQuery Image Slideshow">Responsive jQuery Image Slideshow</a></div>
    </div>
    

        </td>
     </tr>
  </table>
  <!-- End of slider 2-->
  
   <%end if %>
  
  <table class="mainTable" >
                
  <tr>
   <td class= "category">
                 
                  <%
                  If rstCategory.RecordCount > 0 Then
                  %>
                  <table 
                  class="table_outer_border" >                  
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
					 
					 
                      <span class="cssLink"><a href="productsearchRetail.asp?pCategoryID=<%=rstCategory("Category_ID")%>" title="Ship Model - <%=rstCategory("Category_Name")%>"><%=rstCategory("Category_Name")%> </a></span></td>
                    </tr>
                    <%
                      rstCategory.MoveNext
                    Wend
                    rstCategory.Close
                    cnn.Close
                    Set rstCategory = Nothing
                    Set cnn = Nothing
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
             
                 <td class="pageContent">
                
                
                
                
  
  <table class="table98border" > 
  
  <tr>
  <th class="thfeatured" colspan="3">
FEATURED CATEGORIES
  </th>
  </tr>
  

   
       
    <tr >
   
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=39">
      <img src ="../images/categories/architecture.jpg"  id="image1"   />
      
      <br />
      
      
       Architecture</a>
       
       </td>
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=41">
      <img  src ="../images/categories/automobile.jpg"  id="Img1"   /><br />
      
       Automobile<a />
       </td>
         <td height="10" class="tdImgBox " >
         <a href = "productsearchRetail.asp?pCategoryID=40">
      <img  src ="../images/categories/aviation.jpg"  id="Img2"   /><br />
      
       Aviation</a>
       </td>
         
    </tr>
    
    
       
    
       
    <tr >
   
      <td height="10" class="tdImgBox " >
      
      <a href = "productsearchRetail.asp?pCategoryID=27">
      <img src ="../images/categories/battleship.jpg"  id="Img3"   />
      
      <br />
      
      
       Battle Ship Model</a>
       
       </td>
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=10">
      <img  src ="../images/categories/boat_canoe.jpg"  id="Img4"   /><br />
      
       Boats / Canoes Model<a />
       </td>
         <td height="10" class="tdImgBox " >
         <a href = "productsearchRetail.asp?pCategoryID=12">
      <img  src ="../images/categories/cruiseship.jpg"  id="Img5"   /><br />
      
       Cruise Ship Model</a>
       </td>
         
    </tr>
    
       
   
   
       
    <tr >
   
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=17">
      <img src ="../images/categories/displaycase.jpg"  id="Img6"   />
      
      <br />
      
      
       Display Cases</a>
       
       </td>
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=31">
      <img  src ="../images/categories/canoe.jpg"  id="Img7"   /><br />
      
       Full size Canoes, Kayaks<a />
       </td>
         <td height="10" class="tdImgBox " >
         <a href = "productsearchRetail.asp?pCategoryID=37">
      <img  src ="../images/categories/furniture.jpg"  id="Img8"   /><br />
      
       Furnitures</a>
       </td>
         
    </tr>
    
    

       
    <tr >
   
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=42">
      <img src ="../images/categories/nauticalglobe.jpg"  id="Img9"   />
      
      <br />
      
      
       Globes / Globe Bar</a>
       
       </td>
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=38">
      <img  src ="../images/categories/nautical.jpg"  id="Img10"   /><br />
      
       Nautical Gifts<a />
       </td>
         <td height="10" class="tdImgBox " >
         <a href = "productsearchRetail.asp?pCategoryID=7">
      <img  src ="../images/categories/others.jpg"  id="Img11"   /><br />
      
       Other Novelties</a>
       </td>
         
    </tr>
    
    
    
         
          
       
       
    <tr >
   
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=22">
      <img src ="../images/categories/Speedboat.jpg"  id="Img12"   />
      
      <br />
      
      
      Speed Boat Model</a>
       
       </td>
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=19">
      <img  src ="../images/categories/tallshipinter.jpg"  id="Img13"   /><br />
      
       Tall Ship Model- Captain Line<a />
       </td>
         <td height="10" class="tdImgBox " >
         <a href = "productsearchRetail.asp?pCategoryID=1">
      <img  src ="../images/categories/tallship.jpg"  id="Img14"   /><br />
      
       Tall Ship Model - Admiral Line</a>
       </td>
         
    </tr>
         
         
         
         
       
    <tr >
   
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=24">
      <img src ="../images/categories/sloop.jpg"  id="Img15"   />
      
      <br />
      
      
       Sloop </a>
       
       </td>
      <td height="10" class="tdImgBox " >
      <a href = "productsearchRetail.asp?pCategoryID=11">
      <img  src ="../images/categories/schooner.jpg"  id="Img16"   /><br />
      
       Schooner <a />
       </td>
         <td height="10" class="tdImgBox " >
         <a href = "productsearchRetail.asp?pCategoryID=21">
      <img  src ="../images/categories/tallshipxl.jpg"  id="Img17"   /><br />
      
      XL Ship Model - Fleet Admiral</a>
       </td>
         
    </tr>
    
    
    
    
  </table>
  
  <br>
  <br>
  
  <!-- ShareThis BEGIN --><div class="sharethis-inline-share-buttons"></div><!-- ShareThis END -->
  
</td>

              </tr>
			  <tr>
  <td></td>
  
  </tr>
  

  
            </table>
    <bR>
	<BR>
	
  
 <!--#Include file="footerRetail.asp"  -->
</td>

<td class= "mainright">  </td>
</tr>



</table>
</body>
</html>
