<%
'response.Expires=0
'response.CacheControl= "no-cache"
'response.AddHeader "Pragma", "no-cache"


%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<meta name="viewport" content="width=device-width, initial-scale=0.75">

<title>Header</title>
<link rel="stylesheet" type="text/css" href="product_stylesheet.css">

<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','https://www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-1997370-3', 'auto');
  ga('send', 'pageview');

</script>


  <!-- Insert to your webpage before the </head> -->
    <script src="../amazingslider/sliderengine/jquery.js"></script>
    <script src="../amazingslider/sliderengine/amazingslider.js"></script>
    <script src="../amazingslider/sliderengine/initslider-1.js"></script>
	
<script type="text/javascript">
	document.cookie = "screenSize=" + screen.width;
</script>

</head>

<body>


   
	  <hr  align="center"  width="98%"  >
      <table class="table98noborder" >
	  
     
        <tr>
                <td >
                <a href="productsRetail.asp">Home</a>                </td>
                
              
				<td> 
				<a href="productSearchRetail.asp?pCategoryID=-4" target="_self">Best Sellers</a>
				</td>
				
                <td><a href="cartRetail.asp"> <img width="50%" src="../img/icons/shoppingcart.jpg"></a> </td>
                
             			
				<%
				
		          if len(  session("login")) > 0 then%>
		          
		          <td><a href="editProfileRetail.asp"> My Profile</a> </td>
		          
                  <td><a href="logoutRetail.asp">Logout </a> </td>
                    
    <% else%>
  
    <td><a href="loginRetail.asp">Login</a> </td>
    
    	          
    <%
	end if

%>

				
				
				
				
				
        </tr>
      </table>  
	  
	  
	  
   
   
  <hr  align="center" width="98%"  >
  
   <%
   
  if Hour(Now()) mod 5 =0 then
  'change slider every hour
  
  'if 1 =2 then
  ' alternate every hour
 ' eric 09/15/2022
  '
   %>
   
   
   <!-- begin slider thanskgiving-->
  <table class="table98slider" >
  <tr>
  <td>



    <div id="amazingslider-1" style="display:block;position:relative;margin:  0px 0px 0px; ">
        <ul class="amazingslider-slides" style="display:none;">
            <li><a href="aboutus.asp" target="_self"><img src="../amazingslider/adrian/1.jpg" /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=-1"><img src="../amazingslider/adrian/2.jpg"  /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=1" target="_self"><img src="../amazingslider/adrian/3.jpg" /></a></li>
           
            <li><a href="ProductSearchRetail.asp?pCategoryID=31" target="_self">	<img src="../amazingslider/adrian/6.jpg" >	</a></li>
			 <li><a href="registration.asp" target="_self"><img src="../amazingslider/adrian/4.jpg"  /></a></li>
           	<li><a href="Contactus.asp" target="_self"><img src="../amazingslider/adrian/5.jpg" /></a></li>
			
		

			
        </ul>
        <div class="amazingslider-engine" style="display:none;"><a href="http://amazingslider.com" title="Responsive jQuery Image Slideshow">Responsive jQuery Image Slideshow</a></div>
    </div>
    

        </td>
     </tr>
  </table>
  <!-- End of slider 1-->
  
  
  <%
  elseif Hour(Now()) mod 3 =0 then
  'change slider every hour
  'response.write (hour(now))
  
  
'  elseif 1 =3 then
  ' alternate every hour
 ' eric 09/15/2022
  '
   %>
   
   
   <!-- begin slider thanskgiving-->
  <table class="table98slider" >
  <tr>
  <td>



    <div id="amazingslider-1" style="display:block;position:relative;margin:  0px 0px 0px; ">
        <ul class="amazingslider-slides" style="display:none;">
            <li><a href="aboutus.asp" target="_self"><img src="../amazingslider/images/1.jpg" /></a></li>
            <li><a href="Contactus.asp"><img src="../amazingslider/images/2.jpg"  /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=1" target="_self"><img src="../amazingslider/images/3.jpg" /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=-1" target="_self"><img src="../amazingslider/images/4.jpg"  /></a></li>
            <li><a href="https://facebook.com/omhusa" target="_blank"><img src="../amazingslider/images/5.jpg" /></a></li>
            <li><a href="dealerlocator.asp?area=6" target="_self"><img src="../amazingslider/images/6.jpg" /></a></li>
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
            <li><a href="ProductSearchRetail.asp?pCategoryID=1" target="_self"><img src="../amazingslider/jessica/1.jpg"  /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=1" target="_self"><img src="../amazingslider/jessica/2.jpg"  /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=10" target="_self"><img src="../amazingslider/jessica/3.jpg"  /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=41" target="_self"><img src="../amazingslider/jessica/4.jpg"  /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=27" target="_self"><img src="../amazingslider/jessica/5.jpg"  /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=40" target="_self"><img src="../amazingslider/jessica/6.jpg"/></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=42" target="_self"><img src="../amazingslider/jessica/7.jpg" /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=12" target="_self"><img src="../amazingslider/jessica/8.jpg" /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=38" target="_self"><img src="../amazingslider/jessica/9.jpg" /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=22" target="_self"><img src="../amazingslider/jessica/10.jpg"  /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=24" target="_self"><img src="../amazingslider/jessica/11.jpg" /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=31" target="_self"><img src="../amazingslider/jessica/12.jpg" /></a></li>
            <li><a href="ProductSearchRetail.asp?pCategoryID=7" target="_self"><img src="../amazingslider/jessica/13.jpg" /></a></li>
            
            
        </ul>
        <div class="amazingslider-engine" style="display:none;"><a href="http://amazingslider.com" title="Responsive jQuery Image Slideshow">Responsive jQuery Image Slideshow</a></div>
    </div>
    

        </td>
     </tr>
  </table>
  <!-- End of slider 2-->
  
  
  
   <%end if %>
   <br>
   
  
</body>

</html>