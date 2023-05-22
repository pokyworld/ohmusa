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
            cart.elements[i].value = 0;
            continue;
          }
          if (isNaN(cart.elements[i].value.trim()) and cart.elements[i].name <> "promoCode") {
            alert("Quantity  is not valid ");
            cart.elements[i].focus();
            return false;
          }
        }
      }
      return true;
    }
  </script>
  <meta name="viewport" content="width=device-width, initial-scale=0.75">
</head>
<body>
  <table class="fixedTable">
    <tr>
      <td class="mainleft"> </td>
      <td class="maincenter">
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
            (function (i, s, o, g, r, a, m) {
              i['GoogleAnalyticsObject'] = r; i[r] = i[r] || function () {
                (i[r].q = i[r].q || []).push(arguments)
              }, i[r].l = 1 * new Date(); a = s.createElement(o),
                m = s.getElementsByTagName(o)[0]; a.async = 1; a.src = g; m.parentNode.insertBefore(a, m)
            })(window, document, 'script', 'https://www.google-analytics.com/analytics.js', 'ga');
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
          <hr align="center" width="98%">
          <table class="table98noborder">
            <tr>
              <td>
                <a href="productsRetail.asp">Home</a>
              </td>
              <td>
                <a href="productSearchRetail.asp?pCategoryID=-4" target="_self">Best Sellers</a>
              </td>
              <td><a href="cartRetail.asp"> <img width="50%" src="../img/icons/shoppingcart.jpg"></a> </td>
              <td><a href="loginRetail.asp">Login</a> </td>
            </tr>
          </table>
          <hr align="center" width="98%">
          <!-- begin slider 2-->
          <table class="table98slider">
            <tr>
              <td>
                <div id="amazingslider-1" style="display:block;position:relative;margin:  0px 0px 0px; ">
                  <ul class="amazingslider-slides" style="display:none;">
                    <li><a href="ProductSearchRetail.asp?pCategoryID=1" target="_self"><img
                          src="../amazingslider/jessica/1.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=1" target="_self"><img
                          src="../amazingslider/jessica/2.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=10" target="_self"><img
                          src="../amazingslider/jessica/3.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=41" target="_self"><img
                          src="../amazingslider/jessica/4.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=27" target="_self"><img
                          src="../amazingslider/jessica/5.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=40" target="_self"><img
                          src="../amazingslider/jessica/6.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=42" target="_self"><img
                          src="../amazingslider/jessica/7.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=12" target="_self"><img
                          src="../amazingslider/jessica/8.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=38" target="_self"><img
                          src="../amazingslider/jessica/9.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=22" target="_self"><img
                          src="../amazingslider/jessica/10.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=24" target="_self"><img
                          src="../amazingslider/jessica/11.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=31" target="_self"><img
                          src="../amazingslider/jessica/12.jpg" /></a></li>
                    <li><a href="ProductSearchRetail.asp?pCategoryID=7" target="_self"><img
                          src="../amazingslider/jessica/13.jpg" /></a></li>
                  </ul>
                  <div class="amazingslider-engine" style="display:none;"><a href="http://amazingslider.com"
                      title="Responsive jQuery Image Slideshow">Responsive jQuery Image Slideshow</a></div>
                </div>
              </td>
            </tr>
          </table>
          <!-- End of slider 2-->
          <br>
        </body>
        </html>
        <table class="searchTable">
          <tr>
            <td class="cssTextCENTER" height="28" width="100%">
              <form action="productsearchRetail.asp" method="POST" name="SearchForm">
                <select size="1" name="pCategoryID">
                  <option selected value="0">All categories</option>
                  <option value="0">- - - - - - - - -</option>
                  <option value="47">Anne Home</option>
                  <option value="39">Architecture</option>
                  <option value="41">Automobile</option>
                  <option value="40">Aviation</option>
                  <option value="27">Battle Ship Model</option>
                  <option value="10">Boats / Canoes Model</option>
                  <option value="12">Cruise Ship Model</option>
                  <option value="36">Custom Made</option>
                  <option value="17">Display Cases</option>
                  <option value="21">Extra Large Ship Model ...</option>
                  <option value="20">Fishing Boat Model</option>
                  <option value="31">Full Size Boats</option>
                  <option value="37">Furniture</option>
                  <option value="42">Globes / Globe Bar</option>
                  <option value="6">Half-Hull</option>
                  <option value="46">Lighting</option>
                  <option value="50">Metal Home Accessories</option>
                  <option value="43">Nautical Compass</option>
                  <option value="38">Nautical Decor</option>
                  <option value="44">Nautical Telescopes/Bin...</option>
                  <option value="7">Other Novelties</option>
                  <option value="24">Sail Boat Model (1 mast.</option>
                  <option value="11">Sail Boat Model (2 mast...</option>
                  <option value="49">Scrap Metal Sculpture</option>
                  <option value="22">Speed Boat Model</option>
                  <option value="35">Surf Board</option>
                  <option value="1">Tall Ship Model - Admir...</option>
                  <option value="19">Tall Ship Model - Capta...</option>
                  <option value="48">The Perfect Combo</option>
                </select>
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
              <table class="table_outer_border">
                <tr>
                  <th class="thcategoryBGcolor">
                    CATEGORIES</th>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink"><a href="productsearchRetail.asp?pCategoryID=-1"
                        title="Ship Model - New Products "> <strong>New Products!!!</strong> </a></span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=47" title="Ship Model - Anne Home">Anne Home </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=39" title="Ship Model - Architecture">Architecture
                      </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=41" title="Ship Model - Automobile">Automobile </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=40" title="Ship Model - Aviation">Aviation </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=27" title="Ship Model - Battle Ship Model">Battle
                        Ship Model </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=10" title="Ship Model - Boats / Canoes Model">Boats /
                        Canoes Model </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=12" title="Ship Model - Cruise Ship Model">Cruise
                        Ship Model </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=36" title="Ship Model - Custom Made">Custom Made </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=17" title="Ship Model - Display Cases">Display Cases
                      </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=21"
                        title="Ship Model - Extra Large Ship Model - Fleet Admiral *****">Extra Large Ship Model - Fleet
                        Admiral ***** </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=20" title="Ship Model - Fishing Boat Model">Fishing
                        Boat Model </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=31" title="Ship Model - Full Size Boats ">Full Size
                        Boats </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=37" title="Ship Model - Furniture">Furniture </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=42" title="Ship Model - Globes / Globe Bar">Globes /
                        Globe Bar </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=6" title="Ship Model - Half-Hull">Half-Hull </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=46" title="Ship Model - Lighting">Lighting </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=50" title="Ship Model - Metal Home Accessories">Metal
                        Home Accessories </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=43" title="Ship Model - Nautical Compass">Nautical
                        Compass </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=38" title="Ship Model - Nautical Decor">Nautical
                        Decor </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=44"
                        title="Ship Model - Nautical Telescopes/Binocular">Nautical Telescopes/Binocular </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=7" title="Ship Model - Other Novelties">Other
                        Novelties </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=24"
                        title="Ship Model - Sail Boat Model (1 mast)">Sail Boat Model (1 mast) </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=11"
                        title="Ship Model - Sail Boat Model (2 masts +)">Sail Boat Model (2 masts +) </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=49" title="Ship Model - Scrap Metal Sculpture">Scrap
                        Metal Sculpture </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=22" title="Ship Model - Speed Boat Model">Speed Boat
                        Model </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=35" title="Ship Model - Surf Board">Surf Board </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=1"
                        title="Ship Model - Tall Ship Model - Admiral Line ****">Tall Ship Model - Admiral Line ****
                      </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=19"
                        title="Ship Model - Tall Ship Model - Captain Line">Tall Ship Model - Captain Line </a>
                    </span>
                  </td>
                </tr>
                <tr>
                  <td width="100%" align="left">&nbsp;</td>
                </tr>
                <tr>
                  <td align="left" class="tdmargin10">
                    <span class="cssLink">
                      <a href="productsearchRetail.asp?pCategoryID=48" title="Ship Model - The Perfect Combo">The
                        Perfect Combo </a>
                    </span>
                  </td>
                </tr>
              </table>
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
                        <img border="0" src="../images/SALE.jpg"><br />
                      </a>
                    </p>
                    <p align="center">
                      <a href="catalog_r.asp" title="catalog">
                        <img border="0" src="../images/catalog.JPG"><br />
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
                  <th class="thfeatured" colspan="3">
                    SHOPPING CART
                  </th>
                </tr>
                <tr>
                  <td align="left" valign="top">&nbsp;
                  </td>
                  <td>
                    <div align="center">
                      <center>
                        <p align="center">&nbsp;
                        <br /><br />
                        
                        <form method="POST" name="cart" action="cartRetail.asp" id="theForm">
                          <!-- cart line items -->
                          <table class="table98border_aboutus">
                            <!--#include file="./_CartItem.asp"-->
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
                              $ 172.92&nbsp;</td>
                          </tr>
                          <tr>
                            <td>
                              Promo Code <input type="text" name="promoCode" size="20"> (Click update to apply)
                            </td>
                            <td align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                            </td>
                            <td align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                            </td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                              Domestic Ground Shipping Cost&nbsp;&nbsp;&nbsp; </td>
                            <td align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                              &nbsp;0&nbsp;
                              <font color="#FF0000">*</font>
                            </td>
                          </tr>
                          <tr>
                            <td></td>
                            <td align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                              Grand Total&nbsp;&nbsp;&nbsp; </td>
                            <td align="right" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
                              $ 172.92&nbsp;</td>
                          </tr>
                        </table>
                      </center>
                    </div>
                    <p align="center">
                      <input type="hidden" value="1" name="pCartSize">
                      <input type="hidden" name="pUpdate" value="update">
                      <input type="submit" name="update" border="0" src="../images/update.gif" value=" Update "
                        onClick="return validateTextBoxes();">
                      <input type="button" value=" Continue Shopping " onclick="location.href='productsretail.asp'" />
                      <!--<input type="submit" value="Update" name="pUpdate">-->
                    <p align="center"></p>
                    <p align="center">
                      <input type="button" value=" Checkout " onclick="location.href='checkoutRetail.asp'" />
                    </p>
                    <p align="left">
                      <font class="attentionText">Note: <br /></font>
                      <font>
                        * Free domestic ground shipping except orders to Hawaii and Alaska<br>
                        * For items with limited quantity, please contact us to verify availability <br>
                        * Soldout items will be automatically added to back order
                        unless you notify us otherwise.
                      </font>
                    </p>
                    <p align="center">
                      <a href="https://www.positivessl.com" target="_blank"
                        style="font-family: arial; font-size: 10px; color: #212121; text-decoration: none;"><img
                          src="https://www.positivessl.com/images-new/PositiveSSL_tl_white.png" alt="SSL Certificate"
                          title="SSL Certificate" border="0" /></a>
                    </p>
                    </form>
                    <!-- </font> -->
                    <!-- </font> -->
                  </td>
                  <td align="left" valign="top">&nbsp;</td>
                </tr>
              </table><!--end table98 -->
            </td>
          </tr>
        </table>
        <!--end mainTable-->
        <html>
        <head>
          <meta http-equiv="Content-Language" content="en-us">
          <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
          <meta name="GENERATOR" content="Microsoft FrontPage 4.0">
          <meta name="ProgId" content="FrontPage.Editor.Document">
          <title>Footer</title>
          <link rel="stylesheet" type="text/css" href="product_stylesheet.css">
          <!-- Add icon library -->
          <link rel="stylesheet"
            href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
        </head>
        <body>
          <hr align="center" width="98%">
          <table class="table98noborder">
            <tr>
              <td> <a href="Aboutus_r.asp">About Us </a> </td>
              <td> <a href="OMHTIMELINEDESIGN_R.PDF" target="blank">OMH Timeline </a> </td>
              <td> <a href="catalog_r.asp">Catalog</a> </td>
              <td> <a href="faq_r.asp">Terms & Conditions</a> </td>
              <td> <a href="stepstobuild_r.asp">How to Build A Ship Model</a> </td>
              <td> <a href="Contactus_r.asp">Contact Us </a> </td>
            </tr>
          </table>
          <hr align="center" width="98%">
          <table class="table98noborder">
            <tr>
              <td>
                Get Connected on:
                <!-- Add font awesome icons -->
                <a href="http://www.facebook.com/omhusa" target='blank'> <img src="../img/icons/facebook.jpg"> </a>
                <a href="http://www.linkedin.com/company/old-modern-handicrafts-inc" target='blank'> <img
                    src="../img/icons/linkedin.jpg"> </a>
                <a href="http://www.twitter.com/omhinc" target='blank'> <img src="../img/icons/twitter.jpg"> </a>
                <a href="https://www.instagram.com/old.modern.handicrafts/" target='blank'> <img
                    src="../img/icons/instagram.jpg"> </a>
                <a href="http://www.pinterest.com/omhsalesinc" target='blank'> <img
                    src="../img/icons/pininterest.jpg"></a>
                <a href="https://www.youtube.com/channel/UCdp5Y_rJ24oFKSJ2cCxQCrw" target='blank'> <img
                    src="../img/icons/youtube.jpg"></a>
              </td>
              <td>
                To become a dealer, <b>please call (888) 900-1805</b></td>
              <td>
                Copyright© by Old Modern Handicrafts ™ </td>
              <td>
                <a href="#"
                  onclick="window.open('https://www.sitelock.com/verify.php?site=omhusa.com','SiteLock','width=600,height=600,left=160,top=170');"><img
                    class="img-responsive" alt="SiteLock" title="SiteLock"
                    src="//shield.sitelock.com/shield/omhusa.com" /></a>
              </td>
            </tr>
          </table>
        </body>
        </html>
      </td>
      <!--end mainCenter-->
      <td class="mainright">&nbsp;</td>
    </tr>
  </table>
</body>
</html>