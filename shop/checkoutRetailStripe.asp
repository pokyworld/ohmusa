
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
'Included Files: header.htm, footer.htm, animate.js, StyleSheet.css, asp_lib.inc.asp.
'--------------------------------------------------------------
'Updated By   Eric    Date 01/04/2011   Comments
'Display products details
'--------------------------------------------------------------
%>
<!-- #include file="../include/asp_lib.inc.asp" -->
<!-- #include file="../include/sqlCheckInclude.asp" -->

<%

Dim strSQLCateCombo, cnn1, strSQLCmd1

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
   
   
   
'************************************************************************************************************************   
'start check out ASP   




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
dim loggedIn
strLogin = session("login")
if len(strLogin) >0  then
	loggedIn=1
else 
	loggedIn=0
end if


'check if login
'if len(Session("consumer")) < 1  then
   Session("requestLoginURL") = "checkoutRetail.asp"
   'Response.Redirect "loginRetail.asp"
'end if
 

'Create connection to execute insert command
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  set rst = Server.CreateObject("adodb.RecordSet")
  set rstCount = Server.CreateObject("adodb.RecordSet")
  cnn.Open

strSQLCmd="select * from consumer where login = '"&strlogin&"'"  
 rst.open strSqlcmd, cnn,3
 

   
%>



<!-- TWO STEPS TO INSTALL CREDIT CARD VALIDATION:

  1.  Copy the coding into the HEAD of your HTML document
  2.  Add the last code into the BODY of your HTML document  -->

<!-- STEP ONE: Paste this code into the HEAD of your HTML document  -->
  

 
<html>
  <head>
    <title>Old-Modern Handicrafts - View Detail Product</title>
    <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
    <script language="JavaScript1.2" src="../include/javascript.js"></script>
    <script language="JavaScript" src="../include/tools.js"></script>
    <script language="JavaScript" src="checkoutRetail.js"></script>
    <script language="JavaScript" src="checkoutRetailStripe.js"></script>
    <meta name="viewport" content="width=device-width, initial-scale=0.75">
  </head>
  <body>
 <table class="fixedTable" >
  <tr>
    <td class= "mainleft" >  </td>
    <td class = "maincenter" >      
    <!--Include file="HeaderRetail.asp"-->
    
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
                        <a href="catalog_r.asp" title="catalog"> 
                        <img border="0" src="../images/catalog1.gif"><br />
					  </p>
					  
					</a>
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
          SECURED SSL CHECK OUT
     </th>
     </tr>
     
     
     
     <%'***************************************************************************************************%>
     <%'----------------check out  content start--------------%>
  	<%
	
	dim contact, customer, street, city, state, zip, phone, email, billing, logged_in
	
	if not rst.eof then
		contact=rst("contact")
		customer=rst("customer")
		street=rst("street")
		city=rst("city")
		state = rst("state")
		zip= rst("zip")
		phone= rst("phone")
		email= rst("email")
		billing= rst("billing")
		logged_in=1
		
	end if	
		
		
		
	
	%>
	
  	
  	
  <tr>
    <td width="20" align="left" valign="top">
	   
    </td>
    <td  align ="left">
         
       
		
		
		<%
		if logged_in = 1 then %>
		 <b>Shipping address</b>
          <a href="editProfileRetail.asp?callUrl=checkoutRetail.asp">EDIT</a>
          
           
            <table class="table98borderCheckout">
              <tr>
                <td width="45%">Name:<%=contact%> </td>
                <td width="55%">Company Name:<%=customer%></td>
              </tr>
              <tr>
                <td width="45%">Address:<%=street%> </td>
                <td width="55%">City: <%=city%></td>
              </tr>
              <tr>
                <td width="45%">State:<%=state%>&nbsp; </td>
                <td width="55%">Zip:<%=zip%></td>
              </tr>
              <tr>
                <td width="45%">Phone:<%=phone%> </td>
                <td width="55%">Email: <%=email%> </td>
              </tr>
            </table>
           
        
          
         
          <p>
          <b>Credit Card billing address</b>
          <a href="editProfileRetail.asp?callUrl=checkoutRetail.asp">EDIT</a>
   
          
               <table class="table98borderCheckout">
                <tr>
                  <td width="100%"><%=billing%>&nbsp;</td>
                </tr>
              </table>
          
          
    
		  </p>
        <%
			'not log in, check out as guest
			else 		
		%>
			<br>
			
			Please log in <a href="loginRetail.asp">here</a> or enter your shipping address below:
			
		<% end if%>
			
	
          
 <form method="post" action = "checkoutprocessRetail.asp" name="Payment">
         <p> 
          
          <b>Alternate Shipping Address</b> 
		  <input type="hidden" name="loggedIn" value=<%=loggedIn%> >
		  
        <table class="table98borderCheckout"  >
            <tr>
              <td colspan="2">Name
            <input type="text" name="name" size="20"><font color="red">*</font></td>
			
             
            </tr>
			<tr>
              <td colspan="2">
            Company Name<input type="text" name="companyName" size="32"></td>
            </tr>
			
            <tr>
              <td width="40%">Street
              <input type="text" name="address" size="22"><font color="red">*</font></td>
              <td width="60%">City <input type="text" name="city" size="20"><font color="red">*</font></td>
            </tr>
            <tr>
              <td width="40%">Street Line 2
              <input type="text" name="address2" size="20"></td>
              <td width="60%">State/Province
              <input type="text" name="state" size="20"><font color="red">*</font></td>
            </tr>
            <tr>
              <td colspan="2">Zip <input type="text" name="zip" size="9"></td>
			  
			  </tr>
			 <tr>
			   
              <td colspan="2"> 
              Country 
			   <select class="selectlarge" id="selCountry" name="country" size="1">
<option>AFGHANISTAN</option>
<option>ALBANIA</option>
<option>ALGERIA</option>
<option>AMERICAN SAMOA</option>
<option>ANDORRA</option>
<option>ANGOLA</option>
<option>ANGUILLA</option>
<option>ANTIGUA(ANTIGUA / BARBUDA)</option>
<option>ARGENTINA</option>
<option>ARMENIA</option>
<option>ARUBA</option>
<option>AUSTRALIA</option>
<option>AUSTRALIA(CHRISTMAS IS)</option>
<option>AUSTRALIA(COCOS KEELING IS)</option>
<option>AUSTRALIA(NORFOLK IS)</option>
<option>AUSTRIA</option>
<option>AZERBAIJAN</option>
<option>BAHAMAS</option>
<option>BAHRAIN</option>
<option>BANGLADESH</option>
<option>BARBADOS</option>
<option>BELARUS</option>
<option>BELGIUM</option>
<option>BELIZE</option>
<option>BENIN</option>
<option>BERMUDA</option>
<option>BHUTAN</option>
<option>BOLIVIA</option>
<option>BONAIRE(NETHERLANDS ANTILLES)</option>
<option>BOSNIA / HERZEGOVINA</option>
<option>BOTSWANA</option>
<option>BRAZIL</option>
<option>BRUNEI</option>
<option>BULGARIA</option>
<option>BURKINA FASO</option>
<option>BURUNDI</option>
<option>CAMBODIA</option>
<option>CAMEROON</option>
<option>Canada(ST PIERRE / MIQUELON)</option>
<option>CANADA</option>
<option>CANARY ISLANDS</option>
<option>CAPE VERDE</option>
<option>CAYMAN IS</option>
<option>CENTRAL AFRICAN REP</option>
<option>CHAD</option>
<option>CHILE</option>
<option>CHINA</option>
<option>COLOMBIA</option>
<option>COMOROS</option>
<option>CONGO</option>
<option>Congo, Democratic Republic of(DEM REP OF THE CONGO)</option>
<option>COOK IS</option>
<option>COSTA RICA</option>
<option>COTE D IVOIRE</option>
<option>CROATIA</option>
<option>CUBA</option>
<option>CURACAO(NETHERLANDS ANTILLES)</option>
<option>CYPRUS</option>
<option>CZECH REPUBLIC</option>
<option>DENMARK</option>
<option>DJIBOUTI</option>
<option>DOMINICA</option>
<option>DOMINICAN REPUBLIC</option>
<option>EAST TIMOR</option>
<option>ECUADOR</option>
<option>EGYPT</option>
<option>EL SALVADOR</option>
<option>ERITREA</option>
<option>ESTONIA</option>
<option>ETHIOPIA</option>
<option>FALKLAND IS</option>
<option>FAROE IS</option>
<option>FIJI</option>
<option>FIJI(WALLIS / FUTUNA IS)</option>
<option>FINLAND</option>
<option>FRANCE</option>
<option>FRENCH GUIANA</option>
<option>GABON</option>
<option>GAMBIA</option>
<option>GEORGIA</option>
<option>GERMANY</option>
<option>GHANA</option>
<option>GIBRALTAR</option>
<option>GREECE</option>
<option>GREENLAND</option>
<option>GRENADA</option>
<option>GUADELOUPE</option>
<option>GUAM</option>
<option>GUAM(MICRONESIA)</option>
<option>GUATEMALA</option>
<option>GUINEA BISSAU</option>
<option>Guinea Republic(GUINEA)</option>
<option>GUINEA-EQUATORIAL(EQUATORIAL GUINEA)</option>
<option>GUYANA (British)(GUYANA)</option>
<option>HAITI</option>
<option>HONDURAS</option>
<option>HONG KONG</option>
<option>HUNGARY</option>
<option>ICELAND</option>
<option>INDIA</option>
<option>INDONESIA</option>
<option>Iran (Islamic Republic of)(IRAN)</option>
<option>IRAQ</option>
<option>Ireland, Republic of(IRELAND)</option>
<option>ISRAEL(GAZA STRIP)</option>
<option>ISRAEL</option>
<option>ISRAEL(WEST BANK)</option>
<option>Italy(SAN MARINO)</option>
<option>Italy(VATICAN CITY)</option>
<option>ITALY</option>
<option>JAMAICA</option>
<option>JAPAN</option>
<option>JERSEY</option>
<option>JORDAN</option>
<option>KAZAKHSTAN</option>
<option>KENYA</option>
<option>KIRIBATI</option>
<option>KOREA, Republic of(KOREA SOUTH)</option>
<option>KOREA, The D.P.R. of(KOREA NORTH)</option>
<option>KUWAIT</option>
<option>KYRGYZSTAN</option>
<option>Lao People&#39;s Democratic Republic(LAOS)</option>
<option>LATVIA</option>
<option>LEBANON</option>
<option>LESOTHO</option>
<option>LIBERIA</option>
<option>LIBYA</option>
<option>LIECHTENSTEIN</option>
<option>LITHUANIA</option>
<option>LUXEMBOURG</option>
<option>MACAU</option>
<option>Macedonia, Republic of (FYROM)(MACEDONIA)</option>
<option>MADAGASCAR</option>
<option>MALAWI</option>
<option>MALAYSIA</option>
<option>MALDIVES</option>
<option>MALI</option>
<option>MALTA</option>
<option>MARSHALL IS</option>
<option>MARTINIQUE</option>
<option>MAURITANIA</option>
<option>MAURITIUS</option>
<option>MEXICO</option>
<option>Moldova, Republic of(MOLDOVA)</option>
<option>MONACO</option>
<option>MONGOLIA</option>
<option>MONTSERRAT</option>
<option>MOROCCO</option>
<option>MOZAMBIQUE</option>
<option>MYANMAR</option>
<option>NAMIBIA</option>
<option>Nauru, Republic of(NAURU)</option>
<option>NEPAL</option>
<option>NETHERLANDS</option>
<option>NEW CALEDONIA</option>
<option>NEW ZEALAND</option>
<option>NICARAGUA</option>
<option>NIGER</option>
<option>NIGERIA</option>
<option>NIUE</option>
<option>NORWAY</option>
<option>OMAN</option>
<option>PAKISTAN</option>
<option>PANAMA</option>
<option>PAPUA NEW GUINEA</option>
<option>PARAGUAY</option>
<option>PERU</option>
<option>PHILIPPINES</option>
<option>POLAND</option>
<option>PORTUGAL</option>
<option>PUERTO RICO</option>
<option>QATAR</option>
<option>REUNION IS</option>
<option>ROMANIA</option>
<option>Russian Federation(RUSSIA)</option>
<option>RWANDA</option>
<option>SAIPAN(NORTHERN MARIANA IS)</option>
<option>SAMOA</option>
<option>SAO TOME / PRINCIPE</option>
<option>SAUDI ARABIA</option>
<option>SENEGAL</option>
<option>Serbia and Montenegro(YUGOSLAVIA)</option>
<option>SEYCHELLES</option>
<option>SIERRA LEONE</option>
<option>SINGAPORE</option>
<option>SLOVAKIA</option>
<option>SLOVENIA</option>
<option>SOLOMON IS</option>
<option>SOMALIA</option>
<option>Somaliland, Rep of (North Somalia)</option>
<option>South Africa(ST HELENA)</option>
<option>SOUTH AFRICA</option>
<option>SPAIN</option>
<option>SRI LANKA</option>
<option>ST LUCIA</option>
<option>St. BARTHELEMY</option>
<option>St. EUSTATIUS</option>
<option>St. Kitts(ST KITTS / NEVIS)</option>
<option>St. Maarten(NETHERLANDS ANTILLES)</option>
<option>St. Vincent(ST VINCENT/GRENADINE)</option>
<option>SUDAN</option>
<option>SURINAME(SURINAM)</option>
<option>SWAZILAND</option>
<option>SWEDEN</option>
<option>SWITZERLAND</option>
<option>SYRIA</option>
<option>TAHITI(FRENCH POLYNESIA)</option>
<option>TAIWAN</option>
<option>TAJIKISTAN</option>
<option>TANZANIA</option>
<option>THAILAND</option>
<option>TOGO</option>
<option>TONGA</option>
<option>TRINIDAD / TOBAGO</option>
<option>TUNISIA</option>
<option>TURKEY</option>
<option>TURKMENISTAN</option>
<option>TURKS / CAICOS IS</option>
<option>TUVALU</option>
<option>UGANDA</option>
<option>UKRAINE</option>
<option>UNITED ARAB EMIRATES</option>
<option>UNITED KINGDOM(ENGLAND)</option>
<option>UNITED KINGDOM(NORTHERN IRELAND)</option>
<option>UNITED KINGDOM(SCOTLAND)</option>
<option>UNITED KINGDOM(WALES)</option>
<option selected="">UNITED STATES</option>
<option>URUGUAY</option>
<option>UZBEKISTAN</option>
<option>VANUATU</option>
<option>VENEZUELA</option>
<option>VIETNAM</option>
<option>VIRGIN IS BRITISH</option>
<option>VIRGIN IS USA</option>
<option>YEMEN</option>
<option>ZAMBIA</option>
<option>ZIMBABWE</option>

  </select>
  </td>
            </tr>
            <tr>
              <td width="40%">Phone <input type="text" name="phone" size="20"><font color="red">*</font></td>
			  <td width="40%">Email <input type="text" name="email" size="20"><font color="red">*</font></td>
			  
              
            </tr>
            <tr>
              <td width="40%">
              
                
                  Order Message</td>
              <td width="60%">Gift Message</td>
            </tr>
            <tr>
              <td width="40%">
                <textarea rows="5" name="ordermessage" cols="25"></textarea>
              </td>
              <td width="60%">
                <textarea rows="5" name="message" cols="25"></textarea>
              </td>
            </tr>
          </table>
          
          
          
<center>
<table class="table98borderCheckout">
  <tr>
    <td width="29%">Payment Method </td>
    <td width="71%">
<select name="CardType" onChange="javascript:updateDisplay();">
<option value="MasterCard">MasterCard
<option value="VisaCard">Visa
<option value="DiscoverCard">Discover</option>
<option value="AmExCard">Amex</option>
<option value="4">Credit Card On File</option>
</select><img border="0" src="../images/visa.jpg"><img border="0" src="../images/mc.jpg"><img border="0" src="../images/discover.jpg"></td>
  </tr>
  <tr>
    <td width="29%">Card Number</td>
    <td width="71%"> <input name="CardNumber" size="28" maxlength="19"> (No - or 
    space)</td>
    <input name="CardNumberEncode" type="hidden">
    <input name="KEY" type="hidden">


  </tr>
  <tr>
    <td width="100%" colspan="2">Expiration Date: Month
<select name="ExpMon">
<option value="1" selected>1
<option value="2">2
<option value="3">3
<option value="4">4
<option value="5">5
<option value="6">6
<option value="7">7
<option value="8">8
<option value="9">9
<option value="10">10
<option value="11">11
<option value="12">12
</select> Year <input name="ExpYear" size="2" maxlength="2">(YY)<p>CVV
    <input name="verify" size="4" maxlength="4"> (last 3 digit in the back of 
    Visa or MC, or 4 digit in the front of MEX)<p>
    <input name="verifyEncode" type="hidden">
    
    </td>
  </tr>
</table>
<p>
<br>
<input type="submit" value="Check Out" OnClick="return CheckCardNumber(this.form)"></p>

</form>
</center>
 
          
          
         
          
          
                 <p align="center"></td>
    <td>
        &nbsp; &nbsp;</td>
    
  </tr>



<%


  '----------------check out content end--------------
'*********************************************************************************************************   
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
  
  

  
  
  

  
  
  
  
  
  
  
