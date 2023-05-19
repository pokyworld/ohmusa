
<%
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
<!-- #include file="include/asp_lib.inc.asp" -->
<!-- #include file="include/sqlCheckInclude.asp" -->
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
Dim strcontact, strlogin, strEmail, strEmail2, strcompany, strcountry, straddress, strCity, strState,strProvince, StrZip, strPhone, strFax
dim strphone2, strContact2, strweb, strResale, strCheck, strCard, strWire, strNet30
dim strBilling, strTradeshow, strsource, StrotherSource, strComment
dim strdropship, strretail, stronline, strdistributor, strcatalog, strdesigner, strother



Dim rst, cnn, objMail, strMailContent
Dim strSuccess

  
'Get data
strcontact =fixstring(Trim(Request.Form("pname")))
'strSuccess = Request.QueryString("pSuccess")
If len(strcontact)>0 Then
  'Continue getting data
  strlogin = fixstring(Trim(Request.Form("plogin")))
  dim sqlCmdlogin 
  sqlCmdLogin= "select * from dealer where login = '"&strlogin&"'"
  'Create connection to execute insert command
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  set rst = Server.CreateObject("adodb.RecordSet")
  
  dim rstMsg
  set rstMsg=  Server.CreateObject("adodb.RecordSet")
  
  cnn.Open
  rst.open sqlcmdLogin, cnn,3

  rstMsg.open "select * from screenmessage", cnn, 3
  
  if rst.eof then 

 strEmail=   fixstring(Trim(Request.Form("pemail")))
 
 
 strcompany = fixstring(Trim(Request.Form("pCompany")))
 strcountry = fixstring(Trim(Request.Form("pcountry")))
 straddress = fixstring(Trim(Request.Form("pAddress")))
 strCity = fixstring(Trim(Request.Form("pCity")))
 strState = fixstring(Trim(Request.Form("pState")))
 strProvince = fixstring(Trim(Request.Form("pprovince")))

 StrZip = fixstring(Trim(Request.Form("pZip")))
 strPhone = fixstring(Trim(Request.Form("pPhone")))


 strweb= fixstring(Trim(Request.Form("pWebsite")))
 

 strcatalog= fixstring(Trim(Request.Form("c9")))
 strdesigner= fixstring(Trim(Request.Form("c10")))
 strother= fixstring(Trim(Request.Form("c11")))
 

 strTradeshow=fixstring(Trim(Request.Form("ptradeshow")))
 strsource=fixstring(Trim(Request.Form("pSource")))

 strComment=fixstring(Trim(Request.Form("pComment")))
  
  
   strEmail2=   fixstring(Trim(Request.Form("pemail2")))
strFax = fixstring(Trim(Request.Form("pFax")))
 strphone2 = fixstring(Trim(Request.Form("pphone2")))
 strContact2= fixstring(Trim(Request.Form("pcontact2")))
 
strResale= fixstring(Trim(Request.Form("pResale")))
 strdropship= fixstring(Trim(Request.Form("c1")))
 strretail= fixstring(Trim(Request.Form("c2")))
 stronline= fixstring(Trim(Request.Form("c3")))
 strdistributor= fixstring(Trim(Request.Form("c4")))
 strcheck= fixstring(Trim(Request.Form("c5")))
 strCard= fixstring(Trim(Request.Form("c6")))
 strWire= fixstring(Trim(Request.Form("c7")))
 strNet30 = fixstring(Trim(Request.Form("c8")))
  
  StrotherSource=fixstring(Trim(Request.Form("potherSource")))
 strBilling=fixstring(Trim(Request.Form("pbilling")))
         
 if len(strProvince)>0 then
 	strState=strProvince
 end if

  'Prepare SQL statement
  strSQLCmd = "Insert Into dealer (customer, contact, phone, fax, street, city, state, zip, country, email, website, login, "&_
  "phone2, contact2, resale,  billing, show, source, othersource, comment, ccheck, card, wire, net30, dropship, " &_
  "retail, online, distri, ccatalog, design, regdate) " &_
    " values ('" & strcompany & "', '" & strContact & "', '" & strPhone & "', '" & strFax& "', '" &straddress & "', '" &_
     strCity & "', '" & strstate & "', '" & strZip & "', '" & strcountry & "', '" & stremail  & "', '" & strweb & "', '" &_
      strlogin & "', '" & strphone2& "', '" &strcontact2& "', '" & strresale & "', '" & strBilling& "', '" & strtradeShow &_
      "', '" & strSource & "', '" &_
      strOtherSource & "', '" & strComment& "',' " & strCheck& " ',' " & strCard & " ',' " & strWire & " ',' " &_
      strNet30 & " ',' " & strdropship& " ',' " & strRetail& " ',' " & strOnline & " ',' " & strdistributor & " ',' " &_
      strCatalog & " ',' " & strDesigner &" ',' " & date() &"')"
      
 

     
  cnn.Execute strSQLCmd
  
  
  
  'Send mail to webmaster
    
  dim cusType
  if not strComp(Ucase(strdropship), "1") then
  	cusType= "Drop Shipper"
  end if 
  if not strComp(Ucase(strRetail), "1") then
  	cusType=cusType& " Retail Store "
  end if
  if not strComp(Ucase(strOnline), "1") then
  	cusType=cusType& " Online Store "
  end if
  if not strComp(Ucase(strdistributor), "1") then
  	cusType=cusType& " Distributor "
  end if 	
  if not strComp(Ucase(strCatalog), "1") then
  	cusType=cusType& " Catalog Company "
  end if 
 if not strComp(Ucase(strDesigner), "1") then
  	cusType=cusType& " Designer "
 end if 	
  if not strComp(Ucase(strother), "1") then
  	cusType=cusType& " Others "
 end if 
 
 
  strMailContent = "Dear Sales," & "<br><br>"
  strMailContent = strMailContent & "You have received an new dealer application from "
  strMailContent = strMailContent & strCompany & "<br>"
  
  strMailContent = strMailContent & " Please verify their information within 24 hours " & "<br>"
  
  strMailContent = strMailContent & "Customer Type: " &  cusType  & "<br>"   
  strMailContent = strMailContent & "Contact: " & strContact  & "<br>"  
  strMailContent = strMailContent & "Phone: " &strphone & "<br>"
  strMailContent = strMailContent & "Fax: " &strFax & "<br>"
  strMailContent = strMailContent & "Address: " & strAddress & "<br>"
  
   
  strMailContent = strMailContent & "City: " & strCity  & " State: "&strState&"  Zip: " & strZip & "<br>"  
  strMailContent = strMailContent & "Province: " & strProvince &"  Country: " & strCountry  & "<br>"  
  strMailContent = strMailContent & "Email: " &strEmail & "<br>"
  strMailContent = strMailContent & "Web Site: " & strWeb & "<br>"
  
  strMailContent = strMailContent & "Prefered Login Name: " & strLogin  & "<br>"  
  strMailContent = strMailContent & "Alternate Contact: " & strContact2  & " Alternate Phone: " & strPhone2 & "<br>"   
  strMailContent = strMailContent & "Resale / Fed ID: " &strResale & "<br>"
  
  strMailContent = strMailContent & "Billing Address: " & strBilling  & "<br>"  
  strMailContent = strMailContent & "Tradeshow: " & strtradeshow  & " Source: " & strSource&"<br>" 
  strMailContent = strMailContent & "OtherSource: " &strOthersource & "<br>"
 
  dim pmyType
  if not strComp(Ucase(strCheck), "1") then
  	pmyType= "Company Check"
  end if
  if not strComp(Ucase(strCard), "1") then
  	pmyType=pmyType& " Credit Card "
  end if
  if not strComp(Ucase(strNet30), "1") then
  	pmyType=pmyType& " Net 30 "
  end if
  if not strComp(Ucase(strWire), "1") then
  	pmyType=pmyType& " Wire Tranfer "
  end if
  strMailContent = strMailContent & "Payment Type: " & pmyType  & "<br>"  

  
  
  strMailContent = strMailContent & "Comment: " & strComment & "<br>"
  
  
  
  'Set objMail = CreateObject("CDONTS.Newmail")
 ' objMail.From = strEmail
  'objMail.To = "sales05@omh1.com"
  'objMail.cc="service@omh1.com"
 
  'objMail.Subject = "New Dealer Registration:" & strCompany 
  'objMail.BodyFormat = 0 
  'objMail.MailFormat = 0
  'objMail.Body = strMailContent
  'objMail.Send
  'Set objMail = Nothing



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
    dim strSubject

    With cdoMessage 
        Set .Configuration = cdoConfig 
        .From = rstMsg("ms14")
	.To = rstMsg("ms14")
	.cc=rstMsg("ms6")
        .Subject = "New Dealer Registration:" & strCompany 
        .HTMLBody = strMailContent
	.Send 
    End With 
    

 
    Set cdoMessage = Nothing 

    Set cdoConfig = Nothing 
 
 
cnn.Close
  Set cnn = Nothing

  

    Response.Redirect "confirmregistration.asp"
  else
  Response.Redirect "confirmregistration.asp?error=1"
  	
  end if
Else	
%>

                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="product_stylesheet.css">
 <link rel="shortcut icon" type="image/x-icon" href="http://www.omhvn.com/favicon.ico" />
 
  <script language="JavaScript1.2" src="include/javascript.js"></script>
 <script language="JavaScript1.2">
function validateData(){
  var customer = document.dealer.pname.value;
  var login = document.dealer.plogin.value;
  var company = document.dealer.pcompany.value;
  var address = document.dealer.paddress.value;
  var city = document.dealer.pcity.value;
  var state = document.dealer.pstate.value;
  var zip = document.dealer.pzip.value;
  var phone  = document.dealer.pphone.value;
  //var fax = document.dealer.pfax.value;
  var email = document.dealer.pemail.value;
  var email2= document.dealer.pemail2.value;
  //var billing=document.dealer.pbilling.value;

  var verify = document.dealer.C11.value;
  
  //Check if contact name is empty
  if (isBlank(customer)){
    alert("Please fill out the contact name!");
    document.dealer.pname.focus();
    return false;
  }
  
  if (isBlank(login)){
    alert("Please pick a login name!");
    document.dealer.plogin.focus();
    return false;
  }
  if (login.indexOf(" ")>=0) {
  	alert("login name can't contain space");
  	document.dealer.plogin.focus();
  	return false;
 
  }
  
  if (login.indexOf("&")>=0) {
  	alert("login name can't contain space and these special characters: &$#!'*");
  	document.dealer.plogin.focus();
  	return false;
 
  }
    if (login.length<4) {
  	alert("login name must have at least 4 characters");
  	document.dealer.plogin.focus();
  	return false;
 
  }
  
  
  //if (login.indexOf("@")>=0) {
  //	alert("login name can't contain space and special character");
  //	document.dealer.plogin.focus();
  //	return false;
 
  //}

if (login.indexOf("$")>=0) {
 	alert("login name can't contain space and these special characters: &$#!'*");
  	document.dealer.plogin.focus();
  	return false;
 
  }
  
  if (login.indexOf("#")>=0) {
  	alert("login name can't contain space and these special characters: &$#!'*");
  	document.dealer.plogin.focus();
  	return false;
 
  }
  
  if (login.indexOf("!")>=0) {
  		alert("login name can't contain space and these special characters: &$#!'*");
  	document.dealer.plogin.focus();
  	return false;
 
  }




  if (login.indexOf("'")>=0) {
  	alert("login name can't contain space and these special characters: &$#!'*");
  	document.dealer.plogin.focus();
  	return false;
 
  }

  if (login.indexOf("*")>=0) {
 	alert("login name can't contain space and these special characters: &$#!'*");
  	document.dealer.plogin.focus();
  	return false;
 
  }  
 // if (isBlank(company)){
   // alert("Please fill out your company name!");
    // document.dealer.pcompany.focus();
    // return false;
  // }
 
  if (isBlank(address)){
    alert("Please fill out your shipping address!");
    document.dealer.paddress.focus();
    return false;
  }
  if (isBlank(city)){
    alert("Please fill out the city!");
    document.dealer.pcity.focus();
    return false;
  }
  if (isBlank(state)){
    alert("Please select your state!");
    document.dealer.pstate.focus();
    return false;
  }
  
  if (isBlank(zip)){
    alert("Please fill out your zip/postal code!");
    document.dealer.pzip.focus();
    return false;
  }
  
  //Check if phone is empty
 
	if (isBlank(phone)){
    alert("Please fill out your phone #!");
    document.dealer.pphone.focus();
    return false;
  }
  

  

  //Check if Email is empty
  if (isBlank(email)){
    alert("Please enter your email address!");
    document.dealer.pemail.focus();
    return false;
  }
   //Check if confirmed Email is empty
  if (isBlank(email2)){
    alert("Please confirm your email address!");
    document.dealer.pemail2.focus();
    return false;
  }

  //Check if Email is valid
  if (! isEmail(email)){
    alert("Invalid email address!");
    document.dealer.pemail.focus();
    return false;
  }
  
  //Check if confirmed Email is valid
  if (! isEmail(email2)){
    alert("Invalid email address!");
    document.dealer.pemail2.focus();
    return false;
  }

  
  //Check if Email is confirmed correctly
  if (!(email==email2)){
    alert("Email and confirm email do not match!");
    document.dealer.pemail2.focus();
    return false;
  }
  
  
   
   
   
 
	//if (document.dealer.cbilling.checked!=0){
    	//document.dealer.pbilling.value="Same As Shipping Address";
	    //return true;
  //}
  
  
  //Check if billing is OK
 
	//if (isBlank(billing)){
    //alert("Please fill out your complete billing address");
    //document.dealer.pbilling.focus();
    //return false;
  //}



 //Check if verification is OK
 
	if (document.dealer.C11.checked==0){
    alert("Please agree to our terms and conditions");
    document.dealer.C11.focus();
    return false;
  }

  
  return true;
}
        </script>
        
        
        
  </head>
  <body>
 
 
 
 
 
 <table class="fixedTable" >



<tr>
	<td class= "mainleft" >  </td>
	<td class = "maincenter" >   
	
<!--#Include file="Header.asp"  -->
    <table class="searchTable">
        <tr>
                          
                        <td class="cssTextCENTER" height="28" width="100%">
                            <form action="ProductSearch.asp" method="POST" name="SearchForm">
                         
                               
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
                   
                    <%
                    While Not rstCategory.EOF
                    %>
                   
                    
                   
                    <tr>
                      <td width="100%" align="left">&nbsp;</td>
                    </tr>
                    <tr>
                      <td align="left" class = "tdmargin10">
					 
					 
                      <span class="cssLink">
                      <a href="ProductSearch.asp?pCategoryID=<%=rstCategory("Category_ID")%>" title="Ship Model - <%=rstCategory("Category_Name")%>"><%=rstCategory("Category_Name")%> </a>
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
                </td>
                
             <!--end   <td class="category"> -->
    
    
  

    
    <td class="pageContent" align="center">
    
    
    
  
      
      
               <!--start content about us -->
      <table class="table98border_aboutus">
	<tr>
      <th class="thfeatured" colspan="3" >
     CREATE A TRADE ACCOUNT
     </th>
     </tr>
  
	
	<tr>
	<td >
	</td>
										<td  >
                                           
                                      <form method="POST" name="dealer" action="registration.asp"> 
                                              
                                            
                                             
                                              <p class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              
                                              <span style="line-height:150%">
                                              <u>
                                              Contact Name:</u>
                                              <input type="text" name="pname" size="21"></span><span style="line-height:150%; ">&nbsp; *&nbsp;  </span>
                                              </p>
                                              <p class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                                  
                                                  <u>Login 
                                              Name</u><span style="line-height:150%"><input type="text" name="plogin" size="25">*</span></p>
                                           
                                              
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              Email<span style="line-height:150%"><input type="text" name="pemail" size="41"></span><span style="line-height:150%; ">*</span></p>
                                              
                                              
                                              
                                              
                                            
                                              <p style="font-size:10px">
											Upon checking this box, you agree to our  <a href ="faq.asp" target= "blank"> terms and conditions. </a> <br>											
                                            For dealers, please <a href="mailto:sales@omhusa.com">email</a> us 
                                            a copy of your resale certificate (mandatory in CA) or 
                                            business license. 
											
                                            <Br />An opening order or deposit may be required for new drop shipper.</font>
											
											</p>
											
											
											
											
											<span style="line-height:150%">
                                              
                                              </span>
                                              </p>
                                              <p align="center" style="margin-top: -1; margin-bottom: -1">
                                                                                                    I 
                                            agree<span style="line-height:150%">*<input type="checkbox" name="C11" value="1"></span></p>
                                              <p align="center" style="margin-top: -1; margin-bottom: -1">
                                            </p>
                                              <p align="center" style="margin-top: -1; margin-bottom: -1">
                                            &nbsp;
											 <input type="submit" value="Submit" onClick="return validateData()" name="button1">
											<input type="reset" value="Reset" name="B2"></p>
                                              <p align="center" style="margin-top: -1; margin-bottom: -1">&nbsp;
                                            </p>
                                             <p align="center">
                                       <a href="https://www.positivessl.com" target ="_blank"  style="font-family: arial; font-size: 10px; color: #212121; text-decoration: none;"><img src="https://www.positivessl.com/images-new/PositiveSSL_tl_white.png" alt="SSL Certificate" title="SSL Certificate" border="0" /></a>
                                        </p>
                                            </form>
										</td>
										
										<td></td>
	
										
									</tr>
								
	
	
							
						


  

  
  
  

  </table>
					
      
		
		 
      
      		
</td>

<!--end td class pagecontent -->



</tr>
	 
      
</table>
      
        <!--end content contactus -->
    
        
      
      
      
      



<!--#Include file="Footer.asp"  --> 
 
</td>

 <!--end mainCenter -->



<td class = "mainright" >    </td>
</tr>
</table>
 
 
 

  
  
  
  
  
  
  </body>
  
  
  
  
  </html>
  	<% end if %>