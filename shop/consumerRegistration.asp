
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
                                              
                                              <p align="center" >
                                              &nbsp;</p>
                                              
                                              <p class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              <u>Please check all 
                                              that apply to you or your business</u></p>
                                              <p class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              Drop Ship<input type="checkbox" name="C1" value="1">&nbsp;&nbsp;&nbsp;&nbsp;Retail Store&nbsp;
                                              <input type="checkbox" name="C2" value="1">&nbsp;&nbsp;&nbsp;Online
                                              <input type="checkbox" name="C3" value="1">&nbsp;&nbsp;&nbsp;&nbsp;Distributor<span style="line-height:150%"> </span>
                                              <input type="checkbox" name="C4" value="1"></p>
                                              <p class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                                  Mail 
                                              Order
                                              <input type="checkbox" name="C9" value="1"> 
                                              Designer
                                              <input type="checkbox" name="C10" value="1">
											  Others
                                              <input type="checkbox" name="C11" value="1">
											  </p>
                                              <p class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              
                                              <span style="line-height:150%">
                                              <u>
                                              Contact Name:</u>
                                              <input type="text" name="pname" size="21"></span><span style="line-height:150%; ">&nbsp; *&nbsp;  </span>
                                              </p>
                                              <p class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                                  
                                                  <u>Prefer Login 
                                              Name</u><span style="line-height:150%"><input type="text" name="plogin" size="25">*</span></p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              <u>
                                              Co</u><u>mpany 
                                              Name:</u>
                                              
                                              <span style="line-height:150%">
                                              
                                              &nbsp;
    <input type="text" name="pcompany" size="33"></span><span style="line-height:150%; font-size:12.0pt"> </span>
                                                  <span style="line-height:150%; ">
                                                  *</span></p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              Country:
                                         
     <select class="selectlarge" id="selCountry" name="pcountry" size="1">
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
                                              
                                              
                                              
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              <u>Address:</u>
                                              <span style="line-height:150%">
                                              <input type="text" name="paddress" size="40"></span><span style="line-height:150%; ">*</span></p>
                                              
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              
                                              
                                              <span style="line-height:150%">
                                              City: 
                                              <input type="text" name="pcity" size="21"></span><span style="line-height:150%; ">*</span><span style="line-height:150%"> State</span>
                                               <select id="selState" name="pstate" class="select">
					<option value="">Select One</option>
					<option value="Others">Others</option>
<option value="AL">Alabama</option>
<option value="AK">Alaska</option>
<option value="AZ">Arizona</option>
<option value="AR">Arkansas</option>
<option value="CA">California</option>
<option value="CO">Colorado</option>
<option value="CT">Connecticut</option>
<option value="DE">Delaware</option>
<option value="DC">District of Columbia</option>
<option value="FL">Florida</option>
<option value="GA">Georgia</option>
<option value="HI">Hawaii</option>
<option value="ID">Idaho</option>
<option value="IL">Illinois</option>
<option value="IN">Indiana</option>
<option value="IA">Iowa</option>
<option value="KS">Kansas</option>
<option value="KY">Kentucky</option>
<option value="LA">Louisiana</option>
<option value="ME">Maine</option>
<option value="MD">Maryland</option>
<option value="MA">Massachusetts</option>
<option value="MI">Michigan</option>
<option value="MN">Minnesota</option>
<option value="MS">Mississippi</option>
<option value="MO">Missouri</option>
<option value="MT">Montana</option>
<option value="NE">Nebraska</option>
<option value="NV">Nevada</option>
<option value="NH">New Hampshire</option>
<option value="NJ">New Jersey</option>
<option value="NM">New Mexico</option>
<option value="NY">New York</option>
<option value="NC">North Carolina</option>
<option value="ND">North Dakota</option>
<option value="OH">Ohio</option>
<option value="OK">Oklahoma</option>
<option value="OR">Oregon</option>
<option value="PA">Pennsylvania</option>
<option value="RI">Rhode Island</option>
<option value="SC">South Carolina</option>
<option value="SD">South Dakota</option>
<option value="TN">Tennessee</option>
<option value="TX">Texas</option>
<option value="UT">Utah</option>
<option value="VT">Vermont</option>
<option value="VA">Virginia</option>
<option value="WA">Washington</option>
<option value="WV">West Virginia</option>
<option value="WI">Wisconsin</option>
<option value="WY">Wyoming</option>

  </select>
    &nbsp;&nbsp;</p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                                  Zip/Postal<span style="line-height:150%"><input type="text" name="pzip" size="19"></span><span style="line-height:150%; ">*</span></p>
                                              
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              Province
                                              <span style="line-height:150%">
                                              <input type="text" name="Pprovince" size="21"></span></p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              <u>Phone:</u>
                                              <span style="line-height:150%">
                                              <input type="text" name="pphone" size="29"></span><span style="line-height:150%; ">*</span>
											   </p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              <u>Web site</u>:<span style="line-height:150%"><input type="text" name="pwebsite" size="35"></span></p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              Email<span style="line-height:150%"><input type="text" name="pemail" size="41"></span><span style="line-height:150%; ">*</span></p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                                  <span style="line-height:150%; ">Confirm 
                                              Email: </span><span style="line-height:150%">
                                              <input type="text" name="pemail2" size="42"></span><span style="line-height:150%; ">*</span></p>
                                              
											 
                                                                                           
                                              <p class="MsoBodyText" style="margin-top:-1; margin-bottom:-1">
                                              <span style="line-height:150%">
                                              How did you hear about us?</span></p>
                                              
                                              
                                              <p class="MsoBodyText" style="margin-top:-1; margin-bottom:-1"></p>
                                              <![if !mso]><![endif]>
                                              
                                              <p style="margin-top: -1; margin-bottom:-1">
                                              
                                              <span style="line-height:150%">
                                              Trade Show&nbsp;/&nbsp;Show Name:
                                              <input type="text" name="ptradeshow" size="17"></span></p>
                                              <p style="margin-top: -1; margin-bottom:-1">
                                                  <span style="line-height:150%">
                                                  Internet&nbsp; 
                                              Source:<input type="text" name="psource" size="29"></span></p>
                                              
                                              
                                              <p class="MsoBodyText" style="margin-top:-1; margin-bottom:-1">
                                              <span style="line-height:150%">
                                              Other 
                                              Information/Comments:</span></p>
                                              
                                              <p class="MsoBodyText" style="line-height:150%; margin-top:-1; margin-bottom:-1">
                                              <span style="line-height:150%">
                                              <textarea rows="5" name="pcomment" cols="50%"></textarea></span></p>
                                           
                                              
                                              
                                            
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