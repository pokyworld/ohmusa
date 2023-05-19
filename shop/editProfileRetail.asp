<%@ Language=VBScript %>
<%option explicit%>
<!--include virtual="/shop/payment/functions/helpers.inc"-->
<%
' Dim Item
' For Each Item In Request.Form
'   Response.Write "<pre>" & Item & ": " & Request.Form(Item) & "</pre>"
' Next

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





%>



<%
Dim strcontact, strlogin, strEmail, strcompany, strcountry, straddress, strCity, strState, StrZip, strPhone, strFax
dim strphone2, strContact2, strweb, strResale, strCheck, strCard, strWire, strNet30
dim strCreditcard, strExpire, strcvv, strBilling, strTradeshow, strsource, StrotherSource, strComment
dim strdropship, strretail, stronline, strdistributor, strcatalog, strdesigner, strPassword
dim redirectUrl
Dim cnn, rst, objMail, strSQLCmd, strMailContent
Dim Success

  redirectUrl=trim(request.querystring("callUrl"))
    success=trim(request.querystring("success"))
	if len(success)>0 then 
		success=1
	else 
		success=0
	end if
	
  if len(redirectUrl)=0 then
	 redirectUrl= "/shop/productsretail.asp"
 end if
 
 
  
'Get data

dim strAction
strAction = Request.Form("pAction")


strLogin = session("login")
'response.write(len(strLogin))
If len(strLogin)>0 Then
	session("login")=strLogin
else
	response.redirect("/shop/productsretail.asp")
end if

If len(strLogin)>0 and len(strAction)=0 Then
  dim sqlCmd 
  sqlCmd= "select * from consumer where login = '"&strlogin&"'"
  'Create connection to execute insert command
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  set rst = Server.CreateObject("adodb.RecordSet")


 
  cnn.Open

  rst.open sqlcmd, cnn,3
  if (not rst.eof) then 

 	
%>
  

 
                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
 
 <script language="JavaScript1.2" src="../include/javascript.js"></script>
<script lnguage="JavaScript1.2">

function validateData(){
  var customer = document.consumer.pname.value;
  var login = document.consumer.plogin.value;
  var email = document.consumer.pemail.value;
  var company = document.consumer.pcompany.value;
  var country = document.consumer.pcountry.value;
  var address = document.consumer.paddress.value;
  var city = document.consumer.pcity.value;
  var state = document.consumer.pstate.value;
  var zip = document.consumer.pzip.value;
  var billing= document.consumer.pBilling.value;
  var phone  = document.consumer.pphone.value;
 
  
  
  //Check if contact name is empty
  if (isBlank(customer)){
    alert("Please fill out the contact name!");
    document.consumer.pname.focus();
    return false;
  }

  
  if (isBlank(login)){
    alert("Please pick a login name!");
    document.consumer.plogin.focus();
    return false;
  }
  
  if (login.indexOf(" ")>=0) {
  	alert("login name can't contain space and special character");
  	document.consumer.plogin.focus();
  	return false;
 
  }
  if (login.indexOf("'")>=0) {
  	alert("login name can't contain space or special character");
  	document.consumer.plogin.focus();
  	return false;
 
  }

  if (login.indexOf("*")>=0) {
  	alert("login name can't contain space or special character");
  	document.consumer.plogin.focus();
  	return false;
 
  }  


  if (isBlank(country)){
    alert("Please fill out your country!");
    document.consumer.pcountry.focus();
    return false;
  }
  if (isBlank(address)){
    alert("Please fill out your shipping address!");
    document.consumer.paddress.focus();
    return false;
  }
  if (isBlank(city)){
    alert("Please fill out the city!");
    document.consumer.pcity.focus();
    return false;
  }
  if (isBlank(state)){
    alert("Please fill out your state!");
    document.consumer.pstate.focus();
    return false;
  }
  
  if (isBlank(zip)){
    alert("Please fill out your zip/postal code!");
    document.consumer.pzip.focus();
    return false;
  }
  
   if (isBlank(billing)){
    alert("Please fill out your complete billing address!");
    document.consumer.pBilling.focus();
    return false;
  }
  
  
  if (isBlank(phone)){
    alert("Please fill out your phone number!");
    document.consumer.pphone.focus();
    return false;
  }

  
  //Check if Login is empty
  if (isBlank(login)){
    alert("Please enter your email address!");
    document.consumer.plogin.focus();
    return false;
  }
  
  //Check if Login is valid
  if (! isEmail(login)){
    alert("Invalid email address!");
     document.consumer.plogin.focus();
    return false;
  }
  
  //Check if Email is empty
  if (isBlank(email)){
    alert("Please enter your email address!");
    document.consumer.pemail.focus();
    return false;
  }
  
  //Check if Email is valid
  if (! isEmail(email)){
    alert("Invalid email address!");
     document.consumer.pemail.focus();
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
	
<!--#Include file="headerRetail.asp"  -->
  
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
      <th class="thfeatured" colspan = "3" >
          Consumer Profile
     </th>
     </tr>
     
     
     
     <%'***************************************************************************************************%>
     <%'----------------Middle content start--------------%>
  	
  	
  			
									<tr>
									<td >&nbsp;</td>
										<td >
                                           
                                      <form method="POST" name="consumer" > 
                                      
                                             <table>
                                              
                                              <p class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                          
                                              <span style="line-height:150%">
                                              <tr>
											  <td>
                                              
                                              Contact Name</td><td>
                                              <input type="text" name="pname" size="21" value="<%=rst("contact")%>"></span>
											     <span style="line-height:150%; ">&nbsp;*</span>
                                              <br>
											  
											  </td>
											  
                                           
                                              </tr>
											  
                                              
                        
                        <tr>
											    <td>
                                              
                                              Email Address<span style="line-height:150%"></td><td>
                                              <input type="text" name="pemail" id="pemail" value="<%=trim(rst("Email"))%>" size="20">&nbsp;*<br></td>
											  </tr>
                        <tr>
                          <td>&nbsp;</td>
											    <td style="line-height:0;">
                          <%
                            Dim matchChecked, showLogin
                            If LCase(Trim(rst("Email"))) = LCase(Trim(rst("login"))) Then
                              matchChecked = " checked"
                              showLogin = "none"
                            Else
                              matchChecked = ""
                              showLogin = "table-row"
                            End If
                          %>
                                              <input type="checkbox" name="match" id="match"  value="match"<%=matchChecked%> />
                                              <label for="loginMatch"> Use Email for Login</label><br><br>
											  </tr>
                        <tr id="row-login" style="display:<%=showLogin%>">
											  <td>
                                              
                                              Login Name/Email<span style="line-height:150%"></td><td>
                                              <input type="text" name="plogin" id="plogin" value="<%=trim(rst("login"))%>" size="20" />&nbsp;*<br></td>
											  </tr>
											  <tr>
											  <td>
                                              
											  Password
											  </td><td>
                                              
                                              <input type="password" name="ppassword" size="25" value="<%=trim(rst("password"))%>"></span></p></td>
											 </tr>
											 <tr>
											  <td>
                                              
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              
                                              
                                              Company 
                                              Name</td><td>
											  <span style="line-height:150%">
                                              <input type="text" name="pcompany" size="33" value="<%=rst("customer")%>"></span><span style="line-height:150%; "> </span>
											  </td>
                                              </p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              </tr><tr>
											  <td>
                                              
                                              Country</td><td>
                                              
                                              <input type="text" name="pcountry" size="33" value="<%=rst("country")%>"></td>
											  </tr><tr>
											  <td>
                                              
											  
											  
				
											  
											  
                                              </p>
											  
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              
                                              Shipping Address
											  </td>
											  <td>
											  
                                              <input type="text" name="paddress" size="40" value="<%=rst("street")%>">&nbsp;*
                                              </td>
											  </tr><tr>
											  <td>
                                              
                                              </p>
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                                  <span style="line-height:150%">
                                                  City &nbsp;
												  </td><td>
												  <input type="text" name="pcity" size="21" value="<%=rst("city")%>"></span>
                                             
											 <span style="line-height:150%">
											 </td>
											 </tr> <tr>
											  <td>
                                              
                                              State/ Province </td>
											  <td>
                                            </span> 
                                              <span style="line-height:150%">
                                              
										    <input type="text" name="pstate" size="15" value="<%=rst("state")%>"></span>
											</td>
											 </tr> <tr>
											  <td>
                                              
											Zip/Postal</td><td>
    										<span style="line-height:150%">
    										<input type="text" name="pzip" size="19" value="<%=rst("zip")%>"></span><span style="line-height:150%; ">&nbsp;*</span>
											
											</td>
											</tr>  <tr>
											  <td>
                                              
                                              </p>
                                                                                         
                                              
                                              
                                              
                                              
                                            
                                              
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              Billing Address 
											  </td>
											  <td>
											  <input type="checkbox" name ="billingCheckBox"> Same as Shipping
											  <br> 
                                              <span style="line-height:150%">
											   <input type="text" name="pBilling" size="40" value="<%=rst("billing")%>"> </span><span style="line-height:150%; ">
                                              *</span></p>
											  </td>
											  
											  </tr>
											  
											  <tr>
											  <td>
                                              
                                              
                                              
                                              
                                              
                                            
                                              
                                              <p class="MsoNormal" 
                                                  style="line-height: 150%; margin-top:-1; margin-bottom:-1; ">
                                              
                                              Phone
											  </td><td>
                                              <span style="line-height:150%">
                                              <input type="text" name="pphone" size="29" value="<%=rst("phone")%>"></span><span style="line-height:150%; ">&nbsp;*&nbsp;&nbsp; </span>
											  <br>
											  
                                           </p>
                                              <p class="MsoNormal" 
                                                  style="line-height: 150%; margin-top:-1; margin-bottom:-1; ">
                                              
                                              </td>
											  </tr>
											  <tr>
											  <td>
                                              
											  Alt. Phone </td><td>
                                              <span style="line-height:150%">
                                              <input type="text" name="pphone2" size="25" value="<%=trim(rst("phone2"))%>">&nbsp; </span>
											  <br>
											  </td>
											  </tr>
											  <tr>
											  <td>
                                              
                                                  Alt. Contact<span style="line-height:150%"> </td><td>
                                              <input type="text" name="pContact2" size="25" value="<%=rst("contact2")%>"></span></p>
                                              <p class="MsoNormal" 
                                                  style="line-height: 150%; margin-top:-1; margin-bottom:-1;">
                                              
                                              </span>
											  <br>
                                            </td>
											</tr>
											  <tr>
											  <td>
                                              
                                              <p class="MsoNormal" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              <span style="line-height:150%;">
                                              Comments</span> </p></td><td>
                                              <p class="MsoBodyText" style="line-height:150%; margin-top:-1; margin-bottom:-1">
                                              <span style="line-height:150%">
                                              <textarea rows="5" name="pcomment" cols="40"><%=rst("comment")%></textarea></span></p>
                                            </td>
											</tr>
											
											  
											  <tr>
											  <td></td>
											  <td>
                                              <p align="center" style="margin-top: -1; margin-bottom: -1">
                                            &nbsp;
											 <input type="submit" value="Save" onClick="return validateData()" name="pAction"> </p>
											 </td>
											 </tr>
											 
                                            </form>
                                              
											  <br>
											  
											  <%if success=1 then%>
											  <tr>
											  <td></td>
												
											  <td><br><br>
											  
												Thank you! Your information has been updated.
												</td>
												</tr>
												</table>
												
												
												
												<%end if%>
                                              
										</td>
										<td>&nbsp;&nbsp;</td>
									</tr>
			
  	
          <tr>
              <td colspan ="3">
                    &nbsp;&nbsp;&nbsp;
              </td>
          </tr>
						
							
							
									


  	
  	
   <% '----------------Middle content end--------------
  '***************************************************************************************************%> 

   
			
		</table><!--end table98 --> 
    
       </td>
       </tr>
       </table>
        <!--end mainTable-->
        
        <!--#Include file="FooterRetail.asp"  --> 
        
        </td>
        <!--end mainCenter-->
        



<td class = "mainright" >    </td>
</tr>

</table>

<script>
document.addEventListener("DOMContentLoaded", () => {
  console.log("Document loaded");
  var match = document.querySelector("#match");
  var pemail = document.querySelector("#pemail");
  var plogin = document.querySelector("#plogin");
  var rowlogin = document.getElementById("row-login");
  var storedLogin = plogin.value;
  match.addEventListener("change", (e) =>{
    console.log("checked", match.checked);
    // if(match.checked === true) {
    //   match.checked = false;
    // } else {
    //   match.checked = true;
    // }
    if(match.checked === true) {
      rowlogin.style.display="none";
      plogin.value = pemail.value;
    } else {
      rowlogin.style.display="table-row";
      plogin.value = storedLogin;
    }
  });
  pemail.addEventListener("keyup", (e) =>{
    if(match.checked) {
      plogin.value = pemail.value;
    } else {
      plogin.value = storedLogin;
    }
  });
});
  
</script>
  </body>
  </html> 
  
  <% 

cnn.close
set cnn=nothing

end if %><% else
	'edit consumer information and update all consumer information


  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")

'get server smtp information
 dim rstMsg
  set rstMsg=  Server.CreateObject("adodb.RecordSet")
 
  cnn.Open
  rstMsg.open "select * from screenmessage", cnn, 3
	dim smtp, user, pass, sendto
	smtp=rstMSG("ms13")
	user=rstmsg("ms12")
	pass=rstmsg("ms11")
	sendto=rstmsg("ms15")



  
'Get data
 strcontact =fixstring(Trim(Request.Form("pname")))
 strlogin =fixstring(Trim(Request.Form("plogin")))
 strPassword=   Replace(Trim(Request.Form("ppassword")), "'", "''")
 strEmail=   Replace(Trim(Request.Form("pemail")), "'", "''")
 strcompany = Replace(Trim(Request.Form("pcompany")), "'", "''")
 strcountry = Replace(Trim(Request.Form("pcountry")), "'", "''")
 straddress = Replace(Trim(Request.Form("pAddress")), "'", "''")
 strCity = Replace(Trim(Request.Form("pCity")), "'", "''")
 strState = Replace(Trim(Request.Form("pState")), "'", "''")
 StrZip = Replace(Trim(Request.Form("pZip")), "'", "''")
 strPhone = Replace(Trim(Request.Form("pPhone")), "'", "''")

 strphone2 = Replace(Trim(Request.Form("pphone2")), "'", "''")
 strContact2= Replace(Trim(Request.Form("pcontact2")), "'", "''")
 strweb= Replace(Trim(Request.Form("pWebsite")), "'", "''")
 strBilling=Replace(Trim(Request.Form("pBilling")), "'", "''")
 strComment=Replace(Trim(Request.Form("pComment")), "'", "''")
    
response.write 	(Request.Form("billingCheckBox")) 
if strcomp (Request.Form("billingCheckBox"), "on") = 0 then
	strBilling="Same as shipping address"
	
end if




	 
  strSQLCmd = "update consumer set " &_
    "customer = '" & strcompany & "', " &_
    "contact = '"  & strContact & "', " &_
    "phone = '"   &  strPhone &"', " &_
   
    "street = '"   &  straddress &"', " &_
     "city= '"     &  strCity &"', " &_
     "state = '"  &  strstate &"', " &_
     "zip= '"     &  strZip &"', " &_
     "country= '" &  strcountry &"', " &_
     "email = '"  & stremail  &"', " &_
     "website='"   & strweb &"', " &_
     "login = '"  & strlogin &"', " &_
     "phone2 = '" & strphone2&"', " &_
     "contact2='" & strcontact2&"', " &_
     "billing= '" & strBilling&"', " &_
     "comment='" & strComment&"'"
	if len(strPassword) > 0 then
   	 	strSQLCmd = strSQLCmd & ", Password = '" & strPassword & "'"
  	end if
  	 strSQLCmd = strSQLCmd & " where login= '"&session("login")&"'"
	'response.write (strSQlCmd)
  'Create connection to execute insert command
dim cnn2
  Set cnn2 = Server.CreateObject("ADODB.Connection")
  cnn2.ConnectionString = Application.Contents("dbConnStr")
  cnn2.Open
  if len(session("login"))>1 then
  	cnn2.Execute strSQLCmd
  end if
  session("login")=strLogin

  cnn2.Close
  Set cnn2 = Nothing
	
	
	
	
  	
  strMailContent = "Hi," & "<br><br>"
  strMailContent = strMailContent & "The following information has been updated: " & "<br>"
  strMailContent = strMailContent & "Company: " & strCompany & "<br>"
  strMailContent = strMailContent & "Contact: " & strContact  & "<br>"  
  strMailContent = strMailContent & "Phone: " &strphone & "<br>"
  strMailContent = strMailContent & "Address: " & strAddress & "<br>"
  
   
  strMailContent = strMailContent & "City: " & strCity  & " State: "&strState&" Zip: " & strZip & "<br>"  
  strMailContent = strMailContent & "Billing address: " & strBilling  & "<br>"  
 strMailContent = strMailContent & "Country: " & strCountry  & "<br>"  
  strMailContent = strMailContent & "Email: " &strEmail & "<br>"
  strMailContent = strMailContent & "Web Site: " & strWeb & "<br>"
  
  strMailContent = strMailContent & "Prefered Login Name: " & strLogin  & "<br>"  
  strMailContent = strMailContent & "Alternate Contact: " & strContact2  & " Alternate Phone: " & strPhone2 & "<br>"   
   
 
  
  
  strMailContent = strMailContent & "Comment: " & strComment & "<br>"
  
  
   'send email
   'send email using CDO / 02/17/2013 by eric

   dim sch, cdoconfig, cdomessage
   sch = "http://schemas.microsoft.com/cdo/configuration/" 
 
    Set cdoConfig = CreateObject("CDO.Configuration") 


 
    With cdoConfig.Fields 
        .Item(sch & "sendusing") = 2 ' cdoSendUsingPort 
        .Item(sch & "smtpserver") = smtp
	.Item(sch & "smtpauthenticate") =1
	.Item(sch & "sendusername") =user
	.Item(sch & "sendpassword") =pass
        .update 
    End With 
 
    Set cdoMessage = CreateObject("CDO.Message") 
 
    With cdoMessage 
        Set .Configuration = cdoConfig 
        '.From = strEmail
		'new server no longer allow sending email directly from customer
		' use the same email from/to to send email
        .From = user
        .To = user
        .Subject = "Updated profile:" & strCompany 
        .HTMLBody = strMailContent
        .Send 
    End With 
    

 
    Set cdoMessage = Nothing 
    Set cdoConfig = Nothing 

  cnn.Close
  Set cnn = Nothing


  
  
  
 response.redirect("/shop/editProfileRetail.asp?success=1")
	
		
end if%>