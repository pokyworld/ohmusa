
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


dim resetPasswordEmailSent
resetPasswordEmailSent=-1


Dim rst, cnn, objMail, strMailContent, strMailContent2
Dim strSuccess


  
  strEmail=   fixstring(Trim(Request.Form("pemail")))
  
  'use email as login 
  
  strlogin = stremail

  if len(strlogin)> 0 then
  
  dim sqlCmdlogin 
  sqlCmdLogin= "select * from consumer where email like '%" & strEmail&"'"
  'Create connection to execute insert command
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  set rst = Server.CreateObject("adodb.RecordSet")
  
  dim rstMsg
  set rstMsg=  Server.CreateObject("adodb.RecordSet")
  
  cnn.Open
  rst.open sqlcmdLogin, cnn,3

  rstMsg.open "select * from screenmessage", cnn, 3
  
  '---------------------------------------------------------------------------------------------------
  'if email in the system then ok to go forward to send an email to reset password
  
  if not rst.eof  then 

 

 
  
  'Send mail to webmaster
    
  dim cusType
  
 cusType="Consumer"
 
 
 
  strContact=rst("contact")
  
  strMailContent2="Dear " & strContact & ",<br><br> "
  strMailContent2 = strMailContent2  & "Please click on the link below to reset your password<br><br>"
  
  strMailContent2 = strMailContent2  & "https://omhusa.com/passwordReset.asp?login=" & strEmail
  
  
  
  
  
  
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



dim sch, cdoconfig, cdomessage, cdoMessage2
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
 
  

   
    
  Set cdoMessage2 = CreateObject("CDO.Message") 
    With cdoMessage2 
        Set .Configuration = cdoConfig 
        .From = rstMsg("ms14")
	.To = strEmail
	    .Subject = "Password reset from OMHUSA"
        .HTMLBody = strMailContent2
	.Send 
    End With 
    
 
	
    Set cdoMessage = Nothing 
	set cdoMessage2 = nothing
	

    Set cdoConfig = Nothing 
 
 
cnn.Close
  Set cnn = Nothing

  

   resetPasswordEmailSent=1
   
  else
  ' email not in the system
    
	resetPasswordEmailSent=0
  
  	
  end if
 end if
 
%>

                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
 <link rel="shortcut icon" type="image/x-icon" href="http://www.omhvn.com/favicon.ico" />
 
  <script language="JavaScript1.2" src="../include/javascript.js"></script>
 <script language="JavaScript1.2">
 function resetInputFocusOut()
 {
 
 var contactName=document.dealer.pname;
 if (contactName.value==="")
 {
 contactName.value="Contact Name";

  }
   var email=document.dealer.pemail;
 if (email.value==="")
 {
 contactName.value="Email";

  }
  
   
 }
 
function validateData(){
 
  var email = document.dealer.pemail.value;
 

  //Check if Email is empty
  if (isBlank(email)){
    alert("Please enter your email address!");
    document.dealer.pemail.focus();
    return false;
  }
   //Check if confirmed Email is empty
 
  //Check if Email is valid
  if (! isEmail(email)){
    alert("Invalid email address!");
    document.dealer.pemail.focus();
    return false;
  }
  
 
  
  return true;
}
        </script>
        
    <meta name="viewport" content="width=device-width, initial-scale=0.75">    
        
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
                </td>
                
             <!--end   <td class="category"> -->
    
    
  

    
    <td class="pageContent" align="center">
    
    
    
  
      
      
               <!--start content about us -->
      <table class="table98border_aboutus">
	<tr>
      <th class="thfeatured" colspan="3" >
   Forgot Password
     </th>
     </tr>
  
	
	<tr>
	<td width="10%" >
	</td>
										<td  >
                                           
                                      <form method="POST" name="dealer" action="forgotPassword.asp"> 
                                              
                                            
                                             <br>
							
	
						  
						 
	
                                              <p  class="MsoBodyText" style="line-height: 150%; margin-top:-1; margin-bottom:-1">
                                              
                                         
                                             <br><br>
											 <% if resetPasswordEmailSent=1 then%>
												We have sent you an email, please follow the link in the email to reset your password
											<%else%>
											 
                                              Email    &nbsp; <span style="line-height:150%"><input type="text" name="pemail" size="41"></span><span style="line-height:150%; ">*</span></p>
		
											  
                                              
                                            
                                              
                                              
                                        <br><br>
                                              <p  style="margin-top: -1; margin-bottom: -1">
                                            &nbsp;
											 <input type="submit" value="Submit" onClick="return validateData()" name="button1">
											</p>
                                             
                                             <br><br>

                                            </form>
											
											<%end if %>
											 <% if resetPasswordEmailSent=0 then%>
												Your email has not been registered with us, please click <a href="registrationRetail.asp"> here</a> to join.
												
											<%end if%>
											
											<br>
											<br><br><br><br><br><br><br><br><br><br><br><br><br>
										</td>
										
										<td></td>
	
										
									</tr>
								
	
	
							
						


  

  
  
  

  </table>
					
      
		
		 
      
      		
</td>

<!--end td class pagecontent -->



</tr>
	 
      
</table>
      
        <!--end content contactus -->
    
        
      
      
      
      



<!--#Include file="footerRetail.asp"  --> 
 
</td>

 <!--end mainCenter -->



<td class = "mainright" >    </td>
</tr>
</table>
 
 
 

  
  
  
  
  
  
  </body>
  
  
  
  
  </html>
  	