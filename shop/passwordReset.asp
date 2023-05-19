
<%
'--------------------------------------------------------------
'      Coded By: Eric
'       Purpose: 
'   Used Tables: 
'  Invoked From: 
'       Invokes: 
'Included Files: header.htm, footer.htm, animate.js, StyleSheet.css, asp_lib.inc.asp.
'--------------------------------------------------------------
'Updated By   Eric    Date 04/17/2023
'passwod reset
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




Dim rst, cnn, objMail, strMailContent, strMailContent2
Dim Success
dim strlogin , strContact, strEmail, strPass
strlogin = request.querystring("login")
strPass=Trim(Request.Form("password1"))
success=request.querystring("success")
if len(success)> 0 then	
	success=1
else 
	success=0
end if
 
'Get data


If len(strlogin)>0 and len(strPass)>5 Then
    
  dim sqlCmd 
  
  sqlCmd="update consumer set password='" & strPass & "' where login='" & strlogin & "'"
 
  
  'Create connection to execute insert command
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  set rst = Server.CreateObject("adodb.RecordSet")
  
  dim rstMsg
  set rstMsg=  Server.CreateObject("adodb.RecordSet")
  
  cnn.Open
  cnn.Execute SQLCmd
  
  rstMsg.open "select * from screenmessage", cnn, 3
  
  
    'Prepare SQL statement
  strSQLCmd = "select * from  consumer where login='"&trim(strlogin)&"'"
  rst.open strSqlcmd, cnn, 3
  
  
  if not rst.eof then 
	strContact=rst("contact")

  'Send mail to consumer
    
  dim cusType
  
 cusType="Consumer"
 
 
 
  strMailContent = "Hi " & strContact &  ",<br><br>"
  strMailContent = strMailContent & "Your password has been reset. <br><br> "
    
  
  strMailContent = strMailContent & "Login Name: " & strLogin  & "<br>"  
  strMailContent = strMailContent & "Please click on the link below to continue shopping with us<br>"
  
  strMailContent =strMailContent & "https://omhusa.com/consumer/"
  
  
  
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
	.To = rst("email")
	'.cc=rstMsg("ms6")
        .Subject = "Password reset from omhusa.com"
        .HTMLBody = strMailContent
		if len(.to)>0 then
		
		.Send 
		end if
    End With 

	
	
	

 
   
	
    Set cdoMessage = Nothing 

    Set cdoConfig = Nothing 
 
 
cnn.Close
  Set cnn = Nothing

  

    Response.Redirect "passwordReset.asp?success=1&login="&strlogin
  else
  'Response.Redirect "confirmregistrationC.asp?error=1"
  	
  end if
Else	
%>

                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
 <link rel="shortcut icon" type="image/x-icon" href="http://www.omhvn.com/favicon.ico" />
 
  <script language="JavaScript1.2" src="../include/javascript.js"></script>
 <script language="JavaScript1.2">


	
		
function validateData(){
  var password1 = document.passwordReset.password1.value;
  var password2 = document.passwordReset.password2.value;
  

  if (password1!=password2) {
	alert("password confirmation must be the same");
	return false;
 }
  

 //Check if password is valid
  if (!isPassValid(password1)){
  
   //alert("Password length must be between 6-12 characters and without space!");
   document.passwordReset.password1.focus();
  return false;
 }
  
  
 
 
  return true;
}
        </script>
        
		
<%
if isnull(Request.Cookies("screenSize")) or len(trim(Request.Cookies("screenSize")))=0 then

	
	%>
	
	
<script type="text/javascript">
window.onload = function() {
    if(!window.location.hash) {
        window.location = window.location + '#loaded';
        window.location.reload();
    }
}
</script>
<%
		
	
end if
%>

<meta name="viewport" content="width=device-width, initial-scale=0.75">
  
        
        
  </head>
  <body>
 
 
 
 
 
 <table class="fixedTable" >



<tr>
	<td class= "mainleft" >  </td>
	<td class = "maincenter" >   
	
<!--#Include file="HeaderRetail.asp"  -->

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
     Create or Reset your Password
     </th>
     </tr>
  
	
	<tr>
	<td >
	</td>
										<td align="center" >
                                           
                                      <form method="POST" name="passwordReset" action="passwordReset.asp?login=<%=strlogin%>"> 
                                              
                                            <table>
											
                                             
                                         
                                             <tr>
											 
                                             
                                             <td>     
                                                  New Password</td>
												  <td>
												  <input type="password" name="password1" size="25">
											  </td>
											  
											  
											  
                                           
                                              <Tr>
											  <td>
                                             
                                              Confirm password
											  
											  </td>
											  <td>
											  <input type="password" name="password2" size="25">   
                                              
											  </td>
											  
                                              
                                              
                                              
                                              
                                        
											
											
											<tr>
											<td>
											<br>
											<br>
											 <input type="submit" value="Submit" onClick="return validateData()" name="button1">
											
											<br><br>
											<br><br>
											
											
											
											</td>
											</tr>
											<tr>
											
											<td colspan="2">
											
																						<%if success=1 then%>
											
											
												<p  style="margin-top: -1; margin-bottom: -1; font-size: large; "> Thank you, your password has been reset. Please click <a href="loginRetail.asp">here</a> to login </p>
											<% end if%>
											<br><br>
											<br><br>
											<br><br>
											<br><br>
											</td>
											
											
											
											
											

											
                                              <p align="center" style="margin-top: -1; margin-bottom: -1">&nbsp;
                                            </p>
                                         </table>
										 

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
    
        
      
      
      
      



<!--#Include file="FooterRetail.asp"  --> 
 
</td>

 <!--end mainCenter -->



<td class = "mainright" >    </td>
</tr>
</table>
 
 
 

  
  
  
  
  
  
  </body>
  
  
  
  
  </html>
  	<% end if %>