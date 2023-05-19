
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
'verify email
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
Dim strSuccess
dim strlogin , strContact, strEmail, strPass
strlogin = request.querystring("login")

 
'Get data

strSuccess = Request.QueryString("pSuccess")
If len(strlogin)>0 Then
  strPass="omh$@142132" 
  
  dim sqlCmd 
  
  sqlCmd="update consumer set consumer=1, password='" & strPass & "' , emailVerify=1 where DATEDIFF(day, regdate, getdate())<=1 and login='" & strlogin & "'"
 
  
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
  strMailContent = strMailContent & "Thank you for verifying your email, below is your temporary login information: <br><br> "
    
  
  strMailContent = strMailContent & "Login Name: " & strLogin  & "<br>"  
  strMailContent = strMailContent & "password: " &strPass & "<br>"
  
  
  strMailContent = strMailContent & "Please click on the link below to continue shopping with us<br>"
  
  strMailContent =strMailContent & "https://omhusa.com/cartRetail.asp?login=" & strlogin 
  
  
  
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
        .Subject = "Welcome to OMHUSA"
        .HTMLBody = strMailContent
	'.Send 
    End With 
    

 
   
	
    Set cdoMessage = Nothing 

    Set cdoConfig = Nothing 
 
 
cnn.Close
  Set cnn = Nothing

  

    Response.Redirect "passwordReset.asp?login="&strLogin
	
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

        </script>
        
        
        
  </head>
  <body>
 
 
 
 
 
 <table class="fixedTable" >



<tr>
	<td class= "mainleft" >  </td>
	<td class = "maincenter" >   
	
<!--#Include file="HeaderRetail.asp"  -->
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
    
        
      
      
      
      



<!--#Include file="FooterRetail.asp"  --> 
 
</td>

 <!--end mainCenter -->



<td class = "mainright" >    </td>
</tr>
</table>
 
 
 

  
  
  
  
  
  
  </body>
  
  
  
  
  </html>
  	<% end if %>