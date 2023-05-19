<%@ Language=VBScript %>
<%option explicit%><%
'--------------------------------------------------------------
'      Coded By: Eric
'       Purpose: Display all category and search product form.
'   Used Tables: products
'  Invoked From: productsearch
'       Invokes: order.asp
'Included Files: header.asp, footer.asp, animate.js, StyleSheet.css, asp_lib.inc.asp.
'--------------------------------------------------------------
'Updated By   Eric    Date 01/04/2011   Comments
'Display products details
'--------------------------------------------------------------
%>
<!-- #include file="../include/asp_lib.inc.asp" -->
<!-- #include file="../include/sqlCheckInclude.asp" -->



<% 

Dim strSQLCateCombo, cnn1, strSQLCmd1,strSQL

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

Dim i
' Turn on error Handling
On Error Resume Next

Dim strFirst_Name, strLast_Name, strEmail, strFb_Info, strPhone, strRegion
Dim cnn, rst, objMail, strSQLCmd, strMailContent, ipaddress
Dim strSuccess
  
'Get data
strFirst_Name = fixstring(Trim(Request.Form("pFirst_Name")))
strSuccess = fixstring(trim(Request.QueryString("pSuccess")))
ipaddress= fixstring(trim(Request.ServerVariables("remote_addr")))
		
		
		

If Len(strFirst_Name) > 0 Then

  'Continue getting data
  strRegion = fixstring(Trim(Request.Form("region")))
  strLast_Name = fixstring(Trim(Request.Form("pLast_Name")))
  strPhone= fixstring(Trim(Request.Form("pPhone")))
  strEmail = fixstring(Trim(Request.Form("pEmail")))
  strFb_Info = fixstring(Trim(Request.Form("pFb_Info")))

  'Prepare SQL statement
  strSQLCmd = "Insert Into Feedback (First_Name, Last_Name, Email, Fb_Info, Fb_Date, Rp_Date, ipaddress, Showable) " &_
    "values ('" & strFirst_Name & "', '" & strLast_Name & "', '" & strEmail & "', '" & strFb_Info &_
    "', '" & date() & "', '" & date() & "','" & ipaddress & "', 0)"
  
  'Create connection to execute insert command

  Set cnn = Server.CreateObject("ADODB.Connection")
  dim rstMsg
   set rstMsg=  Server.CreateObject("adodb.RecordSet")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  cnn.Open
  
  rstMsg.open "select * from screenmessage", cnn, 3
  'no longer insert feedback into server Eric 02/15/2013 to prevent attack
  
  
  'cnn.Execute strSQLCmd
  
  
  'Send mail
  strMailContent = "Dear Sales," & "<br><br>"
  strMailContent = strMailContent & "You have received a feed back from "
  strMailContent = strMailContent & strFirst_Name & " " & strLast_Name & "<br>"
  strMailContent = strMailContent & "Phone number: " & strPhone& "<br>"
  strMailContent = strMailContent & "Feedback Content: " & "<br>"
  strMailContent = strMailContent & strFb_Info & "<br><br>"
  strMailContent = strMailContent & "Please contact the customer ASAP. Have a wonderful day!"
  
    
 
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
    dim strSubject, feedbackcontainlink
	

    With cdoMessage 
        Set .Configuration = cdoConfig 
        .From = strEmail

	if not strComp(strRegion, "America") then
	 .To = rstMsg("ms14")
	  strSubject="USA: "
	end if
	if not strComp(strRegion, "Europe") then
  	 .To = rstMsg("ms16")
  	 strSubject="Europe: "
	end if
	if not strComp(strRegion, "Asia") then
  	 .To = rstMsg("ms16")
  	 strSubject="Asia: "
	end if
	
	if not strComp(strRegion, "Others") then
  	 .To = rstMsg("ms16")
	 strSubject="Others: "
	end if


        .Subject = strSubject& " Feedback From: " & strFirst_Name & " " & strLast_Name
        .HTMLBody = strMailContent
	.cc=rstMsg("ms17")
	
	if InStr(.htmlbody, "http")=0 then 
        .Send 
	else
		feedbackcontainlink=1
	end if
		
	
    End With 
    
' Error Handler
If Err.Number <> 0 or feedbackcontainlink=1 Then
   ' Error Occurred / Trap it
	Response.Redirect "contactus_r.asp?pSuccess=false"
   On Error Goto 0  ' But don't let other errors hide!
   ' Code to cope with the error here
End If
On Error Goto 0 ' Reset error handling.




 
    Set cdoMessage = Nothing 

    Set cdoConfig = Nothing 
 
  cnn.Close
  Set cnn = Nothing
 

  
Response.Redirect "contactus_r.asp?pSuccess=true"
Else
  %>
 
                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - View Detail Product</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">
   <script language="JavaScript1.2" src="../include/javascript.js"></script>
 <script language="JavaScript1.2">
function validateData(){
  var strFirst_Name = document.Feedback.pFirst_Name.value;
  var strLast_Name = document.Feedback.pLast_Name.value;
  var strEmail = document.Feedback.pEmail.value;
  var strFb_Info = document.Feedback.pFb_Info.value;
  var strPhone = document.Feedback.pPhone.value;
  
  //Check if First_Name is empty
  if (isBlank(strFirst_Name)){
    alert("Please enter your first name!");
    document.Feedback.pFirst_Name.focus();
    return false;
  }
  
  //Check if Last_Name is empty
  if (isBlank(strLast_Name)){
    alert("Please enter your last name!");
    document.Feedback.pLast_Name.focus();
    return false;
  }

//Check if Email is empty
  if (isBlank(strPhone)){
    alert("Please input your phone number!");
    document.Feedback.pPhone.focus();
    return false;
  }
  //Check if Email is empty
  if (isBlank(strEmail)){
    alert("Please input your email!");
    document.Feedback.pEmail.focus();
    return false;
  }
  
  //Check if Email is valid
  if (! isEmail(strEmail)){
    alert("Invalid email address!");
    document.Feedback.pEmail.focus();
    return false;
  }
  
  //Check if Message is empty
  if (isBlank(strFb_Info)){
    alert("Please input your message!");
    document.Feedback.pFb_Info.focus();
    return false;
  }
  
  return true;
}



</script>

<script src="https://www.google.com/recaptcha/api.js" async defer></script>


<script type="text/javascript">
	document.cookie = "screenSize=" + screen.width;
</script>




  </head>
  <body>
 
 
 
 
 
 <table class="fixedTable" >



<tr>
	<td class= "mainleft" >  </td>
	<td class = "maincenter" >   
	
<!--#Include file="headerRetail.asp"  -->
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
<table class="mainTable" >
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
      <table class="table98border_aboutus" >
	<tr>
      <th class="thfeatured" colspan="3"
      >
     CONTACT US
     </th>
     </tr>
     
<tr>
    <td colspan = "3">
    </td>
</tr>
<tr>
 <td width = "5%" >&nbsp; &nbsp;</td>
<td class="tdverylightBorder_contactus" >
  <table class="tableContactUs"
          >
          <tr>
          <th>
          USA Head Quarter</th></tr>
          <tr>
          <td >
          
              1500 S Milliken Ave Ste. G <br />
              Ontario, CA 91761 USA<br />
              <br />
              (909) 598 2525 Office<br />
              (909) 598 2521 Fax<br />
              (888) 900-1805 Toll Free<br />
              <br />
              Email: <a href="mailto:service@omh1.com">service@omhusa.com</a><br />
              <a href="http://www.omhusa.com">http://www.omhusa.com</a>
			  
			  
			  <br>
			  <BR>
			   
			  
			  </td>
			  
          </tr>
          
           <br>
          <tr>
    <td width="100%" align="center">



	</td>
  </tr>
  
  
          </table>
		  
		  


</td>

 <td width="5%" >&nbsp; &nbsp;<br />
    </td>
</tr>
<tr>
<td colspan="3">
</td>
</tr>
<tr><td colspan="3">
</tr>


							 
<tr>
            <td  >   
   </td>
            
    <td class="tdverylightBorder_contactus"   >CONTACT FORM
    
         
   
      <TABLE >
      	<form name="Feedback" method="post" action="contactus_r.asp">		
		<tr>
			<td class="cssInputLeft" width="20%">Region</td>
			<td class="cssInputLeft" width="80%">
            <select size="1" name="region">
            
            <option selected  value="America">North America</option>
            <option value="Asia">Asia</option>
            <option value="Europe">Europe</option>
            <option value="Others">Others</option>
            
            </select>
            </td>
		</tr>
		<tr>
			<td class="cssInputLeft" width="20%">First Name</td>
			<td class="cssInputleft" width="80%">
			<input type="text" name="pFirst_Name" size="20" maxlength="30">*</td>
		</tr>
		<tr>
			<td class="cssInputLeft">Last Name</td>
			<td class="cssInputleft"><input type="text" name="pLast_Name" size="20" maxlength="30">*</td>
		</tr>
		<tr>
			<td class="cssInputLeft">Phone</td>
			<td class="cssInputleft">
            <input type="text" name="pPhone" size="20" maxlength="30">*</td>
		</tr>
		<tr>
			<td class="cssInputLeft">Email</td>
			<td class="cssInputleft"><input type="text" name="pEmail" size="20" maxlength="40">*</td>
		</tr>
		<tr>
			<td class="cssInputLeft" valign="top">Message</td>
			<td class="cssInputleft">
                <textarea name="pFb_Info" style="height: 100; width:30vw"></textarea>* </br>(No hyperlink please)</td>
		</tr>
	
	
      </TABLE>  
	  
		
		
		<% 
		'Announce if feedback has been sent
      If Len(strSuccess) =4 Then
        Response.Write("<p class=""cssTextLeft"">Thank you for your time. We will get back to you soon!</font></p>")
	elseif Len(strSuccess) > 4 then ' false
		Response.Write("<p class=""cssTextLeft""><font color=""#FF0000"">Send email failed! Please email us at service@omh1.com</font></p>")
      End If
      %>
      
     </td>
  <td  ></td>
   </tr>						






<tr>
<td colspan ="3">
</td></tr>

<tr>
			
			<td colspan="3" align="center" >
			<div class="g-recaptcha" data-sitekey="6LfQvY4UAAAAAHD3bIKUSA3Wiwh_bZAIy42xwvOn" width="20%"></div>
			<br>
			
			  <input type="submit" value="Submit" onClick="return validateData()" name="button1">
			  &nbsp;&nbsp;
			  <input type="button" value="Cancel" name="button1" onClick="javascript:document.location.href='contactus_r.asp'">
			  
			  <p >   
					   
					   <a href="#" onclick="window.open('https://www.sitelock.com/verify.php?site=omhusa.com','SiteLock','width=600,height=600,left=160,top=170');" ><img class="img-responsive" alt="SiteLock" title="SiteLock" src="//shield.sitelock.com/shield/omhusa.com" /></a>
					   </p>
					   
			</td>
			
			
		</tr>
	
		
    </form>

  </table>
					

	  	
		 
      
      		
</td>

<!--end td class pagecontent -->



</tr>
	 
      
</table>
      
        <!--end content contactus -->
    
        
      
      
      
      



<!--#Include file="footerretail.asp"  --> 
 
</td>

 <!--end mainCenter -->



<td class = "mainright" >    </td>
</tr>
</table>
 
 
 

  
  
  
  
  
  
  </body>
  
  
  
  <%end if %>