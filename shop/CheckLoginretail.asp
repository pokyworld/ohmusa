<%@ Language=VBScript %>
<%option explicit%>
<%

response.Expires=0
response.CacheControl= "no-cache"
response.AddHeader "Pragma", "no-cache"


%>

<!-- #include file="../include/sqlCheckInclude.asp" -->
<!-- #include file="../include/asp_lib.inc.asp" -->
<%
'--------------------------------------------------------------
'      Coded By: Eric Vuong on 02/28/06.
'       Purpose: Check if user login with correct UserName and Password.
'    parameters:
'                pLoginName
'                pPassword
'   Used Tables: company_profile.
'  Invoked From: login.asp.
'       Invokes: 
'Included Files: 
'--------------------------------------------------------------
'Updated By    Eric Vuong       Date 02/18/2009      Comments
'secured login and fixed log in bug
'--------------------------------------------------------------
%>
<%


'***********************************************************
  Function CheckDatabase(strLoginName,strPassword)
    
    dim conStr, Connect, rst, strSQL, strupdateSQL, bolReturn
    dim screenMsgSql
    dim rstMsg
    
    'Open database
    conStr=Application.Contents("dbConnStr")
    Set Connect = Server.CreateObject("ADODB.Connection")
    Connect.Open conStr
    
     screenMsgSql="select * from screenmessage"
     set rstMsg=Server.CreateObject("ADODB.RecordSet")
     rstMsg.open screenMsgSql, Connect, 3
  
    
    'Create SQL string
    set rst=Server.CreateObject("ADODB.RecordSet")
    strSQL = "Select login, password, contact, active from consumer " &_
             "where login='" & strLoginName & "'"
    rst.Open strSQL,Connect,3 
    
    if rst.RecordCount>0 then
    	
      if  not( StrComp(rst("active"), "1")) then
      	'if active 
      'confirm password
    'remove any space
    dim strpass
    strpass= trim(rst("password"))
    strpassword=trim(strpassword)
      
    'if strComp(rst("Password"),strPassword,0)=0 then
	if strComp(strpass,strPassword)=0 then
	'if not strComp(strPassword,rst("Password"),0) then  
	'if strComp(strPassword,rst("Password"),0)=0 then
        
        'password is correct        
        'return true
        bolReturn = TRUE
        'get the contact name of company
        contact=rst("contact")
        session("login")=strLoginName
		Session("consumer")=strLoginName
		
		session("enablePromo")=rstMsg("enablePromo")
		
        'stamp login date
        strupdateSQL="update consumer set recent_login ='"&DateAdd("h",-3,Now)& "'where login='" & strLoginName & "'"
        Connect.execute strupdateSQL
		rst.close()
		
		
		
		
			dim templogin, productId, quantity, ip_address
		templogin=session("templogin")
		
	if len(templogin) > 0 then 
		
		strSQL="select * from shoppingcart where login= '" & templogin & "'"
		rst.Open strSQL,Connect,3
		
		''loop thru all items that are added to shopping cart before login
		'and delete from existing shopping cart from that user
		
		while not rst.eof
			productId=cint(rst("product_id"))
			quantity=cint(rst("quantity"))
			ip_address=rst("ip_address")
			
			strupdateSQL="delete from shoppingcart where login ='" & strLoginName & "' and product_id =" & productId
			
			Connect.execute strupdateSQL
			rst.movenext
		wend
		
		
		strupdateSQL= "update shoppingcart set logged_in=1, login ='" & strLoginName & "' where login ='" & templogin & "' and logged_in=0"
	
		Connect.execute strupdateSQL
		rst.Close
		Connect.Close 
		response.redirect "cartretail.asp"
		
			
	end if
		
		
		
		
		
                
     	 else
        
        'password is incorrect        
        'return false
        response.write("Password is incorrect")
        '    response.Redirect ("login.asp?loginfailed=1")
        bolReturn = FALSE
        
     	 end if
    
   	   else
   	   'login found but account inactive
		response.write(rstMsg("ms9")) 
   
	   end if

    else 
      
      'Not found login name
      response.write ("Login_Name not found")
      bolReturn = FALSE
   end if

	
	
    Connect.Close 
    
    CheckDatabase = bolReturn
  
  End Function
'*****************************************************************  
  
  dim strLoginName,strPassword,strReturn,strPasswordField, contact
    
  
  'Receive data from login page
  strLoginName = fixString(Trim(Request.Form("pLoginName")))
  strPassword  = fixString(Request.Form("pPassword"))
    

	'check cookies
	'session("testcookies")
if len(session("testcookies"))=0 then

      	response.write("Oops, something went wrong!!! Please try to clear the cache, refresh the page and <a href = 'login.asp'>click here</a> to try again")
else     
		
	
	'Compare LoginName and Password with database
	
  
	if len(strLoginName)>0 AND len(strPassword)>0 then
    
		'call check function
		if CheckDatabase(strLoginName,strPassword) then
      
			'login is valid
			'get session
			'Session("wholesaler") = strLoginName
			Session("consumer") = strLoginName
			Session("contact") = contact
			'assgin null value for timesLogin
			Session("TimesLogin") = 0
			
			
			
		
		
		
			
      
			'Return to previous page
			if (len(Session("requestLoginURL")))>0 then
				Response.Redirect (Session("requestLoginURL"))
			else
				Response.Redirect ("productsRetail.asp?t=" & minute(now())&second(now()))

			end if	
          
		end if
	else
		
     	response.write("please go back and enter your username and password")
  end if
end if
%>