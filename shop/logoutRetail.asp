<%@ Language=VBScript %>
<%option explicit%>
<%

response.Expires=0
response.CacheControl= "no-cache"
response.AddHeader "Pragma", "no-cache"

'--------------------------------------------------------------
'      
'--------------------------------------------------------------
'Updated By           Date       Comments
'
'--------------------------------------------------------------
%>
<%




'Session("wholesaler")=""
'session("login")=""
Session.Contents.RemoveAll()
Session.Abandon
response.redirect("productsRetail.asp")


%>

Go back to <a href="productsretail.asp">products</a>