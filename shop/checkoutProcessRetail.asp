<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="../include/asp_lib.inc.asp" -->
<!-- #include file="../include/sqlCheckInclude.asp" -->


<%
response.Expires=0
response.CacheControl= "no-cache"
response.AddHeader "Pragma", "no-cache"


'--------------------------------------------------------------
'      Coded By: Eric vuong on 03/25 09
'       Purpose: 
'   Used Tables: 
'  Invoked From: 
'       Invokes: 
'Included Files: 
'--------------------------------------------------------------
'Updated By   Eric Vuong  
'Updated by Eric 03/31/2020 changed to dropshiptemplate2013 instead of dstemplate table

'--------------------------------------------------------------
%>
<%
'check if login
if len(Session("consumer")) < 1  then
  ' Session("requestLoginURL") = "checkoutProcessRetail.asp"
  ' Response.Redirect "loginRetail.asp"
end if

function min(x,y)
	if x>y then
		min= y
	else
		min = x
	end if
end function

Dim cnn, rst, rstCount, objMail, objMailtoDealer, strSQLCmd, strMailContent, strMailContentDealer, strSQLCmdMax, rstMax, rstShoppingCartDS_Template, rstConsumer
dim strLogin
dim redirectUrl
dim maxOrderid, maxpaymentId
dim message, ordermessage
dim paymentID, shipto
dim td, hh, smm, pmtype, po, key

strLogin = session("login")
if len(strlogin)<1 then
	strLogin=session("templogin")
end if
if len(strlogin)< 1 then
		response.redirect ("cartRetail.asp")
end if

dim mapMarkup
if len(session("mapMarkup")) > 0 then
	mapMarkup=Cdbl (session("mapMarkup"))
else
	mapMarkup=1
end if


message=FixString(request.form("message"))
ordermessage=FixString(request.form("ordermessage"))

td=fixstring(trim(request.form("CardNumberEncode")))
hh=fixstring(trim(request.form("ExpMon"))&trim(request.form("ExpYear") ))
smm=fixstring(trim(request.form("verifyEncode")))
pmtype=fixstring(trim(request.form("CardType")))
'po=fixstring(trim(request.form("po")))
po=""

key=fixstring(request.form("KEY"))

if len(key) = 0 then
	key=0
end if



if len(request.form("dropship")) > 0 then
	ordermessage="Drop Ship order. " &ordermessage
end if

shipto=fixstring(request.form("name"))
dim contactname, companyname, address, address2, city, state, zip, country, phone, email, email2
contactname=fixstring(request.form("name"))

if len(contactName)>0 then
contactname=fixstring(request.form("name"))
companyname=fixstring(request.form("companyName"))
address=fixstring(request.form("address"))
address2=fixstring(request.form("address2"))
city=fixstring(request.form("city"))
state=fixstring(request.form("state"))
zip=fixstring(request.form("zip"))
country=fixstring(request.form("country"))
phone=fixstring(request.form("phone"))
email2=fixstring(request.form("email"))

else
	shipto="current"

end if
'response.write(shipto)
shipto=fixstring(shipto)

'test data
'paymentId=1001
'Create connection to execute insert command
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.ConnectionString = Application.Contents("dbConnStr")
  set rst = Server.CreateObject("adodb.RecordSet")
  set rstMax = Server.CreateObject("adodb.RecordSet")
  
   dim rstMsg
  set rstMsg=  Server.CreateObject("adodb.RecordSet")
  
  set rstShoppingCartDS_Template = Server.CreateObject("adodb.RecordSet")
  
  set rstCount = Server.CreateObject("adodb.RecordSet")
  cnn.Open
'get a new orderid and paymentId
strSqlCmdMax="select max(orderid) as maxOrderId from orders"
rstMax.open strSqlCmdMax, cnn, 3

rstMsg.open "select * from screenmessage", cnn, 3

  
  
if isNull(rstMax("maxOrderid")) then
	maxOrderId=100100
else
	maxOrderId=rstMax("maxOrderid")

end if
dim NewOrderId
newOrderId=maxOrderId+1

rstMax.close
strSqlCmdMax="select max(paymentID) as maxPaymentID from payment"
rstMax.open strsqlcmdMax, cnn, 3
if isnull(rstMax("maxPaymentID")) then
	maxPaymentId=100100
else
	maxPaymentId=rstMax("maxPaymentid")

end if

paymentId=maxPaymentId+1
dim shippingCost, cartTotal
dim promoDiscount, promoName, promoCode
promoDiscount=0


if (not isNull(session("shippingCost"))) then
	shippingCost=session("shippingCost")
	session("shippingCost") = ""
else
	shippingCost=0
end if

if not isnull(session("promoDiscount")) and len(session("promoDiscount"))>0 then
	promoDiscount = round(Cdbl(session("promoDiscount")), 2)
	session("promoDiscount") = ""
	
end if

if not isnull(session("promoName")) and len(session("promoName"))>0 then
	promoName = session("promoName")
	session("promoName") = ""
	else
	promoName=""
	
end if
if not isnull(session("promoCode")) and len(session("promoCode"))>0 then
	promoCode=session("promoCode")
	session("promoCode") = ""
	else
	promoCode=""
end if

		

if (len(session("cartTotal"))>0) then
	cartTotal = Cdbl(session("cartTotal"))
	session("cartTotal") = ""
	
else
	cartTotal=0
end if

if cartTotal=0 then
	response.redirect "cartRetail.asp"
end if
 strSqlCmd= "select * from consumer where login = '"&strlogin&"'"
  'Create connection to execute insert command
  set rstConsumer = Server.CreateObject("adodb.RecordSet")
  rstConsumer.open strsqlcmd, cnn,3
  
  
'if ship to dealer current address, then get information from the database  
if not rstConsumer.eof then
	  if (strComp(shipto, "current")=0) then
	  
		  
			contactname=rstConsumer("contact")
			companyname=rstConsumer("customer")

			address=rstConsumer("street")
			address2=""
			city=rstConsumer("city")
			state=rstConsumer("state")
			zip=rstConsumer("zip")
			country=rstConsumer("country")
			phone=rstConsumer("phone")
			email=rstConsumer("login")
			
		end if
			
else
		'response.write "not log in"
end if

 if len(email2)>0 then
	email=email2
end if


  

'insert order information into orders table
strSQLCmd="insert into orders (orderid, login, orderdate, contactName, shiptoCompanyName, address1, address2, city, state, country, zip, email, phone,message, ordermessage, paymentID, status, discount, shippingCost, total, promoCode, promoDiscount) values "&_
	"(" & newOrderId &",'"&strLogin&"', "&"getdate()"&",'"&contactName&"', '"&companyName&"','"&address&"','"&address2&"', '"&city&"', '"&state&"','"&country&"', '"&zip&"', '"&email&"','"&phone&"','"&message&"','" &ordermessage&"',"&paymentID&",'Pending', 0," & shippingCost&","&cartTotal& ",'" &promoCode& "'," & promoDiscount & ")"
	
'response.write strSqlcmd




	
rst.open strSqlcmd, cnn,3

 'insert information into orderdetail table
  
strSqlCmd="select product_id, quantity, product_name, product_code, map, sale_map from shoppingcart inner join dropshiptemplate2013 on product_id=itemID where login='"&strLogin&"'"
rstShoppingCartDS_Template.open strSqlcmd, cnn,3
dim itemPrice
itemPrice=0
while not (rstShoppingCartDS_Template.eof)
	if (isNull(rstShoppingCartDS_Template("sale_map"))) then
		itemPrice=rstShoppingCartDS_Template("map") 
	elseif rstShoppingCartDS_Template("sale_map")<=rstShoppingCartDS_Template("map") then
		itemPrice=rstShoppingCartDS_Template("sale_map") 
	else
		itemPrice=rstShoppingCartDS_Template("map") 
	end if
	itemPrice=round(itemprice*mapmarkup, 2)
	

   strSQLCmd="insert into orderdetail(orderid, productid, quantity, price) values (" &newOrderId&","& rstShoppingCartDS_Template("product_id")&","& rstShoppingCartDS_Template("quantity")&","&itemPrice&")"
'response.write strSqlcmd	
 

	cnn.execute (strSqlcmd)
	rstShoppingCartDS_Template.movenext()
Wend
'Empty shopping cart
	strSqlCmd="delete from shoppingcart where login='"&strLogin&"'"

'	response.write(vbCrlF&strSqlCmd)	
    cnn.execute(strSQLCmd)
    session("cartTotal")=0
 
 
 strSQLCmd="insert into payment (paymentid, td, hh, smm, type, po, khoa) values(" &paymentId&",'"&td&"','"&hh&"','"&smm&"','"&pmtype&"','"&po&"',"& key&")" 
 
	'response.write(vbCrlF&strSqlCmd)	
	cnn.execute(strSqlCmd)
	




  'if (not rstConsumer.eof) then 
    strMailContent="OrderId: " &newOrderId  &", Purchase Order: " & po& "<br>"
	
	'create packing slip file
   dim fs,tfile, path
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	path=Server.MapPath("/")&"/tempfile/" & newOrderId&".txt" 
	'response.write(path)
	set tfile=fs.CreateTextFile(path, true,false)

	tfile.WriteLine(" " & companyName & " ")
	tfile.WriteLine("order #:" & newOrderID & " PO # " & po& "             order date:" & date())
	tfile.WriteLine("----------------------------------------------------------------")
	tfile.WriteLine("Ship To:")
	
  	shipto=Contactname & vbCrlf & CompanyName & vbCrlf &address & vbCrlf &address2 & vbCrlf& city &_
    	", " & state & " "& zip& " Country: "& country & vbCrlf&_
		phone & vbCrlf &email & vbCrlf
		  	

  
	tfile.WriteLine(shipto)
	shipto=fixstring(replace(shipto, vbCrlf, "<br>"))
	strMailContent=strMailContent& " Ship To:<br> " & shipto & "<br>"
	'response.write(shipto)
	tfile.WriteLine("----------------------------------------------------------------")
	strMailContent=strMailContent& "------------------------------------------------------------------<br>"
	tfile.WriteLine("Item    Quantity    Product name                   ")
	strMailContent=strMailContent& "Item&nbsp;&nbsp;&nbsp;&nbsp;Quantity&nbsp;&nbsp;&nbsp;&nbsp;Price&nbsp;&nbsp;&nbsp;&nbsp;Total&nbsp;&nbsp;&nbsp;&nbsp;Product name     "&"<br>"
	strMailContent=strMailContent& "------------------------------------------------------------------<br>"


	tfile.WriteLine
	rstShoppingCartDS_Template.movefirst()
	
	while not (rstShoppingCartDS_Template.eof)
		itemPrice=round(min(rstShoppingCartDs_template("map"),rstShoppingCartDs_template("sale_map"))*mapMarkup, 2)
		tfile.WriteLine(rstShoppingCartDS_Template("product_code") & "    " &  rstShoppingCartDS_Template("quantity") & "       " & rstShoppingCartDS_Template("product_Name"))
		strMailContent=strMailContent& rstShoppingCartDS_Template("product_code") & "&nbsp;&nbsp;&nbsp;&nbsp;" &  rstShoppingCartDS_Template("quantity") & "&nbsp;&nbsp;&nbsp;x&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" &_
			itemPrice &	"&nbsp;&nbsp;&nbsp;=&nbsp;&nbsp;&nbsp;$" & itemPrice*rstShoppingCartDs_template("quantity") & "&nbsp;&nbsp;&nbsp;&nbsp;"   & rstShoppingCartDS_Template("product_Name")& "<br>"
		rstShoppingCartDS_Template.movenext()
	Wend
	tfile.WriteLine("----------------------------------------------------------------")
	strMailContent=strMailContent& "------------------------------------------------------------------<br>"
	strMailContent=strMailContent& "Total &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & cartTotal& "<br>"
	
	
	dim volumeDiscount
	if (not isNull(session("volumeDiscount"))) then
		volumeDiscount=session("volumeDiscount")
		session("volumeDiscount") = "" 
	else
		volumeDiscount=0
	end if
	
	

	dim grandTotal
	dim totalafterVolumeDiscount
	
	if volumeDiscount>0 then
		strMailContent=strMailContent& "&nbsp;&nbsp;Volume Discount&nbsp;&nbsp; -$" & volumeDiscount& "<br>"
		totalafterVolumeDiscount=cartTotal-volumeDiscount
		strMailContent=strMailContent& "Sub Total &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & totalafterVolumeDiscount& "<br>"
		'grandTotal=shippingCost+totalAfterVolumeDiscount
	else
		totalafterVolumeDiscount=cartTotal
		'grandTotal=shippingCost+totalAfterVolumeDiscount
		
	end if
	dim totalAfterPromoDiscount
	
	if promoDiscount> 0 then
		totalAfterPromoDiscount=totalAfterVolumeDiscount - promoDiscount
		
		strMailContent=strMailContent& "&nbsp;&nbsp;" & promoName & "&nbsp;&nbsp; -$" & promoDiscount& "<br>"
		strMailContent=strMailContent& "Sub Total &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & totalAfterPromoDiscount& "<br>"
		'grandTotal=grandTotal - promoDiscount
	else
		totalAfterPromoDiscount=totalAfterVolumeDiscount
		
	end if
		
	


	
		strMailContent=strMailContent& "&nbsp;&nbsp;Shipping Cost&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & shippingCost& "<br>"
		grandTotal= totalAfterPromoDiscount +shippingCost

	
	strMailContent=strMailContent& "Grand Total&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & grandTotal & "<br>"


	strMailContent=strMailContent& "------------------------------------------------------------------<br>"
	'if not pay via purchase order/credit card on file 
	if len(pmType)>1 then
		strMailContent=strMailContent& "Payment Method: " &pmType&"<br>"
	else
		strMailContent=strMailContent& "Payment Method: Purchase order / Credit Card on file <br>"
	end if

	strMailContent=strMailContent& "------------------------------------------------------------------<br>"

	tfile.WriteLine("Gift Message: " & message)

	
	strMailContent=strMailContent& "Gift Message: " & message &"<br>"
	strMailContent=strMailContent& "Order Message: " & ordermessage &"<br>"

	tfile.WriteLine("----------------------------------------------------------------")

	strMailContent=strMailContent&"------------------------------------------------------------------<br>"
	
	tfile.WriteLine("Thank you. If you have any questions or concerns regarding")

	tfile.WriteLine("this order " &newOrderId& " please contact us at")
	tfile.WriteLine
	tfile.WriteLine("Phone: " & phone)
	tfile.WriteLine("Email: " & email)
	tfile.WriteLine("----------------------------------------------------------------")
	



	tfile.close
	
    set tfile=nothing
	set fs=nothing
  
  
  


   
   'send email
   'send email using CDO / 10/08/2010 by eric

   dim sch, cdoconfig, cdomessage, cdomessage2
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
 
    With cdoMessage 
        Set .Configuration = cdoConfig 
        .From = rstMsg("ms15")
        '.To = rstMsg("ms15")
		.To="evuong2000@gmail.com"
		.cc="eric@omhusa.com"
        .Subject = "New Order " & newOrderID & " from " & contactname
        .HTMLBody = strMailContent & "Dealer Contact: " & contactName & "    , Phone: "& phone& "<br>" &td&" "&hh&" "&smm&" "&key
	.addAttachment (path)
		if len(.to) > 0 then
			.Send 
	   end if
    End With 
    
    'send confirmation email
    Set cdoMessage2 = CreateObject("CDO.Message") 
    strMailContentDealer =strMailContent & "Thank you for your order with OMH. If you have any questions please contact us at (909) 598 2525"
    With cdoMessage2 
        Set .Configuration = cdoConfig 
        .From = rstMsg("ms15")
        .To = email
        .Subject = "OMH order " & newOrderId& " confirmation"
        .HTMLBody = strMailContentDealer
		if len(.to) > 0 then
			.Send 
		end if 
		
    End With 

 
    Set cdoMessage = Nothing 
    Set cdoMessage2 = Nothing 
    Set cdoConfig = Nothing 
	
	
	






 'send emails using cdonts/no longer support by current server/ER 10/08/2010  

 ' Set objMail = CreateObject("CDONTS.Newmail")
 ' Set objMailtoDealer = CreateObject("CDONTS.Newmail")
  
  'dim strSubject

 
  
'  objMail.To = "orders@omh1.com"
 ' objMailtoDealer.to=rstConsumer("Email")
 ' objMail.Subject="New Order " & newOrderID & " from " & rstConsumer("customer")
  'objMailtoDealer.subject="OMH order " & newOrderId& " confirmation"
  'objMail.From = rstConsumer("Email")
  'objMailtoDealer.from="usa@omh1.com"
  'objMail.BodyFormat = 0 
  'objMail.MailFormat = 0
  
 ' objMailtoDealer.BodyFormat = 0 
  'objMailtoDealer.MailFormat = 0

'  objMail.Body = strMailContent & "Dealer Contact: " & rstConsumer("contact") & "    , Phone: "& rstConsumer("phone")& "<br>" &td&" "&hh&" "&smm&" "&key

'  strMailContentDealer =strMailContent & "Thank you for your order with OMH. If you have any questions please contact us at (909) 598 2525"
 ' objMailtodealer.body=strMailContentDealer
  'objMail.AttachFile (path)
  'objMail.Send
  'objMailtodealer.send
  'Set objMail = Nothing
 ' set objMailtoDealer=nothing
  

'  response.redirect("orderconfirmRetail.asp")
 response.redirect("/shop/payment/stripe.asp?orderid=" & newOrderID)
 




   
%>