<%

Function CreateMsgBody(Order)
  strMailContent="OrderId: " &newOrderId  &", Purchase Order: " & po& "<br>"
  shipto=Contactname & "<br/>" & CompanyName & "<br/>" &address & "<br/>" &address2 & "<br/>"& city &_
    	", " & state & " "& zip& " Country: "& country & "<br/>" &_
		phone & "<br/>" &email & "<br/>"
	shipto=fixstring(replace(shipto, "<br/>", "<br>"))
	strMailContent=strMailContent& " Ship To:<br> " & shipto & "<br>"
	strMailContent=strMailContent& "------------------------------------------------------------------<br>"
  strMailContent=strMailContent& "Item&nbsp;&nbsp;&nbsp;&nbsp;Quantity&nbsp;&nbsp;&nbsp;&nbsp;Price&nbsp;&nbsp;&nbsp;&nbsp;Total&nbsp;&nbsp;&nbsp;&nbsp;Product name     "&"<br>"
	strMailContent=strMailContent& "------------------------------------------------------------------<br>"
	rstShoppingCartDS_Template.movefirst()
  while not (rstShoppingCartDS_Template.eof)
		itemPrice=round(min(rstShoppingCartDs_template("map"),rstShoppingCartDs_template("sale_map"))*mapMarkup, 2)
		tfile.WriteLine(rstShoppingCartDS_Template("product_code") & "    " &  rstShoppingCartDS_Template("quantity") & "       " & rstShoppingCartDS_Template("product_Name"))
		strMailContent=strMailContent& rstShoppingCartDS_Template("product_code") & "&nbsp;&nbsp;&nbsp;&nbsp;" &  rstShoppingCartDS_Template("quantity") & "&nbsp;&nbsp;&nbsp;x&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" &_
		itemPrice &	"&nbsp;&nbsp;&nbsp;=&nbsp;&nbsp;&nbsp;$" & itemPrice*rstShoppingCartDs_template("quantity") & "&nbsp;&nbsp;&nbsp;&nbsp;"   & rstShoppingCartDS_Template("product_Name")& "<br>"
		rstShoppingCartDS_Template.movenext()
	Wend
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
  strMailContent=strMailContent& "Gift Message: " & message &"<br>"
	strMailContent=strMailContent& "Order Message: " & ordermessage &"<br>"
	strMailContent=strMailContent&"------------------------------------------------------------------<br>"
		
  CreateMsgBody = ""
End Function

Function CreatePackingSlipFile(Order)
  dim fs,tfile, path
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	path=Server.MapPath("/")&"/tempfile/" & newOrderId&".txt" 
	RW("FilePath: " & path)
	set tfile=fs.CreateTextFile(path, true,false)

	tfile.WriteLine(" " & companyName & " ")
	tfile.WriteLine("order #:" & newOrderID & " PO # " & po& "             order date:" & date())
	tfile.WriteLine("----------------------------------------------------------------")
	tfile.WriteLine("Ship To:")
	
  shipto=Contactname & vbCrlf & CompanyName & vbCrlf &address & vbCrlf &address2 & vbCrlf& city &_
  ", " & state & " "& zip& " Country: "& country & vbCrlf &_
	phone & vbCrlf &email & vbCrlf
		  	
	tfile.WriteLine(shipto)
	tfile.WriteLine("----------------------------------------------------------------")
	tfile.WriteLine("Item    Quantity    Product name                   ")
	tfile.WriteLine
	rstShoppingCartDS_Template.movefirst()

  while not (rstShoppingCartDS_Template.eof)
		itemPrice=round(min(rstShoppingCartDs_template("map"),rstShoppingCartDs_template("sale_map"))*mapMarkup, 2)
		tfile.WriteLine(rstShoppingCartDS_Template("product_code") & "    " &  rstShoppingCartDS_Template("quantity") & "       " & rstShoppingCartDS_Template("product_Name"))
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
	tfile.WriteLine("Gift Message: " & message)
	tfile.WriteLine("----------------------------------------------------------------")
	tfile.WriteLine("Thank you. If you have any questions or concerns regarding")
	tfile.WriteLine("this order " &newOrderId& " please contact us at")
	tfile.WriteLine
	tfile.WriteLine("Phone: " & phone)
	tfile.WriteLine("Email: " & email)
	tfile.WriteLine("----------------------------------------------------------------")
	tfile.close
	
  set tfile=nothing
	set fs=nothing

  CreatePackingSlipFile = ""
End Function

Function SendConfirmation(Order)

  'if (not rstConsumer.eof) then 
    ' strMailContent="OrderId: " &newOrderId  &", Purchase Order: " & po& "<br>"
	
	'create packing slip file
  '  dim fs,tfile, path
	' set fs=Server.CreateObject("Scripting.FileSystemObject")
	' path=Server.MapPath("/")&"/tempfile/" & newOrderId&".txt" 
	' 'response.write(path)
	' set tfile=fs.CreateTextFile(path, true,false)

	' tfile.WriteLine(" " & companyName & " ")
	' tfile.WriteLine("order #:" & newOrderID & " PO # " & po& "             order date:" & date())
	' tfile.WriteLine("----------------------------------------------------------------")
	' tfile.WriteLine("Ship To:")
	
  ' 	shipto=Contactname & vbCrlf & CompanyName & vbCrlf &address & vbCrlf &address2 & vbCrlf& city &_
  '   	", " & state & " "& zip& " Country: "& country & vbCrlf&_
	' 	phone & vbCrlf &email & vbCrlf
		  	

  
	' tfile.WriteLine(shipto)
	' shipto=fixstring(replace(shipto, vbCrlf, "<br>"))
	' strMailContent=strMailContent& " Ship To:<br> " & shipto & "<br>"
	'response.write(shipto)
	' tfile.WriteLine("----------------------------------------------------------------")
	' strMailContent=strMailContent& "------------------------------------------------------------------<br>"
	' tfile.WriteLine("Item    Quantity    Product name                   ")
	' strMailContent=strMailContent& "Item&nbsp;&nbsp;&nbsp;&nbsp;Quantity&nbsp;&nbsp;&nbsp;&nbsp;Price&nbsp;&nbsp;&nbsp;&nbsp;Total&nbsp;&nbsp;&nbsp;&nbsp;Product name     "&"<br>"
	' strMailContent=strMailContent& "------------------------------------------------------------------<br>"


	' tfile.WriteLine
	' rstShoppingCartDS_Template.movefirst()
	
	' while not (rstShoppingCartDS_Template.eof)
	' 	itemPrice=round(min(rstShoppingCartDs_template("map"),rstShoppingCartDs_template("sale_map"))*mapMarkup, 2)
	' 	tfile.WriteLine(rstShoppingCartDS_Template("product_code") & "    " &  rstShoppingCartDS_Template("quantity") & "       " & rstShoppingCartDS_Template("product_Name"))
	' 	strMailContent=strMailContent& rstShoppingCartDS_Template("product_code") & "&nbsp;&nbsp;&nbsp;&nbsp;" &  rstShoppingCartDS_Template("quantity") & "&nbsp;&nbsp;&nbsp;x&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" &_
	' 		itemPrice &	"&nbsp;&nbsp;&nbsp;=&nbsp;&nbsp;&nbsp;$" & itemPrice*rstShoppingCartDs_template("quantity") & "&nbsp;&nbsp;&nbsp;&nbsp;"   & rstShoppingCartDS_Template("product_Name")& "<br>"
	' 	rstShoppingCartDS_Template.movenext()
	' Wend
	' tfile.WriteLine("----------------------------------------------------------------")
	' strMailContent=strMailContent& "------------------------------------------------------------------<br>"
	' strMailContent=strMailContent& "Total &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & cartTotal& "<br>"
	
	
	' dim volumeDiscount
	' if (not isNull(session("volumeDiscount"))) then
	' 	volumeDiscount=session("volumeDiscount")
	' 	session("volumeDiscount") = "" 
	' else
	' 	volumeDiscount=0
	' end if
	
	

	' dim grandTotal
	' dim totalafterVolumeDiscount
	
	' if volumeDiscount>0 then
	' 	strMailContent=strMailContent& "&nbsp;&nbsp;Volume Discount&nbsp;&nbsp; -$" & volumeDiscount& "<br>"
	' 	totalafterVolumeDiscount=cartTotal-volumeDiscount
	' 	strMailContent=strMailContent& "Sub Total &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & totalafterVolumeDiscount& "<br>"
	' 	'grandTotal=shippingCost+totalAfterVolumeDiscount
	' else
	' 	totalafterVolumeDiscount=cartTotal
	' 	'grandTotal=shippingCost+totalAfterVolumeDiscount
		
	' end if
	' dim totalAfterPromoDiscount
	
	' if promoDiscount> 0 then
	' 	totalAfterPromoDiscount=totalAfterVolumeDiscount - promoDiscount
		
	' 	strMailContent=strMailContent& "&nbsp;&nbsp;" & promoName & "&nbsp;&nbsp; -$" & promoDiscount& "<br>"
	' 	strMailContent=strMailContent& "Sub Total &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & totalAfterPromoDiscount& "<br>"
	' 	'grandTotal=grandTotal - promoDiscount
	' else
	' 	totalAfterPromoDiscount=totalAfterVolumeDiscount
		
	' end if
		
	


	
	' 	strMailContent=strMailContent& "&nbsp;&nbsp;Shipping Cost&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & shippingCost& "<br>"
	' 	grandTotal= totalAfterPromoDiscount +shippingCost

	
	' strMailContent=strMailContent& "Grand Total&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$" & grandTotal & "<br>"


	' strMailContent=strMailContent& "------------------------------------------------------------------<br>"
	' 'if not pay via purchase order/credit card on file 
	' if len(pmType)>1 then
	' 	strMailContent=strMailContent& "Payment Method: " &pmType&"<br>"
	' else
	' 	strMailContent=strMailContent& "Payment Method: Purchase order / Credit Card on file <br>"
	' end if

	' strMailContent=strMailContent& "------------------------------------------------------------------<br>"

	' tfile.WriteLine("Gift Message: " & message)

	
	' strMailContent=strMailContent& "Gift Message: " & message &"<br>"
	' strMailContent=strMailContent& "Order Message: " & ordermessage &"<br>"

	' tfile.WriteLine("----------------------------------------------------------------")

	' strMailContent=strMailContent&"------------------------------------------------------------------<br>"
	
	' tfile.WriteLine("Thank you. If you have any questions or concerns regarding")

	' tfile.WriteLine("this order " &newOrderId& " please contact us at")
	' tfile.WriteLine
	' tfile.WriteLine("Phone: " & phone)
	' tfile.WriteLine("Email: " & email)
	' tfile.WriteLine("----------------------------------------------------------------")
	



	' tfile.close
	
  '   set tfile=nothing
	' set fs=nothing
  
  
  


   
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
	

End Function

Function SendEmail(Subject, Body, To, From, CC)

End Function

%>