<%

Set addresses = Server.CreateObject("Scripting.Dictionary")

Set shippingAddr = new Address
shippingAddr.Line1 = ShipAddrLine1
shippingAddr.Line2 = ShipAddrLine2
shippingAddr.City = ShipAddrCity
shippingAddr.State = ShipAddrState
shippingAddr.Zip = ShipAddrZip
shippingAddr.Country = ShipAddrCountry

addresses.Add "ShippingAddress", shippingAddr

If UseShippingAddrForBillingAddr = True Then
  Set billingAddr = shippingAddr
  Set invoiceAddr = shippingAddr
Else
  Set billingAddr = new Address
  billingAddr.Line1 = BillingAddrLine1
  billingAddr.Line2 = BillingAddrLine2
  billingAddr.City = BillingAddrCity
  billingAddr.State = BillingAddrState
  billingAddr.Zip = BillingAddrZip
  billingAddr.Country = BillingAddrCountry

  Set invoiceAddr = new Address
  invoiceAddr.Line1 = BillingAddrLine1
  invoiceAddr.Line2 = BillingAddrLine2
  invoiceAddr.City = BillingAddrCity
  invoiceAddr.State = BillingAddrState
  invoiceAddr.Zip = BillingAddrZip
  invoiceAddr.Country = BillingAddrCountry
End If

addresses.Add "BillingAddress", billingAddr
addresses.Add "InvoiceAddress", invoiceAddr

' RW("FullName: "& FullName & "<br/>" & _
'   "Login: "& Login & "<br/>" & _
'   "UserId: "& UserId & "<br/>" & _
'   "Email: "& Email & "<br/>" & _
'   "Phone: "& Phone & "<br/>" & _
'   "")
' Response.End

Set newCustomer = SetCustomer(FullName, Login, UserId, StripeCustomerId, Email, Phone, addresses)

' PrintCustomer(newCustomer)

Set addresses = Nothing


%>