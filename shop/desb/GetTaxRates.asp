<%@ Language=VBScript %>
<!--#include file="../../include/asp_lib.inc.asp" -->
<!--#include file="../../include/sqlCheckInclude.asp" -->
<!--#include virtual="/shop/payment/classes/aspJSON1.17.asp"-->
<!--#include virtual="/shop/payment/functions/helpers.inc"-->

<!DOCTYPE html>
<html lang="en">
<head>
</head>
<body>

<%

  ' zip = "92415"
  ' city = "San Bardino"
  ' state = "CA"

  Done = False

  If Done = False Then
    sql = "SELECT TOP 100 id, zip, city, state FROM dbo.zipcodes "
    sql = sql & "WHERE active = 1 AND salestax = 0;"
    ' RW(sql)

    Set cnn = Server.CreateObject("ADODB.Connection")
    cnn.ConnectionString = Application.Contents("dbConnStr")
    cnn.Open
    Set rsZip = cnn.Execute(sql)
    If Not rsZip.BOF And Not rsZip.EOF Then
      RW("Running API...")
      Response.Flush
      Do While Not rsZip.EOF
        id = Trim(rsZip("id")&"")
        zip = Trim(rsZip("zip")&"")
        city = Trim(rsZip("city")&"")
        state = Trim(rsZip("state")&"")
        If Len(zip) >= 1 And Len(city) >= 1 And Len(state) >= 1 Then
          result = GetNijaSalesTaxRate(zip, city, state)
          ' RW("zip: " & zip & ", city: " & city & ", state: " & state & ", salestax : " & FormatNumber(result * 100, 2))
          If result > 0 Then
            sql = "UPDATE dbo.zipcodes SET salestax = " & result & " "
            sql = sql & "WHERE id = " & id
            cnn.Execute(sql)
          End If
        End If
        rsZip.MoveNext
      Loop
    Else
      Done = True
    End If
    RW("Done")
    Set rsZip = Nothing
    Set cnn = Nothing
  End If
%>
  <script>
    document.addEventListener("DOMContentLoaded", () => {
      var done = <%=CInt(Done)%>;
      if(!done === true) {
        location.reload();
      }
    });
  </script>  
</body>
</html>