<%@ Language=VBScript %>
<!--#include virtual="/include/asp_lib.inc.asp" -->
<!--#include virtual="/include/sqlCheckInclude.asp" -->
<!--#include virtual="/shop/payment/functions/helpers.inc"-->
<%

  Dim strSQLCateCombo, cnn1, strSQLCmd1
  Dim rstCategory

  '********************************************************************************************************************************************************************************************************

  ' need these ASP section for category menu

  ' SQL statement for creating combo box. If name has more than 13 char then add ... as a tail.

  sql = "select Left(Category_Name, 23)+Left('...', Len(Category_Name) - Len(Left(Category_Name, 23))), "
  sql = sql & "Category_ID from Category where status <>'inactive' "
  sql = sql & "order by Category_Name asc "
  strSQLCateCombo = sql

  ' Create connection and query category data.

  sql = "select Category_ID, Category_Name from Category where status <>'inactive' "
  sql = sql & "order by upper(Category_Name) asc "
  strSQLCmd1 = sql

  Set cnn1 = Server.CreateObject("ADODB.Connection")
  cnn1.ConnectionString = Application.Contents("dbConnStr")
  cnn1.Open

  Set rstCategory = Server.CreateObject("ADODB.Recordset")
  rstCategory.Open strSQLCmd1, cnn1, 3

  ' end category menu ASP

  '********************************************************************************************************************************************************************************************************
  '************************************************************************************************************************
%>

<html>
  <head>
    <title>Old-Modern Handicrafts |Layout Template #1</title>
    <link rel="stylesheet" type="text/css" href="include/product_stylesheet.css">

    <script language="JavaScript1.2" src="include/javascript.js"></script>
  </head>

  <body>
    <table class="fixedTable">
      <tr>
        <td class="mainleft">&nbsp;</td>
        <td class="maincenter">
          <!--#Include virtual="/shop/payment/include/headerRetail.asp"  -->
          <table class="mainTable">
            <tr>
<% 
          if not isnull(Request.Cookies("screenSize")) and len(trim(Request.Cookies("screenSize")))>0 then
            if (cint((Request.Cookies("screenSize"))) <600) then 
%>
              <td class="category" hidden="true">
<% 
            else 
%>
              <td class="category">
<% 
            end if '// (cint((Request.Cookies("screenSize"))) <600)
          else 
%>
              <td class="category">
<% 
          end if  '// not isnull(Request.Cookies("screenSize")) and len(trim(Request.Cookies
%>


<% 
              If rstCategory.RecordCount > 0 Then
%>
                <table class="table_outer_border">
                  <tr><th class="thcategoryBGcolor">CATEGORIES</th></tr>
                  <tr><td width="100%" align="left">&nbsp;</td></tr>
                  <tr>
                    <td align="left" class="tdmargin10">
                      <span class="cssLink"><a href="../productsearchRetail.asp?pCategoryID=-1" title="Ship Model - New Products "> <strong>New Products!!!</strong> </a></span>
                    </td>
                  </tr>
<% 
                While Not rstCategory.EOF
                  CategoryID = rstCategory("Category_ID")
                  CategoryName = rstCategory("Category_Name")
%>
                  <tr><td width="100%" align="left">&nbsp;</td></tr>
                  <tr>
                    <td align="left" class="tdmargin10">
                      <span class="cssLink">
                        <a href="../productsearchRetail.asp?pCategoryID=<%=CategoryID%>" title="Ship Model - <%=CategoryName%>"> <%=CategoryName%></a>
                      </span>
                    </td>
                  </tr>
<% 
                  rstCategory.MoveNext 
                Wend 
                rstCategory.Close 
                cnn1.Close 
                Set rstCategory=Nothing 
                Set cnn1=Nothing
%>
                </table>
<%
              End If  '// rstCategory.RecordCount> 0

%>
                <br />
                <table class="table_outer_border">
                  <tr>
                    <th class="thcategoryBGcolor">
                      LINKS</th>
                  </tr>

                  <tr>
                    <td width="100%" align="left">&nbsp;</td>
                  </tr>

                  <tr>

                    <td class="tdmargin10">


                      <p align="center">
                        <a href="productsearchRetail.asp?pCategoryID=-3" title="Items on sale">
                          <img border="0" src="../../images/SALE.jpg"><br />

                        </a>
                      </p>

                      <p align="center">
                        <a href="catalog_r.asp" title="catalog">
                          <img border="0" src="../../images/catalog.JPG"><br />
                        </a>
                      </p>

                    </td>
                  </tr>

                  <tr>
                    <td width="100%" align="left">&nbsp;</td>
                  </tr>

                </table>
              </td>

              <!--end  <td class="category"> -->

              <td class="pageContent">
                <!--start content about us -->
                <!--#Include file="template.html"-->
              </td>
            </tr>
          </table><!--end mainTable-->
          <!--Include virtual="/shop/payment/include/FooterRetail.asp"  -->
        </td>
        <!--end mainCenter-->
        <td class="mainright"> </td>
      </tr>
    </table>
  </body>
</html>
