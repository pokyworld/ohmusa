
<%@language=vbscript%>
<%option explicit%>
<!--#include file="../Include/asp_lib.inc.asp"-->
<!-- #include file="../include/sqlCheckInclude.asp" -->

<%
'--------------------------------------------------------------
'      Coded By: Eric Vuong on 02/16/2010
'       Purpose: authorization for download template and pictures
'   Used Tables: 
'  Invoked From: drop ship list.asp
'       Invokes: 
'Included Files: 
'--------------------------------------------------------------
'Updated By           Date       Comments
'
'--------------------------------------------------------------
%>
<%


'check if login
'if len(Session("wholesaler")) > 1 then
if len(Session("wholesaler")) < 1  then
  ' Session("requestLoginURL") = "shipping.asp"
   'Response.Redirect "login.asp"
end if


%>



<%

Dim strSQLCateCombo, cnn1, strSQLCmd1

Dim rstCategory
dim defaultItemPerPage
defaultItemPerPage=50
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
  

 
                                     
  <html>
  <head>
  <title>Old-Modern Handicrafts - Ship Building Steps</title>
 <link rel="stylesheet" type="text/css" href="../product_stylesheet.css">


  </head>
  <body>
 
 
 
 
 
 
 
 <table class="fixedTable" >



<tr>
	<td class= "mainleft" >  </td>
	<td class = "maincenter" >   
	
<!--#Include file="headerretail.asp"  -->
    
<table class="mainTable">
  <tr>
   
    
    <td class="pageContent">
    
    
    
  
      
      
               <!--start content about us -->
      <table class="table98border" >
	<tr>
      <th class="thfeatured" colspan = "3" >
        STEPS TO BUILD OMH SHIP MODELS
     </th>
     </tr>
     <tr>
      <td></td>
     </tr>
     
     
     
     <%'***************************************************************************************************%>
     <%'----------------Middle content start--------------%>
 
 	<tr>
 	
  	<td> &nbsp;&nbsp;</td>
    <td >
	
	<% 
	if not isnull(Request.Cookies("screenSize")) or len(trim(Request.Cookies("screenSize")))>0 then
	
	if (cint((Request.Cookies("screenSize"))) >=500) then
			
	
	%>
	
    <table class="table98border ">
    
    
     <tr>
    
    <td colspan  = "2">
    
          <br />
    
          &nbsp; <b>Step 1: Prepare materials.</b> Fine lumbers are seasoned and kiln dried to the appropriate humidity level before they are used.
          <br />
          <br />
        </td>
    
    </tr>
    
    <tr >
   
   <td class="tdTextStep">
                 
            
            Great products are made from great materials. We acquired the best materials from around the globe
            including clear grade western red cedar from Canada, exotic woods from south America and Asia, U.S. made Gorilla glue, 
            US Hexcel fiber glass, and other top of the line materials. <br />
            <br />
            
            Products made of western red cedar are not only beautiful but also free of cracking or warping.
          
           </td> 
    <td class="tdImgBoxStep">
            <img src = "../images/steps/cedar.jpg"  /></td>
    

    </tr>
    
    
    <tr>
    
    <td colspan  = "2">
    
        <br />
    
        <b>&nbsp;
    
   Step 2:</b> <b>Research.</b> Extensive research through original plans and pictures are completed so that we have all correct information to build an authentic model.
        <br />
    </td>
    
    </tr>
    
    <tr >
    <td class="tdImgBoxStep ">
            <img src = "../images/steps/buoc1_plan1.jpg"  />        </td>
    
    <td class="tdImgBoxStep ">
            <img src = "../images/steps/buoc1_plan2.jpg"   />
    </td>
    
    
    </tr>
    
  
    
    
   <tr>
    
    <td colspan  = "2">
    
        <br />
    
        <b>&nbsp;
    
   Step 3:</b> <b>Start building.</b> Build the keel, bulkheads and gunwales. This step is very important to get the model into ship shape.

        <br />

        </td>
    
    </tr>
    
    
      <tr >
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/buoc 3.1.jpg"  />
    </td>
           
    
    <td class="tdImgBoxStep ">
            <img src = "../images/steps/buoc 3.2.jpg"  />
            </td>
    
    
    </tr>
    
    
       <tr>
    
    <td colspan  = "2">
    
        <br />
    
        <b>&nbsp;
    
    Step 4: Plank.</b> Planks are cut and bent to the shape of the hull. Each plank is then glued to the bulkhead carefully.
					
        <br />
					
        </td>
    
    </tr>
    
    
      <tr >
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/buoc 4.0.jpg"  />
    </td>
           
    
    <td class="tdImgBoxStep ">
            <img src = "../images/steps/buoc 4.2.jpg"  />
            </td>
    
    
    </tr>
    
  
    
    <tr >
<td class="tdTextStep">
    
        <br />
    
        <b>Step 5: Plank some more. </b>A second layer of planking is done. In this particular ship, the second layer consists of many small pieces of wood to form an inlay hull. Planking is a time-consuming process but makes our models much more attractive.         <br />
        </td>
    
    <td class="tdImgBoxStep" > 
            <img src = "../images/steps/buoc 5.1.jpg"  />
    </td>
    
           
    </tr>
    
           
    
    
      <tr >
      <td class="tdTextStep">
          <b>
          Step 6: Glue.</b> Glue and epoxy are poured evenly into the inside of the hull. Top quality, wood glue is used adequately for planking and to make sure the hull will not split due to humidity changes.

        </td>
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/buoc 6.1.jpg"  />
    </td>
    
   
        
        
           
    </tr>
    
              <tr>
    
    <td colspan = "2" >
    
        <br />
    
        <b>&nbsp;
    
  Step 7: Sand.</b> The hull is sanded repeatedly so that the surface is as smooth and shiny as fine furniture.

        <br />

        </td>
    
    </tr>
    
    
      <tr >
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/sanding2.jpg"   />
    </td>
            <td class="tdImgBoxStep " > 
            <img src = "../images/steps/buoc 7.2.jpg"  />
    </td>
    </tr>
     
      <tr>
    
    <td colspan  = "2">
    
        <br />
    
        <b>&nbsp;
    
   Step 8: Install the deck.</b> Notice that the deck is laser cut to imitate the actual deck of the original ship.

        <br />

        </td>
    
    </tr>
    
    
      <tr >
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/buoc 8.3.jpg"  />
    </td>
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/buoc 8.2.jpg"  />
    </td>
    </tr>
    
     <tr >
     <td class="tdTextStep">
         <b>Step 9: </b>Build stern details. The ship stern section includes admiral cabin, chart house and other details.         </td>
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/buoc 9.1.jpg"  />
    </td>
    
    
        
           
    </tr>
    
    
    
    
      <tr >
     <td class="tdTextStep">
           <b>
    Step 10: Build bow details.</b> The ship bow section includes the bow sprit, bow sprit yard, beak-head, and other details.
    
        </td>
        
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/buoc 10.1.jpg"  />
    </td>
   
           
    </tr>

       <tr>
    
    <td colspan  = "2">
        <br />
        <b>&nbsp;
    Step 11: Drill. </b>Gun ports are drilled along the sides of the ship.
        <br />
    </td>
    
    </tr>
    
    
    <tr >
    <td class="tdTextStep">
        <br />
        <b>Step 12: Paint.</b> The hull is painted with several coats of clear or solid color. Please examine the model to see that our paint job is done very carefully.
        <br />
    </td>
        <td class="tdImgBoxStep "> 
            <img src = "../images/steps/paint.jpg"  />
    </td>    
    </tr>
    
           <tr>
    
   <td class="tdTextStep "> 
       <b>
       Step 13: Build masts.</b> Masts, including the main mast, foremast and mizzen mast are built. The yard arms and crow’s nests are also added at this step.
   </td>
     <td class="tdImgBoxStep "> 
            <img src = "../images/steps/mast.jpg"  />
    </td>
 
   
    </tr>
    
    
    <tr >
    <td  colspan ="2" > 
        <br />
        <b>&nbsp;
    Step 14: Rigging. </b>This is a very tedious process that takes our craftsmen many hours to complete. 
        <br />

    </td>
    </tr>
    <tr>
    
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/rigging1.jpg"  />
    </td>
    
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/rigging2.jpg"  />
    </td>
    
    
           
    </tr>
    
    <tr>
  <td class="tdTextStep "> 
    
        <br />
    
        <b>Step 15: Build railings.</b>    
        <br />
        </td>
        
   <td class="tdImgBoxStep "> 
            <img src = "../images/steps/railing1.jpg"  />
    </td>
    
    </tr>
    
    
    <tr >
    <td colspan ="2"> 
        <br />
        <b>&nbsp;
Step 16: Build deck details.</b> Other deck details are added: Lanterns, boat davits, ship’s wheel, life boats, deck rooms, belfry, staircase, skylight...
        <br />
    </td>
           
    </tr> 


     <tr >
     <td class="tdTextStep">
    
         <b>Step 17: Add the sails.</b> The sails are all hand stitched with fine details. They are also seasoned to make the sails look and feel like real. 
    </td>

     <td class="tdImgBoxStep " > 
            <img src = "../images/steps/buoc17.jpg"  />
    </td> 
    
   
           
    </tr>
    
    
    

    
     <tr >
      <td class="tdTextStep">
           <br />
           <b>Step 18: Finishing touches.</b> Finish up with brass sculptures and decorative ornaments. Our brass ornaments are done in-house. The ornaments are all casted from metal by hand by our skillful craftsmen.
           <br />
    </td>
    
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/buoc18.jpg"  />
    </td>
           
    </tr>
    


     <tr >
     
     <td class="tdTextStep">
         <b>Step 19: Quality control.</b>  A final quality control process is conducted to make sure our models are historically accurate, detailed, well-built, and attractive. Each OMH model is uniquely identified by a serial number.</td>
    
    
     <td class="tdImgBoxStep "> 
            <img src = "../images/steps/buoc19.jpg"  />
    </td>
     
           
    </tr>
    
    
    
    <tr >
    <td  colspan ="2" class="tdTextStep"> 
            <br />
            <b>
            Step 20: Packaging. </b>In this final important step, the ship is packed in a sturdy wood crate and then put in a nice carton box with cushion for maximum protection. We also perform a drop test to make sure the package is safe in shipping.             <br />
    </td>
    
    
           
    </tr> 
    <tr>
     <td class="tdImgBoxStep " > 
            <img src = "../images/steps/packaging1.jpg"  />
    </td>
    
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/packaging2.jpg"  />
    </td>
    
    </tr>
    

 <tr >
<td class="tdTextStep">
     <br />
OMH ship models come with a certificate of authenticity signed by our master builder.
       </td>
    
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/certificate.jpg"  />
    </td>
           
    </tr>
    
    
    



















    
    </table>
	
	
	<% 
	
	
	
	else
	
	'----------------------------------------------------------------------------------------------------------------------
	'html for mobile here
	%>
	 <table class="table98border ">
    
    
     <tr>
    
    <td colspan  = "1">
    
          <br />
    
          &nbsp; <b>Step 1: Prepare materials.</b> Fine lumbers are seasoned and kiln dried to the appropriate humidity level before they are used.
          <br />
          <br />
        </td>
    
    </tr>
    
    <tr >
   
   <td class="tdTextStep">
                 
            
            Great products are made from great materials. We acquired the best materials from around the globe
            including clear grade western red cedar from Canada, exotic woods from south America and Asia, U.S. made Gorilla glue, 
            US Hexcel fiber glass, and other top of the line materials. <br />
            <br />
            
            Products made of western red cedar are not only beautiful but also free of cracking or warping.
          
      
   <br>
   
            <img src = "../images/steps/cedar.jpg"  /></td>
    

    </tr>
    
    
    <tr>
    
    <td colspan  = "1">
    
        <br />
    
        <b>&nbsp;
    
   Step 2:</b> <b>Research.</b> Extensive research through original plans and pictures are completed so that we have all correct information to build an authentic model.
        <br />
    </td>
    
    </tr>
    
    <tr >
    <td class="tdImgBoxStep ">
            <img src = "../images/steps/buoc1_plan1.jpg"  />      <br>
			
            <img src = "../images/steps/buoc1_plan2.jpg"   />
    </td>
    
    
    </tr>
    
  
    
    
   <tr>
    
    <td colspan  = "1">
    
        <br />
    
        <b>&nbsp;
    
   Step 3:</b> <b>Start building.</b> Build the keel, bulkheads and gunwales. This step is very important to get the model into ship shape.

        <br />

        </td>
    
    </tr>
    
    
      <tr >
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/buoc 3.1.jpg"  />
   <br>
            <img src = "../images/steps/buoc 3.2.jpg"  />
            </td>
    
    
    </tr>
    
    
       <tr>
    
    <td colspan  = "1">
    
        <br />
    
        <b>&nbsp;
    
    Step 4: Plank.</b> Planks are cut and bent to the shape of the hull. Each plank is then glued to the bulkhead carefully.
					
        <br />
					
        </td>
    
    </tr>
    
    
      <tr >
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/buoc 4.0.jpg"  />
    <br>
            <img src = "../images/steps/buoc 4.2.jpg"  />
            </td>
    
    
    </tr>
    
  
    
    <tr >
<td class="tdTextStep">
    
        <br />
    
        <b>Step 5: Plank some more. </b>A second layer of planking is done. In this particular ship, the second layer consists of many small pieces of wood to form an inlay hull. Planking is a time-consuming process but makes our models much more attractive.         <br />
        <br>
            <img src = "../images/steps/buoc 5.1.jpg"  />
    </td>
    
           
    </tr>
    
           
    
    
      <tr >
      <td class="tdTextStep">
          <b>
          Step 6: Glue.</b> Glue and epoxy are poured evenly into the inside of the hull. Top quality, wood glue is used adequately for planking and to make sure the hull will not split due to humidity changes.
<br>
            <img src = "../images/steps/buoc 6.1.jpg"  />
    </td>
    
   
        
        
           
    </tr>
    
              <tr>
    
    <td colspan = "1" >
    
        <br />
    
        <b>&nbsp;
    
  Step 7: Sand.</b> The hull is sanded repeatedly so that the surface is as smooth and shiny as fine furniture.

        <br />

        </td>
    
    </tr>
    
    
      <tr >
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/sanding2.jpg"   />
   <br>
            <img src = "../images/steps/buoc 7.2.jpg"  />
    </td>
    </tr>
     
      <tr>
    
    <td colspan  = "1">
    
        <br />
    
        <b>&nbsp;
    
   Step 8: Install the deck.</b> Notice that the deck is laser cut to imitate the actual deck of the original ship.

        <br />

        </td>
    
    </tr>
    
    
      <tr >
    <td class="tdImgBoxStep " > 
            <img src = "../images/steps/buoc 8.3.jpg"  />
    <br>
            <img src = "../images/steps/buoc 8.2.jpg"  />
    </td>
    </tr>
    
     <tr >
     <td class="tdTextStep">
         <b>Step 9: </b>Build stern details. The ship stern section includes admiral cabin, chart house and other details.  <br>
            <img src = "../images/steps/buoc 9.1.jpg"  />
    </td>
    
    
        
           
    </tr>
    
    
    
    
      <tr >
     <td class="tdTextStep">
           <b>
    Step 10: Build bow details.</b> The ship bow section includes the bow sprit, bow sprit yard, beak-head, and other details.
    
       <br>
            <img src = "../images/steps/buoc 10.1.jpg"  />
    </td>
   
           
    </tr>

       <tr>
    
    <td colspan  = "1">
        <br />
        <b>&nbsp;
    Step 11: Drill. </b>Gun ports are drilled along the sides of the ship.
        <br />
    </td>
    
    </tr>
    
    
    <tr >
    <td class="tdTextStep">
        <br />
        <b>Step 12: Paint.</b> The hull is painted with several coats of clear or solid color. Please examine the model to see that our paint job is done very carefully.
        <br />
   <br>
            <img src = "../images/steps/paint.jpg"  />
    </td>    
    </tr>
    
           <tr>
    
   <td class="tdTextStep "> 
       <b>
       Step 13: Build masts.</b> Masts, including the main mast, foremast and mizzen mast are built. The yard arms and crow’s nests are also added at this step.
   <br>
            <img src = "../images/steps/mast.jpg"  />
    </td>
 
   
    </tr>
    
    
    <tr >
    <td  colspan ="2" > 
        <br />
        <b>&nbsp;
    Step 14: Rigging. </b>This is a very tedious process that takes our craftsmen many hours to complete. 
        <br />

    </td>
    </tr>
    <tr>
    
    <td class="tdImgBoxStep "> 
            <img src = "../images/steps/rigging1.jpg"  />
   <br> 
            <img src = "../images/steps/rigging2.jpg"  />
    </td>
    
    
           
    </tr>
    
    <tr>
  <td class="tdTextStep "> 
    
        <br />
    
        <b>Step 15: Build railings.</b>    
        <br />
     <br> 
            <img src = "../images/steps/railing1.jpg"  />
    </td>
    
    </tr>
    
    
    <tr >
    <td colspan ="2"> 
        <br />
        <b>&nbsp;
Step 16: Build deck details.</b> Other deck details are added: Lanterns, boat davits, ship’s wheel, life boats, deck rooms, belfry, staircase, skylight...
        <br />
    </td>
           
    </tr> 


     <tr >
     <td class="tdTextStep">
    
         <b>Step 17: Add the sails.</b> The sails are all hand stitched with fine details. They are also seasoned to make the sails look and feel like real. 
  <br> 
            <img src = "../images/steps/buoc17.jpg"  />
    </td> 
    
   
           
    </tr>
    
    
    

    
     <tr >
      <td class="tdTextStep">
           <br />
           <b>Step 18: Finishing touches.</b> Finish up with brass sculptures and decorative ornaments. Our brass ornaments are done in-house. The ornaments are all casted from metal by hand by our skillful craftsmen.
           <br />
   <br>
            <img src = "../images/steps/buoc18.jpg"  />
    </td>
           
    </tr>
    


     <tr >
     
     <td class="tdTextStep">
         <b>Step 19: Quality control.</b>  A final quality control process is conducted to make sure our models are historically accurate, detailed, well-built, and attractive. Each OMH model is uniquely identified by a serial number.
		 <br>
            <img src = "../images/steps/buoc19.jpg"  />
    </td>
     
           
    </tr>
    
    
    
    <tr >
    <td  colspan ="1" class="tdTextStep"> 
            <br />
            <b>
            Step 20: Packaging. </b>In this final important step, the ship is packed in a sturdy wood crate and then put in a nice carton box with cushion for maximum protection. We also perform a drop test to make sure the package is safe in shipping.             <br />
    </td>
    
    
           
    </tr> 
    <tr>
     <td class="tdImgBoxStep " > 
            <img src = "../images/steps/packaging1.jpg"  />
    <br>
            <img src = "../images/steps/packaging2.jpg"  />
    </td>
    
    </tr>
    

 <tr >
<td class="tdTextStep">
     <br />
OMH ship models come with a certificate of authenticity signed by our master builder.
       <br> 
            <img src = "../images/steps/certificate.jpg"  />
    </td>
           
    </tr>
    
    
    



















    
    </table>
	
	

    <%end if
	
end if%>      
    </td>
    <td>&nbsp;
    </td>
 	</tr>
 	
 
 	
  	
   <% '----------------Middle content end--------------
  '***************************************************************************************************%> 

   
	
		<tr>
		<td colspan ="3">&nbsp; </td></tr>	
		</table><!--end table98 --> 
    
       </td>
       </tr>
       </table>
        <!--end mainTable-->
        
        <!--#Include file="Footerretail.asp"  --> 
        
        </td>
        <!--end mainCenter-->
        


<td class = "mainright" >    </td>
</tr>

</table>
 
 
 

  
  
  
  
  
  
  </body>
  
  
  
  
  

  
  
  
  
  
  
  
