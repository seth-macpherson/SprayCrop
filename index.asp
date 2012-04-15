<%Option Explicit%>
<% 
	DIM conn,rs,sql
    if not session("login") then 
        response.redirect("loginForm.asp") 
    end if 
%> 

<html>
<head>
	<title><%=Application("BusinessName")%>&nbsp;-&nbsp;<%=Application("ProgramName")%></title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
<script language="JavaScript">
<!--
function getLayer(layerName)
{

 return (document.all) ? document.all[layerName].style :
        (document.layers) ? document.layers[layerName] :
        document.getElementById(layerName);
}
function toggleLayer(layerName)
{
 with (getLayer(layerName)){
 	if (document.all)
	  display = (display == "none") ? "block" : "none";
	else
	  style.display = (style.display == "none") ? "block" : "none";
  }
}
function show(layerName)
{
 with (getLayer(layerName)){
 	if (document.all)
	  display =  "block";
	else
	  style.display =  "block";
  }
}
function hide(layerName)
{
 with (getLayer(layerName)){
 	if (document.all)
	  display =  "none";
	else
	  style.display =  "none";
  }
}


//-->
</script>	
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">

<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Welcome</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="ffffff" align="center">
	<tr>
		<td height="14" valign="top" class="content"><br>

            <!--
			<table width="500" align="center" border="0" cellpadding="2" cellspacing="2" bgcolor="#FFFFCC" class="grower-table" align="center">
				<tr>
					<td>
						<p>Don't have Acrobat Reader for viewing PDFs? Get it here <font face="Verdana, Arial, Helvetica, sans-serif" size="-1"></font>. <a href="http://www.adobe.com/products/acrobat/readstep2.html" target="_blank"><img src="./images/getreader.gif" width="88" height="31" border="0"></a></p>
					</td>
				</tr>
				<tr bgcolor="ffffff">
					<td>
						<p>&nbsp;</p>
					</td>
				</tr>
				<tr>
					<td>
						<p>Want to convert Grower Reports to PDF?<br><br>
							Download and install the free CutePDF and the associated 
							GNU Ghostwriter converter software both available at:<br><br>
							<a href="http://www.cutepdf.com/Products/CutePDF/writer.asp">http://www.cutepdf.com/Products/CutePDF/writer.asp</a></p>
					</td>
				</tr>
				<tr bgcolor="ffffff">
					<td>
						<p>&nbsp;</p>
					</td>
				</tr>
				<tr>
					<td>
					<p><a href="SprayProgramInstructions.doc" target="_blank">Click here to download Spray instructions as a word document.</a><br><br>
					<a href="SprayProgramInstructions.pdf" target="_blank">Click here to download Spray instructions in pdf format.</a>
					</p>
					</td>
				</tr>


              </table>
            
            -->
            
        </td>
      </tr>
    
</table>
                <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
            <br /><br>
                
<%	if Request("ServerVar") = "SHOW" then
		call ServerVars()
	end if
	%>
	
<!--#include file="i_adminfooter.asp" -->

</body>
</html> 

<%
Sub OldContent()
%>
          <p><h1>Spray Program Instructions</h1><br><br>



Your username is your Duckwall Pooley Grower Number.  Decide your password with your Duckwall Pooley Field Consultant and they will enter your password into the Spray Program's security settings.  You will
then be ready to enter your spray records. <em><strong>Please protect your password as it protects the security of your spray records.</strong></em><br><br>
If you wish to report your sprays to another packing house, contact your Duckwall Pooley Field Consultant and they’ll set you up in the
system to be able to use the other packing house’s grower number for sprays to crops that aren’t processed at Duckwall Pooley. In order to get the spray record info to the other packing house,
you will need to print a report and give it to the other company.<br><br>

Be sure to use the correct grower number for the
corresponding packing house and crop.
<br><br>
<h2>TO ENTER SPRAY RECORDS</h2><br><br>
<ol>
<li>	From the DuckwallPooley.com website, click on the <strong>Spray Program</strong> link.</li>
<li>	Enter your Username (your Grower Number), and Password (chosen by you and given to your Field Consultant prior to use).</li>
<li>	Click on the <a href="enterspraydata.asp">Enter Spray Data</a> link located in the blue navigation bar on the left-side of your Spray Program webpage (located below the duck).</li>
<li>	Fill in all the information on the form.  Use the Tab Key or click with your mouse to move to the different fields <strong><em>(Do NOT use the BACKSPACE KEY!!  If you use the Backspace key, you will be ejected from the spray record and will lose your data entry)</em></strong>.  
<ol type="a">
<li>NOTE: <a href="#" onMouseOver="show('d1')" onMouseOut="hide('d1')">required fields</a> are marked with an asterisk and the field color is shaded with Duckwall blue.  </li>
<div style="position:absolute">
<div id="d1" style="position: relative; background-color: cccccc; display: none; border: thick solid Black; padding: 0; top: -120px; width:200px; padding-bottom: 5px; padding-left: 5px; padding-right: 5px; padding-top: 5px;" name="d1">
If these fields are not completed, your spray record data will not be posted to the database.  Required fields MUST be filled in.</div></div>
<li>Specific instructions for each field (each piece of information) are located below.</li></ol></li>
<li>Once the form is completed, press the <br>
<img src="images/addsprayrecord2.gif" alt="[Add Spray Record]" width="205" height="35" border="0" align="right"> <br> button at the bottom of the screen.  <strong>This is an important step, as your spray information will not be recorded in the Duckwall Pooley database unless you press the Add Spray Record button.</strong>
The <br>
<img src="images/AddSavesprayrecord.gif" alt="[Add Spray Record and Save Spray Data]" width="372" height="23" border="0" align="right"><br> button not only saves
your record in the database, but it also allows your entered data to remain displayed on the screen, so that you can select a different Grower Number and/or Location and record the same spray data to a different site(s).</li>
<li>From the Enter Spray Data page, you will see up to 20 of your previous entries.  To view more entries, Click on the <a href="grower_report.asp"><strong>Grower Report</strong></a> link located in the blue navigation bar on the left-side of your Spray Program webpage (located below the duck).  </li>
<li>After clicking on the Add Spray Record button or the Add Spray Record and Save Spray Data, you may continue to enter new records by scrolling to the top of the page and beginning the data-entry process again.  Or you may exit the program by clicking on the <strong>Log Out</strong> link in the upper right-hand corner of your screen.  If you do not Log Out, AND you use the <strong>Duckwall-Pooley Website</strong> link located in the blue navigation bar on the left-side of your Spray Program webpage (located below the duck), you will still have access to the Spray Program - and so will anyone else who uses your computer.
  To improve the security of your records, ALWAYS Log Out AND close your <a href="#" onMouseOver="show('d2')" onMouseOut="hide('d2')">web browser</a>.<br><br>
<div style="position:absolute">
<div id="d2" style="position: relative; background-color: cccccc; display: none; border: thick solid Black; padding: 0; top: -120px; width:200px; padding-bottom: 5px; padding-left: 5px; padding-right: 5px; padding-top: 5px;" name="d2">
A web browser is a program you use to use the Internet.  Examples are:  Internet Explorer, Netscape, FireFox, and Opera.</div></div>
</ol><br><bre>
<hr width="100%" size="2">
<h1>SPECIFIC INSTRUCTIONS FOR EACH FIELD</h1><br><br>
<div align="center">
<div align="left" style="background-color: Cccccc; border: thick solid Navy; width:75%;padding-bottom: 5px; padding-left: 5px; padding-right: 5px;">
<strong>NOTE: </strong> Not all fields are <strong>required</strong> to be filled out.  Required fields are marked with an asterisk and the field color is shaded with Duckwall blue - these fields must be filled in or your spray record data entry will not be successful.  <br><br>


Filling in non-required fields provides you with greater information about your spray application, but aren't required for reporting purposes.</div></div><br><br>

<strong>Grower:</strong>  When you log on with your userid and password, you will see only your farm name(s) in this field.  If you have multiple grower numbers, click on the blue down-arrow at the right edge of the field data entry box to see all of your grower accounts.  <br><br>
<img src="images/selectagrower.gif" alt="" width="385" height="49" border="0"><br>
 
Click on your desired grower id to select it.  Press the tab key on the keyboard to move to the next field, or just click in the next field you in which you wish to enter data.<br><br>
<div align="center">
<div align="left" style="background-color: Cccccc; border: thick solid Navy; width:75%;padding-bottom: 5px; padding-left: 5px; padding-right: 5px;">
<strong>NOTE: </strong> Several fields have been programmed to allow you to create your own drop down menu of selection options. <br><br>

Once you enter a New Supervisor, for example, that supervisor's name will be in the database.<br><br>

The next time you enter spray records, you'll be able to click on the drop down arrow button and click on your choice, rather than having to type it each time.<br><br>

The fields that have this drop down menu option programmed in for you are:
<ul>
<li>Supervisor </li>
<li>Supervisor License </li>
<li>Applicator</li>
<li>Applicator License</li>
<li>Chemical Supplier</li>
<li>Recommended by</li>
<br>
To Enter New information, click in the data entry box below the Enter New label, and type in your desired name or description. Concisely name your items using appropriate descriptive names.
<br><br><li>Location (adding locations to your drop down menu is done differently than the previous fields. Directions are included later on this instruction sheet, and also when you click on the Locations link on the blue navigation bar on the left.)</li>
</ul><br><br>

</div></div><br><br>

<strong>Supervisor:</strong>  Enter the name of the individual who is supervising the spray. <br><br>
<strong>Supervisor License:</strong>  Enter the Supervisor's license number. <br><br>
<strong>Applicator:</strong>  Enter the name of the individual who is applying the spray.<br><br>
<strong>Applicator License:</strong>  Enter the Applicator's license number.  This is not a required field because if an applicator is not licensed and they are supervised, the supervisor license number fulfills this requirement.<br><br>
<strong>Chemical Supplier:</strong>  Enter the name of the firm from whom you purchased the chemical you are applying.<br><br>
<strong>Recommended by:</strong>  Enter the name of the individual that recommended this application.<br><br>
<strong>Spray Start Date:</strong>  click on the calendar icon to the right of the data entry box.  Click on the desired date - that will select that date and fill in the box for you.  <br><br>

<div align="center">
<div align="left" style="background-color: Cccccc; border: thick solid Navy; width:75%;padding-bottom: 5px; padding-left: 5px; padding-right: 5px;">
<strong>NOTES regarding calendars: </strong>
<li>Once you see the calendar, if you accidentally double click - or click somewhere off the calendar - it will disappear!  Look in your task bar at the bottom of your screen and you'll see a program button for the calendar.  Click on that button and the calendar will appear again.)</li>
<li>Click on the single arrows to display different months if needed.  Click on the double arrows to display different years.  </li>
</div></div><br><br>
<strong>Spray End Date: </strong> click on the calendar icon to the right of the data entry box.  Click on the desired date - that will select that date and fill in the box for you.
<br><br>
<img src="images/timefinishedspraying.gif" alt="" width="217" height="85" hspace="10" border="0" align="right"><strong>Time Finished Spraying:</strong>  This data is required ONLY if you will be running the Central or Field Posting Reports.  To enter the time, use the hh:mm[AM/PM] format.  For example, 2:30 pm would be entered as 02:30PM.<br><br>
<br><strong>Crop: </strong> Select the correct crop from the drop down box that appears when you click on the down arrow next to the Crop box.  Your screen will refresh, as the database is loading only the varieties and spray products that are relevant for the crop you've selected.<br><br>
<strong>Variety: </strong> click on each appropriate variety's checkbox. You may select 0-4 varieties. <b>[This is an Optional Field]</b><br><br>
<strong>Comments: </strong> Use the comments box if you'd like to record more information about the crop.<br><br>
<strong>Bartlett:</strong>  Please select Yes if this spray was applied to Bartletts.  Please select No if the spray was not applied to Bartletts.
[This field is not repetitive; it helps simplify the number of Crop options, and enables Bartlett-specific reporting for the canners that require that information.]<br><br>
<strong>Harvest Date:</strong>  click on the calendar icon to the right of the data entry box.  Click on the desired date - that will select that date and fill in the box for you.  <strong>[This is an Optional Field]</strong><br><br>
<strong>Method:</strong>  Select from the following options, either with your mouse or by typing the first letter(s) of your desired option: 
<li>Ground
<li>Air
<li>Hand gun, or
<li>Other.  <strong><em>(IF YOU CHOOSE OTHER, PLEASE TYPE YOUR METHOD OF APPLICATION IN THE COMMENTS BOX AT THE BOTTOM OF THE DATA ENTRY PAGE.)</em></strong><br><br>
<strong>Stage:</strong>  Select from the following options, either with your mouse or by typing the first letter(s) of your desired option: <br><br>
<li>Dormant
<li>Delayed Dormant
<li>Pink
<li>Blossom
<li>Petal Fall
<li>Shuck Fall
<li>Cover
<li>Preharvest
<li>Post Harvest<br><br>
<strong>Target:</strong>  Select from the following options, either with your mouse or by typing the first letter(s) of your desired option: 
<li>Insects
<li>Disease
<li>Nutrients
<li>Weeds
<li>Other<strong><em>(IF YOU CHOOSE OTHER, PLEASE TYPE THE TARGET IN THE COMMENTS BOX AT THE BOTTOM OF THE DATA ENTRY PAGE.)</em></strong>
<li>Thinning
<li>Growth Regulator<br><br>
<strong>Product Name and Formulation:</strong>  Click on the blue down-arrow at the right edge of the SELECT A PRODUCT data entry box.  Click on the appropriate Product to select the applied product.  <strong><em>Be sure you match the EPA number from your selection in the database with the EPA number of your applied product.</em></strong><br><br>
<li>You may enter up to six sprays for one application.
<li>If the product you applied is not in this list, you will need to ask your Field Consultant to update this product list.
<li>If you mistakenly enter a product, but have not yet entered the spray record, you may get rid of that product by clicking on the product name down arrow and selecting SELECT A PRODUCT, located at the very top of the product list.  Even though there may be numbers
in that line on your Enter Spray Record page, when you click the button to add the spray record, you’ll see that unwanted
spray was not recorded in any way.<br><br>
<strong>Unit and Max App Use and Max App Seas</strong>   Once you select the Product, the below Unit and Max App Use and Max App Seas are automatically filled in from the database's product information records.<br><br>
<li>Unit:  the units of measure for each specific product.
<li>Max App Use:  the maximum quantity of application per each use.  From the manufacturer's records 
<li>Max App Seas:  the maximum quantity of application per each season.  From the manufacturer's records<br><br> 
<strong>Location:</strong>  Once you have entered a description of a spray block and variety, that description will be in a drop down menu for you to select (either by using your mouse or typing the first letter(s) of that description).<br><br>
<strong>Enter New Location:</strong>  If the location description you desire is not listed in the Location box, click on the Location link located
on the blue navigation bar on the left.  Click in the Add Location box and type in your location description, then click the Insert button with your mouse. Concisely name the location by block and appropriate descriptive name.<br><br>

<strong>Edit or Delete a Location:</strong> To edit a location, click on the Location link
located on the blue navigation bar on the left. Then click on the word “Edit” on the line of the location you
would like to change. Wait for the screen to refresh. You’ll know it is refreshed when your location text is placed in
the Location box and the name of that box has changed from Add Location to Update Location.  Make your desired
edits in the Update Location box and then click Update. The screen will refresh again, and your
newly edited location name will be placed in the location table and will also be
available in your drop down menu on the Enter Spray Data page. <br><br>
To delete a location, click on the Location link located on the blue navigation bar on the left.  Then click on the link “Make InActive.”
This will remove this location from your drop down selection box on your Enter Spray Data page.
<br><br>


<strong>Weather:</strong>  Enter short descriptive weather descriptions for your information [This is an Optional Field].<br><br>
<strong>Acres Treated:</strong>  Enter a number to indicate how many acres to which the spray was applied.  NOTE:  This field will automatically fill in for the next spray you enter during the same data entry session.  ANOTHER NOTE:  if you used the Add Spray Record and Save Spray Data button, you will need to enter the number of Acres Treated, it will not 'stick' from the last data entry.<br><br>
<strong>Rate per Acre: </strong> Enter a number to indicate the rate of application per acre.
Total Material Applied  Once you enter the Acres Treated and the Rate per Acre, the Total Material Applied is automatically calculated here.<br>
OR<br>
Once you enter the Acres Treated you may enter Total Material Applied and the Rate per Acre will automatically be calculated for you.<br><br>
<strong>Comments:</strong>  Please type in additional detail if above fields don't adequately describe the spray.  <strong>[This is an Optional Field]</strong><br><br>

<img src="images/addsprayrecord.gif" alt="" width="205" height="35" border="0" align="left">Click on this button when you have entered all data for the individual spray application record. Your screen will refresh once you click on this button.
If your data entry was complete, a message of Success! will be displayed in red text near the top of your screen.  If some fields were not properly filled out, a message
of Unsuccessful will be displayed along with directions on what to fix.  Once you fulfill the program’s requirements
and click on the Add Spray Record button, the screen will refresh – and you’ll
get either the Success! or Unsuccessful message.  Your data will only be recorded in the database if you receive a
message of Success! <br><br>

NOTE:  This program allows you to enter up to six spray applications at one time.  However, when the spray applications are recorded, each different spray will be recorded in the database as its own individual record.  And each individual record will be displayed separately in the recorded spray records at the bottom of your spray data entry screen and in the printed reports.<br><br>

<img src="images/AddSaveSprayRecord.gif" alt="Add Record and Save" width="372" height="23" border="0" align="left">

Click on this button when you have entered all data for the individual spray application record <b>AND</b> you wish to record the same or
similar sprays applied to a different Grower Number and/or Location. When you click on this button, most of your
entered data will remain displayed on the screen so that you do not have to enter that information again.<br>


To select a different Grower Number, click on the down arrow next to the Grower box <img src="images/selectagrower.gif" alt="Select Grower" width="385" height="49" border="0">, and click on the desired grower number. (If
the number you need is not displayed, contact your Duckwall-Pooley Field Consultant.)  Wait for the screen to refresh.<br>

You will have to enter data in the following fields:<br>

<li>Weather
<li>Location
<li># of Acres Treated
<br>
Once your data entry is done, select either Add Spray Record or the other Add Spray Record and Save Spray Data button and, if you receive
the message of Success!, your records will be saved in the database.<br><br>

<img src="images/cancel.gif" alt="" width="85" height="30" border="0" align="left"> Click on this button to delete your data entry and exit out of the Spray Record Program.<br><br><br>
<strong>To Exit out of the Program</strong><br><br>
Click on the Log Out link at the very upper right-hand corner of any Spray Program page.<br><br>
If you do not Log Out, AND you use the<strong> Duckwall-Pooley Website</strong> link located in the blue navigation bar on the left-side of your Spray Program webpage, you will still have access to the Spray Program - and so will anyone else who uses your computer.  To improve the security of your records, ALWAYS Log Out AND close your web browser.<br><br>
<strong>TO PRINT SPRAY RECORDS</strong>
<ol>
<li>	Click on the Grower Report link located in the blue navigation bar on the left-side of your Spray Program webpage (located below the duck). </li> 
<li>	Fill in all the Search criteria on the form.  Use the Tab Key or click with your mouse to move to the different fields.</li> 
<li>	Click on the Search Now button.</li> 
<li>	All of your records that meet the criteria you entered are displayed below.  You may either view them on your screen, or you may print preformatted reports.</li> 
<li>	ONLY after you click on the Search Now button are your options shown about the variety of reports you can print.</li> 
<li>	TO PRINT THE REPORT you generated, you have two options:  

a) to print the report and include the comments for each record, or 
b) to print the report without the comments.
<ol type="a">
<li>a.	click on the VIEW PRINTABLE
<ol type="i"><li>
i.	select the print option (either using the File/Print commands in the browser's drop down menus or by clicking on your printer icon.)</li></ol></li>
<li>b.	click on the VIEW PRINTABLE with Comments
<ol type="i"><li>
i.	select the print option (either using the File/Print commands in the browser's drop down menus or by clicking on your printer icon.)</li></ol></li> 
<li>	TO PRINT A CENTRAL POSTING REPORT<br>
<ol type="a">
<li>    Follow steps 1-3 listed above in "To Print Spray Records"
<li>	Enter the data in the grey box after you read and fully understand the Grower Report Instructions:  </li>
<img src="images/centralposting.gif" alt="" width="637" height="338" border="0"><br>
<li>	Then click in the grey box on the Step 3: Click Here to View Printable Central Posting Report link.</li>
<li>Your requested report will pop up in a new browser window.  In the new window, select File (in the uppermost menu bar), and then Page Setup.  Change the Orientation to Landscape.</li>
<li>Now you are ready to print the report.  Click on the Print button in your browser's icon bar, or select File and then select Print.</li></ol></li>
</ol><br><br>
<strong>TO REVIEW, EDIT, OR DELETE SPRAY RECORDS</strong><br><br>
<ol>
<li>Click on the Enter Spray Data link located in the blue navigation bar on the left-side of your Spray Program webpage (located below the duck). </li> 
<li>	Scroll down below the data entry area.  You'll see this text:  <br>
Up to last 20 entered records in your login.</li>
<li>	In the top line of each record, there are option links.  If you want to Edit that record, click on the Edit link.  If you want to Delete that record, click on the Delete link. </li> 
<li>If you click on the Edit link, all data from that record will automatically populate the above spray data entry area.  You may make whatever changes to your spray record by clicking on the area you want to change and typing in new data or clicking on different selections from the drop down boxes.</li>
<li>	Once you've finished editing that record, click on the Update button located at the bottom of the data entry area.</li>
</ol>

<img src="images/editsprayrecord.gif" alt="" width="613" height="446" border="0">
</p>



<%
End Sub

Sub ServerVars()
    DIM name
    For Each name In Request.ServerVariables
        Response.Write ("<br><b>" & name & " :</b> " & Request.ServerVariables(name))
    Next
End Sub
%>