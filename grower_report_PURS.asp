<%
Option Explicit
if not session("login") or not listContains("1,2,3", session("accessid")) then
	response.redirect("index.asp")
end if

REM NOTE uses i_SprayRecordPURS.asp!!!
%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayRecordPURS.asp"-->
<!--#include file="i_Crop.asp"-->
<!--#include file="i_Varieties.asp"-->
<!--#include file="i_Growers.asp"-->
<!--#include file="i_Method.asp"-->
<!--#include file="i_SprayList.asp"-->
<!--#include file="i_Stage.asp"-->
<!--#include file="i_Target.asp"-->
<!--#include file="i_Units.asp"-->
<!--#include file="i_Method.asp"-->
<!--#include file="i_SprayYears.asp"-->
<%
'CREATED by LocusInteractive on 08/02/2005
'MODIFIED kmiers for PURS reporting & cleanup code.
Dim errorFound, formError, errorMessage, tempErrorMessage, _
	urlSearch, formSearch, _
	formCropID, urlCropID, formVarietyID, urlVarietyID, formBartlet, urlBartlet
Dim conn, sql, rs, rsSelect, rsCrop, i, j, searchQueryString, searchQueryString2, page, searching, _
	startrow, endrow, recordsPerPage, nextPage, _
	thisGrowerID, thisLocation, thisSprayDate, thisCount, tablestarted, _
	searchGrower, searchHighSprayDate, searchLowSprayDate, _
	searchSprayYear, PURS, sXML, bUploaded

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")
set rsCrop = Server.CreateObject("ADODB.RecordSet")
'Initialize variables
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"

PURS = Request("PURS")
bUploaded = false

formSearch=Request.Form.Item("Search")
urlSearch=Request.QueryString("Search")

if Request.QueryString("Uploaded") = 1 then
	bUploaded = true
end if

'See if ID was passed through URL or FORM
IF urlSearch = "" THEN urlSearch = 0 END IF
IF formSearch = "" THEN formSearch = urlSearch End IF
urlSearch = formSearch

formCropID = Request.Form.Item("CropID")
urlCropID=Request.QueryString("CropID")
'if formCropID = "" then formCropID = urlCropID end if
IF urlCropID = "" THEN urlCropID = 0 END IF
IF formCropID = "" THEN formCropID = urlCropID End IF
urlCropID = formCropID

formVarietyID = Request.Form.Item("VarietyID")
urlVarietyID=Request.QueryString("VarietyID")
'if formVarietyID = "" then formVarietyID = urlVarietyID end if
IF urlVarietyID = "" THEN urlVarietyID = 0 END IF
IF formVarietyID = "" THEN formVarietyID = urlVarietyID End IF
urlVarietyID = formVarietyID

formBartlet = Request.Form.Item("Bartlet")
urlBartlet=Request.QueryString("Bartlet")
'if formBartlet = "" then formBartlet = urlBartlet end if
'IF urlBartlet = "" THEN urlBartlet = "" END IF
IF formBartlet = "" THEN formBartlet = urlBartlet End IF
urlBartlet = formBartlet

searching=Request.Form.Item("searching")
if Request.QueryString("searching") <> "" AND searching = "" THEN 
	searching = Request.QueryString("searching")
END IF
urlSearch = 1
formSearch = 1

searchGrower = Request.Form.Item("searchGrower")
IF Request.QueryString("searchGrower") <> "" AND searchGrower = "" THEN 
	searchGrower = Request.QueryString("searchGrower")
END IF
IF searchGrower = "" THEN
	searchGrower = session("growerID")
END IF

searchHighSprayDate=Request.Form.Item("searchHighSprayDate")
if Request.QueryString("searchHighSprayDate") <> "" AND searchHighSprayDate = "" THEN 
	searchHighSprayDate = Request.QueryString("searchHighSprayDate")
END IF
if not isDate(searchHighSprayDate) then searchHighSprayDate = "" end if

searchLowSprayDate=Request.Form.Item("searchLowSprayDate")
if Request.QueryString("searchLowSprayDate") <> "" AND searchLowSprayDate = "" THEN 
	searchLowSprayDate = Request.QueryString("searchLowSprayDate")
END IF
if not isDate(searchLowSprayDate) then searchLowSprayDate = "" end if

searchSprayYear=Request.Form.Item("searchSprayYear")
if Request.QueryString("searchSprayYear") <> "" AND searchSprayYear = "" THEN 
	searchSprayYear = Request.QueryString("searchSprayYear")
END IF
if searchSprayYear = "" THEN
	sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
	set rs = conn.execute(sql)
	searchSprayYear = rs(0)
END IF

page=Request.Form.Item("page")
if Request.QueryString("page") <> "" AND page = "" THEN 
	page = Request.QueryString("page")
END IF
if Request.QueryString("newPage") <> "" THEN 
	page = Request.QueryString("newPage")
END IF
if page = "" then
	page = 1
end if

searchQueryString = "searchGrower=" & RemoveWhitespace(searchGrower) & _
	"&searchHighSprayDate=" & RemoveWhitespace(searchHighSprayDate) & _
	"&searchLowSprayDate=" & RemoveWhitespace(searchLowSprayDate) & _
	"&page=" & page & "&searching=" & searching & _
	"&searchSprayYear=" & searchSprayYear & _
	"&formCropID=" & formCropID  & "&formVarietyID=" & formVarietyID & "&formBartlet=" & formBartlet
searchQueryString2 = "searchGrower=" & RemoveWhitespace(searchGrower) & _
	"&searchHighSprayDate=" & RemoveWhitespace(searchHighSprayDate) & _
	"&searchLowSprayDate=" & RemoveWhitespace(searchLowSprayDate) & _
	"&page=" & page & "&searching=" & searching & _
	"&searchSprayYear=" & searchSprayYear & _
	"&CropID=" & formCropID  & "&VarietyID=" & formVarietyID & "&formBartlet=" & formBartlet

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("grower_report_PURS.asp?" & searchQueryString)
END IF

	if formSearch and PURS <> 1 then
%>
<html>
<head>
	<title><%=Application("BusinessName")%>&nbsp;-&nbsp;<%=Application("ProgramName")%>&nbsp;-&nbsp;PURS Grower Report</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
	<script language="JavaScript" src="datepicker.js"></script>	
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center>
<tr><td><img src="images/spacer.gif" height="4" width="1" border="0">
<h1> > PURS Grower Report</h1><br><img src="images/spacer.gif" height="4" width="1" border="0">
</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
	<tr>
		<td colspan="3" class="bodytext">

<form action="grower_report_PURS.asp?searching=1" method="post" name="frmsearch">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<table width="100%" border="0" cellpadding="2" cellspacing="0">
							<tr bgcolor="#cccccc">
								<td align="left" class="bodytext" colspan="3"><strong>
SELECT SEARCH CRITERIA</strong> &nbsp;&nbsp; &nbsp;&nbsp;<a href="grower_report_PURS.asp?search=1">Reset Search</a><br>
Hold down the "CTRL" key to multiple select.<br>
* indicates in-active data</td>
							</tr>

							<tr valign=top>
							
							<td>
                            <img src="images/spacer.gif" width="150" height="1"><br><span class="subtitle"><label for="GrowerName">Grower</label>:</span>
                            <%if session("growerid")=0 then %>
                                <%
	                                set rsSelect = GetAllGrowers()
                                %><br />
                                <select name="SearchGrower" size="12" multiple>
                                <%
                                IF not rsSelect.EOF THEN
                                DO WHILE not rsSelect.eof 
                                %>
                                <option value="<%=trim(rsSelect.Fields("GrowerID"))%>" <%if listContains(trim(SearchGrower), trim(rsSelect.Fields("GrowerID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%'=rsSelect.Fields("GrowerNumber")&"&nbsp;&nbsp;&nbsp;"%><%=rsSelect.Fields("GrowerName")%> </option>
                                <%
                                rsSelect.MoveNext
                                LOOP
                                END IF
                                %>
                                </select>
                            <%else%>
                                <%=session("growername")%>
                            <%end if%>
                            </td>
                            

								<td valign="top">
									<span class="subtitle">Beginning Spray Date:</span><br>
									<table border=0 cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<input type="text" value="<%=SearchLowSprayDate%>" name="SearchLowSprayDate" size="10" maxlength="21" class="bodytext">
											</td>
											<td>
												<a href="javascript:show_calendar('frmsearch.SearchLowSprayDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a>
											</td>
										</tr>
									</table>
									<br>
									<span class="subtitle">Ending Spray Date:</span><br>
									<table border=0 cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<input type="text" value="<%=SearchHighSprayDate%>" name="SearchHighSprayDate" size="10" maxlength="21" class="bodytext">
											</td>
											<td>
												<a href="javascript:show_calendar('frmsearch.SearchHighSprayDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a>
											</td>
										</tr>
									</table>
								</td>
								<td valign="top">
									<strong>Spray Year</strong>
									<select name="searchSprayYear" size="4" >
<%
		set rsSelect = GetAllSprayYears()
		IF not rsSelect.EOF THEN
			DO WHILE not rsSelect.eof 
%>
										<option value="<%=trim(rsSelect.Fields("SprayYearID"))%>" <%if listContains(trim(searchSprayYear), trim(rsSelect.Fields("SprayYearID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("SprayYear")%> </option>
<%
				rsSelect.MoveNext
			LOOP
		END IF
%>
									</select>
								</td>
							</tr>
							<tr>
								<td valign="top" colspan="1">
									<table border="1">
										<tr>
											<td nowrap>
												<span class="subtitle"><label for="Bartlet">Bartlett:</label></span> <span class="bodytext"><span class="bodytext"><input type="radio" value="1" name="Bartlet" <% if ListContains(formBartlet,"True") OR ListContains(formBartlet,"1")  THEN %>Checked<% END IF %> style="background-color:beige">YES <input type="radio" value="0" name="Bartlet" <% if ListContains(formBartlet,"False") OR ListContains(formBartlet,"0")  THEN %>Checked<% END IF %> style="background-color:beige">NO&nbsp;<input type="radio" value="" name="Bartlet" <% if formBartlet = ""  THEN %>Checked<% END IF %> style="background-color:beige">Any</span>
												<br>
												<strong>Crops</strong><br>
<%
		set rsSelect = GetActiveCrops()
		i = 0
		IF not rsSelect.EOF THEN
			DO WHILE not rsSelect.eof 
				i = i + 1
%>
												<input type="checkbox" name="CropID" value="<%=rsSelect.Fields("CropID")%>"<%if listContains(formCropID, trim(rsSelect.Fields("CropID"))) then response.write("checked") end if%> style="background-color:beige"><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Crop")%><br>
<%
				set rsCrop = GetActiveVarietiesByCropID(rsSelect.Fields("CropID"))
				IF not rsCrop.EOF THEN
					response.write ("&nbsp;&nbsp;varieties:<br>")
					DO WHILE not rsCrop.eof %>
			&nbsp;&nbsp;<input type="checkbox" name="VarietyID" value="<%=rsCrop.Fields("VarietyID")%>"
<%if listContains(formVarietyID, trim(rsCrop.Fields("VarietyID"))) then 
response.write("checked") 
end if%>>
<%if not rsCrop.Fields("Active") then %>*<%end if%><%=rsCrop.Fields("Variety")%><br>
<%
						rsCrop.MoveNext
					LOOP
				END IF
				i = 0
				rsSelect.MoveNext
			LOOP
		END IF
%>
											</td>
										</tr>
									</table>
								</td>
								<td valign="top" colspan="2" bgcolor="#cccccc">
									<b>Instructions for Uploading Reports to PURS:</b>
									<br><br><b>Prerequisites:</b> Setup your reporting account on the State of Oregon PURS Reporting site.
									<br><a href="http://www.oregon.gov/ODA/PEST/purs_index.shtml" target="_new">OR ODA PURS</a>
									<ol>
										<li>Enter the selection criteria on this page for the usage you wish to report.
										<li>Click on the REPORT NOW button.
										<li>Review the resulting report for accuracy.
										<li>Click here to <a href="grower_report_PURS.asp?PURS=1&<%=searchQueryString2%>" target="OregonPURS_XML">Generate XML Report for these records.</a><br>
											<font size=-1>The XML Report will open in a new browser window.</font>
										<li>Save this XML Document.
											<ol type="A">
												<li>In the new browser window containing the XML document,
												<li>Save the page as a local file,
												<ol type="i">
													<li>Using Firefox: from the Firefox menubar, select "File", then "Save Page As",
													<li>Using Internet Explorer: from the IE menubar, select "File", then "Save As",
													<li>It is NOT recommended that you leave the default file name of "grower_report_PURS.asp"
												</ol>
												<li>Enter a unique file name,
												<li>then Save this file where you can find it for the following steps.
												<li>Close the new browser window containing the XML document.
											</ol>
										<li>Login to the state reporting site:<br>
											<a href="https://purs-reports.oda.state.or.us/cgi-bin/WebObjects/PURSReporter.woa" target="_new">OR State ODA PURS Reporting Site</a>
										<li>Click on Proceed to EDS, then click on Proceed to Upload.
										<li>On the EDS upload page "Browse" and select the file you saved.
										<li>Validate the uploaded file.
										<li>Upon a successful validation, click on Complete EDS Upload.  Once you see your newly posted xml file 
										posted on the PURS EDS Files History screen, return to the Spray Program and press this link
										<a href="grower_report_PURS.asp?Uploaded=1&<%=searchQueryString2%>">Update Spray Records as Reported.</a>
										(This will mark your selected spray data, in the Spray Program database, as reported to PURS, so
										you won't be able to report that data twice.)
										<li>Ta Da!!!
									</ol>
									<br>
									NOTES:<br>
									PURS is not set up to reject duplicate spray record entries. Please be viligant and post each of your xml files 
									only once.  If you mistakenly post a duplicate file, you can remove duplicates by using the Remove button on the 
									EDS Files Upload History page.<br><br>
									Once you've clicked on Update Spray Records, and thus marked them as Reported in the Spray Program, those
									records will appear with a green background when those records match the search criteria in your process
									to report records to PURS.<br><br>
									It is very important you assure your records are accurate before you post them 
									to PURS.  Watch out for this scenario.  You report a spray record to PURS.  Then 
									you edit that record in the Spray Program.  That edit will not be 
									automatically reported to PURS.  You will need to go to the PURS site and edit 
									that record on the PURS site.
								</td>
							</tr>
						</table>
						<br>
						<input type="submit" name="Go" value="--- REPORT NOW ---" class="bodytext">
</form>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br><br>

<%
	else
		if PURS <> 1 then
%>
<table width="500" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<strong>CURRENT SEARCH</strong> <a href="grower_report_PURS.asp?search=1&<%=searchQueryString%>">Edit Search</a>&nbsp;&nbsp;<a href="grower_report_PURS.asp?search=0">Reset Search</a><br>
			<table width="500" border="4" cellpadding="2" cellspacing="0">
				<tr bgcolor="#cccccc">
					<td valign="top">
<%
		if searchGrower <> "" then
			set rs = GetGrowersByID(searchGrower) %>
						<strong>Growers:</strong><br> 
<%
			IF not rs.EOF THEN
				DO WHILE not rs.eof 
					response.write(rs.Fields("GrowerName") & "<br>")
					rs.MoveNext
				LOOP
			END IF
		end if
%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%
		end if
	end if 'end form search

	IF searching THEN

		set rs = GetSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchSprayYear,formCropID,formVarietyID,formBartlet)

		if bUploaded then
			IF not rs.EOF THEN
				DO WHILE not rs.eof
					call PURS_ReportedSprayRecord(rs.Fields("SprayRecordID"))
					rs.MoveNext
				LOOP
				rs.MoveFirst
			END IF
			set rs = GetSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchSprayYear,formCropID,formVarietyID,formBartlet)
		end if

		thisCount = rs.RecordCount
		if PURS <> 1 then

%><br><br>

<strong><%=CStr(thisCount)%> records returned.</strong><br>
<!--
<a href="pu_growerreport.asp?<%=searchQueryString%>" target="printable" class="bodytext">VIEW PRINTABLE</a><br>
<a href="pu_growerreport.asp?<%=searchQueryString%>&showComments=TRUE" target="printable" class="bodytext">VIEW PRINTABLE with Comments</a><br>
-->
<%
			recordsPerPage = 200
			if thisCount > recordsPerPage then
				startrow = (page - 1) * recordsPerPage + 1
				endrow = startrow + recordsPerPage
				if startrow > thisCount then
					startrow = 1
					endrow = thisCount + 1
				end if
				if endrow > thisCount then
					endrow = thisCount + 1
				end if
				response.write (recordsPerPage & " records per page.")
				For i = 0 to thisCount Step recordsPerPage
					nextpage = (i + recordsPerPage)/recordsPerpage
					if trim(page) = trim(nextpage) then
						response.write("  &nbsp;&nbsp;<strong>" & nextpage & "</strong>")
					else
						response.write(" &nbsp;&nbsp;<a href=grower_report_PURS.asp?newPage=" & nextpage & "&"  & searchQuerySTring &  ">" & nextpage & "</a> ")
					end if
				Next	
			else
				startrow = 1
				endrow = thisCount + 1
			end if 

%>
<table border="0">
	<tr>
		<td>
<%
		end if ' PURS <> 1

i = 0
j = 0
thisGrowerID = 0
thisSprayDate = ""
tablestarted = false
IF not rs.EOF THEN
	Dim sGrowerID, sGrowerName, sGrowerNumber, sAddress, sCity, sState, sZipCode, _
		sSupervisor, sLicenseNumber, sFieldman, sSprayStartDate, sSprayEndDate, sLocation, _
		sChemicalSupplier, sSprayRecID, sSprayPURSReport, sReported
	Set sGrowerID = rs.Fields("GrowerID")
	Set sGrowerName = rs.Fields("GrowerName")
	Set sGrowerNumber = rs.Fields("GrowerNumber")
	Set sAddress = rs.Fields("Address")
	Set sCity = rs.Fields("City")
	Set sState = rs.Fields("State")
	Set sZipCode = rs.Fields("ZipCode")
	Set sSupervisor = rs.Fields("Supervisor")
	Set sLicenseNumber = rs.Fields("LicenseNumber")
	Set sFieldman = rs.Fields("Fieldman")
	Set sSprayStartDate = rs.Fields("SprayStartDate")
	Set sSprayEndDate = rs.Fields("SprayEndDate")
	Set sLocation = rs.Fields("Location")
	Set sChemicalSupplier = rs.Fields("ChemicalSupplier")
	Set sSprayRecID = rs.Fields("SprayRecordID")
	Set sSprayPURSReport = rs.Fields("PURS_Report")
	Set sReported = rs.Fields("PURS_Reported")

	if PURS <> 1 then
		Response.Write("<br><a href=""grower_report_PURS.asp?PURS=1&" & searchQueryString2 & """ target=""OregonPURS_XML"">Generate XML Report for these records.</a>")
		Response.Write("<br>The XML Report will open in a new browser window.")
	else
		Response.Clear
		Response.ContentType = "text/xml"
		sXML = "" ' initial XML report
	end if

	DO WHILE not rs.eof
		j = j + 1

		IF PURS <> 1 then
			if j < endrow and j >= startrow THEN
				i = i + 1
				if thisGrowerID <> sGrowerID then
					thisGrowerID = sGrowerID
					thisSprayDate = "999"
					thisLocation = "999"
					if tablestarted then 
						tablestarted=false
						Response.Write("</table><br><br>")
					end if
%>
			<br><br>
			<table>
				<tr>
					<td valign="top">
						<h1><strong><%=sGrowerName%>&nbsp;#<%=sGrowerNumber%></strong>&nbsp;&nbsp;</h1><br>
<%=sAddress%><br>
<%=sCity%> &nbsp;&nbsp;<%=sState%> &nbsp;&nbsp;<%=sZipCode%> &nbsp;&nbsp;<br>
<strong>Applicator/Supervisor</strong> <%=sSupervisor%> &nbsp;&nbsp;<br>
<strong>Supervisor License</strong> <%=sLicenseNumber%> &nbsp;&nbsp;<br>
<strong>Fieldman</strong> <%=sFieldman%> &nbsp;&nbsp;<br>
					</td>
					<td valign="top" align="right">
						<strong>Report Date</strong> <%=now()%>
					</td>
				</tr>
			</table>
<br>
<%
				end if

				if (thisSprayDate = "999" or _
					thisSprayDate <> sSprayStartDate or _
					thisLocation = "999" or _
					thisLocation <> sLocation) then
						thisSprayDate = sSprayStartDate
						thisLocation = sLocation

						if tablestarted then 
							tablestarted=false
							Response.Write("</table><br><br>")
						end if
%>
<h2><strong>Spray Date: </strong><%=sSprayStartDate%><%if sSprayEndDate <> "" then%>
-<%=sSprayEndDate%>
<% end if %> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
<strong>Location:</strong> 
<%=sLocation%> &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
<strong>Chemical Supplier:</strong> 
<%=sChemicalSupplier%> &nbsp;&nbsp;

<table border="1" cellspacing="0" cellpadding="1">
	<tr>
		<td>
			<strong>Product</strong>
		</td>
		<td>
			<strong>Crop</strong>
		</td>
		<td>
			<strong>Bartlett</strong>
		</td>
		<td>
			<strong>Stage</strong>
		</td>
		<td>
			<strong>Weather</strong>
		</td>
		<td>
			<strong>Target</strong>
		</td>
		<td>
			<strong>Harvest</strong>
		</td>
		<td>
			<strong>Method</strong>
		</td>
		<td>
			<strong>#Acres Treated</strong>
		</td>
		<td>
			<strong>Rate/Acre</strong>
		</td>
		<td>
			<strong>Total Mtl Appld</strong>
		</td>
		<td>
			<strong>Units of Product</strong>
		</td>
	</tr>
<%
						tablestarted = true
				end if

	if rs.Fields("PURS_Reported") = true then
'		Response.Write("<tr class=""NOT_PURS_Reported"">")
		Response.Write("<tr bgcolor=""#00FF33"">")
	else
'		Response.Write("<tr class=""PURS_Reported"">")
		Response.Write("<tr bgcolor=""#FFFFCC"">")
	end if
%>

		<td nowrap><%=rs.Fields("Name")%>
<%
if rs.Fields("Applicator") <> "" then
%>
<br><strong>Applicator: </strong> <%=rs.Fields("Applicator")%> <strong>License:</strong> <%=rs.Fields("ApplicatorLicense")%>
<%
end if
if rs.Fields("RecommendedBy") <> "" then%>
<br><strong>Recommended By: </strong> <%=rs.Fields("RecommendedBy")%>
<%
end if
%>
</td><td><%=rs.Fields("Crop")%>
<%
if rs.Fields("Variety") <> "" then
	response.write("<br>varieties:" & rs.Fields("Variety"))
end if
if rs.Fields("Variety2") <> "" then
	response.write("," & rs.Fields("Variety2"))
end if
if rs.Fields("Variety3") <> "" then
	response.write("," & rs.Fields("Variety3"))
end if
if rs.Fields("Variety4") <> "" then
	response.write("," & rs.Fields("Variety4"))
end if%>
</td>
<td><%=rs.Fields("Bartlet")%></td>
<td><%=rs.Fields("Stage")%></td>
<td><%=rs.Fields("Weather")%>&nbsp;</td>
<td><%=rs.Fields("Target")%></td>
<td><%=rs.Fields("HarvestDate")%>&nbsp;</td>
<td><%=rs.Fields("Method")%></td>
<td><%=rs.Fields("AcresTreated")%></td>
<td><%=(int(rs.Fields("RateAcre")*100)/100)%></td>
<td><%=(int(rs.Fields("AcresTreated")*rs.Fields("RateAcre")*100)/100)%></td>
<td><%=rs.Fields("Unit")%></td>
</tr>
<% 
			end if
		END IF

		if PURS = 1 then
			if sSprayPURSReport and not sReported then
			sXML = writeXML(sXML, "", "<UseReport id=""DP_G_" & trim(sGrowerNumber) & "_SR_" & trim(CStr(sSprayRecID)) & """>", "")
'			sXML = writeXML(sXML, "", "<UseReport id=""ACME" & trim(CStr(j)) & """>", "")

			sXML = writeXML(sXML, "<SiteCategory>", "Agriculture", "</SiteCategory>")
			sXML = writeXML(sXML, "<SpecificSite>", "Fruits/Nuts", "</SpecificSite>")
'1/12/2007 changed to use sSprayEndDate instead of sSprayStartDate
			sXML = writeXML(sXML, "<useDate>", Month(sSprayEndDate) & "/" & Day(sSprayEndDate) & "/" & Year(sSprayEndDate), "</useDate>")
			sXML = writeXML(sXML, "<UseLocation><UseLocationWaterBasin><waterBasin>", "Middle Columbia", _
				"</waterBasin></UseLocationWaterBasin></UseLocation>")

			sXML = writeXML(sXML, "<products>", "", "")
			sXML = writeXML(sXML, "<UseReportProduct>", "", "")

			sXML = writeXML(sXML, "<epaProductNumber>", rs.Fields("PURS_EPA_Number"), "</epaProductNumber>")
			sXML = writeXML(sXML, "<productName>", rs.Fields("PURS_Name"), "</productName>")
			sXML = writeXML(sXML, "<purpose>", rs.Fields("PURS_Target"), "</purpose>")
			sXML = writeXML(sXML, "<quantity>", CStr(int(rs.Fields("AcresTreated")*rs.Fields("RateAcre")*100)/100), "</quantity>")
			sXML = writeXML(sXML, "<quantityUnit>", rs.Fields("PURSUnit"), "</quantityUnit>")

			sXML = writeXML(sXML, "</UseReportProduct>", "", "")
			sXML = writeXML(sXML, "</products>", "", "")

			sXML = writeXML(sXML, "", "</UseReport>", "")
			end if
		end if

		rs.MoveNext
	LOOP

	if PURS = 1 then

'		sXML = "&lt;?xml version=""1.0"" encoding=""UTF-8""?&gt;" & vbCrLf & _
'				"&lt;!DOCTYPE Submission SYSTEM ""http://purs-reports.oda.state.or.us/WebObjects/PURSReporter.woa/Contents/WebServerResources/purs.dtd""&gt;" & vbCrLf & _
'				"&lt;Submission version=""1.0""&gt;" & vbCrLf & _
'				sXML & vbCrLf & _
'				"&lt;/Submission&gt;"

		sXML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
				"<!DOCTYPE Submission SYSTEM ""http://purs-reports.oda.state.or.us/WebObjects/PURSReporter.woa/Contents/WebServerResources/purs.dtd"">" & vbCrLf & _
				"<Submission version=""1.0"">" & vbCrLf & _
				sXML & vbCrLf & _
				"</Submission>"

REM KILROY NEED TO CHECK PURS_Report flag of SprayList records.
'		Response.Write("<br>XML === <pre>" & sXML & "</pre> === XML")
'		Response.Write("<pre>" & sXML & "</pre>")
		Response.Write(sXML)
	end if

Else
%>
<tr><td class="bodytext" colspan="27">No Records Selected</td></tr>
<%	end if
	if PURS <> 1 then
%>
</td></tr>
</table>
<%
	end if
	end if
	if PURS <> 1 then
 %>
<!--#include file="i_adminfooter.asp" -->

</td></tr>

</table>
<%
	end if
	set rs = nothing
	EndConnect(conn)

	if PURS <> 1 then

%>
</body>
</html>

<%
	end if

function writeXML(ByRef pXML, ByVal pBTag, ByVal pTagVal, ByVal pETag)
	dim sTemp, bReplace
	bReplace = false
	sTemp = vbCrLf
	if bReplace then
		if trim(pBTag) > "" then
			sTemp = sTemp & Replace(Replace(pBTag,"<","&lt;"), ">", "&gt;")
		end if
		if trim(pTagVal) > "" then
			sTemp = sTemp & Replace(Replace(pTagVal,"<","&lt;"), ">", "&gt;")
		end if
		if trim(pETag) > "" then
			sTemp = sTemp & Replace(Replace(pETag,"<","&lt;"), ">", "&gt;")
		end if
		writeXML = pXML & sTemp
	else
		writeXML = pXML & pBTag & pTagVal & pETag & sTemp
	end if
end function
%>