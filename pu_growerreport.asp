<%Option Explicit%>
<%if not session("login") or not listContains("1,2,3", session("accessid")) then
	response.redirect("index.asp")
end if%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayRecord.asp"-->

<%
'CREATED by LocusInteractive on 08/02/2005
Dim conn, sql, rs, rsSelect, counter, searchQueryString, page, searching, thisLocation

dim showComments, searchGrower, searchHighSprayDate, searchHighHarvestDate, _
	searchLowSprayDate, searchLowHarvestDate, searchSprayYear, _
	formCropID, formVarietyID, formBartlet, lastComment

showComments = Request.QueryString("showComments")
if showComments = "" then
	showComments = FALSE
END IF
formBartlet = Request.QueryString("formBartlet")


	IF Request.QueryString("searchGrower") THEN 
		searchGrower = Request.QueryString("searchGrower")
	END IF
	IF searchGrower = "" THEN
		searchGrower = session("growerID")
	END IF

searchSprayYear=Request.Form.Item("searchSprayYear")
if Request.QueryString("searchSprayYear") <> "" AND searchSprayYear = "" THEN 
	searchSprayYear = Request.QueryString("searchSprayYear")
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

searchHighHarvestDate=Request.Form.Item("searchHighHarvestDate")
if Request.QueryString("searchHighHarvestDate") <> "" AND searchHighHarvestDate = "" THEN 
	searchHighHarvestDate = Request.QueryString("searchHighHarvestDate")
END IF
if not isDate(searchHighHarvestDate) then searchHighHarvestDate = "" end if

searchLowHarvestDate=Request.Form.Item("searchLowHarvestDate")
if Request.QueryString("searchLowHarvestDate") <> "" AND searchLowHarvestDate = "" THEN 
	searchLowHarvestDate = Request.QueryString("searchLowHarvestDate")
END IF
if not isDate(searchLowHarvestDate) then searchLowHarvestDate = "" end if


searchQueryString = "searchGrower=" & RemoveWhitespace(searchGrower)  & "&searchHighSprayDate=" & RemoveWhitespace(searchHighSprayDate) & "&searchHighHarvestDate=" & RemoveWhitespace(searchHighHarvestDate) & "&searchLowSprayDate=" & RemoveWhitespace(searchLowSprayDate) & "&searchLowHarvestDate=" & RemoveWhitespace(searchLowHarvestDate) & ",searchSprayYear=" & searchSprayYear 

'Initialize Form Fields


'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")




%>
<html>
<head>
	<title><%=Application("BusinessName")%>&nbsp;-&nbsp;<%=Application("ProgramName")%>&nbsp;-&nbsp;Grower Report</title>

	<style type="text/css">
body {
	background: white;
	}
td { 
	page-break-inside: avoid; 
	page-break-before: avoid; 
	page-break-after: avoid; 
	font-family : Arial, Sans-serif; 
	font-size : 11px; 
	} 


tr { 
	page-break-inside: avoid; 
	page-break-before: avoid; 
	page-break-after: avoid; 
	} 
	
table { 
	page-break-inside: avoid; 
	page-break-before: avoid; 
	page-break-after: avoid; 
	} 

	</style>
</head>

<body bgcolor="#ffffff" leftmargin="10" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0"  onload="if (window.focus)self.focus();">

<%	
	formCropID = Request.QueryString("CropID")
	formVarietyID = Request.QueryString("formVarietyID")


	set rs = GetCountSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate,searchSprayYear,formCropID,formVarietyID,formBartlet)
	dim thisCount
	 thisCount = rs(0)

	set rs = GetSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate,searchSprayYear,formCropID,formVarietyID,formBartlet)
	Dim i
	i = 0

dim j,thisGrowerID,thisSprayDate,tablestarted
i = 0
j = 0
thisGrowerID = 0
thisSprayDate = ""
tablestarted = false

IF not rs.EOF THEN
	DO WHILE not rs.eof 
		j = j + 1
		i = i + 1

		if thisGrowerID <> rs.Fields("GrowerID") then
			thisGrowerID = rs.Fields("GrowerID")
			thisSprayDate = ""

			if tablestarted then 
				tablestarted=false
				i=1
%>
</table></td></tr></table><p style="page-break-before: always">

<%
			end if
%>

<table width="800" border="0"><tr><td valign="top" colspan="2"><!--img src="images/printheader.jpg" alt="" width="650" height="100" border="0"--></td></tr>
<tr><td valign="top">

<strong><%=rs.Fields("GrowerName")%><!--&nbsp;#<%=rs.Fields("GrowerNumber")%>--></strong>&nbsp;&nbsp;<br>
<%=rs.Fields("Address")%><br>
<%=rs.Fields("City")%> &nbsp;&nbsp;<%=rs.Fields("State")%> &nbsp;&nbsp;<%=rs.Fields("ZipCode")%> &nbsp;&nbsp;<br>
<strong>Supervisor</strong> <%=rs.Fields("Supervisor")%> &nbsp;&nbsp;<%=rs.Fields("LicenseNumber")%><br>
<strong>Fieldman</strong> <%=rs.Fields("Fieldman")%> &nbsp;&nbsp;<br>
</td><td valign="bottom" align="right"><!--Report Date:<strong> <%=now()%></strong>--><br>Spray Year:
<%
			sql = "SELECT SprayYear FROM SprayYears WHERE SprayYearID = " & rs.Fields("SprayYearID")
			set rsSelect = conn.execute(sql)
%>
 <strong><%=rsSelect(0)%></td></tr></table>
 <table width="800" border="1" cellspacing="0" cellpadding="3" style="page-break-inside: avoid; page-break-before: avoid; page-break-after: avoid;">
<tr><td align="center"><strong>Product/EPA #</strong></td>
<td align="center"><strong><span style="layout-flow: vertical-ideographic;">Crop</strong></td>
<td><strong><span style="layout-flow: vertical-ideographic;">Bartlett</strong></td>
<td align="center">
<strong><span style="layout-flow: vertical-ideographic;">Stage</strong></td>
<td align="center">
<strong><span style="layout-flow: vertical-ideographic;">Weather</span></strong></td>
<td align="center">
<strong><span style="layout-flow: vertical-ideographic;">Target</strong></td>
<td align="center">
<strong><span style="layout-flow: vertical-ideographic;">Harvest<br>Date</strong></td>
<td align="center">
<strong><span style="layout-flow: vertical-ideographic;">Method</strong></td>
<td align="center">
<strong><span style="layout-flow: vertical-ideographic;">#Acres<br>Treated</span></strong></td>
<td align="center">
<strong><span style="layout-flow: vertical-ideographic;">Rate/Acre</span></strong></td>
<td nowrap align="center">
<strong><span style="layout-flow: vertical-ideographic;">Total Mat.<br>Applied</span></strong></td>
<td align="center">
<strong><span style="layout-flow: vertical-ideographic;">Units</span></strong></td></tr>
<%
			tablestarted = true
		end if

		if thisSprayDate <> rs.Fields("SprayStartDate") or  thisLocation <> rs.Fields("Location") then
			thisSprayDate = rs.Fields("SprayStartDate")
			thisLocation = rs.Fields("Location")
%>
<tr><td colspan="13"><strong>Spray Date: <%=rs.Fields("SprayStartDate")%>
<%			if rs.Fields("SprayEndDate") <> "" then
				Response.Write("-" & rs.Fields("SprayEndDate"))
			end if%> 
			 : <%=rs.Fields("TimeFinishedSpraying")%> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
<br><strong>Location: </strong> <%=rs.Fields("Location")%>
<%			if rs.Fields("Applicator") <> "" then%>
<br><strong>Applicator: </strong> <%=rs.Fields("Applicator")%> &nbsp;#<%=rs.Fields("ApplicatorLicense")%>
<%			end if
			if rs.Fields("ChemicalSupplier") <> "" then%>
<br><strong>Chemical Supplier: </strong> <%=rs.Fields("ChemicalSupplier")%> &nbsp;&nbsp;&nbsp;&nbsp;
<%			end if
			if rs.Fields("RecommendedBy") <> "" then%>
<br><strong>Recommended By: </strong> <%=rs.Fields("RecommendedBy")%> &nbsp;&nbsp;<br>
<%			end if%>
</td></tr>
<%		end if %>

<tr>
<td>
<%=rs.Fields("Name")%><br>
<%=rs.Fields("ActiveInd")%><br>
REI: <%=rs.Fields("REI")%><br>
PHI: <%=rs.Fields("PHI")%>

</td>
<td><%=rs.Fields("Crop")%>
<%		if rs.Fields("Variety") <> "" then
			response.write(" - " & rs.Fields("Variety"))
		end if
		if rs.Fields("Variety2") <> "" then
			response.write("," & rs.Fields("Variety2"))
		end if
		if rs.Fields("Variety3") <> "" then
			response.write("," & rs.Fields("Variety3"))
		end if
		if rs.Fields("Variety4") <> "" then
			response.write("," & rs.Fields("Variety4"))
		end if %>
</td>
<td><%if rs.Fields("Bartlet") then response.write("Y") else response.write("N") end if%>&nbsp;</td>
<td><%=rs.Fields("Stage")%></td><td><%=rs.Fields("Weather")%>&nbsp;</td><td><%=rs.Fields("Target")%></td>
<td>
<%		if isDate(rs.Fields("HarvestDate")) then
			response.write(month(rs.Fields("HarvestDate")) & "/" & day(rs.Fields("HarvestDate")) & "/" & right(year(rs.Fields("HarvestDate")),2))
		end if%>&nbsp;</td>
<td><%=rs.Fields("Method")%></td><td><%=rs.Fields("AcresTreated")%></td><td><%=(int(rs.Fields("RateAcre")*100)/100)%></td><td><%=(int(rs.Fields("AcresTreated")*rs.Fields("RateAcre")*100)/100)%></td><td><%=rs.Fields("Unit")%></td></tr>
<%		REM show comments only on last spray product line for multiple product applications.
		lastComment = rs.Fields("Comments")
		rs.MoveNext
		if not rs.EOF then
			if thisSprayDate <> rs.Fields("SprayStartDate") or _
				thisLocation <> rs.Fields("Location") then
					if lastComment <> "" AND showComments then%>
<tr><td align="right">Comments:</td><td colspan="11"><%=lastComment%></td></tr>
<%					end if
			end if
		else
			if lastComment <> "" AND showComments then%>
<tr><td align="right">Comments:</td><td colspan="11"><%=lastComment%></td></tr>
<%			end if
		end if

'		rs.MoveNext
		if not rs.EOF THEN
			if thisGrowerID <> rs.Fields("GrowerID") or thisLocation <> rs.Fields("Location") or thisSprayDate <> rs.Fields("SprayStartDate") then
'changed from 10 to 8 1/12/2007
				if i > 8 then ' time for a page break%>
</table><p style="page-break-before: always">

<%
					thisGrowerID = 9999
					thisSprayDate = ""
					tablestarted = false
					i=0
				end if 'time for a page break
			end if
		end if
	LOOP

END IF
set rs = nothing
EndConnect(conn)
%>

</body>
</html>
