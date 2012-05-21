<%Option Explicit%>
<%if not session("login") or not listContains("1,2,3", session("accessid")) then
	response.redirect("index.asp")
end if%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayRecord.asp"-->
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
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError,thisLocation
Dim urlSprayRecordID,formSprayRecordID,urlAdding,formAdding,urlSearch,formSearch,formCropID,formCropID2,formCropID3,formCropID4,urlCropID,formVarietyID,formBartlet,urlBartlet
Dim conn,sql,rs,rsSelect,counter,searchQueryString,page,searching,formrei,formdonotenteruntil,formcomments,urlrei,urldonotenteruntil,urlcomments,rsCrop,urlVarietyID

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")
set rsCrop = Server.CreateObject("ADODB.RecordSet")
'Initialize variables
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"
formSprayRecordID=Request.Form.Item("SprayRecordID")
urlSprayRecordID=Request.QueryString("SprayRecordID")

'See if ID was passed through URL or FORM
IF urlSprayRecordID = "" THEN urlSprayRecordID = 0 END IF
IF formSprayRecordID = "" THEN formSprayRecordID = urlSprayRecordID End IF
urlSprayRecordID = formSprayRecordID

formAdding=Request.Form.Item("Adding")
urlAdding=Request.QueryString("Adding")

'See if ID was passed through URL or FORM
IF urlAdding = "" THEN urlAdding = 0 END IF
IF formAdding = "" THEN formAdding = urlAdding End IF
urlAdding = formAdding


formSearch=Request.Form.Item("Search")
urlSearch=Request.QueryString("Search")

'See if ID was passed through URL or FORM
IF urlSearch = "" THEN urlSearch = 0 END IF
IF formSearch = "" THEN formSearch = urlSearch End IF
urlSearch = formSearch

     if request.servervariables("request_method")="POST" and request.Form("changerole")>"" then

        dim g: g=split(request.Form("growerrole"),"|")

        if isarray(g) then
            session("growerid")=g(0)
            session("growername")=g(1)
        end if

    end if

formCropID = Request.Form.Item("CropID")
urlCropID=Request.QueryString("CropID")
if formCropID = "" then formCropID = urlCropID end if
'See if ID was passed through URL or FORM
IF urlCropID = "" THEN urlCropID = 0 END IF
IF formCropID = "" THEN formCropID = urlCropID End IF
urlCropID = formCropID
if right(formCropID,1)="," then formCropID=left(formCropID,len(formCropID)-1)
formcropid=replace(formcropid," ","")

'# debug June-2010
'# response.write formCropID
'# response.end

formVarietyID = Request.Form.Item("VarietyID")
urlVarietyID=Request.QueryString("VarietyID")
if formVarietyID = "" then formVarietyID = urlVarietyID end if
'See if ID was passed through URL or FORM
IF urlVarietyID = "" THEN urlVarietyID = 0 END IF
IF formVarietyID = "" THEN formVarietyID = urlVarietyID End IF
urlVarietyID = formVarietyID

formBartlet = Request.Form.Item("Bartlet")
urlBartlet=Request.QueryString("Bartlet")
if formBartlet = "" then formBartlet = urlBartlet end if
'See if ID was passed through URL or FORM
IF urlBartlet = "" THEN urlBartlet = "" END IF
IF formBartlet = "" THEN formBartlet = urlBartlet End IF
urlBartlet = formBartlet


formdonotenteruntil = Request.Form.Item("donotenteruntil")
urldonotenteruntil=Request.QueryString("donotenteruntil")
if formdonotenteruntil = "" then formdonotenteruntil = urldonotenteruntil end if
'See if ID was passed through URL or FORM
IF urldonotenteruntil = "" THEN urldonotenteruntil = "" END IF
IF formdonotenteruntil = "" THEN formdonotenteruntil = urldonotenteruntil End IF
urldonotenteruntil = formdonotenteruntil


formREI = Request.Form.Item("REI")
urlREI=Request.QueryString("REI")
if formREI = "" then formREI = urlREI end if
'See if ID was passed through URL or FORM
IF urlREI = "" THEN urlREI = "" END IF
IF formREI = "" THEN formREI = urlREI End IF
urlREI = formREI


formcomments = Request.Form.Item("comments")
urlcomments=Request.QueryString("comments")
if formcomments = "" then formcomments = urlcomments end if
'See if ID was passed through URL or FORM
IF urlcomments = "" THEN urlcomments = "" END IF
IF formcomments = "" THEN formcomments = urlcomments End IF
urlcomments = formcomments

dim searchGrower,searchHighSprayDate,searchHighHarvestDate,searchLowSprayDate,searchLowHarvestDate,searchSprayYear

searching=Request.Form.Item("searching")
if Request.QueryString("searching") <> "" AND searching = "" THEN
	searching = Request.QueryString("searching")
END IF
urlSearch = 1
formSearch = 1
'IF session("growerid") = 0 THEN
'	searchGrower=Request.Form.Item("searchGrower")
'	if Request.QueryString("searchGrower") <> "" AND searchGrower = "" THEN
'		searchGrower = Request.QueryString("searchGrower")
'	END IF
'ELSE
'	searchGrower = session("growerID")
'END IF

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
	searching=1
	page = Request.QueryString("newPage")
END IF
if page = "" then
	page = 1
end if

searchQueryString = "searchGrower=" & RemoveWhitespace(searchGrower)  & "&searchHighSprayDate=" & RemoveWhitespace(searchHighSprayDate) & "&searchHighHarvestDate=" & RemoveWhitespace(searchHighHarvestDate) & "&searchLowSprayDate=" & RemoveWhitespace(searchLowSprayDate) & "&searchLowHarvestDate=" & RemoveWhitespace(searchLowHarvestDate) & "&page=" & page & "&searching=" & searching & "&searchSprayYear=" & searchSprayYear & "&rei=" & formREI & "&donotenteruntil=" & formdonotenteruntil  & "&CropID=" & formCropID  & "&formVarietyID=" & formVarietyID   & "&formcomments=" & formcomments & "&formBartlet=" & formBartlet

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("grower_report.asp?" & searchQueryString)
END IF

'Initialize Form Fields



%>
<html>
<head>
	<title><%=Application("BusinessName")%>&nbsp;-&nbsp;<%=Application("ProgramName")%>&nbsp;-&nbsp;Grower Report</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
	<script language="JavaScript" src="datepicker.js"></script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">

<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center>
<tr><td><img src="images/spacer.gif" height="4" width="1" border="0">
<h1> > Grower Report</h1><br><img src="images/spacer.gif" height="4" width="1" border="0">
</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr>
<td colspan="2" class="bodytext">
<!--SEARCH-->

<%if formSearch then %>

<form action="grower_report.asp?searching=1" method="post" name="frmsearch">
<table width="100%" border="1" cellpadding="0" cellspacing="0">
<tr><td>
<table width="100%" border="0" cellpadding="5" cellspacing="0">
<tr bgcolor="#cccccc">
<td align="left" class="bodytext" colspan="2"><strong>
SELECT SEARCH CRITERIA</strong> &nbsp;&nbsp; &nbsp;&nbsp;<a href="grower_report.asp?search=1">Reset Search</a><br>
Hold down the "CTRL" key to multiple select.<br>
* indicates in-active data</td><td>&nbsp;</td>
</tr>


<tr valign="top">
<td><br>
<img src="images/spacer.gif" width="150" height="1"><br><span class="subtitle"><label for="GrowerName">Grower</label>:</span>

<%if session("growerid")=0 then %>

    <%
	    set rsSelect = GetAllGrowers()
    %><br />
    <select name="SearchGrower" size="12" multiple>
    <option value="0" <%if SearchGrower="0" or SearchGrower="" then response.write "selected"%>>---SELECT ALL---</option>
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


<%

        dim roleConn: set roleConn = Connect()
        dim rsRoles: set rsRoles = conn.execute("exec growerunit$bygrower " & session("growerid"))
        with response

        .Write "<br><input type=hidden name=changerole />"
        .Write "<select name=growerrole onchange=form.changerole.value=1;form.submit(); style=background-color:beige;>"
        do until rsRoles.eof

            .Write "<option value="""&rsRoles("growerid")&"|"&rsRoles("growername")&""""
            if cint(session("growerid"))=rsRoles("growerid") then .Write " selected"
            .Write ">"
            .Write rsRoles("growername")
            .write "</option>"

        rsRoles.movenext
        loop
        .Write "</select>"

        end with
       	EndConnect(roleConn)

%>


<%end if%>
</td>

<td valign="top">
<br />

    <table width="100%">
    <tr valign=top>
    <td>

    <span class="subtitle"><label for="SearchHighSprayDate">Beginning Spray Date</label>:</span><br> <table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=SearchLowSprayDate%>" name="SearchLowSprayDate" size="10" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmsearch.SearchLowSprayDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table><br><span class="subtitle">Ending Spray Date:  <table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=SearchHighSprayDate%>" name="SearchHighSprayDate" size="10" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmsearch.SearchHighSprayDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table></span>

    </td>
    <td>
    <span class="subtitle"><label for="SearchHighHarvestDate">Beginning  Harvest Date</label>:</span><br>
    <span class="bodytext"><table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=SearchLowHarvestDate%>" name="SearchLowHarvestDate" size="10" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmsearch.SearchLowHarvestDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table></span><br><span class="subtitle">Ending Harvest Date: </span><span class="bodytext"><table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=SearchHighHarvestDate%>" name="SearchHighHarvestDate" size="10" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmsearch.SearchHighHarvestDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table></span><br><br>


    </td>
    </tr>
    </table>

</td>

<td valign="top"><br />
<%
	set rsSelect = GetAllSprayYears()
%>
<strong class=subtitle>Spray Year:</strong><br />
<select name="searchSprayYear" size="4" >
<%
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof
%>
<option value="<%=trim(rsSelect.Fields("SprayYearID"))%>" <%if listContains(trim(searchSprayYear), trim(rsSelect.Fields("SprayYearID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("SprayYear")%> </option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></td>
</tr>
<tr>
<td valign="top" colspan=3>
<strong class=subtitle>Crops:</strong>
<br />
<table width="100%" cellpadding=5 border=0>

    <%
	    set rsSelect = GetActiveCrops()
    %>
    <%
    i = 0
    IF not rsSelect.EOF THEN
    DO WHILE not rsSelect.eof
    i = i + 1
    %>
   <tr valign=top><td nowrap>

    <input type="checkbox" name="CropID" value="<%=rsSelect.Fields("CropID")%>"<%if listContains(formCropID, trim(rsSelect.Fields("CropID"))) or formCropID="0" then response.write("checked") end if%> style="background-color:beige"><%if not rsSelect.Fields("Active") then %>*<%end if%>
	<b><%=rsSelect.Fields("Crop")%></b></td><td align=left>
    <span style="font-size:8pt;">
	    <%
	    set rsCrop = GetActiveVarietiesByCropID(rsSelect.Fields("CropID"))
	    IF not rsCrop.EOF THEN
		    response.write ("")
		    DO WHILE not rsCrop.eof %>
				    &nbsp;&nbsp;<input style="padding:0px;margin:0px;" type="checkbox" name="VarietyID" value="<%=rsCrop.Fields("VarietyID")%>"
    <%if listContains(formVarietyID, trim(rsCrop.Fields("VarietyID"))) then
    response.write("checked")
    end if%>>
    <%if not rsCrop.Fields("Active") then %>*<%end if%><%=rsCrop.Fields("Variety")%>&nbsp;
    <%
	    rsCrop.MoveNext
	    LOOP
    END IF

    %>



    <%
    'if (i = 1) then
	response.write ("</td></tr>")
    'end if
    'if i=1 then i=-1
    rsSelect.MoveNext
    LOOP
    END IF
    %>
</table>
<br /><span class="subtitle"><label for="Bartlet">Bartlett:</label></span> <span class="bodytext"><span class="bodytext"><input type="radio" value="1" name="Bartlet" <% if ListContains(formBartlet,"True") OR ListContains(formBartlet,"1")  THEN %>Checked<% END IF %> style="background-color:beige">YES <input type="radio" value="0" name="Bartlet" <% if ListContains(formBartlet,"False") OR ListContains(formBartlet,"0")  THEN %>Checked<% END IF %> style="background-color:beige">NO&nbsp;<input type="radio" value="" name="Bartlet" <% if formBartlet = ""  THEN %>Checked<% END IF %> style="background-color:beige">Any</span>
<br />&nbsp;
</td>
</tr>

<tr>
<td valign="top" bgcolor="cccccc" colspan="3" class="bodytext"><strong>CENTRAL POSTING INSTRUCTIONS<br>
*Only required if generating a Central Report </strong><br><br>
<strong>Step 1) Enter the REI Information:</strong><br>
<em>IMPORTANT</em><br>
<li><em>Enter the LONGEST REI of all sprays applied in the application.</em></li>
<li><em>Growers are responsible for checking labels for Re-Entry Intervals.</em></li><br>
Restricted Entry Interval:<input type="text" name="rei" value="<%=formrei%>" class="bodytext" size="10" maxlength="30"><br>
Do not enter until date: <input type="text" name="donotenteruntil" value="<%=formdonotenteruntil%>" class="bodytext" size="10" maxlength="30"><br>
Comments:<br><textarea name="comments" class="bodytext" rows="3" cols="35"><%=formComments%></textarea><br>
<strong>Step 2)</strong>Enter search criteria and click "Search Now"<br>
<strong>Step 3)</strong>
<% IF searching THEN%> <a href="pu_centralposting.asp?<%=searchQuerySTring%>" target="printable" class="bodytext"><% END IF%>Click Here to view the Printable Central Posting Report<% IF searching THEN%> </a><% END IF%><br>
<strong>Step 4)</strong> In the report popup window, choose "File" then "Page Setup" to set the ORIENTATION to LANDSCAPE<br>
</td></tr>
</table><br /><br />

<input type="submit" name="Go" value="--- SAVE REI INFO and SEARCH NOW---" class="bodytext">
</form>
</td>
</tr>
</table>
</td>
</tr>
</table>
<br><br>

<%else%>
<table width="500" border="1" cellpadding="0" cellspacing="0">
<tr><td>
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr bgcolor="#cccccc">
<strong>CURRENT SEARCH</strong> <a href="grower_report.asp?search=1&<%=searchQuerySTring%>">Edit Search</a>&nbsp;&nbsp;<a href="grower_report.asp?search=0">Reset Search</a><br>
</td></tr>
<tr><td valign="top">
<%if searchGrower <> "" then
	set rs = GetGrowersByID(searchGrower)%>
<strong>Growers:</strong><br>
<%
IF not rs.EOF THEN
DO WHILE not rs.eof
	response.write(rs.Fields("GrowerName") & "<br>")
rs.MoveNext
LOOP
END IF
%>
<%end if%>

</td></tr></table>
</td></tr></table>

</form>
</td>
</tr>
</table>
</td>
</tr>
</table>
<br><br>
<%end if 'end form search%>


<!-- KILROY
<% response.write("Session.growerid: " + CStr(session("growerid")) + vbCRLF) %>
<% response.write("searchGrower: " + CStr(searchGrower) + vbCRLF) %>
<% response.write("Request.Form.Item(searchGrower): " + Request.Form.Item("searchGrower") + vbCRLF) %>
<% response.write("Request.QueryString(searchGrower): " + Request.QueryString("searchGrower") + vbCRLF) %>
-->
<!--
IF session("growerid") = 0 THEN
	searchGrower=Request.Form.Item("searchGrower")
	if Request.QueryString("searchGrower") <> "" AND searchGrower = "" THEN
		searchGrower = Request.QueryString("searchGrower")
	END IF
ELSE
	searchGrower = session("growerID")
END IF
-->

<%
IF searching THEN

	set rs = GetCountSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate,searchSprayYear,formCropID,formVarietyID,formBartlet)
	dim thisCount
	 thisCount = rs(0)

	set rs = GetSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate,searchSprayYear,formCropID,formVarietyID,formBartlet)
	Dim i
	i = 0

	Dim target_rs, targets, target_dictionary, spray_rec_id,a

	Set target_dictionary = Server.CreateObject("Scripting.Dictionary")

	IF not rs.EOF THEN
		DO WHILE not rs.EOF
		spray_rec_id = rs.Fields("SprayRecordID")
		set target_rs = conn.execute("SELECT t.target FROM Targets t INNER JOIN SprayRecordTargets srt ON t.targetID = srt.targetID WHERE srt.SprayRecordID = " & spray_rec_id)

		targets = ""
		Do While Not target_rs.EOF
	    	targets = targets & target_rs.Fields("Target")
			target_rs.MoveNext
			If target_rs.EOF = False Then
	    	targets = targets & ", "
			End If
		Loop

		if target_dictionary.Exists(spray_rec_id) = false then
			target_dictionary.Add spray_rec_id, targets
		end if

		set targets = nothing
		set spray_rec_id = nothing
		set target_rs = nothing

		rs.MoveNext
		LOOP

	END IF

	' go back to the start of the rs
	rs.MoveFirst

	i = 0


	'Response.Write("<p>Key values:</p>")
	'a=target_dictionary.Items
	'for i=0 to target_dictionary.Count-1
	'  Response.Write(a(i))
	'  Response.Write("<br />")
	'next


%><br><br>

<strong><%=thisCount%> records returned.</strong><br>
<a href="pu_growerreport.asp?<%=searchQueryString%>" target="printable" class="bodytext">VIEW PRINTABLE</a><br>
<a href="pu_growerreport.asp?<%=searchQueryString%>&showComments=TRUE" target="printable" class="bodytext">VIEW PRINTABLE with Comments</a><br>



<%
Dim startrow,endrow,recordsPerPage,nextPage
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
for i = 0 to thisCount Step recordsPerPage
	nextpage = (i + recordsPerPage)/recordsPerpage
	if trim(page) = trim(nextpage) then
		response.write("  &nbsp;&nbsp;<strong>" & nextpage & "</strong>")
	else
		response.write(" &nbsp;&nbsp;<a href=grower_report.asp?newPage=" & nextpage & "&"  & searchQuerySTring &  ">" & nextpage & "</a> ")
	end if
Next
else
	startrow = 1
	endrow = thisCount + 1

end if

%>

<table border="0"><tr><td>

<%

dim j,thisGrowerID,thisSprayDAte,tablestarted
i = 0
j = 0
thisGrowerID = 0
thisSprayDAte = ""
tablestarted = false
IF not rs.EOF THEN
DO WHILE not rs.eof
j = j + 1

if j < endrow and j >= startrow THEN
i = i + 1
if thisGrowerID <> rs.Fields("GrowerID") then
thisGrowerID = rs.Fields("GrowerID")
thisSprayDate = "999"
thisLocation = "999"
%>
<%if tablestarted then
	tablestarted=false%>
</table><br><br>
<%end if%>
<br><br>
<table><tr><td valign="top"><h1><strong><%=rs.Fields("GrowerName")%>&nbsp;#<%=rs.Fields("GrowerNumber")%></strong>&nbsp;&nbsp;</h1><br>
<%=rs.Fields("Address")%><br>
<%=rs.Fields("City")%> &nbsp;&nbsp;<%=rs.Fields("State")%> &nbsp;&nbsp;<%=rs.Fields("ZipCode")%> &nbsp;&nbsp;<br>
<strong>Applicator/Supervisor</strong> <%=rs.Fields("Supervisor")%> &nbsp;&nbsp;<br>
<strong>Supervisor License</strong> <%=rs.Fields("LicenseNumber")%> &nbsp;&nbsp;<br>


<strong>Fieldman</strong> <%=rs.Fields("Fieldman")%> &nbsp;&nbsp;<br>
</td><td valign="top" align="right"><strong>Report Date</strong> <%=now()%></td></tr></table>
<br>
<%end if

if (thisSprayDate = "999" OR thisSprayDAte <> rs.Fields("SprayStartDate")  or thisLocation = "999" or thisLocation <> rs.Fields("Location")) then
thisSprayDate = rs.Fields("SprayStartDate")
thisLocation = rs.Fields("Location")
%>
<%if tablestarted then
	tablestarted=false%>
</table><br><br>
<%end if%>
<h2><strong>Spray Date: </strong><%=rs.Fields("SprayStartDate")%><%if rs.Fields("SprayEndDate") <> "" then%>

-<%=rs.Fields("SprayEndDate")%>
 : <%=rs.Fields("TimeFinishedSpraying")%>

<%end if%> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
<strong>Location:</strong>
<%=rs.Fields("Location")%> &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
<strong>Chemical Supplier:</strong>
<%=rs.Fields("ChemicalSupplier")%> &nbsp;&nbsp;

<table border="1" cellspacing="0" cellpadding="1">
<tr><td><strong>Product/EPA #</strong></td><td><strong>Crop</strong></td><td><strong>Bartlett</strong></td><td><strong>Stage</strong></td><td><strong>Weather</strong></td><td><strong>Target(s)</strong></td><td><strong>Harvest</strong></td><td><strong>Method</strong></td><td><strong>#Acres Treated</strong></td><td><strong>Rate/Acre</strong></td><td><strong>Total Mtl Appld</strong></td><td><strong>Units of Product</strong></td></tr>
<%
tablestarted = true
end if %>

<tr><td nowrap><%=rs.Fields("Name")%><br />

<%if rs.Fields("PackerNUmber") <> "" then%>
<br><strong>Packer: </strong> <%=rs.Fields("PackerNumber")%>
<%end if%>
<%if rs.Fields("Applicator") <> "" then%>
<br><strong>Applicator(s): </strong><br />
&mdash; <%= replace(left(rs.Fields("Applicator"),len(rs.Fields("Applicator"))-1),";","<br />&mdash;")%>
<br /><strong>License:</strong> <%=rs.Fields("ApplicatorLicense")%>
<%end if%>
<%if rs.Fields("RecommendedBy") <> "" then%>
<br><strong>Recommended By: </strong> <%=rs.Fields("RecommendedBy")%>
<%end if%>
</td><td><%=rs.Fields("Crop")%>
<%if rs.Fields("Variety") <> "" then
	response.write("<br>varieties:<br>" & rs.Fields("Variety"))
end if%>
<%if rs.Fields("Variety2") <> "" then
	response.write("<br>" & rs.Fields("Variety2"))
end if%>
<%if rs.Fields("Variety3") <> "" then
	response.write("<br>" & rs.Fields("Variety3"))
end if%>
<%if rs.Fields("Variety4") <> "" then
	response.write("<br>" & rs.Fields("Variety4"))
end if%>

</td><td><%=rs.Fields("Bartlet")%></td><td><%=rs.Fields("Stage")%></td><td><%=rs.Fields("Weather")%>&nbsp;</td>
<td>

<% spray_rec_id = rs.Fields("SprayRecordID") %>
<%= target_dictionary.Item(spray_rec_id) %></td>

<td><%=rs.Fields("HarvestDate")%>&nbsp;</td><td><%=rs.Fields("Method")%></td><td><%=rs.Fields("AcresTreated")%></td><td><%=(int(rs.Fields("RateAcre")*100)/100)%></td><td><%=(int(rs.Fields("AcresTreated")*rs.Fields("RateAcre")*100)/100)%></td><td><%=rs.Fields("Unit")%></td></tr>
<%
END IF
rs.MoveNext
LOOP

Else
%>
<tr><td class="bodytext" colspan="27">No Records Selected</td></tr>
<%	end if %>
</td></tr>
</table>
<%	end if %>

<!--#include file="i_adminfooter.asp" -->

</td></tr>

</table>
<%
	set rs = nothing
	EndConnect(conn)
%>
</body>
</html>
