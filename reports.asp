<%Option Explicit%>
<%if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayRecord.asp"-->
<!--#include file="i_Crop.asp"-->
<!--#include file="i_Growers.asp"-->
<!--#include file="i_Method.asp"-->
<!--#include file="i_SprayList.asp"-->
<!--#include file="i_Stage.asp"-->
<!--#include file="i_Target.asp"-->
<!--#include file="i_Units.asp"-->
<!--#include file="i_Method.asp"-->
<%
'CREATED by LocusInteractive on 08/02/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlSprayRecordID,formSprayRecordID,urlAdding,formAdding,urlSearch,formSearch
Dim conn,sql,rs,rsSelect,counter,searchQueryString,page,searching

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

dim searchGrower,searchHighSprayDate,searchHighHarvestDate,searchLowSprayDate,searchLowHarvestDate

searching=Request.Form.Item("searching")
if Request.QueryString("searching") <> "" AND searching = "" THEN 
	searching = Request.QueryString("searching")
END IF
urlSearch = 1
formSearch = 1
searchGrower=Request.Form.Item("searchGrower")
if Request.QueryString("searchGrower") <> "" AND searchGrower = "" THEN 
	searchGrower = Request.QueryString("searchGrower")
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

searchQueryString = "searchGrower=" & RemoveWhitespace(searchGrower)  & "&searchHighSprayDate=" & RemoveWhitespace(searchHighSprayDate) & "&searchHighHarvestDate=" & RemoveWhitespace(searchHighHarvestDate) & "&searchLowSprayDate=" & RemoveWhitespace(searchLowSprayDate) & "&searchLowHarvestDate=" & RemoveWhitespace(searchLowHarvestDate) & "&page=" & page & "&searching=" & searching

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("reports.asp?" & searchQueryString)
END IF

'Initialize Form Fields


'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")




%>
<html>
<head>
	<title>SprayRecord List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->
<table width="594" border="1" cellspacing="0" cellpadding=" 0" bordercolor="#013166" bgcolor="#beige"><tr><td><img src="images/spacer.gif" height="4" width="1" border="0"><br>
&nbsp;&nbsp;<h1>Grower Reports<h1><br><img src="images/spacer.gif" height="4" width="1" border="0"></td></tr></table>
<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr>
<td colspan="2" class="bodytext">
<!---SEARCH--->

<%if formSearch then %>

<form action="reports.asp?searching=1" method="post" name="frmsearch">
<table width="500" border="1" cellpadding="0" cellspacing="0">
<tr><td>
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr bgcolor="#cccccc">
<td align="left" class="bodytext" colspan="2"><strong>
SELECT SEARCH CRITERIA</strong> &nbsp;&nbsp; &nbsp;&nbsp;<a href="reports.asp?search=1">Reset Search</a><br>
Hold down the "CTRL" key to multiple select.<br>
* indicates in-active data</td><td>&nbsp;</td>
</tr>


<tr><td valign="top"><span class="subtitle"><label for="GrowerName">Grower</label>:</span>
<%
	set rsSelect = GetAllGrowers()
%>
<select name="SearchGrower" size="12" multiple>
<%
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=trim(rsSelect.Fields("GrowerID"))%>" <%if listContains(trim(SearchGrower), trim(rsSelect.Fields("GrowerID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("GrowerName")%> </option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></span></td>
<td valign="top">
<span class="subtitle"><label for="SearchHighSprayDate">SprayDate</label>:</span><br><span class="bodytext">High:  <input type="text" value="<%=SearchHighSprayDate%>" name="SearchHighSprayDate" size="10" maxlength="21" class="bodytext"></span><br>Low: <input type="text" value="<%=SearchLowSprayDate%>" name="SearchLowSprayDate" size="10" maxlength="21" class="bodytext"></td>
<td valign="top">
<span class="subtitle"><label for="SearchHighHarvestDate">HarvestDate</label>:</span><br>High: <span class="bodytext"><input type="text" value="<%=SearchHighHarvestDate%>" name="SearchHighHarvestDate" size="10" maxlength="21" class="bodytext"></span><br>
Low: <span class="bodytext"><input type="text" value="<%=SearchLowHarvestDate%>" name="SearchLowHarvestDate" size="10" maxlength="21" class="bodytext"></span></td>

</tr>
</table>
<input type="submit" name="Go" value="--- SEARCH NOW---" class="bodytext">
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
<strong>CURRENT SEARCH</strong> <a href="reports.asp?search=1&<%=searchQuerySTring%>">Edit Search</a>&nbsp;&nbsp;<a href="reports.asp?search=0">Reset Search</a><br>
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
<%end if 'end form search%>






<%
IF searching THEN

	set rs = GetCountSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate)
	dim thisCount
	 thisCount = rs(0)

	set rs = GetSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate)
	Dim i
	i = 0
%><br><br>
<%
	Dim Field
 For Each Field In rs.Fields 
     'Response.Write "<br>" & Field.name & "" 
Next 
%>
<strong><%=thisCount%> records returned.</strong> -- <a href="javascript:void(0);" onclick="window.open('pu_growerreport.asp?<%=searchQuerySTring%>','printable','width=550,height=350,scrollbars=yes,resizable=yes');" class="bodytext">VIEW PRINTABLE</a><br>

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
		response.write(" &nbsp;&nbsp;<a href=reports.asp?newPage=" & nextpage & "&"  & searchQuerySTring &  ">" & nextpage & "</a> ")
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
if thisGrowerID <> rs.Fields("Growers.GrowerID") then
thisGrowerID = rs.Fields("Growers.GrowerID")
thisSprayDAte = ""
%>
<%if tablestarted then 
	tablestarted=false%>
</table></td></tr><tr><td>
<%end if%>
<br><br>
<table><tr><td valign="top"><h1><strong><%=rs.Fields("GrowerName")%>&nbsp;#<%=rs.Fields("GrowerNumber")%></strong>&nbsp;&nbsp;</h1><br>
<%=rs.Fields("Address")%><br>
<%=rs.Fields("City")%> &nbsp;&nbsp;<%=rs.Fields("State")%> &nbsp;&nbsp;<%=rs.Fields("ZipCode")%> &nbsp;&nbsp;<br>
<strong>Applicator/Supervisor</strong> <%=rs.Fields("ApplicatorSupervisor")%> &nbsp;&nbsp;<br>
<strong>ChemicalSupplier</strong> <%=rs.Fields("ChemicalSupplier")%> &nbsp;&nbsp;<br>
<strong>Fieldman</strong> <%=rs.Fields("Fieldman")%> &nbsp;&nbsp;<br>
</td><td valign="top" align="right"><strong>Report Date</strong> <%=now()%></td></tr></table>
<br>
<%end if
if thisSprayDAte <> rs.Fields("SprayDate") then
thisSprayDate = rs.Fields("SprayDate")
%>
<%if tablestarted then 
	tablestarted=false%>
</table></td></tr><tr><td>
<%end if%>
<h2><strong>Spray Date: <%=rs.Fields("SprayDate")%></h2><br>
<table border="1" cellspacing="0" cellpadding="1">
<tr><td><strong>Product</strong></td><td><strong>Crop</strong></td><td><strong>Bartlet</strong></td><td><strong>Stage</strong></td><td><strong>Location</strong></td><td><strong>Target</strong></td><td><strong>Harvest</strong></td><td><strong>Method</strong></td><td><strong>#Acres Treated</strong></td><td><strong>Rate/Acre</strong></td><td><strong>Total Mtl Appld</strong></td><td><strong>Units of Product</strong></td></tr>
<%
tablestarted = true
end if %>

<tr><td nowrap><%=rs.Fields("Name")%></td><td><%=rs.Fields("Crop")%></td><td><%=rs.Fields("Bartlet")%></td><td><%=rs.Fields("Stage")%></td><td><%=rs.Fields("Location")%></td><td><%=rs.Fields("Target")%></td><td><%=rs.Fields("HarvestDate")%>&nbsp;</td><td><%=rs.Fields("Method")%></td><td><%=rs.Fields("AcresTreated")%></td><td><%=rs.Fields("RateAcre")%></td><td><%=rs.Fields("TotalMaterialApplied")%></td><td><%=rs.Fields("Unit")%></td></tr>
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


</td></tr>

</table>
<%
	set rs = nothing
	EndConnect(conn)
%>
<!--#include file="i_adminfooter.asp" -->
</body>
</html>
