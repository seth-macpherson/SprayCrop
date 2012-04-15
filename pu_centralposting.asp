<%Option Explicit%>
<%if not session("login") or not listContains("1,2,3", session("accessid")) then
	response.redirect("index.asp")
end if%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayRecord.asp"-->

<%
'CREATED by LocusInteractive on 08/02/2005
Dim conn,sql,rs,rsSelect,counter,searchQueryString,page,searching

dim showComments,searchGrower,searchHighSprayDate,searchHighHarvestDate,searchLowSprayDate,searchLowHarvestDate,searchSprayYear,formCropID,formVarietyID,formBartlet,thisLocation

showComments = Request.QueryString("showComments")
if showComments = "" then
	showComments = FALSE
END IF
formBartlet = Request.QueryString("formBartlet")

'IF session("growerid") = 0 THEN
'	searchGrower=Request.Form.Item("searchGrower")
'	if Request.QueryString("searchGrower") <> "" AND searchGrower = "" THEN 
'		searchGrower = Request.QueryString("searchGrower")
'	END IF
'ELSE
'	searchGrower = session("growerID")
'END IF
' rem fixed 7/25/06 kmiers
	IF Request.QueryString("searchGrower") THEN 
		searchGrower = Request.QueryString("searchGrower")
	END IF
	IF searchGrower = "" THEN
		searchGrower = session("growerID")
	END IF
'

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

	formCropID = Request.QueryString("formCropID")
	formVarietyID = Request.QueryString("formVarietyID")


searchQueryString = "searchGrower=" & RemoveWhitespace(searchGrower)  & "&searchHighSprayDate=" & RemoveWhitespace(searchHighSprayDate) & "&searchHighHarvestDate=" & RemoveWhitespace(searchHighHarvestDate) & "&searchLowSprayDate=" & RemoveWhitespace(searchLowSprayDate) & "&searchLowHarvestDate=" & RemoveWhitespace(searchLowHarvestDate) & ",searchSprayYear=" & searchSprayYear 

'Initialize Form Fields


'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")




%>
<html>
<head>
	<title><%=Application("BusinessName")%>&nbsp;-&nbsp;<%=Application("ProgramName")%>&nbsp;-&nbsp;Central Posting Report</title>

	<style type="text/css">
body {
	background: white;
	}
td { 
      page-break-inside: avoid; 
	font-family : Arial, Sans-serif; 
	font-size : 20px; 
	} 


tr { 
page-break-inside: avoid; 
page-break-before: avoid; 
page-break-after: avoid; 
} 
	

	</style>
</head>

<body bgcolor="#ffffff" leftmargin="10" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0"  onload="if (window.focus)self.focus();">


<%

	set rs = GetCountSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate,searchSprayYear,formCropID,formVarietyID,formBartlet)
	dim thisCount
	 thisCount = rs(0)

	set rs = GetSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate,searchSprayYear,formCropID,formVarietyID,formBartlet)
	Dim i
	i = 0
%>
<%
%>



<%

dim j,thisGrowerID,thisSprayDAte,tablestarted,thisLastDate
i = 0
j = 0
thisGrowerID = 0
thisSprayDAte = ""
tablestarted = false
IF not rs.EOF THEN
DO WHILE not rs.eof 
j = j + 1
i = i + 1

if thisGrowerID <> rs.Fields("GrowerID") then
thisGrowerID = rs.Fields("GrowerID")
thisSprayDAte = ""
%>
<%if tablestarted then 
	tablestarted=false
	i=1%>
</table></td></tr></table>
<font size="16" face="arial"><%=Request.QueryString("formcomments")%></font><p style="page-break-before: always">

<%end if%>





<table width="900" border="0"><tr><td valign="top" colspan="2"><font size="13pt" face="arial">Central Posting Report</font><br></td></tr>
<tr><td valign="top">

<font size="+3"  face="arial"><%=rs.Fields("GrowerName")%></font><strong>&nbsp;&nbsp;<font size="+3"  face="arial">#<%=rs.Fields("GrowerNumber")%></font></strong>&nbsp;&nbsp;<br>
<%=rs.Fields("Address")%>,&nbsp;&nbsp;<%=rs.Fields("City")%> &nbsp;&nbsp;<%=rs.Fields("State")%> &nbsp;&nbsp;<%=rs.Fields("ZipCode")%> &nbsp;&nbsp;<br>
<strong>Applicator/Supervisor</strong> <%=rs.Fields("Supervisor")%> &nbsp;&nbsp;<br>
<strong>ChemicalSupplier</strong> <%=rs.Fields("ChemicalSupplier")%> &nbsp;&nbsp;<br>
<strong>Fieldman</strong> <%=rs.Fields("Fieldman")%> &nbsp;&nbsp;<br>
</td><td valign="bottom" align="right">Report Date:<strong> <%=now()%></strong><br>Spray Year:
<%
	sql = "SELECT SprayYear FROM SprayYears WHERE SprayYearID = " & rs.Fields("SprayYearID")
	set rsSelect = conn.execute(sql)
	%>
 <strong><%=rsSelect(0)%></td></tr></table>
 <table border="1" width="900" cellpadding="3" cellspacing="0">
<tr><td align="center"><strong>Area Treated</strong></td>
<td align="center"><strong><span >Product Name <br> EPA Reg Number</strong></td>
<td align="center"><strong><span >Active Ingredient: Common or Chemical Name</strong></td>
<td align="center">
<strong><span >Application Date</strong></td>
<td align="center">
<strong><span >Restricted Entry Interval</span></strong></td>
<td align="center">
<strong><span >Do Not Enter Until:</strong></td></tr>
<%
tablestarted = true
end if

if thisSprayDAte <> rs.Fields("SprayStartDate") or  thisLocation <> rs.Fields("Location") then
thisSprayDate = rs.Fields("SprayStartDate")
thisLocation = rs.Fields("Location")
%>
<tr><td colspan="4"><strong>Spray Date: <%=rs.Fields("SprayStartDate")%>-<%=rs.Fields("SprayEndDate")%></td>
<%if i = 1 then%>
<%thisCount = thisCount + 1%>
<td rowspan = <%=thisCount%> height="100px"><font size="50pt"><%=Request.QueryString("REI")%></font></td>
<td rowspan = <%=thisCount%>><font size="50pt"><%=Request.QueryString("donotenteruntil")%></font></td>
<%end if%>
</tr>

<%
end if 
%>


</tr>


<tr><td><%=rs.Fields("Location")%>&nbsp;</td>
<td><%=rs.Fields("Name")%>&nbsp;</td>
<td><%=rs.Fields("ActiveInd")%>&nbsp;</td>
<td><%=rs.Fields("SprayStartDate")%>
<%
thisLastDate = rs.Fields("SprayStartDate")

if rs.Fields("SprayEndDate") <> "" THEN%>
-<%=rs.Fields("SprayEndDate")%>
<%
	thisLastDate = rs.Fields("SprayEndDate")
end if
%>
&nbsp;<%=rs.Fields("TimeFinishedSpraying")%></td>

</tr>
<%
rs.MoveNext
if not rs.EOF THEN
if thisGrowerID = rs.Fields("GrowerID") then
if i = 5  or thisLocation <> rs.Fields("Location") or thisSprayDate <> rs.Fields("SprayStartDate") then ' time for a page break%>
</table>
<font size="25" face="arial"><%=Request.QueryString("formcomments")%></font>
<p style="page-break-before: always">

<%
thisGrowerID = 9999
thisSprayDate = ""
tablestarted = false
i=0
END IF 'time for a page break
END IF
END IF
LOOP

Else
%>


<%	end if %>
</table>

<font size="25" face="arial"><%=Request.QueryString("formcomments")%></font>



<%
	set rs = nothing
	EndConnect(conn)
%>

<!--#include file="i_adminfooter.asp" -->
</body>
</html>
