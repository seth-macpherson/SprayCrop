<%Option Explicit%>
<%if not session("login") or not listContains("3", session("accessid")) then
	response.redirect("index.asp")
end if
    Response.Expires = 0
%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_Growers.asp"-->
<!--#include file="i_GrowerLocations.asp"-->
<%
'MODIFIED from Locus Interactive CREATED by Kim Miers on 07/16/2006
Dim conn, rsGrower, counter

set conn = Connect()

if request.servervariables("request_method")="POST" then
    
    dim gr: gr=split(request.Form("grower"),"|")
    
    if isarray(gr) then 
        session("growerid")=gr(0)
        session("growername")=gr(1)
    end if
    
end if

set rsGrower = GetGrowersByID(session("growerid"))
%>
<html>
<head>
	<title>Roles</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center>
<tr><td><img src="images/spacer.gif" height="4" width="1" border="0">
<h1> > Roles</h1><br><img src="images/spacer.gif" height="4" width="1" border="0">
</td></tr></table>

<form name=frm method=post>
<table width="95%" border="0" bgcolor="FFFFFF" align="center">

<!--<tr><td bgcolor="FFFFFF" class="bodytext" align="right"><a href="GrowerLocations.asp#Instructions">Instructions for maintaining your locations.</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td>
</tr>-->

<tr><td bgcolor="FFFFFF" class="bodytext">Current Role: <b><%=rsGrower.Fields("GrowerNumber")%>&nbsp;<%=rsGrower.Fields("GrowerName")%></b><br>
</td>
</tr>
<tr><td bgcolor="FFFFFF" class="bodytext">&nbsp;<br></td>
</tr>
<tr>
<td colspan="2" class="bodytext">

<table width=600 border=1 cellpadding=3 cellspacing=0><tr bgcolor=#dddddd><td><b>Grower</b></td><td align=center>&nbsp;</td></tr>
<%
dim rsGC: set rsGC = conn.execute("exec growerunit$bygrower " & session("growerid")) 

with response
do until rsGC.eof
    
    .Write "<tr " 
    if cint(session("growerid"))=rsgc("growerid") then .Write "bgcolor=#eeeeee"
    .Write "><td>" & rsgc("growername") & "</td>"
    .Write"<td align=center>&nbsp;"
        .Write "<input type=radio onclick=document.frm.submit(); name=grower value="""&rsgc("growerid")&"|"&rsgc("growername")&""""
        if cint(session("growerid"))=rsgc("growerid") then .Write " checked"
        .Write ">"
    .write "</td></tr>"
    
rsGC.movenext
loop
end with

%>
</table>


<br /><br />&nbsp;
</td></tr></table>


</form>

<%
	EndConnect(conn)
%>

<!--#include file="i_adminfooter.asp" -->

</body>
</html>
