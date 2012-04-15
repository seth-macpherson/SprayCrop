<%Option Explicit%>
<%if not session("login") or not listContains("1,2,3", session("accessid")) then
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
    
    conn.execute "exec growercrop$upd " & session("growerid")
    
    dim arrcrop:arrcrop=split(request.Form("cropid"),",")
    dim pid,pno,cid
    
    for counter=lbound(arrcrop) to ubound(arrcrop)
        
       cid=trim(arrcrop(counter))
       
       conn.execute "exec growercrop$add " & session("growerid") & "," & cid & ",'"&request.Form("packer"&cid)&"'"

    next
    
    if request.Form("isdefault") <> "" then _
        conn.execute "exec growercrop$def " & session("growerid") & "," & request.Form("isdefault")
    
end if

set rsGrower = GetGrowersByID(session("growerid"))
%>
<html>
<head>
	<title>Crops</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center>
<tr><td><img src="images/spacer.gif" height="4" width="1" border="0">
<h1> > Crops</h1><br><img src="images/spacer.gif" height="4" width="1" border="0">
</td></tr></table>

<form name=frm method=post>
<table width="95%" border="0" bgcolor="FFFFFF" align="center">

<!--<tr><td bgcolor="FFFFFF" class="bodytext" align="right"><a href="GrowerLocations.asp#Instructions">Instructions for maintaining your locations.</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td>
</tr>-->

<tr><td bgcolor="FFFFFF" class="bodytext">Grower: <b><%=rsGrower.Fields("GrowerNumber")%>&nbsp;<%=rsGrower.Fields("GrowerName")%></b><br>
</td>
</tr>
<tr><td bgcolor="FFFFFF" class="bodytext">&nbsp;<br></td>
</tr>
<tr>
<td colspan="2" class="bodytext">

<table width=600 border=1 cellpadding=3 cellspacing=0><tr bgcolor=#dddddd><td><b>Crop</b></td><td align=center><b>Default Packer?</b></td><td align=center><b>Default Crop?</b></td></tr>
<%
dim rsGC: set rsGC = conn.execute("exec growercrop$list " & session("growerid")) 

with response
do until rsGC.eof
    
    .Write "<tr " 
    if rsgc("isdefault") then .Write "bgcolor=#eeeeee"
    .Write "><td><input onclick=document.frm.submit() type=checkbox name=cropid value="& rsgc("cropid") 
    if not isnull(rsgc("growercropid")) then .Write " checked"
    .Write ">" & rsgc("crop") & "</td>"

    .Write"<td align=center>"
    IF TRUE THEN
    
    	    dim sql,rsselect:set rsSelect = GetActivePackers()
            .Write "<SELECT name=packer"& rsgc("cropid") &" style=background-color:beige;width:150px;>"
            .Write "<option value="""">Other/Unspecified</option>"
            IF not rsSelect.EOF THEN
            DO WHILE not rsSelect.eof 
            .write "<option value="&rsSelect.Fields("PackerNumber")     
            if rsSelect.Fields("PackerID")=rsgc("packerid") then
                .write " selected" 
            end if
            .Write ">"
            .write string(6-len(rsSelect.Fields("PackerNumber")),"0")&rsSelect.Fields("PackerNumber")
            .Write "</option>"
            rsSelect.MoveNext
            LOOP
            END IF
            .Write "</select>"

    ELSEIF FALSE THEN

        if not isnull(rsgc("growercropid")) then
        if not isnull(rsgc("packerid")) then 
            .Write rsgc("packername") & " ("&rsgc("packernumber")&")"
            .Write " <input type=hidden name=packer"& rsgc("cropid") &" value="""&rsgc("packernumber")&""">"
        else
            .Write "<input type=text name=""packer"& rsgc("cropid") &""" size=10>"
        end if
        else
            .Write "&nbsp;"
        end if
        
    END IF
    .Write "</td>"
    
    .Write"<td align=center>&nbsp;"
    if not isnull(rsgc("growercropid")) then 
        .Write "<input type=radio onclick=document.frm.submit() name=isdefault value="&rsgc("cropid")
        if rsgc("isdefault") then .Write " checked"
        .Write ">"
    end if
    .write "</td></tr>"
    
rsGC.movenext
loop
end with

%>
</table>

<br /><br />&nbsp;
<input type="submit" name="frmaction" value="Save">

</td></tr></table>
</form>

<%
	EndConnect(conn)
%>

<!--#include file="i_adminfooter.asp" -->

</body>
</html>
