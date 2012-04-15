<%Option Explicit%>
<%if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_growerusers.asp"-->
<%
'CREATED by LocusInteractive on 08/02/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlgroweruserID,formgroweruserID,formgrowerID
Dim conn,sql,rs,counter

'Initialize variables
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"
formgroweruserID=Request.Form.Item("groweruserID")
urlgroweruserID=Request.QueryString("groweruserID")

'See if ID was passed through URL or FORM
IF urlgroweruserID = "" THEN urlgroweruserID = 0 END IF
IF formgroweruserID = "" THEN formgroweruserID = urlgroweruserID End IF
urlgroweruserID = formgroweruserID

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Response.Redirect("growerusers_list.asp")
END IF

'Initialize Form Fields
DIM formUsername,formPassword,formAccessID
formUsername = Request.Form.Item("Username")
formPassword = Request.Form.Item("Password")
formAccessID = Request.Form.Item("AccessID")
formgrowerID = request.Form.Item("growerID")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")

IF Request.QuerySTring("task") = "d" and urlgroweruserID <> "" THEN
	Deletegrowerusers(urlgroweruserID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("growerusers_list.asp")
END IF 
IF Request.QuerySTring("task") = "activate" and urlgroweruserID <> "" THEN
	Activategrowerusers(urlgroweruserID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("growerusers_list.asp")
END IF 
IF Request.QuerySTring("task") = "deactivate" and urlgroweruserID <> "" THEN
	DeActivategrowerusers(urlgroweruserID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("growerusers_list.asp")
END IF 


'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("Username"), "varchar","Username", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Password"), "varchar","Password", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("AccessID"), "int","AccessID", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		
	IF NOT ValidateDatatype(Request.Form.Item("growerID"), "int","growerID", true) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		urlgroweruserID = UPDATEgrowerusers(formgroweruserID,formUsername,formPassword,formgrowerid,formAccessID)
		set rs = nothing
		EndConnect(conn)
		Response.Redirect("growerusers_list.asp")
		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlgroweruserID = Insertgrowerusers(formUsername,formPassword,formgrowerid,formAccessID)
		set rs = nothing
		EndConnect(conn)
		Response.Redirect("growerusers_list.asp")
	END IF 'insert	
END IF 'form submitted 

IF formgroweruserID  <>  0 and not errorFound THEN
set rs = GetgrowerusersByID(formgroweruserID)
	IF NOT rs.eof THEN
		formUsername = rs.Fields("Username")
		formPassword = rs.Fields("Password")
		formgrowerID = rs.Fields("growerID")
		formAccessID = rs.Fields("AccessID")
	END IF
END IF%>
<html>
<head>
	<title>Grower Users List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Grower Users</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="ffffff" class="bodytext">
<%if listContains("1,2", session("accessid")) then %>
<br>
<a href="growerusers_list.asp#edit">Add Grower User</a><br />&nbsp;</td>
</tr>
<%end if %>
<tr>
<td colspan="2" class="bodytext">
<table width="100%" border="1" cellpadding="2" cellspacing="0">
<% if  delerror then%>
<tr>
<td colspan="6" class="bodytext" valign="top"><font color="red"><%= delerrormessage %></font></td>
</tr>
<% end if %>
<% if  errorFound then%>
<tr>
<td colspan="6" class="bodytext" valign="top"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></td>
</tr>
<% end if %>
<tr bgcolor=#eeeeee>
<!--<td  valign="top">&nbsp;</td>-->
	<td  valign="top"><h2>Username</h2></td>
	<td  valign="top"><h2>Grower</h2></td>


<td  valign="top" bgcolor=#eeeeee><h2>&nbsp;</h2></td>

</tr>
<%
if listContains("1", session("accessid")) then
	set rs = GetAllgrowerusers()
elseif listContains("2", session("accessid")) then
    set rs = GetPackerGrowerUsers()
end if

%>
<%IF not rs.EOF THEN
DO WHILE not rs.eof %>
<tr>
<!--<td class = "bodytext" valign="top"><%=rs.Fields("groweruserID")%></td>-->
	<td class="bodytext" valign="top"><%=rs.Fields("Username")%>&nbsp;</td>

	<td class="bodytext" valign="top"><%=rs.Fields("growername") %></td>

<td  bgcolor=#eeeeee valign="top" class="bodytext"><a href="growerusers_list.asp?groweruserID=<%=rs.Fields("groweruserID")%>#edit" class="bodytext">Edit</a><br>
<%if listContains("1", session("accessid")) then %><a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="growerusers_list.asp?groweruserID=<%=rs.Fields("groweruserID")%>&task=d" class="bodytext">Delete</a><%end if %>
</td>
</tr> 
<% 
rs.MoveNext
LOOP

else
%>
<tr><td class="bodytext" colspan="6">No Records Selected</td></tr>
<%	end if %>
</table>

<% if errorFound then%>
<tr>
<td>&nbsp;</td>
<td class="bodytext"><font color="red"><% =errorMessage%></font></td>
</tr>
<% End If %>

<%if (listContains("1,2", session("accessid")) or urlgroweruserID <> 0 ) then %>

<a name="edit"></a>
<form action="growerusers_list.asp" method="post" name="frmsearch">
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr>
<td>&nbsp;</td><td align="left" class="bodytext">* indicates required field</td>
</tr>

<tr><td valign="top" align="right"><input type="hidden" value="<% =urlgroweruserID%>" name="groweruserID"><span class="subtitle"><label for="Username">Username</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formUsername%>" name="Username"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Password">Password</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="password" value="<%=formPassword%>" name="Password"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>

<%if listContains("1,2", session("accessid"))  then%>

<tr><td valign="top" align="right"><span class="subtitle"><label for="grower">grower</label>:</span></td><td>

<%

	    dim rsSelect: set rsSelect = GetAllgrowers()
    %>
    <SELECT name="growerID" style="background-color:beige;">
    <%IF session("growerid") = 0 THEN%>
    <option value=""></option>
    <%
    END IF
    IF not rsSelect.EOF THEN
    DO WHILE not rsSelect.eof 
    %>
    <option value="<%	response.write(rsSelect.Fields("growerID"))%>" 
    <%if trim(formgrowerID) = trim(rsSelect.Fields("growerID")) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("growerName")%></option>
    <%
    rsSelect.MoveNext
    LOOP
    END IF
    %>
    </select>
</td></tr>

<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="AccessID">AccessID</label>:</span></td>
<td valign="top"><span class="bodytext"><input type=hidden name=accessID value=3 /> default

<!--
<input type="radio" value="1" name="accessID" <%if formAccessID = 1 then response.write "checked" end if %>>Super Admin - access to Application Setup and growerusers<br>
<input type="radio" value="2" name="accessID" <%if formAccessID = 2 then response.write "checked" end if %>>Admin - access to Application Setup<br>
<input type="radio" value="3" name="accessID" <%if formAccessID = 3 then response.write "checked" end if %>>User - access only to Enter Spray Data, Review, and Reports<br>
-->

</span></td>
</tr>

<%else %>
    <tr><td colspan=2><input type=hidden name=growerid value=<%=formgrowerID %> /><input type=hidden name=accessID value=3 /> </td></tr>
<%end if %>
<tr>
<td>&nbsp;</td>
<td><% IF  urlgroweruserID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>
<%end if %>

</table>
</form>
<%
	set rs = nothing
	EndConnect(conn)
%>
<!--#include file="i_adminfooter.asp" -->
</body>
</html>
