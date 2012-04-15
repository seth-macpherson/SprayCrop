<%Option Explicit%>
<%if not session("login") or not listContains("1", session("accessid")) then
	response.redirect("index.asp")
end if%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_Administrators.asp"-->
<%
'CREATED by LocusInteractive on 08/02/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlAdministratorID,formAdministratorID
Dim conn,sql,rs,counter

'Initialize variables
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"
formAdministratorID=Request.Form.Item("AdministratorID")
urlAdministratorID=Request.QueryString("AdministratorID")

'See if ID was passed through URL or FORM
IF urlAdministratorID = "" THEN urlAdministratorID = 0 END IF
IF formAdministratorID = "" THEN formAdministratorID = urlAdministratorID End IF
urlAdministratorID = formAdministratorID

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Response.Redirect("administrators_list.asp")
END IF

'Initialize Form Fields
DIM formUsername,formPassword,formAccessID
formUsername = Request.Form.Item("Username")
formPassword = Request.Form.Item("Password")
formAccessID = Request.Form.Item("AccessID")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")

IF Request.QuerySTring("task") = "d" and urlAdministratorID <> "" THEN
	DeleteAdministrators(urlAdministratorID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("administrators_list.asp")
END IF 
IF Request.QuerySTring("task") = "activate" and urlAdministratorID <> "" THEN
	ActivateAdministrators(urlAdministratorID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("administrators_list.asp")
END IF 
IF Request.QuerySTring("task") = "deactivate" and urlAdministratorID <> "" THEN
	DeActivateAdministrators(urlAdministratorID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("administrators_list.asp")
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

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		urlAdministratorID = UPDATEAdministrators(formAdministratorID,formUsername,formPassword,formAccessID)
		set rs = nothing
		EndConnect(conn)
		Response.Redirect("administrators_list.asp")
		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlAdministratorID = InsertAdministrators(formUsername,formPassword,formAccessID)
		set rs = nothing
		EndConnect(conn)
		Response.Redirect("administrators_list.asp")
	END IF 'insert	
END IF 'form submitted 

IF formAdministratorID  <>  0 and not errorFound THEN
set rs = GetAdministratorsByID(formAdministratorID)
	IF NOT rs.eof THEN
		formUsername = rs.Fields("Username")
		formPassword = rs.Fields("Password")
		formAccessID = rs.Fields("AccessID")
	END IF
END IF%>
<html>
<head>
	<title>Administrators List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Administrators</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="ffffff" class="bodytext"><br>
Add/edit/delete Administrators.<br><br>
<a href="administrators_list.asp#edit">Add Administrators</a></td>
</tr>
<tr>
<td colspan="2" class="bodytext">
<table width="90%" border="1" cellpadding="2" cellspacing="0">
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
<tr>
<td  valign="top">&nbsp;</td>
<td  valign="top"><h2>Edit</h2></td>

	<td  valign="top"><h2>Username</h2></td>
	<td  valign="top"><h2>AccessID</h2></td>
</tr>
<%
	set rs = GetAllAdministrators()
%>
<%IF not rs.EOF THEN
DO WHILE not rs.eof %>
<tr>
<td class = "bodytext" valign="top"><%=rs.Fields("AdministratorID")%></td>
<td  valign="top" class="bodytext"><a href="administrators_list.asp?AdministratorID=<%=rs.Fields("AdministratorID")%>#edit" class="bodytext">Edit</a><br>
<a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="administrators_list.asp?AdministratorID=<%=rs.Fields("AdministratorID")%>&task=d" class="bodytext">Delete</a></td>
	<td class="bodytext" valign="top"><%=rs.Fields("Username")%>&nbsp;</td>
	<td class="bodytext" valign="top"><%if rs.Fields("AccessID") = 1 then response.write("Super Admin") else if rs.Fields("AccessID") = 2 then response.write ("Admin") else if rs.Fields("AccessID") = 3 then response.write("User") end if end if end if%>&nbsp;</td>
</tr> 
<% 
rs.MoveNext
LOOP

Else
%>
<tr><td class="bodytext" colspan="6">No Records Selected</td></tr>
<%	end if %>
</table>
<a name="edit"></a>
<form action="Administrators_list.asp" method="post" name="frmsearch">
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr>
<td>&nbsp;</td><td align="left" class="bodytext">* indicates required field</td>
</tr>
<% if errorFound then%>
<tr>
<td>&nbsp;</td>
<td class="bodytext"><font color="red"><% =errorMessage%></font></td>
</tr>
<% End If %>
<input type="hidden" value="<% =urlAdministratorID%>" name="AdministratorID">
<tr><td valign="top" align="right"><span class="subtitle"><label for="Username">Username</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formUsername%>" name="Username"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Password">Password</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="password" value="<%=formPassword%>" name="Password"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="AccessID">AccessID</label>:</span></td>
<td valign="top"><span class="bodytext"><input type=hidden name=accessID value=1 /> default

<!--
<input type="radio" value="1" name="accessID" <%if formAccessID = 1 then response.write "checked" end if %>>Super Admin - access to Application Setup and Administrators<br>
<input type="radio" value="2" name="accessID" <%if formAccessID = 2 then response.write "checked" end if %>>Admin - access to Application Setup<br>
<input type="radio" value="3" name="accessID" <%if formAccessID = 3 then response.write "checked" end if %>>User - access only to Enter Spray Data, Review, and Reports<br>
-->

</span></td>
</tr>
<tr>
<td>&nbsp;</td>
<td><% IF  urlAdministratorID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>
<%
	set rs = nothing
	EndConnect(conn)
%>
<!--#include file="i_adminfooter.asp" -->
</body>
</html>
