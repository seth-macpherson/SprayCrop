<%Option Explicit%>
<%if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_packerusers.asp"-->
<%
'CREATED by LocusInteractive on 08/02/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlpackeruserID,formpackeruserID,formpackerID
Dim conn,sql,rs,counter

'Initialize variables
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"
formpackeruserID=Request.Form.Item("packeruserID")
urlpackeruserID=Request.QueryString("packeruserID")

'See if ID was passed through URL or FORM
IF urlpackeruserID = "" THEN urlpackeruserID = 0 END IF
IF formpackeruserID = "" THEN formpackeruserID = urlpackeruserID End IF
urlpackeruserID = formpackeruserID

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Response.Redirect("packerusers_list.asp")
END IF

'Initialize Form Fields
DIM formUsername,formPassword,formAccessID
formUsername = Request.Form.Item("Username")
formPassword = Request.Form.Item("Password")
formAccessID = Request.Form.Item("AccessID")
formPackerID = request.Form.Item("PackerID")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")

IF Request.QuerySTring("task") = "d" and urlpackeruserID <> "" THEN
	Deletepackerusers(urlpackeruserID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("packerusers_list.asp")
END IF 
IF Request.QuerySTring("task") = "activate" and urlpackeruserID <> "" THEN
	Activatepackerusers(urlpackeruserID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("packerusers_list.asp")
END IF 
IF Request.QuerySTring("task") = "deactivate" and urlpackeruserID <> "" THEN
	DeActivatepackerusers(urlpackeruserID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("packerusers_list.asp")
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
	IF NOT ValidateDatatype(Request.Form.Item("PackerID"), "int","PackerID", true) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		urlpackeruserID = UPDATEpackerusers(formpackeruserID,formUsername,formPassword,formpackerid,formAccessID)
		set rs = nothing
		EndConnect(conn)
		Response.Redirect("packerusers_list.asp")
		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlpackeruserID = Insertpackerusers(formUsername,formPassword,formpackerid,formAccessID)
		set rs = nothing
		EndConnect(conn)
		Response.Redirect("packerusers_list.asp")
	END IF 'insert	
END IF 'form submitted 

IF formpackeruserID  <>  0 and not errorFound THEN
set rs = GetpackerusersByID(formpackeruserID)
	IF NOT rs.eof THEN
		formUsername = rs.Fields("Username")
		formPassword = rs.Fields("Password")
		formPackerID = rs.Fields("PackerID")
		formAccessID = rs.Fields("AccessID")
	END IF
END IF%>
<html>
<head>
	<title>packerusers List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Packer Users</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="ffffff" class="bodytext"><br />
<a href="packerusers_list.asp#edit">Add Packer User</a>
<br />&nbsp;
</td>
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
<tr bgcolor=#dddddd>
    <!--<td  valign="top">&nbsp;</td>-->

	<td  valign="top"><h2>Username</h2></td>
	<td  valign="top"><h2>Packer</h2></td>
    <td  valign="top"><h2>&nbsp;</h2></td>

</tr>
<%
	set rs = GetAllpackerusers()
%>
<%IF not rs.EOF THEN
DO WHILE not rs.eof %>
<tr>
<!--<td class = "bodytext" valign="top"><%=rs.Fields("packeruserID")%></td>-->
	<td class="bodytext" valign="top"><%=rs.Fields("Username")%>&nbsp;</td>
	<td class="bodytext" valign="top"><%=rs.Fields("packername") %></td>
    <td bgcolor=#eeeeee valign="top" class="bodytext"><a href="packerusers_list.asp?packeruserID=<%=rs.Fields("packeruserID")%>#edit" class="bodytext">Edit</a><br>
    <a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="packerusers_list.asp?packeruserID=<%=rs.Fields("packeruserID")%>&task=d" class="bodytext">Delete</a></td>

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
<form action="packerusers_list.asp" method="post" name="frmsearch">
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
<input type="hidden" value="<% =urlpackeruserID%>" name="packeruserID">
<tr><td valign="top" align="right"><span class="subtitle"><label for="Username">Username</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formUsername%>" name="Username"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Password">Password</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="password" value="<%=formPassword%>" name="Password"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>

<%if listContains("1", session("accessid")) then %> 
<tr><td valign="top" align="right"><span class="subtitle"><label for="Packer">Packer</label>:</span></td><td>

    <%
	    dim rsSelect: set rsSelect = GetAllPackers()
    %>
    <SELECT name="PackerID" style="background-color:beige;">
    <%IF session("packerid") = 0 THEN%>
    <option value=""></option>
    <%
    END IF
    IF not rsSelect.EOF THEN
    DO WHILE not rsSelect.eof 
    %>
    <option value="<%	response.write(rsSelect.Fields("PackerID"))%>" 
    <%if trim(formPackerID) = trim(rsSelect.Fields("PackerID")) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("PackerName")%></option>
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
<td valign="top"><span class="bodytext"><input type=hidden name=accessID value=2 /> default

<!--
<input type="radio" value="1" name="accessID" <%if formAccessID = 1 then response.write "checked" end if %>>Super Admin - access to Application Setup and packerusers<br>
<input type="radio" value="2" name="accessID" <%if formAccessID = 2 then response.write "checked" end if %>>Admin - access to Application Setup<br>
<input type="radio" value="3" name="accessID" <%if formAccessID = 3 then response.write "checked" end if %>>User - access only to Enter Spray Data, Review, and Reports<br>
-->

</span></td>
</tr>
<%else %> 
<tr><td colspan=2 height=1><input type=hidden name=accessID value=2 /><input type=hidden name=packerid value=<%=session("packerid") %> /></td></tr>
<%end if %>

<tr>
<td>&nbsp;</td>
<td><% IF  urlpackeruserID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
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
