<%Option Explicit%>
<%if not session("login") or not listContains("1", session("accessid")) then
	response.redirect("index.asp")
end if
    Response.Expires = 0
%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_Growers.asp"-->
<!--#include file="i_GrowerLocations.asp"-->
<%
'MODIFIED from Locus Interactive CREATED by Kim Miers on 07/16/2006
Dim errorFound, formError, errorMessage, tempErrorMessage, delError
Dim urlID, formID
Dim conn, rsGrower, counter

'Initialize variables
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"
formID=Request.Form.Item("ID")
urlID=Request.QueryString("ID")
'See if ID was passed through URL or FORM
IF urlID = "" THEN urlID = 0 END IF
IF formID = "" THEN formID = urlID End IF
urlID = formID

REM Get GrowerID
DIM urlGrowerID, formGrowerID

formGrowerID = Request.Form("HF_GrowerID")
urlGrowerID = Request.QueryString("GrowerID")
'Response.Write("<br>urlGrowerID: " & urlGrowerID)
if formGrowerID = "" THEN 
	formGrowerID = urlGrowerID
END IF
if formGrowerID = "" AND session("growerid") <> 0 then
	formGrowerID = session("growerID")
end if
REM end Get GrowerID

IF Request.QueryString("Err") = "DUP" THEN
	errorFound = TRUE
	errorMessage = errorMessage + "<br>Duplicate Location Entered for this Grower. Try Again."
END IF

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
REM KILROY
	Response.Redirect("GrowerUnit.asp")
END IF

'Initialize Form Fields
DIM formUnit, formLicenseNo, formActive, lGrowerID
formUnit = Request.Form.Item("unit")
formLicenseNo = Request.Form.Item("licenseno")
formActive = Request.Form.Item("Active")

'initialize the connection
set conn = Connect()

'IF Request.QueryString("task") = "d" and urlID <> "" THEN
'	DeleteGrowersLocation(urlID)
'	Response.Redirect("GrowerUnit.asp?GrowerID=" & formGrowerID)
'END IF 
IF Request.QueryString("task") = "activate" and urlID <> "" THEN
	conn.execute "exec growerunit$activate " & urlid & ", 1"
	Response.Redirect("GrowerUnit.asp")
END IF 
IF Request.QueryString("task") = "deactivate" and urlID <> "" THEN
	conn.execute "exec growerunit$activate " & urlid & ", 0"
	Response.Redirect("GrowerUnit.asp")
END IF


'Form Was Submitted
'dim item
'for each item in Request.Form
'	Response.Write("<br>Form Items: " & item & " = " & CStr(Request.Form(item)))
'next

IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" OR Request.Form.Item("f_action") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("unit"), "nvarchar", "Supplier", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF

	'Update record
'	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
	IF NOT errorFound AND (Request.Form("update") <> "" OR Request.Form.Item("f_action") = "UPDATE") THEN  
		conn.execute "exec growerunit$upd " & formid & ", '" & formUnit & "'"
		EndConnect(conn)
		lGrowerID = Request.Form.item("HF_Growerid")
		Response.Redirect("GrowerUnit.asp")
	END IF 
	'INSERT
'	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
	IF NOT errorFound AND (Request.Form("insert") <> "" OR Request.Form.Item("f_action") = "INSERT") THEN
'		session("growerID") = Request.Form.item("HF_Growerid")
'Response.Write("<br>Insert new grower location")
'Response.Flush
		dim newRecID
		lGrowerID = Request.Form.item("HF_Growerid")
		conn.execute "exec growerunit$add '" & formUnit & "'"
'		EndConnect(conn)
		'if newRecID = 0 then
		'	Response.Redirect("GrowerUnit.asp?GrowerID=" & lGrowerID & "&Err=DUP")
		'else
			Response.Redirect("GrowerUnit.asp")
		'end if
	END IF 'insert	
END IF 'form submitted 

dim rsEditGrowerLocation
IF formID <> 0 and not errorFound THEN
	set rsEditGrowerLocation = conn.execute("exec growerunit$get " & formid)
	IF NOT rsEditGrowerLocation.eof THEN
		formUnit = rsEditGrowerLocation.Fields("unit")
		'formActive = rsEditGrowerLocation.Fields("active")
	END IF
	rsEditGrowerLocation.Close
	set rsEditGrowerLocation = Nothing
END IF

%>
<html>
<head>
	<title>Growers Orgs</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center>
<tr><td><img src="images/spacer.gif" height="4" width="1" border="0">
<h1> > Grower Orgs</h1><br><img src="images/spacer.gif" height="4" width="1" border="0">
</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">

<!--<tr><td bgcolor="FFFFFF" class="bodytext" align="right"><a href="GrowerLocations.asp#Instructions">Instructions for maintaining your locations.</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td>
</tr>-->

<tr>
<td colspan="2" class="bodytext">
<table width="90%" border="1" cellpadding="2" cellspacing="0">
<% if delerror then %>
<tr>
<td colspan="6" class="bodytext" valign="top"><font color="red"><%= delerrormessage %></font></td>
</tr>
<%	end if
	if errorFound then %>
<tr>
<td colspan="6" class="bodytext" valign="top"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></td>
</tr>
<%	end if %>
<tr bgcolor=#dddddd>
	<td valign="top">&nbsp;</td>
	<td valign="top"><h2>Org</h2></td>
	<td valign="top"><h2>&nbsp;</h2></td>

</tr>
<%
	dim i, rsGrowerLocs
	set rsGrowerLocs = conn.execute("exec growerunit$list") 
	IF not rsGrowerLocs.EOF THEN
		i = 1
		DO WHILE not rsGrowerLocs.eof 
%>
<tr>
	<td class = "bodytext" valign="top"><%=i%></td>
	<td class="bodytext" valign="top"><%=rsGrowerLocs.Fields("unit")%>&nbsp;</td>
	<td valign="top" class="bodytext" bgcolor=#eeeeee><a href="GrowerUnit.asp?ID=<%=rsGrowerLocs.Fields("unitid")%>&GrowerID=<%=formGrowerID%>#edit" class="bodytext">Edit</a> </td>

</tr> 
<% 
			rsGrowerLocs.MoveNext
			i = i + 1
		LOOP
	Else
%>
<tr><td class="bodytext" colspan="6">No Records Selected</td></tr>
<%	end if
	rsGrowerLocs.Close
	set rsGrowerLocs = Nothing %>
</table>
<a name="edit"></a>
<form action="GrowerUnit.asp" method="post">
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr>
<td>&nbsp;</td><td align="left" class="bodytext">* indicates required field</td>
</tr>
<% If errorFound Then %>
<tr>
<td>&nbsp;</td>
<td class="bodytext"><font color="red"><% =errorMessage%></font></td>
</tr>
<% End If %>
<input type="hidden" name="ID" value="<% =urlID%>">
<%	IF  urlID <> 0 THEN %>
	<input type="hidden" name="f_action" value="UPDATE">
<%	ELSE %>
	<input type="hidden" name="f_action" value="INSERT">
<%	END IF %>
<tr>
	<td valign="top" align="right">
		<span class="subtitle"><label for="supplier"><% IF  urlID <> 0 THEN%>Update<% ELSE %>Add<% END IF %> Org</label>:</span>
	</td>
	<td valign="top">
		<span class="bodytext"><input type="text" value="<%=formUnit%>" name="unit"  class="bodytext" size="25" maxlength="150"></span>
	</td>
</tr>

<tr>
	<td>&nbsp;</td>
	<td><% IF  urlID <> 0 THEN %><input type="submit" name="update" value="Update"><% ELSE %><input type="submit" name="insert" value="Insert"><% END IF %>&nbsp;&nbsp;<input type="reset" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>

</td>
</tr>
</table>

<%
	EndConnect(conn)
%>

<!--#include file="i_adminfooter.asp" -->
</body>
</html>
