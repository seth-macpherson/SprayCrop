<%Option Explicit%>
<%if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayList.asp"-->
<!--#include file="i_Units.asp"-->
<%
'CREATED by LocusInteractive on 07/21/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlID,formID
Dim conn,sql,rs,counter,rsSelect,onloadstring
onloadstring = "if (window.focus)self.focus();"

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

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	onloadstring = "javascript:window.close();"
END IF


'Initialize Form Fields
DIM formName,formUnitID,formREI,formPHI,formMAXUseApp,formMAXUseSeason
formName = Request.Form.Item("Name")
formUnitID = Request.Form.Item("UnitID")
formREI = Request.Form.Item("REI")
formPHI = Request.Form.Item("PHI")
formMAXUseApp = Request.Form.Item("MAXUseApp")
formMAXUseSeason = Request.Form.Item("MAXUseSeason")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")

'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("Name"), "nvarchar","Name", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("UnitID"), "int","Units", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("REI"), "nvarchar","REI", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("PHI"), "nvarchar","PHI", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("MAXUseApp"), "nvarchar","MAXUseApp", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		
	IF NOT ValidateDatatype(Request.Form.Item("MAXUseSeason"), "nvarchar","MAXUseSeason", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		

	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlID = InsertSprayList(formName,formUnitID,formREI,formPHI,formMAXUseApp,formMAXUseSeason)
		onloadstring = "javascript:window.opener.refreshdata(" & urlID  & ");window.close();"
	END IF 'insert	
END IF 'form submitted 

%>
<html>
<head>
	<title>SprayList List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0"  onload="<%=onloadstring%>">
<table width="594" border="1" cellspacing="0" cellpadding=" 0" bordercolor="#013166" bgcolor="#beige"><tr><td><img src="images/spacer.gif" height="4" width="1" border="0"><br>
&nbsp;&nbsp;<h1>Spray List<h1><br><img src="images/spacer.gif" height="4" width="1" border="0"></td></tr></table>
<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="FFFFFF" class="bodytext"><br>
<form action="pu_SprayList_list.asp" method="post" name="frmsearch">
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
<input type="hidden" value="<% =urlID%>" name="ID">
<tr><td valign="top" align="right"><span class="subtitle"><label for="Name">Name</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formName%>" name="Name"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Units">Units</label>:</span></td>
<td valign="top">
<%
	set rsSelect = GetAllUnits()
%>
<SELECT name="UnitID">
<option value="">---Unit---</option>
<%
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("UnitID")%>"<%if trim(formUnitID) = trim(rsSelect.Fields("UnitID")) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Unit")%></option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="REI">REI</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formREI%>" name="REI"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="PHI">PHI</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formPHI%>" name="PHI"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="MAXValue">MAX Use/Application</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formMAXUseApp%>" name="MAXUseApp"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="MAXValue">MAX Use/Season</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formMAXUseSeason%>" name="MAXUseSeason"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr>
<td>&nbsp;</td>
<td><% IF  urlID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>
<%
	EndConnect(conn)
	set rs = nothing
	set rsSelect = nothing
%>
<!--#include file="i_adminfooter.asp" -->
</body>
</html>
