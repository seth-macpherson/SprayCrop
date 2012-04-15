<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="include/i_data.asp"-->
<%
'CREATED by LocusInteractive on 08/04/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount
Dim urlSprayYearID,formSprayYearID
Dim conn,sql,rs,counter

'Initialize variables
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"
formSprayYearID=Request.Form.Item("SprayYearID")
urlSprayYearID=Request.QueryString("SprayYearID")

'See if ID was passed through URL or FORM
IF urlSprayYearID = "" THEN urlSprayYearID = 0 END IF
IF formSprayYearID = "" THEN formSprayYearID = urlSprayYearID End IF
urlSprayYearID = formSprayYearID

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Server.Transfer("SprayYears_ins_thanks.asp")
END IF

'Initialize Form Fields
DIM formSprayYear,formActive
formSprayYear = Request.Form.Item("SprayYear")
formActive = Request.Form.Item("Active")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")

'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("SprayYear"), "varchar","SprayYear", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Active"), "bit","Active", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN 
		sql = "UPDATE Test SET	 SprayYear = '" & EscapeQuotes(formSprayYear) & "',Active = " & formActive & " where SprayYearID = " & formSprayYearID
		conn.execute sql, , 129
		Server.Transfer("SprayYears_ins_thanks.asp")
		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		sql = "INSERT INTO SprayYears(SprayYear,Active)VALUES('" & formSprayYear & "'," formActive " & formActive & ")"
		conn.execute sql, , 129
		sql = "SELECT MAX(SprayYearID) AS insertid FROM SprayYears"
		set rs = conn.execute(sql)
		urlSprayYearID = rs.Fields("insertid")
		Server.Transfer("SprayYears_ins_thanks.asp")
	END IF 'insert
END IF 'form submitted 

IF urlSprayYearID  <>  0 and not errorFound THEN
	sql = "SELECT SprayYears.SprayYearID, SprayYears.SprayYear, SprayYears.Active FROM SprayYears WHERE SprayYears.SprayYearID = " & urlSprayYearID & " ORDER BY  SprayYears.SprayYearID"
	set rs = conn.execute(sql)
	IF NOT rs.eof THEN
		formSprayYear = rs.Fields("SprayYear")
		formActive = rs.Fields("Active")
	END IF
END IF%>
<form action="SprayYears_ins.asp" method="post" name="frmsearch">
<table width="100%" border="0" cellpadding="2" cellspacing="0">
<tr>
<td>&nbsp;</td><td align="left" class="bodytext">* indicates required field</td>
</tr>
<% if errorFound then%>
<tr>
<td>&nbsp;</td>
<td class="bodytext"><font color="red"><% =errorMessage%></font></td>
</tr>
<% End If %>
<input type="hidden" value="<% =urlSprayYearID%>" name="SprayYearID">
<tr><td valign="top" align="right"><span class="subtitle"><label for="SprayYear">SprayYear</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formSprayYear%>" name="SprayYear"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr>
<td valign="top" align="right"><span class="subtitle"><label for="Active">Active:</label></span></td>
<td valign="top"><span class="bodytext"><span class="bodytext"><input type="radio" value="1" name="Active" <% if ListContains(formActive,"True") OR ListContains(formActive,"1")  THEN %>Checked<% END IF %>>YES&nbsp;&nbsp;<input type="radio" value="0" name="Active" <% if ListContains(formActive,"False") OR ListContains(formActive,"0")  THEN %>Checked<% END IF %>>NO</span></td>
</tr>
<tr>
<td>&nbsp;</td>
<td><% IF  urlSprayYearID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>
</cfoutput>
