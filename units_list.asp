<%Option Explicit%>
<%if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_Units.asp"-->
<%
'CREATED by LocusInteractive on 07/21/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlID,formID
Dim conn,sql,rs,counter

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
	Response.Redirect("units_list.asp")
END IF

'Initialize Form Fields
DIM formUnit, formPURSUnit, formActive, formSortOrder
formUnit = Request.Form.Item("Unit")
formPURSUnit = Request.Form.Item("PURSUnit")
formActive = Request.Form.Item("Active")
formSortOrder = Request.Form.Item("SortOrder")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
'RESORT
IF Request.Form.Item("Submit") = "resort" THEN
	DIM thisSort,thisItem,thisID,i
	' update each record with the sort number assigned
	For Each thisItem IN Request.Form
		IF left(thisItem,3) = "ord" THEN
			thisSort = Request.Form.Item(thisItem)
			thisID = right(thisItem,len(thisItem)-3)
			IF NOT IsNumeric(thisSort) THEN
				thisSort = 0
			END IF
			sql = "UPDATE Units SET SortOrder =" & thisSort & " WHERE UnitID = " & thisID
			conn.execute sql, , 129
		END IF
	NEXT
	' now make the numbers sequential
	sql = "SELECT UnitID,SortOrder FROM Units ORDER BY SortOrder"
	set rs = conn.execute(sql)
	i = 0
	IF not rs.EOF THEN
		DO WHILE not rs.eof		
			i = i + 1
			sql = "UPDATE Units SET SortOrder = " & i & " WHERE UnitID = " & rs.Fields("UnitID")
			conn.execute sql, , 129
			rs.MoveNext
		LOOP
	END IF
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("units_list.asp")
END IF
	
IF Request.QuerySTring("task") = "d" and urlID <> "" THEN
	DeleteUnits(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("units_list.asp")
END IF 
IF Request.QuerySTring("task") = "activate" and urlID <> "" THEN
	ActivateUnits(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("units_list.asp")
END IF 
IF Request.QuerySTring("task") = "deactivate" and urlID <> "" THEN
	DeActivateUnits(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("units_list.asp")
END IF 


'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("Unit"), "nvarchar","Unit", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Active"), "bit","Active", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		

	'UPDATE
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		urlID = UPDATEUnits(formID,formUnit,formPURSUnit)
		Response.Redirect("units_list.asp")
	END IF 

	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlID = InsertUnits(formUnit,formPURSUnit)
		Response.Redirect("units_list.asp")
	END IF 'insert	

END IF 'form submitted 

IF formID  <>  0 and not errorFound THEN
set rs = GetUnitsByID(formID)
	IF NOT rs.eof THEN
		formUnit = rs.Fields("Unit")
		formPURSUnit = rs.Fields("PURSUnit")
		formActive = rs.Fields("Active")
		formSortOrder = rs.Fields("SortOrder")
	END IF
END IF%>
<html>
<head>
	<title>Units List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Units</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="FFFFFF" class="bodytext"><br>
Add/edit/delete Units.<br><br>
<a href="units_list.asp#edit">Add Units</a></td>
</tr>
<tr>
<td colspan="2" class="bodytext">
<table width="90%" border="1" cellpadding="2" cellspacing="0">
<%
	if  delerror then
%>
<tr>
<td colspan="6" class="bodytext" valign="top"><font color="red"><%= delerrormessage %></font></td>
</tr>
<%	end if
	if  errorFound then
%>
<tr>
<td colspan="6" class="bodytext" valign="top"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></td>
</tr>
<% end if %>
<tr>
<td  valign="top">&nbsp;</td>
<td  valign="top"><h2>Edit</h2></td>


<form action="units_list.asp" method="post">
	<td valign="top"><h2>Sort Order</h2><br><input type="submit" name="submit" value="resort"></td>
	<td valign="top"><h2>Activate</h2></td>
	<td valign="top"><h2>Unit</h2></td>
	<td valign="top"><h2>PURS Unit</h2></td>
</tr>
<%
	set rs = GetAllUnits()
	i = 0
	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			i = i + 1
%>
<tr>
<td class = "bodytext" valign="top"><%=i%></td>
<td  valign="top" class="bodytext"><a href="units_list.asp?ID=<%=rs.Fields("UnitID")%>#edit" class="bodytext">Edit</a><br>
<!--- <a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="units_list.asp?ID=<%=rs.Fields("UnitID")%>&task=d" class="bodytext">Delete</a> ---></td>
<td class="bodytext" nowrap><input type="text" length="3" size="1" name="ord<%=rs.Fields("UnitID")%>" value="<%=rs.Fields("SortOrder")%>"></td>
<td class="bodytext"><% IF rs.Fields("Active") =  0 then%>Not Active<br><a href="units_list.asp?task=activate&ID=<%=rs.Fields("UnitID")%>" onclick="javascript: return confirm('Are you sure you want to activate this record?');">Make Active</a><%else%>Active<br><a href="units_list.asp?task=deactivate&ID=<%=rs.Fields("UnitID")%>" onclick="javascript: return confirm('Are you sure you want to DeActivate this record?');">Make InActive</a><%end if%></td>
	<td class="bodytext" valign="top"><%=rs.Fields("Unit")%>&nbsp;</td>
	<td class="bodytext" valign="top"><%=rs.Fields("PURSUnit")%>&nbsp;</td>
</tr>
<% 
			rs.MoveNext
		LOOP
	Else
%>
<tr><td class="bodytext" colspan="6">No Records Selected</td></tr>
<%	end if %>
</table>
</form>
<a name="edit"></a>
<form action="Units_list.asp" method="post" name="frmsearch">
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
<tr><td valign="top" align="right"><span class="subtitle"><label for="Unit">Unit</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formUnit%>" name="Unit"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="PURSUnit">PURS Unit</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formPURSUnit%>" name="PURSUnit"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr>
<td>&nbsp;</td>
<td><% IF  urlID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>

</td></tr></table>
<%
	EndConnect(conn)
	set rs = nothing
%>
<!--#include file="i_adminfooter.asp" -->
</body>
</html>
