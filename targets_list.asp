<%
Option Explicit

if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if
%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_Target.asp"-->
<%
'CREATED by LocusInteractive on 07/21/2005
'modified kmiers 11/17/2006 cleanup code, add PURS_Target
Dim errorFound, errorMessage, tempErrorMessage, delError
Dim urlID, formID
Dim conn, sql, rs

'Initialize variables
errorFound = FALSE
errorMessage = "The following errors have occurred:"

formID=Request.Form.Item("ID")
urlID=Request.QueryString("ID")

'See if ID was passed through URL or FORM
IF urlID = "" THEN urlID = 0 END IF
IF formID = "" THEN formID = urlID End IF
urlID = formID

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Response.Redirect("targets_list.asp")
END IF

'Initialize Form Fields
DIM formTarget, formPURSTarget, formActive, formSortOrder
formTarget = Request.Form.Item("Target")
formPURSTarget = Request.Form.Item("PURSTarget")
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
			sql = "UPDATE Targets SET SortOrder =" & thisSort & " WHERE TargetID = " & thisID
			conn.execute sql, , 129
		END IF
	NEXT
	' now make the numbers sequential
	sql = "SELECT TargetID,SortOrder FROM Targets ORDER BY SortOrder"
	set rs = conn.execute(sql)
	i = 0
	IF not rs.EOF THEN
		DO WHILE not rs.eof		
			i = i + 1
			sql = "UPDATE Targets SET SortOrder = " & i & " WHERE TargetID = " & rs.Fields("TargetID")
			conn.execute sql, , 129
			rs.MoveNext
		LOOP
	END IF
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("targets_list.asp")
END IF

IF Request.QueryString("task") = "d" and urlID <> "" THEN
	DeleteTarget(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("targets_list.asp")
END IF 
IF Request.QueryString("task") = "activate" and urlID <> "" THEN
	ActivateTarget(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("targets_list.asp")
END IF 
IF Request.QueryString("task") = "deactivate" and urlID <> "" THEN
	DeActivateTarget(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("targets_list.asp")
END IF 


'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("Target"), "nvarchar","Target", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Active"), "bit","Active", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		

	'UPDATE
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		urlID = UPDATETarget(formID,formTarget,formPURSTarget)
		EndConnect(conn)
		set rs = nothing
		Response.Redirect("targets_list.asp")
	END IF 

	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlID = InsertTarget(formTarget,formPURSTarget)
		EndConnect(conn)
		set rs = nothing
		Response.Redirect("targets_list.asp")
	END IF 'insert	

END IF 'form submitted 

IF formID  <>  0 and not errorFound THEN
	set rs = GetTargetByID(formID)
	IF NOT rs.eof THEN
		formTarget = rs.Fields("Target")
		formPURSTarget = rs.Fields("PURS_Target")
		formActive = rs.Fields("Active")
		formSortOrder = rs.Fields("SortOrder")
	END IF
END IF%>
<html>
<head>
	<title>Target List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Targets</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
	<tr>
		<td bgcolor="FFFFFF" class="bodytext"><br>
Add/edit/delete Target.<br><br>
<a href="targets_list.asp#edit">Add Target</a></td>
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
<%
	end if
	if  errorFound then
%>
				<tr>
					<td colspan="6" class="bodytext" valign="top"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></td>
				</tr>
<%
	end if
%>
				<tr>
					<td valign="top">&nbsp;</td>
					<td valign="top"><h2>Edit</h2></td>
<form action="targets_list.asp" method="post">
					<td valign="top"><h2>Sort Order</h2><br><input type="submit" name="submit" value="resort"></td>
					<td valign="top"><h2>Activate</h2></td>
					<td valign="top"><h2>Target</h2></td>
					<td valign="top"><h2>PURS Target</h2></td>
					<td valign="top"><h2>SortOrder</h2></td>
				</tr>
<%
	set rs = GetAllTarget()
	i = 0

	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			i = i + 1
%>
				<tr>
					<td valign="top" class="bodytext"><%=i%></td>
					<td valign="top" class="bodytext"><a href="targets_list.asp?ID=<%=rs.Fields("TargetID")%>#edit" class="bodytext">Edit</a><br>
<!--- <a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="targets_list.asp?ID=<%=rs.Fields("TargetID")%>&task=d" class="bodytext">Delete</a> ---></td>
					<td class="bodytext" nowrap><input type="text" length="3" size="1" name="ord<%=rs.Fields("TargetID")%>" value="<%=rs.Fields("SortOrder")%>"></td>
					<td class="bodytext"><% IF rs.Fields("Active") = 0 then%>Not Active<br><a href="targets_list.asp?task=activate&ID=<%=rs.Fields("TargetID")%>" onclick="javascript: return confirm('Are you sure you want to activate this record?');">Make Active</a><%else%>Active<br><a href="targets_list.asp?task=deactivate&ID=<%=rs.Fields("TargetID")%>" onclick="javascript: return confirm('Are you sure you want to DeActivate this record?');">Make InActive</a><%end if%></td>
					<td class="bodytext" valign="top"><%=rs.Fields("Target")%>&nbsp;</td>
					<td class="bodytext" valign="top"><%=rs.Fields("PURS_Target")%>&nbsp;</td>
					<td class="bodytext" valign="top"><%=rs.Fields("SortOrder")%>&nbsp;</td>
				</tr>
<%
			rs.MoveNext
		LOOP
	ELSE
%>
				<tr><td class="bodytext" colspan="6">No Records Selected</td></tr>
<%
	END IF
%>
			</table>
</form>
<a name="edit"></a>
<form action="targets_list.asp" method="post" name="frmsearch">
<input type="hidden" value="<%=urlID%>" name="ID">
<table width="500" border="0" cellpadding="2" cellspacing="0">
	<tr>
		<td>&nbsp;</td>
		<td align="left" class="bodytext">* indicates required field</td>
	</tr>
<% If errorFound Then %>
	<tr>
		<td>&nbsp;</td>
		<td class="bodytext"><font color="red"><% =errorMessage%></font></td>
	</tr>
<% End If %>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="Target">Target</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formTarget%>" name="Target" class="bodytext" size="25" maxlength="150"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="PURSTarget">PURS Target</label>:</span></td>
		<td valign="top">
			<span class="bodytext">
				<select name="PURSTarget">
<%	if formPURSTarget = "Big game repellant" then %>
					<option value="Big game repellant" class="bodytext" selected>Big game repellant</option>
<%	else %>
					<option value="Big game repellant" class="bodytext">Big game repellant</option>
<%	end if
	if formPURSTarget = "Bird control" then %>
					<option value="Bird control" class="bodytext" selected>Bird control</option>
<%	else %>
					<option value="Bird control" class="bodytext">Bird control</option>
<%	end if
	if formPURSTarget = "Desiccation/defoliation" then %>
					<option value="Desiccation/defoliation" class="bodytext" selected>Desiccation/defoliation</option>
<%	else %>
					<option value="Desiccation/defoliation" class="bodytext">Desiccation/defoliation</option>
<%	end if
	if formPURSTarget = "Disease control" then %>
					<option value="Disease control" class="bodytext" selected>Disease control</option>
<%	else %>
					<option value="Disease control" class="bodytext">Disease control</option>
<%	end if
	if formPURSTarget = "Fish control" then %>
					<option value="Fish control" class="bodytext" selected>Fish control</option>
<%	else %>
					<option value="Fish control" class="bodytext">Fish control</option>
<%	end if
	if formPURSTarget = "Insect control" then %>
					<option value="Insect control" class="bodytext" selected>Insect control</option>
<%	else %>
					<option value="Insect control" class="bodytext">Insect control</option>
<%	end if
	if formPURSTarget = "Marine-fouling organism control" then %>
					<option value="Marine-fouling organism control" class="bodytext" selected>Marine-fouling organism control</option>
<%	else %>
					<option value="Marine-fouling organism control" class="bodytext">Marine-fouling organism control</option>
<%	end if
	if formPURSTarget = "Moss control" then %>
					<option value="Moss control" class="bodytext" selected>Moss control</option>
<%	else %>
					<option value="Moss control" class="bodytext">Moss control</option>
<%	end if
	if formPURSTarget = "Plant growth regulation" then %>
					<option value="Plant growth regulation" class="bodytext" selected>Plant growth regulation</option>
<%	else %>
					<option value="Plant growth regulation" class="bodytext">Plant growth regulation</option>
<%	end if
	if formPURSTarget = "Predator control" then %>
					<option value="Predator control" class="bodytext" selected>Predator control</option>
<%	else %>
					<option value="Predator control" class="bodytext">Predator control</option>
<%	end if
	if formPURSTarget = "Research" then %>
					<option value="Research" class="bodytext" selected>Research</option>
<%	else %>
					<option value="Research" class="bodytext">Research</option>
<%	end if
	if formPURSTarget = "Rodent control" then %>
					<option value="Rodent control" class="bodytext" selected>Rodent control</option>
<%	else %>
					<option value="Rodent control" class="bodytext">Rodent control</option>
<%	end if
	if formPURSTarget = "Slug control" then %>
					<option value="Slug control" class="bodytext" selected>Slug control</option>
<%	else %>
					<option value="Slug control" class="bodytext">Slug control</option>
<%	end if
	if formPURSTarget = "Weed control" then %>
					<option value="Weed control" class="bodytext" selected>Weed control</option>
<%	else %>
					<option value="Weed control" class="bodytext">Weed control</option>
<%	end if
	if formPURSTarget = "Wood preservation" then %>
					<option value="Wood preservation" class="bodytext" selected>Wood preservation</option>
<%	else %>
					<option value="Wood preservation" class="bodytext">Wood preservation</option>
<%	end if
	if formPURSTarget = "Other" then %>
					<option value="Other" class="bodytext" selected>Other</option>
<% else %>
					<option value="Other" class="bodytext">Other</option>
<% end if %>
				</select>
			</span>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><% IF urlID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
	</tr>
</table>
</form>
<%
	EndConnect(conn)
	set rs = nothing
%>

</td></tr></table>


<!--#include file="i_adminfooter.asp" -->
</body>
</html>