<!--#include file="include/page_init.asp"-->
<%
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "CACHE-CONTROL", "NO-CACHE"
Response.Buffer = true

if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if
%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_Crop.asp"-->
<!--#include file="i_Varieties.asp"-->
<%
'CREATED by LocusInteractive on 07/21/2005
Dim errorFound, formError, errorMessage, tempErrorMessage, delError
Dim urlID, formID
Dim conn, sql, rs, counter, formSortOrder, rsCrop

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
	Response.Redirect("crops_list.asp")
END IF

'Initialize Form Fields
DIM formCrop, formSiteCategory, formSpecificSite, formActive
formCrop = Request.Form.Item("Crop")
formSiteCategory = Request.Form.Item("PURS_SiteCategory")
formSpecificSite = Request.Form.Item("PURS_SpecificSite")
formActive = Request.Form.Item("Active")
formSortOrder = Request.Form.Item("SortOrder")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsCrop = Server.CreateObject("ADODB.RecordSet")

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
			sql = "UPDATE Crops SET SortOrder =" & thisSort & " WHERE CropID = " & thisID
			conn.execute sql, , 129
		END IF
	NEXT
	' now make the numbers sequential
	sql = "SELECT CropID,SortOrder FROM Crops ORDER BY SortOrder"
	set rs = conn.execute(sql)
	i = 0
	IF not rs.EOF THEN
		DO WHILE not rs.eof		
			i = i + 1
			sql = "UPDATE Crops SET SortOrder = " & i & " WHERE CropID = " & rs.Fields("CropID")
			conn.execute sql, , 129
			rs.MoveNext
		LOOP
	END IF
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("crops_list.asp")
END IF

IF Request.QuerySTring("task") = "d" and urlID <> "" THEN
	DeleteCrop(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("crops_list.asp")
END IF 

IF Request.QuerySTring("task") = "activate" and urlID <> "" THEN
	ActivateCrop(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("crops_list.asp")
END IF 

IF Request.QuerySTring("task") = "deactivate" and urlID <> "" THEN
	DeActivateCrop(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("crops_list.asp")
END IF 


'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("Crop"), "nvarchar","Crop", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Active"), "bit","Active", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		urlID = UPDATECrop(formID, formCrop, formSiteCategory, formSpecificSite)
		Response.Redirect("crops_list.asp")
		'END UPDATE
	END IF

	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlID = InsertCrop(formCrop, formSiteCategory, formSpecificSite)
		EndConnect(conn)
		set rs = nothing
		Response.Redirect("crops_list.asp")
	END IF 'insert	

END IF 'form submitted 

IF formID <> 0 and not errorFound THEN
	set rs = GetCropByCropID(formID)
	IF NOT rs.eof THEN
		formCrop = rs.Fields("Crop")
		formSiteCategory = rs.Fields("PURS_SiteCategory")
		formSpecificSite = rs.Fields("PURS_SpecificSite")
		formActive = rs.Fields("Active")
	END IF
END IF%>
<html>
<head>
	<title>Crop List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Crops</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
	<tr>
		<td bgcolor="FFFFFF" class="bodytext"><br>
			Add/edit/delete Crop.<br><br>
			<a href="crops_list.asp#edit">Add Crop</a>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="bodytext">
			<table width="90%" border="1" cellpadding="2" cellspacing="0">
<% if  delerror then%>
				<tr>
					<td colspan="7" class="bodytext" valign="top">
						<font color="red"><%= delerrormessage %></font>
					</td>
				</tr>
<%
	end if
	if  errorFound then %>
				<tr>
					<td colspan="7" class="bodytext" valign="top">
						<font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font>
					</td>
				</tr>
<% end if %>
				<tr>
					<td  valign="top">&nbsp;</td>
					<td  valign="top"><h2>Edit</h2></td>
						<form action="crops_list.asp" method="post">
					<td valign="top"><h2>Sort Order</h2><br><input type="submit" name="submit" value="resort"></td>
					<td valign="top"><h2>Activate</h2></td>
					<td valign="top"><h2>Crop</h2></td>
					<td valign="top"><h2>PURS Site Category</h2></td>
					<td valign="top"><h2>PURS Specific Site</h2></td>
				</tr>
<%
	set rs = GetAllCrops()
	i = 0

	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			i = i + 1 %>
				<tr>
					<td class = "bodytext" valign="top"><%=i%></td>
					<td valign="top" class="bodytext">
						<a href="crops_list.asp?ID=<%=rs.Fields("CropID")%>#edit" class="bodytext">Edit Crop</a><br><br>
						<a href="varieties_list.asp?CropID=<%=rs.Fields("CropID")%>">Edit Varieties</a>
<!--- <a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="crops_list.asp?ID=<%=rs.Fields("CropID")%>&task=d" class="bodytext">Delete</a> --->
					</td>
					<td class="bodytext" nowrap>
						<input type="text" length="3" size="1" name="ord<%=rs.Fields("CropID")%>" value="<%=rs.Fields("SortOrder")%>">
					</td>
					<td class="bodytext">
						<% IF rs.Fields("Active") =  0 then%>Not Active<br><a href="crops_list.asp?task=activate&ID=<%=rs.Fields("CropID")%>" onclick="javascript: return confirm('Are you sure you want to activate this record?');">Make Active</a><%else%>Active<br><a href="crops_list.asp?task=deactivate&ID=<%=rs.Fields("CropID")%>" onclick="javascript: return confirm('Are you sure you want to DeActivate this record?');">Make InActive</a><%end if%>
					</td>
					<td class="bodytext" valign="top"><%=rs.Fields("Crop")%><br>
						&nbsp;&nbsp;<strong>Varieties</strong><br>
	<%
'may have better performance getting ALL Varieties, then filter recordset.
			set rsCrop = GetAllVarietiesByCropID(rs.Fields("CropID"))
			IF not rsCrop.EOF THEN
				DO WHILE not rsCrop.eof 
					Response.write("&nbsp;&nbsp;" & rsCrop.Fields("Variety") & "<br>")
					rsCrop.MoveNext
				LOOP
			END IF
%>
					</td>
					<td class="bodytext" valign="top"><%=rs.Fields("PURS_SiteCategory")%></td>
					<td class="bodytext" valign="top"><%=rs.Fields("PURS_SpecificSite")%></td>
				</tr> 
<%
			rs.MoveNext
		LOOP
	ELSE
%>
				<tr><td class="bodytext" colspan="7">No Records Selected</td></tr>
<%
	END IF %>
			</table>
<a name="edit"></a>
<form action="crops_list.asp" method="post" name="frmsearch">
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
<input type="hidden" value="<% =urlID%>" name="ID">
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="Crop">Crop</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formCrop%>" name="Crop" class="bodytext" size="25" maxlength="50"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="PURS_SiteCategory">PURS Site Category</label>:</span></td>
		<td valign="top"><span class="bodytext">
	<!-- Could make the make/model javascript -->
			<select id="PURS_SiteCategory" name="PURS_SiteCategory" class="bodytext">
				<option value="Agriculture" <%if trim(formSiteCategory) = "Agriculture" then Response.write("selected")%>>Agriculture</option>
				<option value="Forestry" <%if trim(formSiteCategory) = "Forestry" then Response.write("selected")%>>Forestry</option>
			</select>
			</span>
		</td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="PURS_SpecificSite">PURS Specific Site</label>:</span></td>
		<td valign="top"><span class="bodytext">
	<!-- Could make the make/model javascript -->
			<select id="PURS_SpecificSite" name="PURS_SpecificSite" class="bodytext">
				<option value="Field Crops" <%if trim(formSpecificSite) = "Field Crops" then Response.write("selected")%>>Agriculture - Field Crops</option>
				<option value="Fruits/Nuts" <%if trim(formSpecificSite) = "Fruits/Nuts" then Response.write("selected")%>>Agriculture - Fruits/Nuts</option>
				<option value="Livestock/Poultry" <%if trim(formSpecificSite) = "Livestock/Poultry" then Response.write("selected")%>>Agriculture - Livestock/Poultry</option>
				<option value="Nursery/Christmas Trees" <%if trim(formSpecificSite) = "Nursery/Christmas Trees" then Response.write("selected")%>>Agriculture - Nursery/Christmas Trees</option>
				<option value="Oil Crops" <%if trim(formSpecificSite) = "Oil Crops" then Response.write("selected")%>>Agriculture - Oil Crops</option>
				<option value="Pasture/Forage/Hay" <%if trim(formSpecificSite) = "Pasture/Forage/Hay" then Response.write("selected")%>>Agriculture - Pasture/Forage/Hay</option>
				<option value="Seed Crops" <%if trim(formSpecificSite) = "Seed Crops" then Response.write("selected")%>>Agriculture - Seed Crops</option>
				<option value="Vegetables" <%if trim(formSpecificSite) = "Vegetables" then Response.write("selected")%>>Agriculture - Vegetables</option>
				<option value="Other" <%if trim(formSpecificSite) = "Other" then Response.write("selected")%>>Agriculture - Other</option>
			</select>
			</span>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><% IF  urlID <> 0 THEN %>
			<input type="submit" name="update" value="Update" class="bodytext">
	<% ELSE %>
			<input type="submit" name="insert" value="Insert" class="bodytext">
	<% END IF %>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
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
