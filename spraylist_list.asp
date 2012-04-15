<%Option Explicit%>
<%if not session("login") or not listContains("1,2,3", session("accessid")) then
	response.redirect("index.asp")
end if%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayList.asp"-->
<!--#include file="i_SprayRecord.asp"-->
<!--#include file="i_Units.asp"-->
<!--#include file="i_Crop.asp"-->
<%
'CREATED by LocusInteractive on 07/21/2005
'MODIFIED kalanmiers 12/4/2006
'	added fields for PURS reporting, general cleanup!
Dim errorFound, formError, errorMessage, tempErrorMessage, delErrorFound, delErrorMessage, _
	urlID, formID, iReportable, _
	conn, sql, rs, rsSelect, rsCrop, sAction, bAddOnly, onloadstring

'Initialize variables
onloadstring = ""
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"
formID=Request.Form.Item("ID")
urlID=Request.QueryString("ID")

'See if ID was passed through URL or FORM
IF urlID = "" THEN urlID = 0 END IF
IF formID = "" THEN formID = urlID End IF
urlID = formID

sAction = trim(Request("Action"))
bAddOnly = false
if sAction = "AddOnly" then
	'onloadstring = "if (window.focus)self.focus();"
	bAddOnly = true
end if

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	if bAddOnly then
		onloadstring = "location.href='spraylist_list.asp';"
	else
		Response.Redirect("spraylist_list.asp")
	end if
END IF

iReportable = 0

if Request.QueryString("PURS") = 1 then
	iReportable = 1
elseif Request.QueryString("PURS") = 2 then
	iReportable = 2
end if


'Initialize Form Fields
DIM formName, formPURSName, formPURSEPANumber, formPURSReport, _
	formUnitID, formREI, formPHI, formMAXUseApp, formMAXUseSeason, formActiveInd,_
	formReEntryIntervalHours, formReEntryIntervalDays, formPreharvestInterval, formCropIDs

formName = Request.Form.Item("Name")
formPURSName = Request.Form.Item("PURS_Name")
formPURSEPANumber = Request.Form.Item("PURS_EPA_Number")
formPURSReport = Request.Form.Item("PURS_Report")
formUnitID = Request.Form.Item("UnitID")
formREI = Request.Form.Item("REI")
formPHI = Request.Form.Item("PHI")
formMAXUseApp = Request.Form.Item("MAXUseApp")
formMAXUseSeason = Request.Form.Item("MAXUseSeason")
formActiveInd = Request.Form.Item("ActiveInd")
formReEntryIntervalHours = Request.Form.Item("ReEntryIntervalHours")
formReEntryIntervalDays = Request.Form.Item("ReEntryIntervalDays")
formPreharvestInterval = Request.Form.Item("PreharvestInterval")
formCropIDs = Request.Form.Item("CropIDs")

'tjpbeg
dim formSearchName
formSearchName = Request.Form.Item("searchName")
'tjpend

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")
set rsCrop = Server.CreateObject("ADODB.RecordSet")

IF Request.QuerySTring("task") = "d" and urlID <> "" THEN
	set rs = GetRecordCountBySprayListID(urlID)
	'response.write(rs(0))
	IF rs(0) > 0 THEN
		set rs = GetSprayListByID(urlID)
		delerrorFound = true
		delerrorMessage = "Spray Product '" & rs.Fields("Name") & "' is in use by a Spray Record.<br>Please use Make InActive instead."
	ELSE
		DeleteSprayList(urlID)
		EndConnect(conn)
		set rs = nothing
		Response.Redirect("spraylist_list.asp")
	END IF
END IF 
IF Request.QuerySTring("task") = "activate" and urlID <> "" THEN
	ActivateSprayList(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("spraylist_list.asp")
END IF 
IF Request.QuerySTring("task") = "deactivate" and urlID <> "" THEN
	DeActivateSprayList(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("spraylist_list.asp")
END IF 


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

		Dim tArray,count
	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  

		urlID = UPDATESprayList(formID, formName, formPURSName, formPURSEPANumber, formPURSReport, _
			formUnitID, formREI, formPHI, formMAXUseApp, formMAXUseSeason, formActiveInd, _
			formReEntryIntervalHours, formReEntryIntervalDays, formPreharvestInterval)

		DeleteCropToSprayBySpray(urlID)
		response.write("<br>" & formCropIDs)
		tArray = Split(formCropIDs, ",")
	  	IF isarray(tArray) THEN
	    	FOR count = LBound(tArray) to UBound(tArray)
				call InsertCropToSpray(urlID,tArray(Count))
			NEXT
		END IF



Response.Redirect("spraylist_list.asp")
		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlID = InsertSprayList(formName, formPURSName, formPURSEPANumber, formPURSReport, _
			formUnitID, formREI, formPHI, formMAXUseApp, formMAXUseSeason, formActiveInd, _
			formReEntryIntervalHours, formReEntryIntervalDays, formPreharvestInterval)
		if bAddOnly then
			onloadstring = "javascript:window.opener.refreshdata(" & urlID  & ");window.close();"
		end if

		DeleteCropToSprayBySpray(urlID)
		tArray = Split(formCropIDs, ",")
	  	IF isarray(tArray) THEN
	    	FOR count = LBound(tArray) to UBound(tArray)
				CALL InsertCropToSpray(urlID,tArray(Count))
			NEXT
		END IF

		EndConnect(conn)
		set rs = nothing
		Response.Redirect("spraylist_list.asp")
	END IF 'insert	
END IF 'form submitted 

IF formID <> 0 and not errorFound THEN

	set rs = GetSprayListByID(formID)

	IF NOT rs.eof THEN
		formName = rs.Fields("Name")
		formPURSName = rs.Fields("PURS_Name")
		formPURSEPANumber = rs.Fields("PURS_EPA_Number")
		formPURSReport = rs.Fields("PURS_Report")
		formUnitID = rs.Fields("UnitID")
		formREI = rs.Fields("REI")
		formPHI = rs.Fields("PHI")
		formMAXUseApp = rs.Fields("MAXUseApp")
		formMAXUseSeason = rs.Fields("MAXUseSeason")
		formActiveInd = rs.Fields("ActiveInd")
		formReEntryIntervalHours = rs.Fields("ReEntryIntervalHours")
		formReEntryIntervalDays = rs.Fields("ReEntryIntervalDays")
		formPreharvestInterval = rs.Fields("PreharvestInterval")
	END IF

	set rsCrop = GetCropsBySprayListID(formID)

	IF NOT rsCrop.EOF THEN
		formCropIDs = ""
		DO WHILE not rsCrop.eof 
			formCropIDs= ListAppend(formCropIDs,rsCrop.Fields("CropID"))
			rsCrop.MoveNext
		LOOP
	END IF
END IF%>
<html>
<head>
	<title>SprayList List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0" onload="<%=onloadstring%>">

<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Spray List</h1><br>&nbsp;</td></tr></table>

<form method=post name=frm>
<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<%
	if not bAddOnly then
%>
	<tr valign=middle>
		<td bgcolor="FFFFFF" class="bodytext"><br>
			<%if listContains("1", session("accessid")) or listContains("2", session("accessid")) then 'or session("username")="mal" then %>
			    <a href="spraylist_list.asp?Action=AddOnly">Add Product</a>&nbsp;&nbsp;&nbsp;&nbsp;
			<%end if %>
			<a href="spraylist_list.asp">All Products</a>&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="spraylist_list.asp?PURS=1">Reportable Sprays</a>&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="spraylist_list.asp?PURS=2">NON Reportable Sprays</a>&nbsp;&nbsp;&nbsp;&nbsp;
			
			<%'tjpbeg%>
            <input type="text" name="searchName" value="<%=formSearchName%>" maxlength="50"> <input type=submit value="Search" />
            <%'tjpend%>
            
			&nbsp;&nbsp;&nbsp;&nbsp;<input type=button value=Print onclick="window.print();" />
		</td>
	</tr>
<%
	end if
%>
	<tr>
	<td colspan="1" class="bodytext">
<table width="100%" border="1" cellpadding="2" cellspacing="0">
<%
	if  delErrorFound then
%>
	<tr>
		<td colspan="14" class="bodytext" valign="top"><font color="red"><strong><%= delerrormessage %></strong></font></td>
	</tr>
<%
	end if
	if  errorFound then
%>
	<tr>
		<td colspan="14" class="bodytext" valign="top"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></td>
	</tr>
<%
	end if
	if not bAddOnly then
%>
	<tr bgcolor=#cccccc>
    	<td valign="top">Active<br />Ingredient</td>
		<td valign="top">PURS?</td>
		<td valign="top">REI</td>
		<td valign="top">PHI</td>
		<td valign="top" nowrap>MAX Use<br />Application</td>
		<td valign="top" nowrap>MAX Use<br />Season</td>
		<td valign="top">Units</td>
		<td valign="top">Crops</td>
		
<!--- 		<td valign="top">Reentry Interval Days</td>
	<td valign="top">Reentry Interval Hours</td>
	<td valign="top">Preharvest Interval</td>
 --->
	</tr>
<%

'	if not bAddOnly then

		select case iReportable
			case 1
				set rs = GetAllReportableSprayList()
			case 2
				set rs = GetAllNONReportableSprayList()
			case else
				set rs = GetAllSprayList()
		end select
		
		
        'tjpbeg
        if formSearchName <> "" then
	        ' with PURS name
	        rs.Filter = "Name LIKE '%" & formSearchName & "%' OR PURS_Name LIKE '%" & formSearchName & "%' OR ActiveInd LIKE '%" & formSearchName & "%'"
	        ' without PURS name
	        'rs.Filter = "Name LIKE '%" & formSearchName & "%'"
        end if
        dim recordCount
        recordCount = rs.RecordCount
        'tjpend


	
'	set rs = GetAllSprayList()
		dim i
		i = 0

		IF not rs.EOF THEN
			DO WHILE not rs.eof 
				i = i + 1
				if (i mod 10 = -1) then 'JMS
%>
	<tr bgcolor=#cccccc>
		<td valign="top">Active<br />Ingredient</td>
		<td valign="top">PURS?</td>
		<td valign="top">REI</td>
		<td valign="top">PHI</td>
		<td valign="top" nowrap>MAX Use<br />Application</td>
		<td valign="top" nowrap>MAX Use<br />Season</td>
		<td valign="top">Units</td>
		<td valign="top">Crops</td>
		
<!--- 	<td valign="top">Reentry Interval Days</td>
	<td valign="top">Reentry Interval Hours</td>
	<td valign="top">Preharvest Interval</td> --->
	</tr>


<%				end if %>
	    
	    <tr bgcolor=beige>
		<!--<td rowspan=2 class = "bodytext" valign="top">#<%=rs.Fields("SprayListID")%></td>-->
		<td colspan=7 class="bodytext" valign="top"><b><%=rs.Fields("Name")%></b>&nbsp;(PURS: <%=rs.Fields("PURS_Name")%>) (EPA #<%=rs.Fields("PURS_EPA_Number")%>)</td>
		<td colspan=2 bgcolor=#eeeeee valign="top" class="bodytext">
		  <%if listContains("1", session("accessid")) or listContains("2", session("accessid")) then '# Apr-2011: or session("username")="mal" then %>
		    <a href="spraylist_list.asp?ID=<%=rs.Fields("SprayListID")%>#edit" class="bodytext">Edit</a>&nbsp;
		    <a onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="spraylist_list.asp?ID=<%=rs.Fields("SprayListID")%>&task=d" class="bodytext">Delete</a>&nbsp;
		    <% IF rs.Fields("Active") =  0 then%><a href="spraylist_list.asp?task=activate&ID=<%=rs.Fields("SprayListID")%>" onclick="javascript: return confirm('Are you sure you want to activate this record?');">Activate</a><%else%><a href="spraylist_list.asp?task=deactivate&ID=<%=rs.Fields("SprayListID")%>" onclick="javascript: return confirm('Are you sure you want to DeActivate this record?');">InActivate</a><%end if%>
          <%end if%>&nbsp;
		</td>
		</tr>
		
		<tr>
		<td class="bodytext" valign="top"><%=rs.Fields("ActiveInd")%>&nbsp;</td>
		<td class="bodytext" valign="top"><%if rs.Fields("PURS_Report") then response.Write "Yes": else response.Write "No"%>&nbsp;</td>
		<td class="bodytext" valign="top"><%=rs.Fields("REI")%>&nbsp;</td>
		<td class="bodytext" valign="top"><%=rs.Fields("PHI")%>&nbsp;</td>
		<td class="bodytext" valign="top"><%=rs.Fields("MAXUseApp")%>&nbsp;</td>
		<td class="bodytext" valign="top"><%=rs.Fields("MAXUseSeason")%>&nbsp;</td>
		<td class="bodytext" valign="top"><%=rs.Fields("Unit")%>&nbsp;</td>
		<td class="bodytext" valign="top">
<%
				set rsCrop = GetCropsBySprayListID(rs.Fields("SprayListID"))
				IF not rsCrop.EOF THEN
					DO WHILE not rsCrop.eof 
						Response.write(rsCrop.Fields("Crop") & "<br>")
						rsCrop.MoveNext
					LOOP
				END IF
%>
&nbsp;
	
		</td>
		
<!--- 	<td class="bodytext" valign="top"><%=rs.Fields("ReentryIntervalDays")%>&nbsp;</td>
	<td class="bodytext" valign="top"><%=rs.Fields("ReentryIntervalHours")%>&nbsp;</td>
	<td class="bodytext" valign="top"><%=rs.Fields("PreharvestInterval")%>&nbsp;</td> --->
	</tr> 
<% 
				rs.MoveNext
			LOOP

		ELSE
%>
	<tr>
		<td class="bodytext" colspan="14">No Records Selected</td>
	</tr>
<%		END IF
	end if
%>
</table>
</form>

 <%if listContains("1", session("accessid")) or listContains("2", session("accessid")) then 'session("username")="mal" then %>
 
<a name="edit"></a>
<form action="SprayList_list.asp" method="post" name="frmsearch">
<table width="500" border="0" cellpadding="2" cellspacing="0">
	<tr>
		<td>&nbsp;</td>
		<td align="left" class="bodytext">* indicates required field</td>
	</tr>
<%	If errorFound Then %>
	<tr>
		<td>&nbsp;</td>
		<td class="bodytext"><font color="red"><% =errorMessage%></font></td>
	</tr>
<%	End If
	if bAddOnly then %>
<input type="hidden" name="Action" value="AddOnly">
<%	end if %>
<input type="hidden" value="<% =urlID%>" name="ID">
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="Name">Name</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formName%>" name="Name" class="bodytext" size="30" maxlength="150"></span></td>
	</tr>

	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="PURS_Name">PURS Name</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formPURSName%>" name="PURS_Name" class="bodytext" size="50" maxlength="150"></span></td>
	</tr>

	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="PURS_EPA_Number">PURS EPA Number</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formPURSEPANumber%>" name="PURS_EPA_Number" class="bodytext" size="30" maxlength="150"></span></td>
	</tr>

	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="PURS_Report">PURS Report</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="checkbox" name="PURS_Report" value="1" <%if formPURSReport then Response.Write("checked")%> class="bodytext"></span></td>
	</tr>

	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="Units">Units</label>:</span></td>
		<td valign="top">
<%
	set rsSelect = GetAllUnits()
%>
			<SELECT name="UnitID">
				<option value="">---Unit---</option>
<%
	IF not rsSelect.EOF THEN
		DO WHILE not rsSelect.eof 
			if rsSelect.Fields("Active") then
%>
				<option value="<%=rsSelect.Fields("UnitID")%>"<%if trim(formUnitID) = trim(rsSelect.Fields("UnitID")) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Unit")%></option>
<%
			end if
			rsSelect.MoveNext
		LOOP
	END IF
%>
			</SELECT>
		</td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="Units">Crops</label>:</span></td>
		<td valign="top">
<%
	set rsCrop = GetAllCrops()
	IF not rsCrop.EOF THEN
		DO WHILE not rsCrop.eof 
%>
<input type="checkbox" name="CropIDs" value="<%=rsCrop.Fields("CropID")%>"<%if ListContains(formCropIDs,rsCrop.Fields("CropID"))  then response.write("Checked") end if%>><%if not rsCrop.Fields("Active") then %>*<%end if%><%=rsCrop.Fields("Crop")%></option>
<%
			rsCrop.MoveNext
		LOOP
	END IF
%>
		</td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="REI">REI</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formREI%>" name="REI" class="bodytext" size="25" maxlength="150"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="PHI">PHI</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formPHI%>" name="PHI" class="bodytext" size="25" maxlength="150"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="MAXValue">MAX Use/Application</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formMAXUseApp%>" name="MAXUseApp" class="bodytext" size="25" maxlength="150"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="MAXValue">MAX Use/Season</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formMAXUseSeason%>" name="MAXUseSeason" class="bodytext" size="25" maxlength="150"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="ActiveInd">Active Ingredient</label>:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formActiveInd%>" name="ActiveInd" class="bodytext" size="55" maxlength="250"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="ReEntryIntervalDays">ReEntry Interval Days</label> number only:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formReEntryIntervalDays%>" name="ReEntryIntervalDays" class="bodytext" size="25" maxlength="150"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="ReEntryIntervalHours">ReEntry Interval Hours</label> number only:</span></td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formReEntryIntervalHours%>" name="ReEntryIntervalHours" class="bodytext" size="25" maxlength="150"></span></td>
	</tr>
	<tr>
		<td valign="top" align="right"><span class="subtitle"><label for="PreharvestInterval">Preharvest Interval</label>:</span> number only</td>
		<td valign="top"><span class="bodytext"><input type="text" value="<%=formPreharvestInterval%>" name="PreharvestInterval" class="bodytext" size="25" maxlength="150"></span></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><% IF  urlID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
	</tr>

</table>
<%end if %>

</form>
<%
	EndConnect(conn)
	set rs = nothing
	set rsSelect = nothing
%>
<!--#include file="i_adminfooter.asp" -->
</body>
</html>
