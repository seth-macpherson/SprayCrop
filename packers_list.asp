<%Option Explicit
if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if

Const QS_Error = "ERROR"
Const DUPLICATE_packer_NUMBER = "DUPLICATE_ID"

%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_packers.asp"-->
<%
'CREATED by LocusInteractive on 07/21/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlID,formID
Dim conn,sql,rs,counter
Dim urlAction, lAddingNewpacker

'Initialize variables
errorFound = FALSE
formError = FALSE
lAddingNewpacker = FALSE
errorMessage = "The following errors have occurred:"
formID=Request.Form.Item("ID")
urlID=Request.QueryString("ID")

urlAction = Request.QueryString("Action")
If urlAction = "add" Then
	lAddingNewpacker = TRUE
End If

'See if ID was passed through URL or FORM
IF urlID = "" THEN urlID = 0 END IF
IF formID = "" THEN formID = urlID End IF
urlID = formID

REM added display of error message for duplicate packer number
If Request.QueryString(QS_ERROR) = DUPLICATE_packer_NUMBER Then
		errorFound = TRUE
		errorMessage = errorMessage + "<br>Duplicate " & Application("packerNumber") & " Entered.<br>Record NOT added." 
End If

'Initialize Form Fields
DIM formpackerNumber,formpackerName,formEmail,formUsername,formPassword,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formCell,formFieldman,formApplicatorSupervisor,formSupervisorLicense,formApplicator,formApplicatorLicense,formChemicalSupplier,formRecommendedBy,formInternalNote,formActive,formCreateDate,urlSearchName,urlSearchNumber,formSearchName,formSearchNumber,urlSearchString,formAdditionalpackerNumbers
DIM formfullrights,formgrowerlimit,formlogofileext

formpackerNumber = Request.Form.Item("packerNumber")
formpackerName = Request.Form.Item("packerName")
formFullRights = cint(Request.Form.Item("fullrights"))
formGrowerLimit = Request.Form.Item("growerlimit")
formlogofileext = Request.Form.Item("logofileext")

formEmail = Request.Form.Item("Email")
formAddress = Request.Form.Item("Address")
formCity = Request.Form.Item("City")
formState = Request.Form.Item("State")
formZipCode = Request.Form.Item("ZipCode")
formContact = Request.Form.Item("Contact")
formTelephone1 = Request.Form.Item("Telephone1")
formTelephone2 = Request.Form.Item("Telephone2")
formFax = Request.Form.Item("Fax")
formActive = Request.Form.Item("Active")
formCreateDate = Request.Form.Item("CreateDate")
formSearchName = Request.Form.Item("SearchName")
formSearchNumber = Request.Form.Item("SearchNumber")
urlSearchName = Request.QueryString("SearchName")
urlSearchNumber = Request.QueryString("SearchNumber")
IF formSearchName = "" THEN formSearchName = urlSearchName END IF
IF formSearchNumber = "" THEN formSearchNumber = urlSearchNumber END IF
urlSearchString = "SearchName=" & formSearchName & "&SearchNumber=" & formSearchNumber

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Response.Redirect("packers_list.asp?" & urlSearchString)
END IF

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")

IF Request.QueryString("task") = "d" and urlID <> "" THEN
	Deletepackers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("packers_list.asp?" & urlSearchString)
END IF 
IF Request.QueryString("task") = "activate" and urlID <> "" THEN
	Activatepackers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("packers_list.asp?" & urlSearchString)
END IF 
IF Request.QueryString("task") = "deactivate" and urlID <> "" THEN
	DeActivatepackers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("packers_list.asp?" & urlSearchString)
END IF 


'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("packerName"), "nvarchar","packerName", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Address"), "nvarchar","Address", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("City"), "nvarchar","City", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("State"), "nvarchar","State", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("ZipCode"), "float","ZipCode", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Contact"), "nvarchar","Contact", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Telephone1"), "nvarchar","Telephone1", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Telephone2"), "nvarchar","Telephone2", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Fax"), "nvarchar","Fax", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
		

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		urlID = UPDATEpackers(formID,formpackerName,formEmail,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formfullrights,formgrowerlimit)
		EndConnect(conn)
		set rs = nothing

		'Response.Redirect("packers_list.asp?" & urlSearchString)
		Response.Redirect("packers_list.asp?" & urlSearchString & "&ID="&formID) 

		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
'		urlID = Insertpackers(formAdditionalpackerNumbers,formpackerNumber,formpackerName,formEmail,formUsername,formPassword,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formCell,formFieldman,formApplicatorSupervisor,formSupervisorLicense,formApplicator,formApplicatorLicense,formChemicalSupplier,formRecommendedBy,formInternalNote)
REM MODIFIED i_packers.asp Kim Miers 7/13/2006
		urlID = Insertpacker(formpackerName,formEmail,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formfullrights,formgrowerlimit)
		EndConnect(conn)
		set rs = nothing
		IF urlID = 0 THEN
			Response.Redirect("packers_list.asp?" + QS_Error + "=" + DUPLICATE_packer_NUMBER)
		ELSE
			Response.Redirect("packers_list.asp?" & urlSearchString)
		END IF
	END IF 'insert	
END IF 'form submitted 

IF formID <> 0 and not errorFound THEN
	set rs = GetpackersByID(formID)
	IF NOT rs.eof THEN
		formpackerNumber = rs.Fields("packerNumber")
		formpackerName = rs.Fields("packerName")
        formgrowerlimit = rs.Fields("growerlimit")
        formfullrights = rs.Fields("fullrights")
        formlogofileext = rs.Fields("logofileext")
                
		formEmail = rs.Fields("Email")
		formAddress = rs.Fields("Address")
		formCity = rs.Fields("City")
		formState = rs.Fields("State")
		formZipCode = rs.Fields("ZipCode")
		formContact = rs.Fields("Contact")
		formTelephone1 = rs.Fields("Telephone1")
		formTelephone2 = rs.Fields("Telephone2")
		formFax = rs.Fields("Fax")
		formActive = rs.Fields("Active")
		formCreateDate = rs.Fields("CreateDate")
	END IF
END IF%>
<html>
<head>
	<title>Packers List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Packers</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="FFFFFF" class="bodytext"><br>
<!--<a href="packers_list.asp?urlSearchString#edit">Add packers</a><br><br>-->
<a href="packers_list.asp?action=add#add">Add a packer</a><br><br>
<%
	Response.flush
%>
<form action="packers_list.asp">
Search by Name: <input type="text" name="searchName" value="<%=formSearchName%>" maxlength="50"> or <%=Application("packerNumber")%>: <input type="text" name="searchNumber" value="<%=formSearchNumber%>" maxlength="50"> <input type="submit" value="Search" name="search">
</form></td>
</tr>
<tr>
<td colspan="2" class="bodytext">
<table width="90%" border="1" cellpadding="2" cellspacing="0">
<% if  delerror then%>
<tr>
<td colspan="21" class="bodytext" valign="top"><font color="red"><%= delerrormessage %></font></td>
</tr>
<% end if
 if  errorFound then%>
<tr>
<td colspan="21" class="bodytext" valign="top"><a href="#edit"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></a></td>
</tr>
<% end if

REM modification by Kim Miers 7/16/2006 to not show list if editing a packer.
IF formID = 0 and not errorFound and not lAddingNewpacker THEN
%>
<tr bgcolor=#dddddd>
<td  valign="top"><h2>Packer</h2></td>
<td  valign="top"><h2>Contact</h2></td>

	<td valign="top"><h2>Activate</h2></td>

</tr>
<%
	set rs = GetpackersByNameNumber(formSearchName,formSearchNumber)
	DIM i
	i = 0

	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			i = i + 1
%>
<tr>
<!--<td class = "bodytext" valign="top"><%=i%></td>-->

	<td class="bodytext" valign="top"><%=rs.Fields("packerName")%><br />#<%=rs.Fields("packerNumber")%>
	<%if rs.Fields("fullrights") then response.write ", Full Rights"%>
	<br>

    <br><%=rs.Fields("Address")%><br><%=rs.Fields("City")%>&nbsp;<br><%=rs.Fields("State")%>&nbsp;<%=rs.Fields("ZipCode")%>&nbsp;<br></td>
	<td class="bodytext" valign="top" nowrap>
	<%=rs.Fields("Contact")%>&nbsp;<br><br />
	<%if rs.Fields("Telephone1") <> "" then response.write("<br>Tel1: " & rs.Fields("Telephone1"))%>
	<%if rs.Fields("Telephone2") <> "" then response.write("<br>Tel2: " & rs.Fields("Telephone2"))%>&nbsp;
	<%if rs.Fields("Fax") <> "" then response.write("<br>Fax: " & rs.Fields("Fax"))%>&nbsp;
	
    <%IF rs.Fields("Email") <> "" THEN%>
    <a href="mailto:<%=rs.Fields("Email")%>"><%=rs.Fields("Email")%></a><br>
    <%END IF%>

	</td>
<td  valign="top" class="bodytext" bgcolor=#eeeeee>
    <% IF rs.Fields("Active") =  0 then%><a href="packers_list.asp?<%=urlSearchString%>&task=activate&ID=<%=rs.Fields("packerID")%>" onclick="javascript: return confirm('Are you sure you want to activate this record?');">Make Active</a><%else%><a href="packers_list.asp?<%=urlSearchString%>&task=deactivate&ID=<%=rs.Fields("packerID")%>" onclick="javascript: return confirm('Are you sure you want to DeActivate this record?');">Make InActive</a><%end if%><br />
    <a href="packers_list.asp?<%=urlSearchString%>&ID=<%=rs.Fields("packerID")%>#edit" class="bodytext">Edit Packer</a><br>
    <!--- <a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="packers_list.asp?<%=urlSearchString%>&ID=<%=rs.Fields("packerID")%>&task=d" class="bodytext">Delete</a> ---></td>

</tr> 
<% 
			rs.MoveNext
		LOOP
	Else
%>
<tr><td class="bodytext" colspan="21">No Records Selected</td></tr>
<%	
	end if
end if
%>
</table>

<% IF  urlID <> 0 OR request.QueryString("action")="add" THEN
%>
	
<a name="add"></a>
<a name="edit"></a>
<form action="packers_list.asp" method="post" name="frmsearch">
<table width="500" border="0" cellpadding="2" cellspacing="0">

<%

	IF  urlID <> 0 THEN
%>
<tr>
<th colspan=2 align="center">EDIT PACKER</th>
</tr>
<%  ELSE %>
<tr>
<th colspan=2 align="center">ADD PACKER</th>
</tr>
<%END IF%>
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
<input type="hidden" value="<% =formSearchName%>" name="SearchName">
<input type="hidden" value="<% =formSearchNumber%>" name="SearchNumber">

<tr><td valign="top" align="right" colspan=2>&nbsp;</td></tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="packerName">* Packer Name</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formpackerName%>" name="packerName"  class="bodytext" size="40" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"></td>
<td valign="top"><span class="subtitle"><label for="FullRights"> Full Rights?</label></span>&nbsp;&nbsp;&nbsp;<span class="bodytext"><input onclick="this.form.update.click();" type="radio" value="1" <%if formfullrights then response.write "checked"%> name="FullRights"> Yes <input type="radio" onclick="this.form.update.click();" value="0" <%if not formfullrights then response.write "checked"%> name="FullRights"> No</span></td>
</tr>
<tr><td valign="top" align="right"></td>
<td valign="top"><span class="subtitle"><label for="FullRights"> Grower Limit:</label></span> <span class="bodytext"><input type="text" value="<%=formgrowerlimit%>" name="GrowerLimit" size=2></span></td>
</tr>

<%if urlid<>0 and formfullrights then %>
<tr><td valign="top" align="right" colspan=2>&nbsp;</td></tr>
<tr><td valign="top" align="center" colspan=2>
    <%
    if formlogofileext>"" then
        response.Write "<img id=logo src=""logos/p"& formpackernumber & formlogofileext&""" />"
    end if 
    %>
</td></tr>
<tr><td valign="top" align="center" colspan=2><a href="javascript:void(window.open('packerlogoupload.aspx?pnum=<%=formpackernumber%>','_logo','width=450,height=450'));">Upload new logo file...</a></td></tr>
<tr><td valign="top" align="right" colspan=2>&nbsp;</td></tr>
<%end if %>

<tr><td valign="top" align="right" colspan=2>&nbsp;</td></tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Email">Email</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formEmail%>" name="Email"  class="bodytext" size=40 maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Address">Address</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formAddress%>" name="Address"  class="bodytext" size=40 maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="City">City</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formCity%>" name="City"  class="bodytext" size=40 maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="State">State</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formState%>" name="State"  class="bodytext" size=40 maxlength="150"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="ZipCode">ZipCode</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formZipCode%>" name="ZipCode" size="10" maxlength="38" class="bodytext"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Contact">Contact</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formContact%>" name="Contact"  class="bodytext" size=40 maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Telephone1">Telephone1</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formTelephone1%>" name="Telephone1"  class="bodytext" size=40 maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Telephone2">Telephone2</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formTelephone2%>" name="Telephone2"  class="bodytext" size=40 maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Fax">Fax</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formFax%>" name="Fax"  class="bodytext" size=40 maxlength="150"></span></td>
</tr>

<tr>
<td>&nbsp;</td>
<td><% IF  urlID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>
<%end if %>

<!--#include file="i_adminfooter.asp" -->
<%
	set rs = nothing
	EndConnect(conn)
%>
</body>
</html>

