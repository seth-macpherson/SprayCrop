<%Option Explicit
if not session("login") or not listContains("1,2", session("accessid")) then
	response.redirect("index.asp")
end if

Const QS_Error = "ERROR"
Const DUPLICATE_GROWER_NUMBER = "DUPLICATE_ID"

%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_Growers.asp"-->
<%
'CREATED by LocusInteractive on 07/21/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlID,formID
Dim conn,sql,rs,counter
Dim urlAction, lAddingNewGrower

'Initialize variables
errorFound = FALSE
formError = FALSE
lAddingNewGrower = FALSE
errorMessage = "The following errors have occurred:"
formID=Request.Form.Item("ID")
urlID=Request.QueryString("ID")

urlAction = Request.QueryString("Action")
If urlAction = "add" Then
	lAddingNewGrower = TRUE
End If

'See if ID was passed through URL or FORM
IF urlID = "" THEN urlID = 0 END IF
IF formID = "" THEN formID = urlID End IF
urlID = formID

REM added display of error message for duplicate grower number
If Request.QueryString(QS_ERROR) = DUPLICATE_GROWER_NUMBER Then
		errorFound = TRUE
		errorMessage = errorMessage + "<br>Duplicate " & Application("GrowerNumber") & " Entered.<br>Record NOT added." 
End If

'Initialize Form Fields
DIM formGrowerNumber,formGrowerName,formEmail,formUsername,formPassword,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formCell,formFieldman,formApplicatorSupervisor,formSupervisorLicense,formApplicator,formApplicatorLicense,formChemicalSupplier,formRecommendedBy,formInternalNote,formActive,formCreateDate,urlSearchName,urlSearchNumber,formSearchName,formSearchNumber,urlSearchString,formAdditionalGrowerNumbers
DIM formfName, formlName, formUnit

formAdditionalGrowerNumbers = Request.Form.Item("AdditionalGrowerNumbers")
formGrowerNumber = Request.Form.Item("GrowerNumber")
formGrowerName = Request.Form.Item("GrowerName")
formfName = Request.Form.Item("fname")
formlName = Request.Form.Item("lname")

formEmail = Request.Form.Item("Email")
formPassword = Request.Form.Item("Password")
formAddress = Request.Form.Item("Address")
formCity = Request.Form.Item("City")
formState = Request.Form.Item("State")
formZipCode = Request.Form.Item("ZipCode")
formContact = Request.Form.Item("Contact")
formTelephone1 = Request.Form.Item("Telephone1")
formTelephone2 = Request.Form.Item("Telephone2")
formFax = Request.Form.Item("Fax")
formCell = Request.Form.Item("Cell")
formFieldman = Request.Form.Item("Fieldman")
formApplicatorSupervisor = Request.Form.Item("ApplicatorSupervisor")
formSupervisorLicense = Request.Form.Item("SupervisorLicense")
formApplicator = Request.Form.Item("Applicator")
formApplicatorLicense = Request.Form.Item("ApplicatorLicense")
formChemicalSupplier = Request.Form.Item("ChemicalSupplier")
formRecommendedBy = Request.Form.Item("RecommendedBy")
formInternalNote = Request.Form.Item("InternalNote")
formActive = Request.Form.Item("Active")
formCreateDate = Request.Form.Item("CreateDate")
formSearchName = Request.Form.Item("SearchName")
formSearchNumber = Request.Form.Item("SearchNumber")
formUnit = Request.Form.Item("unit")
urlSearchName = Request.QueryString("SearchName")
urlSearchNumber = Request.QueryString("SearchNumber")
IF formSearchName = "" THEN formSearchName = urlSearchName END IF
IF formSearchNumber = "" THEN formSearchNumber = urlSearchNumber END IF
urlSearchString = "SearchName=" & formSearchName & "&SearchNumber=" & formSearchNumber

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Response.Redirect("growers_list.asp?" & urlSearchString)
END IF

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")

IF Request.QueryString("task") = "d" and urlID <> "" THEN
	DeleteGrowers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("growers_list.asp?" & urlSearchString)
END IF 
IF Request.QueryString("task") = "activate" and urlID <> "" THEN
	ActivateGrowers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("growers_list.asp?" & urlSearchString)
END IF 
IF Request.QueryString("task") = "deactivate" and urlID <> "" THEN
	DeActivateGrowers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("growers_list.asp?" & urlSearchString)
END IF 


'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("GrowerNumber"), "nvarchar","GrowerNumber", false) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("GrowerName"), "nvarchar","Grower", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("fName"), "nvarchar","First Name", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
    IF NOT ValidateDatatype(Request.Form.Item("lName"), "nvarchar","Last Name", TRUE) THEN
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
	IF NOT ValidateDatatype(Request.Form.Item("Cell"), "nvarchar","Cell", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Fieldman"), "nvarchar","Fieldman", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("ApplicatorSupervisor"), "nvarchar","ApplicatorSupervisor", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("ChemicalSupplier"), "nvarchar","ChemicalSupplier", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("RecommendedBy"), "nvarchar","RecommendedBy", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("InternalNote"), "text","InternalNote", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
		

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		urlID = UPDATEGrowers(formID,formAdditionalGrowerNumbers,formGrowerNumber,formGrowerName,formEmail,formUsername,formPassword,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formCell,formFieldman,formApplicatorSupervisor,formSupervisorLicense,formApplicator,formApplicatorLicense,formChemicalSupplier,formRecommendedBy,formInternalNote,formActive,formCreateDate,formUnit)
		EndConnect(conn)
		set rs = nothing

		Response.Redirect("growers_list.asp?" & urlSearchString)
		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
'		urlID = InsertGrowers(formAdditionalGrowerNumbers,formGrowerNumber,formGrowerName,formEmail,formUsername,formPassword,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formCell,formFieldman,formApplicatorSupervisor,formSupervisorLicense,formApplicator,formApplicatorLicense,formChemicalSupplier,formRecommendedBy,formInternalNote)
REM MODIFIED i_growers.asp Kim Miers 7/13/2006
		urlID = InsertGrower(formAdditionalGrowerNumbers,formGrowerNumber,formGrowerName,formEmail,formUsername,formPassword,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formCell,formFieldman,formApplicatorSupervisor,formSupervisorLicense,formApplicator,formApplicatorLicense,formChemicalSupplier,formRecommendedBy,formInternalNote,formfName,formlName)
		EndConnect(conn)
		set rs = nothing
		IF urlID = 0 THEN
			Response.Redirect("growers_list.asp?" + QS_Error + "=" + DUPLICATE_GROWER_NUMBER)
		ELSE
			Response.Redirect("growers_list.asp?" & urlSearchString)
		END IF
	END IF 'insert	
END IF 'form submitted 

IF formID <> 0 and not errorFound THEN
	set rs = GetGrowersByID(formID)
	IF NOT rs.eof THEN
		formAdditionalGrowerNumbers = rs.Fields("AdditionalGrowerNumbers")
		formGrowerNumber = rs.Fields("GrowerNumber")
		formGrowerName = rs.Fields("GrowerName")
		formEmail = rs.Fields("Email")
		formPassword = rs.Fields("GrowerPassword")
		formAddress = rs.Fields("Address")
		formCity = rs.Fields("City")
		formState = rs.Fields("State")
		formZipCode = rs.Fields("ZipCode")
		formContact = rs.Fields("Contact")
		formTelephone1 = rs.Fields("Telephone1")
		formTelephone2 = rs.Fields("Telephone2")
		formFax = rs.Fields("Fax")
		formCell = rs.Fields("Cell")
		formFieldman = rs.Fields("Fieldman")
		formApplicatorSupervisor = rs.Fields("ApplicatorSupervisor")
		formSupervisorLicense = rs.Fields("SupervisorLicense")
		formApplicator = rs.Fields("Applicator")
		formApplicatorLicense = rs.Fields("ApplicatorLicense")
		formChemicalSupplier = rs.Fields("ChemicalSupplier")
		formRecommendedBy = rs.Fields("RecommendedBy")
		formInternalNote = rs.Fields("InternalNote")
		formActive = rs.Fields("Active")
		formCreateDate = rs.Fields("CreateDate")
		formUnit = rs.fields("unitid")
	END IF
END IF%>
<html>
<head>
	<title>Growers List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Growers</h1><br>&nbsp;</td></tr></table>

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="FFFFFF" class="bodytext">
<br>
<!--<a href="growers_list.asp?urlSearchString#edit">Add Growers</a><br><br>-->
<a href="growers_list.asp?action=add#add">Add a Grower</a><br><br>
<%
	Response.flush
%>
<form action="growers_list.asp">
Search by Name: <input type="text" name="searchName" value="<%=formSearchName%>" maxlength="50"> <input type="submit" value="Search" name="search">
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

REM modification by Kim Miers 7/16/2006 to not show list if editing a Grower.
IF formID = 0 and not errorFound and not lAddingNewGrower THEN
%>
<tr bgcolor=#dddddd>
<td  valign="top">&nbsp;</td>
<td  valign="top"><h2>Grower</h2></td>

	<td valign="top"><h2>Contact</h2></td>
	<td  valign="top"><h2>&nbsp;</h2></td>
	<!--<td  valign="top"><h2>Other</h2></td>-->
</tr>
<%
	set rs = GetGrowersByNameNumber(formSearchName,formSearchNumber)
	DIM i
	i = 0

	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			i = i + 1
%>
<tr>
<td class = "bodytext" valign="top"><%=i%></td>
	<td class="bodytext" valign="top"><%=rs.Fields("GrowerName")%><br />

&nbsp;<br><%=rs.Fields("Address")%><br><%=rs.Fields("City")%>&nbsp;<br><%=rs.Fields("State")%>&nbsp;<%=rs.Fields("ZipCode")%>&nbsp;
</td>
	<td class="bodytext" valign="top" nowrap><%=rs.Fields("Contact")%>&nbsp;<br><%if rs.Fields("Telephone1") <> "" then response.write("<br>Tel1: " & rs.Fields("Telephone1"))%>&nbsp;
	<%if rs.Fields("Telephone2") <> "" then response.write("<br>Tel2: " & rs.Fields("Telephone2"))%>&nbsp;
	<%if rs.Fields("Fax") <> "" then response.write("<br>Fax: " & rs.Fields("Fax"))%>&nbsp;
	<%if rs.Fields("Cell") <> "" then response.write("<br>Cell: " & rs.Fields("Cell"))%>&nbsp;
	<br>
<%IF rs.Fields("Email") <> "" THEN%>
<a href="mailto:<%=rs.Fields("Email")%>"><%=rs.Fields("Email")%></a><br>
<%END IF%>

	<!--<hr>
Terms Agreed: <%'if rs.Fields("TermsAgreed") then response.Write "Yes": else response.Write "No"%><br> By: <%'=rs.Fields("AgreedBy")%><br>On: <%'=rs.Fields("AgreedDate")%>
-->
</td>
	<!--
	<td class="bodytext" valign="top" nowrap>Fieldman: <%=rs.Fields("Fieldman")%>&nbsp;<br>
	<strong>Default values for data entry:</strong><br>
	Supervisor:  <%=rs.Fields("ApplicatorSupervisor")%><br>
	Supervisor License:  <%=rs.Fields("SupervisorLicense")%><br>
	Applicator:  <%=rs.Fields("Applicator")%><br>
	Applicator License:  <%=rs.Fields("ApplicatorLicense")%><br>
	Chemical Supplier:  <%=rs.Fields("ChemicalSupplier")%><br>
	Recommended By:  <%=rs.Fields("RecommendedBy")%><br>
	<br>
	Additional <%=Application("GrowerNumber")%>s:<br>
	<%= rs.Fields("AdditionalGrowerNumbers") %>
	
	</td>
	-->
<td  valign="top" class="bodytext" bgcolor=#eeeeee>
<% IF rs.Fields("Active") =  0 then%><a href="growers_list.asp?<%=urlSearchString%>&task=activate&ID=<%=rs.Fields("GrowerID")%>" onclick="javascript: return confirm('Are you sure you want to activate this record?');">Make Active</a><%else%><a href="growers_list.asp?<%=urlSearchString%>&task=deactivate&ID=<%=rs.Fields("GrowerID")%>" onclick="javascript: return confirm('Are you sure you want to DeActivate this record?');">Make InActive</a><%end if%><br />
<a href="growers_list.asp?<%=urlSearchString%>&ID=<%=rs.Fields("GrowerID")%>#edit" class="bodytext">Edit Grower</a><br><a href="GrowerLocations.asp?GrowerID=<%=rs.Fields("GrowerID")%>">Manage Locations</a>
<!--- <a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="growers_list.asp?<%=urlSearchString%>&ID=<%=rs.Fields("GrowerID")%>&task=d" class="bodytext">Delete</a> ---></td>

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
<a name="add"></a>
<a name="edit"></a>
<form action="Growers_list.asp" method="post" name="frmsearch">
<table width="500" border="0" cellpadding="2" cellspacing="0">

<%
	IF  urlID <> 0 THEN
%>
<tr>
<th colspan=2 align="center">EDITING GROWER</th>
</tr>
<%  ELSE %>
<tr>
<th colspan=2 align="center">ADD A GROWER</th>
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
<!--
<tr><td valign="top" align="right"><span class="subtitle"><label for="GrowerNumber">*<%=Application("GrowerNumber")%></label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formGrowerNumber%>" name="GrowerNumber"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>-->

<tr><td valign="top" align="right" colspan=2>&nbsp;</td></tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="GrowerName">*Grower</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formGrowerName%>" name="GrowerName"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>

<tr><td valign="top" align="right"><span class="subtitle"><label for="GrowerName">Grower Org</label>:</span></td>
<td valign="top"><span class="bodytext"><select name=unit><option value=0></option>
<%
dim rsUnits
set rsUnits = conn.execute("exec growerunit$list") 
dO WHILE not rsunits.eof  %>
<option value="<%=rsunits("unitid")%>" <%if rsunits("unitid")=formUnit then response.write "selected"%>><%=rsunits("unit")%></option>
<%
rsUnits.movenext
loop 
%>
</span></td>
</tr>

<%if urlid=0 then%>

<tr><td valign="top" align="right"><span class="subtitle"><label for="fname">*First Name</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formfName%>" name="fname"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>

<tr><td valign="top" align="right"><span class="subtitle"><label for="lname">*Last Name</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formlName%>" name="lname"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>

<%else %>

<tr><td valign="top" align="right" colspan=2><input type="hidden" value="BLAH" name="fname"><input type="hidden" value="BLAH" name="lname"></td>
</tr>

<%end if%>

<tr><td valign="top" align="right" colspan=2>&nbsp;</td></tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Email">Email</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formEmail%>" name="Email"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<!--<tr><td valign="top" align="right"><span class="subtitle"><label for="Password">Password</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formPassword%>" name="Password"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>-->
<!--<tr><td valign="top" align="right"><span class="subtitle"><label for="AdditionalGrowerNumber">Additional <%=Application("GrowerNumber")%>s</label>:<br>
<strong>Enter comma delimited list only.</strong></span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formAdditionalGrowerNumbers%>" name="AdditionalGrowerNumbers"  class="bodytext" size="25" maxlength="150"></span></td>-->
<tr><td valign="top" align="right"><span class="subtitle"><label for="Address">Address</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formAddress%>" name="Address"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="City">City</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formCity%>" name="City"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="State">State</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formState%>" name="State"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="ZipCode">Zip Code</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formZipCode%>" name="ZipCode" size="10" maxlength="38" class="bodytext"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Contact">Contact</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formContact%>" name="Contact"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Telephone1">Phone 1</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formTelephone1%>" name="Telephone1"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Telephone2">Phone 2</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formTelephone2%>" name="Telephone2"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Fax">Fax</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formFax%>" name="Fax"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Cell">Cell</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formCell%>" name="Cell"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<!--
<tr><td valign="top" align="right"><span class="subtitle"><label for="Fieldman">Fieldman</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formFieldman%>" name="Fieldman"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>

<tr><td colspan="2" align="center"><br><strong>Default Values for Data Entry</strong></td></tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="ApplicatorSupervisor">Supervisor</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formApplicatorSupervisor%>" name="ApplicatorSupervisor"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="SupervisorLicense">Supervisor License</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formSupervisorLicense%>" name="SupervisorLicense"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Applicator">Applicator</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formApplicator%>" name="Applicator"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="ApplicatorLicense">Applicator License</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formApplicatorLicense%>" name="ApplicatorLicense"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="ChemicalSupplier">ChemicalSupplier</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formChemicalSupplier%>" name="ChemicalSupplier"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="RecommendedBy">RecommendedBy</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formRecommendedBy%>" name="RecommendedBy"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
-->
<tr><td valign="top" align="right"><br><br>
<span class="subtitle"><label for="InternalNote">InternalNote</label>:</span></td>
<td valign="top"><br><br><span class="bodytext"><textarea cols="30" rows="4" name="InternalNote" class="bodytext"><%=formInternalNote%></textarea></span></td>
</tr>
<tr>
<td>&nbsp;</td>
<td><% IF  urlID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>
<!--#include file="i_adminfooter.asp" -->
<%
	set rs = nothing
	EndConnect(conn)
%>
</body>
</html>



