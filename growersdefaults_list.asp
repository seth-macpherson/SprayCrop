<%Option Explicit%>
<%if not session("login") or not listContains("3", session("accessid")) then
	response.redirect("index.asp")
end if%>

	
<!--#include file="include/i_data.asp"-->
<!--#include file="i_Growers.asp"-->
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



'Initialize Form Fields
DIM formGrowerNumber,formGrowerName,formEmail,formUsername,formPassword,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formCell,formFieldman,formApplicatorSupervisor,formSupervisorLicense,formApplicator,formApplicatorLicense,formChemicalSupplier,formRecommendedBy,formInternalNote,formActive,formCreateDate,urlSearchName,urlSearchNumber,formSearchName,formSearchNumber,urlSearchString,formAdditionalGrowerNumbers
formAdditionalGrowerNumbers = Request.Form.Item("AdditionalGrowerNumbers")
formGrowerNumber = Request.Form.Item("GrowerNumber")
formGrowerName = Request.Form.Item("GrowerName")
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

IF Request.QuerySTring("task") = "d" and urlID <> "" THEN
	DeleteGrowers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("growers_list.asp?" & urlSearchString)
END IF 
IF Request.QuerySTring("task") = "activate" and urlID <> "" THEN
	ActivateGrowers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("growers_list.asp?" & urlSearchString)
END IF 
IF Request.QuerySTring("task") = "deactivate" and urlID <> "" THEN
	DeActivateGrowers(urlID)
	EndConnect(conn)
	set rs = nothing
	Response.Redirect("growers_list.asp?" & urlSearchString)
END IF 


'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

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
		urlID = UpdateGrowersByGrower(formID,formEmail,formUsername,formPassword,formAddress,formCity,formState,formZipCode,formContact,formTelephone1,formTelephone2,formFax,formCell,formFieldman,formApplicatorSupervisor,formSupervisorLicense,formApplicator,formApplicatorLicense,formChemicalSupplier,formRecommendedBy)
		EndConnect(conn)
		set rs = nothing

		Response.Redirect("Growersdefaults_list.asp?" & urlSearchString)
		'END UPDATE
	END IF 
	
END IF 'form submitted 

IF formID  <>  0 and not errorFound THEN
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
	END IF
END IF%>
<html>
<head>
	<title>Growers List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->
<table width="594" border="1" cellspacing="0" cellpadding=" 0" bordercolor="#013166" bgcolor="#beige"><tr><td><img src="images/spacer.gif" height="4" width="1" border="0"><br>
&nbsp;&nbsp;<h1>Growers<h1><br><img src="images/spacer.gif" height="4" width="1" border="0"></td></tr></table>
<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr>
<td colspan="2" class="bodytext">
<table width="90%" border="1" cellpadding="2" cellspacing="0">

<% if  errorFound then%>
<tr>
<td colspan="21" class="bodytext" valign="top"><a href="#edit"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></a></td>
</tr>
<% end if %>
<tr>
<td  valign="top">&nbsp;</td>
<td  valign="top"><h2>Edit</h2></td>
	<td  valign="top"><h2>GrowerNumber</h2><br><h2>GrowerName</h2><br><h2>Email</h2><br><h2>Address</h2></td>
	<td  valign="top"><h2>Contact</h2></td>
	<td  valign="top"><h2>Other</h2></td>
</tr>
<%
	set rs = GetActiveGrowers()
	DIM i
	i = 0
%>
<%IF not rs.EOF THEN
DO WHILE not rs.eof 
i = i + 1%>
<tr>
<td class = "bodytext" valign="top"><%=i%></td>
<td  valign="top" class="bodytext"><a href="Growersdefaults_list.asp?<%=urlSearchString%>&ID=<%=rs.Fields("GrowerID")%>#edit" class="bodytext">Edit</a><br>
<!--- <a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="growers_list.asp?<%=urlSearchString%>&ID=<%=rs.Fields("GrowerID")%>&task=d" class="bodytext">Delete</a> ---></td>
	<td class="bodytext" valign="top"><%=rs.Fields("GrowerNumber")%>&nbsp;<br><%=rs.Fields("GrowerName")%>
<%IF rs.Fields("Email") <> "" THEN%>
<a href="mailto:<%=rs.Fields("Email")%>"><%=rs.Fields("Email")%></a><br>
<%END IF%>

&nbsp;<br><%=rs.Fields("Address")%><br><%=rs.Fields("City")%>&nbsp;<br><%=rs.Fields("State")%>&nbsp;<%=rs.Fields("ZipCode")%>&nbsp;</td>
	<td class="bodytext" valign="top"><%=rs.Fields("Contact")%>&nbsp;<br>Tel:<%=rs.Fields("Telephone1")%>&nbsp;
	<%if rs.Fields("Telephone2") <> "" then response.write("<br>Tel2: " & rs.Fields("Telephone2"))%>&nbsp;
	<%if rs.Fields("Fax") <> "" then response.write("<br>Fax: " & rs.Fields("Fax"))%>&nbsp;
	<%if rs.Fields("Cell") <> "" then response.write("<br>Cell: " & rs.Fields("Cell"))%>&nbsp;</td>
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
</tr> 
<% 
rs.MoveNext
LOOP

Else
%>
<tr><td class="bodytext" colspan="21">No Records Selected</td></tr>
<%	end if %>
</table>
<a name="edit"></a>
<%if formID <> 0 then %>
<form action="Growersdefaults_list.asp" method="post" name="frmsearch">
<table width="500" border="0" cellpadding="2" cellspacing="0">
<% if errorFound then%>
<tr>
<td>&nbsp;</td>
<td class="bodytext"><font color="red"><% =errorMessage%></font></td>
</tr>
<% End If %>
<input type="hidden" value="<% =urlID%>" name="ID">
<input type="hidden" value="<% =formSearchName%>" name="SearchName">
<input type="hidden" value="<% =formSearchNumber%>" name="SearchNumber">
<tr><td valign="top" align="right"><span class="subtitle"><label for="GrowerNumber"><%=Application("GrowerNumber")%></label>:</span></td>
<td valign="top"><span class="bodytext"><%=formGrowerNumber%></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="GrowerName">Grower Name</label>:</span></td>
<td valign="top"><span class="bodytext"><%=formGrowerName%>"</span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Email">Email</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formEmail%>" name="Email"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Password">Password</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="password" value="<%=formPassword%>" name="Password"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
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
<span class="subtitle"><label for="ZipCode">ZipCode</label>(number only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formZipCode%>" name="ZipCode" size="10" maxlength="38" class="bodytext"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Contact">Contact</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formContact%>" name="Contact"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Telephone1">Telephone1</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formTelephone1%>" name="Telephone1"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Telephone2">Telephone2</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formTelephone2%>" name="Telephone2"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Fax">Fax</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formFax%>" name="Fax"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Cell">Cell</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formCell%>" name="Cell"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
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

<tr>
<td>&nbsp;</td>
<td><input type="submit" name="update" value="Update" class="bodytext">&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>
<%
	end if
	set rs = nothing
	EndConnect(conn)
%>
<!--#include file="i_adminfooter.asp" -->
</body>
</html>

