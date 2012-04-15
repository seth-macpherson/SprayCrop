<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="include/i_data.asp"-->
<%
'CREATED by LocusInteractive on 07/21/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount
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
	Server.Transfer("SprayRecord_ins_thanks.asp")
END IF

'Initialize Form Fields
DIM formGrowerName,formGrowerNumber,formSprayDateDate,formCrop,formBartlet,formStage,formLocation,formMethod,formAcresTreated,formRateAcre,formTotalMaterialApplied,formProductNameFormulation,formUnitsOfProduct,formIFPRating,formTarget,formHarvestDate,formComments,formCreateDate
formGrowerName = Request.Form.Item("GrowerName")
formGrowerNumber = Request.Form.Item("GrowerNumber")
formSprayDateDate = Request.Form.Item("SprayDateDate")
formCrop = Request.Form.Item("Crop")
formBartlet = Request.Form.Item("Bartlet")
formStage = Request.Form.Item("Stage")
formLocation = Request.Form.Item("Location")
formMethod = Request.Form.Item("Method")
formAcresTreated = Request.Form.Item("AcresTreated")
formRateAcre = Request.Form.Item("RateAcre")
formTotalMaterialApplied = Request.Form.Item("TotalMaterialApplied")
formProductNameFormulation = Request.Form.Item("ProductNameFormulation")
formUnitsOfProduct = Request.Form.Item("UnitsOfProduct")
formIFPRating = Request.Form.Item("IFPRating")
formTarget = Request.Form.Item("Target")
formHarvestDate = Request.Form.Item("HarvestDate")
formComments = Request.Form.Item("Comments")
formCreateDate = Request.Form.Item("CreateDate")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")

'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("GrowerName"), "nvarchar","GrowerName", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("GrowerNumber"), "int","GrowerNumber", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("SprayDateDate"), "datetime","SprayDateDate", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Crop"), "int","Crop", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Bartlet"), "bit","Bartlet", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Stage"), "int","Stage", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Location"), "nvarchar","Location", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Method"), "int","Method", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("AcresTreated"), "real","AcresTreated", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("RateAcre"), "real","RateAcre", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("TotalMaterialApplied"), "int","TotalMaterialApplied", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("ProductNameFormulation"), "int","ProductNameFormulation", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("UnitsOfProduct"), "nvarchar","UnitsOfProduct", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("IFPRating"), "nvarchar","IFPRating", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Target"), "int","Target", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("HarvestDate"), "datetime","HarvestDate", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Comments"), "nvarchar","Comments", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("CreateDate"), "datetime","CreateDate", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF 		

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN 
		sql = "UPDATE Test SET	 GrowerName = '" & EscapeQuotes(formGrowerName) & "',GrowerNumber = " & formGrowerNumber & ",SprayDateDate = '" & formSprayDateDate & "',Crop = " & formCrop & ",Bartlet = " & formBartlet & ",Stage = " & formStage & "	 ,Location = '" & EscapeQuotes(formLocation) & "',Method = " & formMethod & ",AcresTreated = " & formAcresTreated & ",RateAcre = " & formRateAcre & ",TotalMaterialApplied = " & formTotalMaterialApplied & ",ProductNameFormulation = " & formProductNameFormulation & "	 ,UnitsOfProduct = '" & EscapeQuotes(formUnitsOfProduct) & "'	 ,IFPRating = '" & EscapeQuotes(formIFPRating) & "',Target = " & formTarget & ",HarvestDate = '" & formHarvestDate & "'	 ,Comments = '" & EscapeQuotes(formComments) & "',CreateDate = '" & formCreateDate & "' where ID = " & formID
		conn.execute sql, , 129
		Server.Transfer("SprayRecord_ins_thanks.asp")
		'END UPDATE
	END IF 
	'INSERT

'rem what is "</cfif>" ???
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		sql = "INSERT INTO SprayRecord(GrowerName,GrowerNumber,SprayDateDate,Crop,Bartlet,Stage,Location,Method,AcresTreated,RateAcre,TotalMaterialApplied,ProductNameFormulation,UnitsOfProduct,IFPRating,Target,HarvestDate,Comments,CreateDate)VALUES('" & formGrowerName & "'," formGrowerNumber " & formGrowerNumber & "," & formSprayDateDate & "
			</cfif>," formCrop " & formCrop & "," formBartlet " & formBartlet & "," formStage " & formStage & ",'" & formLocation & "'," formMethod " & formMethod & "," formAcresTreated " & formAcresTreated & "," formRateAcre " & formRateAcre & "," formTotalMaterialApplied " & formTotalMaterialApplied & "," formProductNameFormulation " & formProductNameFormulation & ",'" & formUnitsOfProduct & "','" & formIFPRating & "'," formTarget " & formTarget & "," & formHarvestDate & "
			</cfif>,'" & formComments & "'," & formCreateDate & "
			</cfif>)"
		conn.execute sql, , 129
		sql = "SELECT MAX(ID) AS insertid FROM SprayRecord"
		set rs = conn.execute(sql)
		urlID = rs.Fields("insertid")
		Server.Transfer("SprayRecord_ins_thanks.asp")
	END IF 'insert
END IF 'form submitted 

IF urlID  <>  0 and not errorFound THEN
	sql = "SELECT SprayRecord.ID, SprayRecord.GrowerName, SprayRecord.GrowerNumber, SprayRecord.SprayDateDate, SprayRecord.Crop, SprayRecord.Bartlet, SprayRecord.Stage, SprayRecord.Location, SprayRecord.Method, SprayRecord.AcresTreated, SprayRecord.RateAcre, SprayRecord.TotalMaterialApplied, SprayRecord.ProductNameFormulation, SprayRecord.UnitsOfProduct, SprayRecord.IFPRating, SprayRecord.Target, SprayRecord.HarvestDate, SprayRecord.Comments, SprayRecord.CreateDate FROM SprayRecord WHERE SprayRecord.ID = " & urlID & " ORDER BY  SprayRecord.ID"
	set rs = conn.execute(sql)
	IF NOT rs.eof THEN
		formGrowerName = rs.Fields("GrowerName")
		formGrowerNumber = rs.Fields("GrowerNumber")
		formSprayDateDate = rs.Fields("SprayDateDate")
		formCrop = rs.Fields("Crop")
		formBartlet = rs.Fields("Bartlet")
		formStage = rs.Fields("Stage")
		formLocation = rs.Fields("Location")
		formMethod = rs.Fields("Method")
		formAcresTreated = rs.Fields("AcresTreated")
		formRateAcre = rs.Fields("RateAcre")
		formTotalMaterialApplied = rs.Fields("TotalMaterialApplied")
		formProductNameFormulation = rs.Fields("ProductNameFormulation")
		formUnitsOfProduct = rs.Fields("UnitsOfProduct")
		formIFPRating = rs.Fields("IFPRating")
		formTarget = rs.Fields("Target")
		formHarvestDate = rs.Fields("HarvestDate")
		formComments = rs.Fields("Comments")
		formCreateDate = rs.Fields("CreateDate")
	END IF
END IF%>
<form action="SprayRecord_ins.asp" method="post" name="frmsearch">
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
<input type="hidden" value="<%=urlID%>" name="ID">
<tr><td valign="top" align="right"><span class="subtitle"><label for="GrowerName">GrowerName</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formGrowerName%>" name="GrowerName"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="GrowerNumber">GrowerNumber</label>(integer only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formGrowerNumber%>" name="GrowerNumber" size="10" maxlength="11" class="bodytext"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="SprayDateDate">SprayDateDate</label>:<br>(mm/dd/yyyy)</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formSprayDateDate%>" name="SprayDateDate" size="17" maxlength="21" class="bodytext"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="Crop">Crop</label>(integer only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formCrop%>" name="Crop" size="10" maxlength="11" class="bodytext"></span></td>
</tr>
<tr>
<td valign="top" align="right"><span class="subtitle"><label for="Bartlet">Bartlet:</label></span></td>
<td valign="top"><span class="bodytext"><span class="bodytext"><input type="radio" value="1" name="Bartlet" <% if ListContains(formBartlet,"True") OR ListContains(formBartlet,"1")  THEN %>Checked<% END IF %>>YES&nbsp;&nbsp;<input type="radio" value="0" name="Bartlet" <% if ListContains(formBartlet,"False") OR ListContains(formBartlet,"0")  THEN %>Checked<% END IF %>>NO</span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="Stage">Stage</label>(integer only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formStage%>" name="Stage" size="10" maxlength="11" class="bodytext"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Location">Location</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formLocation%>" name="Location"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="Method">Method</label>(integer only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formMethod%>" name="Method" size="10" maxlength="11" class="bodytext"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="AcresTreated">AcresTreated</label>(number only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formAcresTreated%>" name="AcresTreated" size="10" maxlength="38" class="bodytext"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="RateAcre">RateAcre</label>(number only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formRateAcre%>" name="RateAcre" size="10" maxlength="38" class="bodytext"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="TotalMaterialApplied">TotalMaterialApplied</label>(integer only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formTotalMaterialApplied%>" name="TotalMaterialApplied" size="10" maxlength="11" class="bodytext"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="ProductNameFormulation">ProductNameFormulation</label>(integer only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formProductNameFormulation%>" name="ProductNameFormulation" size="10" maxlength="11" class="bodytext"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="UnitsOfProduct">UnitsOfProduct</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formUnitsOfProduct%>" name="UnitsOfProduct"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr><td valign="top" align="right"><span class="subtitle"><label for="IFPRating">IFPRating</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formIFPRating%>" name="IFPRating"  class="bodytext" size="25" maxlength="50"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="Target">Target</label>(integer only):</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formTarget%>" name="Target" size="10" maxlength="11" class="bodytext"></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="HarvestDate">HarvestDate</label>:<br>(mm/dd/yyyy)</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formHarvestDate%>" name="HarvestDate" size="17" maxlength="21" class="bodytext"></span></td>
</tr>
<tr><td valign="top" align="right">
<span class="subtitle"><label for="Comments">Comments</label>:</span></td>
<td valign="top"><span class="bodytext"><textarea cols="30" rows="4" name="Comments" class="bodytext"><%=formComments%></textarea></span></td>
</tr>
<tr>
<td valign="top" align="right">
<span class="subtitle"><label for="CreateDate">CreateDate</label>:<br>(mm/dd/yyyy)</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formCreateDate%>" name="CreateDate" size="17" maxlength="21" class="bodytext"></span></td>
</tr>
<tr>
<td>&nbsp;</td>
<td><% IF  urlID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>
</cfoutput>
