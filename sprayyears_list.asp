<%Option Explicit%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayYears.asp"-->
<%
'CREATED by LocusInteractive on 08/04/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
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
	Response.Redirect("sprayyears_list.asp")
END IF

'Initialize Form Fields
DIM formSprayYear,formActive
formSprayYear = Request.Form.Item("SprayYear")
formActive = Request.Form.Item("Active")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")

IF Request.QuerySTring("task") = "d" and urlSprayYearID <> "" THEN
	DeleteSprayYears(urlSprayYearID)
	Response.Redirect("sprayyears_list.asp")
END IF 
IF Request.QuerySTring("task") = "activate" and urlSprayYearID <> "" THEN
	ActivateSprayYears(urlSprayYearID)
	Response.Redirect("sprayyears_list.asp")
END IF 
IF Request.QuerySTring("task") = "deactivate" and urlSprayYearID <> "" THEN
	DeActivateSprayYears(urlSprayYearID)
	Response.Redirect("sprayyears_list.asp")
END IF 


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
		urlSprayYearID = UPDATESprayYears(formSprayYearID,formSprayYear)
Response.Redirect("sprayyears_list.asp")
		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
		urlSprayYearID = InsertSprayYears(formSprayYear)
Response.Redirect("sprayyears_list.asp")
	END IF 'insert	
END IF 'form submitted 

IF formSprayYearID  <>  0 and not errorFound THEN
set rs = GetSprayYearsByID(formSprayYearID)
	IF NOT rs.eof THEN
		formSprayYear = rs.Fields("SprayYear")
	END IF
END IF%>
<html>
<head>
	<title>SprayYears List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Spray Years</h1><br>&nbsp;</td></tr></table>

<table width="95%" align=center border="0" bgcolor="ffffff">
<tr><td bgcolor="ffffff" class="bodytext"><br>
<a href="sprayyears_list.asp#edit">Add Spray Year</a></td>
</tr>
<tr>
<td colspan="2" class="bodytext">
<br>
<table width="90%" border="1" cellpadding="2" cellspacing="0">
<% if  delerror then%>
<tr>
<td colspan="5" class="bodytext" valign="top"><font color="red"><%= delerrormessage %></font></td>
</tr>
<% end if %>
<% if  errorFound then%>
<tr>
<td colspan="5" class="bodytext" valign="top"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></td>
</tr>
<% end if %>
<tr>
<!--<td  valign="top">&nbsp;</td>-->
<!--<td  valign="top"><h2>Edit</h2></td>-->

	<td  valign="top"><h2>SprayYear</h2></td>
	<td valign="top"><h2>Activate</h2></td>

</tr>
<%
	set rs = GetAllSprayYears()
%>
<%IF not rs.EOF THEN
DO WHILE not rs.eof %>
<tr>
<!--<td class = "bodytext" valign="top"><%=rs.Fields("SprayYearID")%></td>-->
<!--<td  valign="top" class="bodytext">
<a href="sprayyears_list.asp?SprayYearID=<%=rs.Fields("SprayYearID")%>#edit" class="bodytext">Edit</a><br
<a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="sprayyears_list.asp?SprayYearID=<%=rs.Fields("SprayYearID")%>&task=d" class="bodytext">Delete</a>
</td>-->
<td class="bodytext" valign="top"><%=rs.Fields("SprayYear")%>&nbsp;</td>
<td class="bodytext"><% IF rs.Fields("Active") =  0 then%>Not Active<br><a href="sprayyears_list.asp?task=activate&SprayYearID=<%=rs.Fields("SprayYearID")%>" onclick="javascript: return confirm('Are you sure you want to activate this record?');">Make Active</a><%else%>Active<br><%end if%></td>
</tr> 
<% 
rs.MoveNext
LOOP

Else
%>
<tr><td class="bodytext" colspan="5">No Records Selected</td></tr>
<%	end if %>
</table>
<a name="edit"></a>
<form action="SprayYears_list.asp" method="post" name="frmsearch">
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr>
<td>&nbsp;</td><td align="left" class="bodytext">All Active Spray List Products from the current active Year will populate the added Spray Year.</td>
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
<td>&nbsp;</td>
<td><% IF  urlSprayYearID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Insert" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>

</table>
</form>

<!--#include file="i_adminfooter.asp" -->
