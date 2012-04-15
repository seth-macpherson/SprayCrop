<%Option Explicit%>
<%if not session("login") or not listContains("3,2", session("accessid")) then
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
DIM formName,formAgreeToTerms
formName = Request.Form.Item("Name")
formAgreeToTerms = Request.Form.Item("AgreeToTerms")

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")


'Form Was Submitted
IF  Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF Request.Form.Item("AgreeToTerms") = "" THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>You must agree to the Terms"
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Name"), "nvarchar","Name", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF

	'Update record
	if session("growerid")<>0 then
	    IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		    urlID = AgreeToGrowerTerms(formName)
		    EndConnect(conn)
		    set rs = nothing
		    session("termsagreed") = 1
		    Response.Redirect("enterspraydata.asp")
		    'END UPDATE
	    END IF 
	else
	    IF NOT errorFound AND Request.Form.Item("update") <> "" THEN  
		    urlID = AgreeToPackerTerms(formName)
		    EndConnect(conn)
		    set rs = nothing
		    session("termsagreed") = 1
		    Response.Redirect("grower_report.asp")
		    'END UPDATE
	    END IF 
	end if
	
END IF	
%>
<html>
<head>
	<title><%=Application("BusinessName")%>&nbsp;-&nbsp;<%=Application("ProgramName")%>&nbsp;-&nbsp;Agree to Terms</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">

<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><img src="images/spacer.gif" height="4" width="1" border="0"><br>
<h1>> Terms of Use</h1><br><img src="images/spacer.gif" height="4" width="1" border="0"></td></tr></table><br />

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<% if  errorFound then%>
<tr>
<td colspan="21" class="bodytext" valign="top"><a href="#edit"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></a></td>
</tr>
<% end if %>
<tr><td bgcolor="FFFFFF" class="bodytext">

<strong>1.	The Site</strong><br>
Welcome to the restricted portion of the <%=Application("CLIENT_NAME")%> website (the "Service"), owned by the <%=Application("CLIENT_NAME")%>.  The Service allows growers to post reports about their pesticide use, and easily calculate when it is safe for workers to return after spraying.  The Service is provided as a courtesy to you by <%=Application("CLIENT_NAME")%>.  <br><br>

<strong>2.	No Guaranteed Availability</strong><br>
<%=Application("CLIENT_NAME")%> reserves the right at any time and from time to time to modify or discontinue the Service (or any part thereof), temporarily or permanently, with or without notice to you. You agree that <%=Application("CLIENT_NAME")%> will not be liable to you for any modification, suspension or discontinuance of the Service. <br><br>

<strong>3.	Compliance with Applicable Laws</strong><br>
You are solely responsible for your compliance with all applicable federal, state and local laws and regulations, including but not limited to the EPA Worker Protection Standards set out at 40 CFR § 170, and any additional requirements imposed by pesticide labels or special permits.  <%=Application("CLIENT_NAME")%> specifically disclaims any and all warranties, express or implied, regarding use of the Service to ensure compliance with the laws of any jurisdiction.  You assume the entire risk as to whether or how your use of the Service may satisfy or help you satisfy your obligations under any applicable law or regulation.  For additional information about the regulations relating to notice of pesticide application, see the EPA booklet "The Worker Protection Standard for Agricultural Pesticides" and the Oregon Farmers Handbook, both available online:<br><br>

<a href="http://www.epa.gov/oecaagct/htc.html" target="_blank">http://www.epa.gov/oecaagct/htc.html</a> <br><br>

<a href="http://www.oregon.gov/ODA/pub_fh_index.shtml</a>" target="_blank">http://www.oregon.gov/ODA/pub_fh_index.shtml</a><br><br>

<strong>4.	Termination of Service to You</strong><br>
You agree that <%=Application("CLIENT_NAME")%>, in its sole discretion, may terminate your password, account (or any part thereof) or use of the Service, and remove and discard any content within the Service, for any reason, including, without limitation, our belief that you have violated or acted inconsistently with the letter or spirit of these Terms. You agree that any termination of your access to the Service under any provision of these Terms may occur without prior notice to you, and you also agree that <%=Application("CLIENT_NAME")%> will not be liable to you for any termination of your access to the Service.  <br><br>

<strong>5.	Restrictions on Use</strong><br>
You agree not to make any use of the Service or transmit any content that is: (1) unlawful under the laws of any jurisdiction to which you or <%=Application("CLIENT_NAME")%> are subject;  (2) harmful, threatening, harassing, defamatory, invasive of the privacy of another; (3) insider information, or any other proprietary or confidential information;  (4)  an infringement of any patent, trademark, trade secret, copyright or other intellectual property right;  (5) falsified, including without limitation the use of forged headers or otherwise manipulated identifiers in order to disguise its origin; or (6) containing or transmitting software viruses or any other malicious computer code, files or programs designed to interrupt, destroy or limit the functionality of any computer software.  You agree to indemnify and hold harmless <%=Application("CLIENT_NAME")%> from any liability incurred as the result of your violation of these Terms.    <br><br>

<strong>6.	Disclaimer of Warranties and Limitation of Liability</strong><br>
<strong>YOU AGREE THAT THE SERVICE IS PROVIDED TO YOU "AS IS" AND "AS AVAILABLE" WITHOUT ANY WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, ACCURACY AND NON-INFRINGEMENT. </strong> <%=Application("CLIENT_NAME")%> does not warrant that the Service will be available at any given time, secure, accurate or free of error.  You use the Service at your own risk, and you assume the risk that the Service may provide incorrect information to you or your workers, as well as the risk that any material downloaded by you from the Service may cause loss of data or damage to your computer system.  <br>
 <strong>YOU UNDERSTAND AND AGREE THAT IN NO EVENT WILL <%=Application("CLIENT_NAME")%> BE LIABLE FOR ANY DIRECT OR INDIRECT DAMAGES, EVEN IF <%=Application("CLIENT_NAME")%> IS AWARE OF THE POSSIBILITY OF SUCH DAMAGES, INCLUDING WITHOUT LIMITATION LOSS OF PROFITS OR FOR ANY OTHER SPECIAL, CONSEQUENTIAL, EXEMPLARY OR INCIDENTAL DAMAGES, HOWEVER CAUSED, WHETHER BASED UPON CONTRACT, NEGLIGENCE, STRICT LIABILITY IN TORT, WARRANTY, OR ANY OTHER LEGAL THEORY, ARISING OUT OF OR RELATED TO YOUR USE OF THE SERVICE.  THE PARTIES INTEND THAT THIS LIMITATION SHOULD APPLY EVEN IF IT CAUSES ANY WARRANTY TO FAIL OF ITS ESSENTIAL PURPOSE.</strong>   <br><br>

<strong>7.	Indemnity.</strong><br>

You agree to indemnify, defend and hold harmless <%=Application("CLIENT_NAME")%> and its officers, directors, and employees from and against all fines, suits, proceedings, claims, causes of action, demands, or liabilities of any kind or of any nature arising out of or in connection with your use of the Service, including without limitation claims alleging that workers were misinformed about the safety of areas sprayed with pesticides.  <br><br>

<strong>8.	Severability & Waiver</strong><br>
The invalidity of any term or provision of these Terms will not affect the validity of any other provision.  Waiver by <%=Application("CLIENT_NAME")%> of strict performances of any provision of these Terms will not be a waiver of or prejudice <%=Application("CLIENT_NAME")%>'s right to require strict performance of the same provision in the future or of any other provision of these Terms.<br><br>

<strong>9.	Entire Agreement</strong><br>
These Terms constitute the entire agreement between the parties as to their subject matter, and there are no other terms, conditions, or obligations between the parties relating to the use of the Service, other than those contained in these Terms.  No modification of these Terms will be valid unless in writing and signed by both parties.<br><br>

</tr>
<tr>
<td colspan="2" class="bodytext">
<table width="90%" border="0" cellpadding="2" cellspacing="0">

<form action="agreetoterms.asp" method="post" name="frmsearch">
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr>
<td>&nbsp;</td><td align="left" class="bodytext">All fields are required.</td>
</tr>
<% if errorFound then%>
<tr>
<td>&nbsp;</td>
<td class="bodytext"><a name="edit"><font color="red">
<% =errorMessage%></font></a></td>
</tr>
<% End If %>
<tr><td valign="top" align="right"><span class="subtitle"><label for="Name">Enter your full name</label>:</span></td>
<td valign="top"><span class="bodytext"><input type="text" value="<%=formName%>" name="Name"  class="bodytext" size="25" maxlength="150"></span></td>
</tr>
<tr><td valign="top" colspan="2" align="center"><input type="checkbox" name="agreetoterms" value="1"><span class="subtitle"><label for="GrowerName">I have read, understand and agree to the above terms.  </label></span></td>

</tr>

<tr>
<td>&nbsp;</td>
<td><input type="submit" name="update" value="Update" class="bodytext"></td>
</tr>

</table>
</form>
<%
	set rs = nothing
	EndConnect(conn)
%>

</td></tr>

</table>

<!--#include file="i_adminfooter.asp" -->

</td></tr>

</table>

</body>
</html>

