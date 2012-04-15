<%Option Explicit%>
<% 
	DIM conn,rs,sql
    if not session("login") then 
        response.redirect("loginForm.asp") 
    end if 
%> 

<html>
<head>
	<title><%=Application("CLIENT_NAME")%> - News/Info</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->
<table width="594" border="1" cellspacing="0" cellpadding=" 0" bordercolor="#013166" bgcolor="#99CCCC"><tr><td><img src="images/spacer.gif" height="4" width="1" border="0"><br>
&nbsp;&nbsp;<h1>News &amp; Information<h1><br><img src="images/spacer.gif" height="4" width="1" border="0"></td></tr></table>
<table width="95%" border="0" bgcolor="ffffff" align="center">
	<tr>
		<td width="550" height="15" valign="top" bgcolor="#FFFFFF" class="main-title"><!-- InstanceBeginEditable name="EditRegion1-title" --><h1 </h1><!-- InstanceEndEditable --></td>
	</tr>
	<tr>
		<td height="14" valign="top" class="content"><!-- InstanceBeginEditable name="EditRegion2-content" -->
<!--
          <p>Click here to read the <a href="../pdf/Harvest%20letter%202005%20doc.pdf" target="_blank" style="color: #000000; font-weight: bold">2005 Harvest Letter</a>. Here are the highlights:</p>
          <ul>
            <li> General Receiving Guidelines</li>
            <li> Ziram/Nutra-Phos 24 Applications</li>
            <li> General Sanitation</li>
            <li> Plastic Bins</li>
            </ul>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>Note: This PDF document requires Acrobat Reader for viewing. Click below <font face="Verdana, Arial, Helvetica, sans-serif" size="-1"></font> for a free download. <a href="http://www.adobe.com/products/acrobat/readstep2.html" target="_blank"><img src="./images/getreader.gif" width="88" height="31" border="0"></a></p>
-->
        <!-- InstanceEndEditable --></td>
	</tr>
</table>
<!--#include file="i_adminfooter.asp" -->
</body>
</html> 