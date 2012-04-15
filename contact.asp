<%@ LANGUAGE = VBScript %>
<%
Option Explicit
Response.Buffer = true
%>

<html>
<head>
	<title>Applied Spray Program | Contact Us</title>
    <link href="dropdown.css" rel="stylesheet" type="text/css">
    <link rel=stylesheet type="text/css" href="li_admin.css">
    <SCRIPT language="JavaScript" SRC="dropdown.js"></SCRIPT>
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">

<table width="1000" align="center" cellspacing="0" cellpadding="5" border="0" bgcolor=ffffff>
	<tr valign=bottom>
    	<td height=50><a href="default.asp">
        <%if session("fullrights") and not isnull(session("logofile")) and session("logofile") >"" then%>
    		<img src="/logos/p<%=session("logofile")%>" border=0 hspace="0" />
		<%else %>
    		<img src="images/logo_small.png" width="42" height="42" border="0"><img src="/images/agsoft.jpg" border=0 hspace="10" />
	    <%end if %>
    	
    	</a></td>
		<td align="right" style="color:black;">You are not logged in | <a href="loginform.asp">Login</a></td>
	</tr>
</table>
<table width="1000" align="center" cellspacing="0" cellpadding="0" border="0" bgcolor=013166>
	<tr>
		<td align="right" style="font-size:6pt;" height=8></td>
	</tr>
</table>
<table width="1000" height=600 align="center" cellspacing="0" cellpadding="0" border="0" bgcolor=beige>

    <tr bgcolor=#8DB33C><td height=25>
    
        <table align=center border=0 cellpadding="0" cellspacing="0">
        <tr valign="middle"> 
        <td width="10">&nbsp;</td>
        <td width="1">|</td>
        <td width="100" align=center nowrap><a href="default.asp" class="navtext">Welcome</a></td>
        <td width="1">|</td>
        <td width="100" align=center nowrap><a href="learn.asp" class="navtext">Learn More</a></td>
        <td width="1">|</td>
        <td width="100" align=center><a href="contact.asp" class="navtext">Contact Us</a></td>
        <td width="1">|</td>
        <td width="75" align=center><a href="loginform.asp" class="navtext">Login</a></td>
        <td width="1">|</td>
        <td>&nbsp;</td>
        </tr>
        </table>

    </td></tr>
    
	<tr valign="top">
		<td valign="top">
                   
        <br />
        
            <table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Contact Us <br>
                    &nbsp;</h1></td></tr></table>

                <table cellpadding=10 cellspacing=0 border=0 width="100%"><tr><td>
                <p>For more information on the SprayCrop Program or to schedule a demo, please contact us at Unison AgSoft:</p>
                
                <blockquote>
                  <p>Kent  Heighton <br>
                      <a href="mailto:&#107;&#101;&#110;&#116;&#64;&#103;&#111;&#114;&#103;&#101;&#116;&#101;&#99;&#46;&#99;&#111;&#109">&#107;&#101;&#110;&#116;&#64;&#103;&#111;&#114;&#103;&#101;&#116;&#101;&#99;&#46;&#99;&#111;&#109</a><br>
                    541-386-7409 <br>
                    PO Box 801<br>
                    Hood River, OR 97031 </p>
                  <p>Heidi Ribkoff<br>
                    <a href="mailto:&#104;&#101;&#105;&#100;&#105;&#114&#64;&#103;&#111;&#114;&#103;&#101;&#116;&#101;&#99;&#46;&#99;&#111;&#109">&#104;&#101;&#105;&#100;&#105;&#114&#64;&#103;&#111;&#114;&#103;&#101;&#116;&#101;&#99;&#46;&#99;&#111;&#109</a><br>
                  </p>
                </blockquote>
                <p>&nbsp;</p>
                
                </td></tr></table>



		</td>
	</tr>
</table>

<table width="1000" align="center" cellspacing="0" cellpadding="0" border="0" bgcolor=013166>
	<tr>
		<td align="right" style="font-size:6pt;" height=8></td>
	</tr>
</table>

<table width="1000" align="center" cellspacing="0" cellpadding="5" border="0" bgcolor=ffffff>
	<tr>
		<td align="left" style="color:black;">Copyright &copy; <%=year(now())%> Unison AgSoft.  All Rights Reserved.</td>
	</tr>
</table>

</body>
</html>
