<%@ LANGUAGE = VBScript %>
<%

Option Explicit
Response.Buffer = true

dim conn,urlrs: set conn = Connect()
set urlrs = conn.execute("exec packer$whois '" & request.ServerVariables("server_name") & "'")

if not urlrs.eof then
    session("fullrights")=true
    session("logofile")=urlrs.Fields("packernumber")&urlrs.fields("logofileext")
    'response.Redirect "loginform.asp?whois"
end if

select case session("accessid")
case 1: response.Redirect "sprayrecords_list.asp"
case 2: response.Redirect "sprayrecords_list.asp"
case 3: response.Redirect "enterspraydata.asp"
case else:
    '# do nothing
end select


%>

<!--#include file="include/i_data.asp"-->

<html>
<head>
	<title>Applied Spray Program | Agricultural Crop Spray Application</title>
    <link href="dropdown.css" rel="stylesheet" type="text/css">
    <link rel=stylesheet type="text/css" href="li_admin.css">
    <SCRIPT language="JavaScript" SRC="dropdown.js"></SCRIPT>
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">

<table width="1000" align="center" cellspacing="0" cellpadding="5" border="0" bgcolor=ffffff>
	<tr valign=bottom>
    	<td height=50>
    	<a href="default.asp">
		<%if session("fullrights") and not isnull(session("logofile")) and session("logofile") >"" then%>
    		<img src="/logos/p<%=session("logofile")%>" border=0 hspace="0" />
		<%else %>
    		<img src="images/logo_small.png" width="42" height="42" border="0"><img src="/images/agsoft.jpg" border=0 hspace="10" />
	    <%end if %>
		</a>
    	</td>
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
        
            <table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Welcome</h1><br>&nbsp;</td></tr></table>

                <table cellpadding=10 cellspacing=0 border=0 width="100%"><tr><td>
                <p>
                The SprayCrop program is a web-based database for agricultural crop spray application & reporting records.                </p>
                
                <ul>
                <li>Access SprayCrop instantly via the Internet, no bulky program on your computer, makes updates seamless 
                <li>Cost effective and Easy to use
                <li>All data is kept safe, secure and private 
                <li>Grower tested and approved 
                <li>Easily upload spray reports to State of Oregon PURS
                </ul>

                <p>The SprayCrop Program:
                </p>
                
                <ol>
                <li>Records required reportable data you enter about your crop care applications 
                <li>Allows easy printing of reports
                <li>Provides easy electronic upload of your spray data to the State of Oregon PURS program
                <li>Allows you to customize your own blocks & properties to track applications as you wish
                </ol>

                <p>The SprayCrop Program is designed for easy data entry &amp; programmed to enable easy duplication of repetitive information. It is grower tested and approved as a secure, cost-effective record keeping and reporting tool.</p>
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
