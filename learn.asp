<%@ LANGUAGE = VBScript %>
<%
Option Explicit
Response.Buffer = true
%>

<html>
<head>
<title>Applied Spray Program | Learn More</title>
    <link href="dropdown.css" rel="stylesheet" type="text/css">
    <link rel=stylesheet type="text/css" href="li_admin.css">
    <SCRIPT language="JavaScript" SRC="dropdown.js"></SCRIPT>
<style type="text/css" media="screen">
<!--
@import url("p7tp/p7tp_08.css");
-->
</style>
<script type="text/javascript" src="p7tp/p7tpscripts.js"></script>
</head>

<body bgcolor="FFFFFF" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0" onLoad="P7_initTP(8,0)">

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
        
            <table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><h1> > Learn More About the SprayCrop Program <br>
                    &nbsp;</h1></td></tr></table>

                <table cellpadding=10 cellspacing=0 border=0 width="100%"><tr><td>
                <p><strong>SprayCrop Program  Attributes:</strong></p>
                <ul>
                  <li>Web-based, accessible by any computer connected to the Internet</li>
                  <li>No bulky program on your computer, makes program updates seamless</li>
                  <li>Password protected so your data is safe, secure  and private</li>
                  <li>Record key, reportable, information  about crop care applications</li>
                  <li>Also record non-required spray records for your information</li>
                  <li>Produce printed reports for your records, state reporting, and  field-posting notices</li>
                  <li>Easy electronic upload of your spray records to the State of Oregon Pesticide Use Reporting System</li>
                  <li>Current Spray Lists updated monthly</li>
                  <li>Easy to use, proven by 2 years use by over 100 growers in the Hood River   Valley</li>
                  <li>Designed to minimize required typing time</li>
                  <li>Programmed to enable easy duplication of repetitive data</li>
                  <li>Customizable fields are retained in drop-down selection boxes for your   subsequent records entry</li>
                  <li>Three custom versions available for Growers and  Packer/Shipper/Associations</li>
                  <li>Training packages available</li>
                  <li>Unlimited phone support</li>
                  </ul> 
                  <div id="p7TP1" class="p7TPpanel">
                    <div class="p7TPheader">
                      <h3><strong>Available Versions</strong></h3>
</div>
                    <div class="p7TPwrapper">
                      <div class="p7TP_tabs">
                        <div id="p7tpb1_1" class="down"><a class="down" href="javascript:;">Version 1 </a></div>
                        <div id="p7tpb1_2"><a href="javascript:;">Version 2 </a></div>
                        <div id="p7tpb1_3"><a href="javascript:;">Version 3 </a></div>
                        <br class="p7TPclear">
                      </div>
                      <div class="p7TPcontent">
                        <div id="p7tpc1_1">
                          <h4>Version 1 - Grower...</h4>
                          <p>As a Grower you have the capability to  enter your spray records easily via our web-based program. Make reports to  print for your records, post in the field or send to your packers or  wholesalers. Your data is safe and secure and only accessible to you. </p>
                          <p>Easily meet your Oregon  reporting requirement by uploading your entered spray data to the State of Oregon Pesticide Use Reporting System.</p>
                          <p>Spray data recorded includes dates, crop  and variety, applicator and applicator license, supervisor and their applicator  license, weather, location, spray name with EPA number, quantity, number of  acres receiving application, target, and application method.&nbsp; An ample comment box is available for you to  record specific remarks and observations important to you.</p>
                          <p>The SprayCrop program allows you to  customize your locations, blocks, applicators, supervisors, and weather.&nbsp; This program is readily accessible via any  computer with Internet access.&nbsp; Once  purchased, you&rsquo;ll receive a user logon name and password that will give you  access to efficient spray recordkeeping and reporting.</p>
                          <p>You, the user, are ultimately responsible  for accurate data entry and record keeping and reporting.&nbsp; To assist you, the SprayCrop program</p>
                          <ol>
                            <li>Will flag over applications.&nbsp; This flag is meant to help alert you of  potential overuse and/or data entry error.</li>
                            <li>Will allow entry of only spray  products that are entered in our system as approved for specific crops.</li>
                            <li>Has certain required reporting information. An error message will specify missing data and the record will not be successfully recorded until this error is corrected.</li>
                          </ol>
                          <p>Unlimited phone support is available to you  after the initial training. The spray list is regularly updated and you receive  all product updates and new features.</p>
                        </div>
                        <div id="p7tpc1_2">
                          <h4>Version 2 - Packer/Shipper/Associate...</h4>
                          <p><strong>View Only Rights </strong></p>
                          <p>As a Packer, you have the capability to  generate reports of spray data entered by any growers associated to you (by the  Growers themselves). Keep this data safe and secure and generate reports as  needed for your wholesalers or your own needs. You do not have the capability  to enter the data, however, if need be you can submit the data to PURS. If need  be, at a glance you can see what sprays your Growers are applying and when they  are being applied. </p>
                          <p>The reports that you generate will include  all of the data entered by your associated Growers. Spray data recorded  includes dates, crop and variety, applicator and applicator license, supervisor  and their applicator license, weather, location, spray name with EPA number,  quantity, number of acres receiving application, target, and application  method. When the grower has submitted their information to PURS, it is also  noted on the reports. Any comments made by the Grower are also noted in the  reports. Either print or save these reports to your local computer to keep as a  record or submit to those that need the information. Unlimited phone support is  available to you after the initial training. The spray list is regularly  updated and you receive all product updates and new features.</p>
                        </div>
                        <div id="p7tpc1_3">
                          <h4>Version 3 - Packer/Shipper/Associate and Grower Package...</h4>
                          <p><strong>Full Rights </strong></p>
                          <p>As a Packer/Distributor you are able to  assign your Growers the use of the SprayCrop program, to generate reports as  needed for your records and to submit to wholesalers. All of the capabilities  that Version 2 offers are also available to you (see above for details). As the  administrator, you are also able to add users both at your Packer level and at  the Grower level. Your data and that of your Growers will be safe and secure.  You may also have the capability to add individual Grower spray data, edit the  spray list and selected areas, such as the crop varieties. There are several  levels of annual license fee depending on the number of Growers that you have.</p>
                          <p>You will also have the option of adding  your Company logo to the SprayCrop program Web pages used by you and your growers.</p>
                          <p>Unlimited phone support is available to you  (you will be able to support your Growers) after the initial training. The  spray list is regularly updated and you receive all product updates and new features. </p>
                        </div>
                      </div>
                    </div>
                    <!--[if lte IE 6]>
<style type="text/css">.p7TPpanel div,.p7TPpanel a{height:1%;}.p7TP_tabs a{white-space:nowrap;}</style>
<![endif]-->
</div>
                <p>&nbsp; </p>
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
