<%Option Explicit%>
<% 
	DIM conn,rs,sql
    if not session("login") then 
        response.redirect("loginForm.asp") 
    end if 
%> 

<html>
<head>
	<title>GBFS Admin - Recipes List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->
<table width="594" border="1" cellspacing="0" cellpadding=" 0" bordercolor="#013166" bgcolor="#beige"><tr><td><img src="images/spacer.gif" height="4" width="1" border="0"><br>
&nbsp;&nbsp;<h1>Grower's Resource Section<h1><br><img src="images/spacer.gif" height="4" width="1" border="0"></td></tr></table>
<table width="95%" border="0" bgcolor="ffffff" align="center">

      <tr>
        <td height="14" valign="top" class="content"><!-- InstanceBeginEditable name="EditRegion2-content" -->
          <p style="color: #B33E44; font-size: 14px; font-weight: bold">Links for Growers</p>
          <ul>
            <li><a href="http://wwwsqfi.com" target="_blank" style="color: #000000; font-weight: bold">Safe Quality Food (SQF) Institute</a> The SQF Program is a fully integrated food safety and quality management protocol designed specifically for the food sector. The SQF program is owned by the Food Marketing Institute (FMI).
                <ul>
                  <li><a href="http://www.sqfi.com/trademark/SQF1000_Certification_Trademark.pdf" target="_blank" style="color: #000000; font-weight: bold">SQF 1000 Certification Trademark Rules</a> </li>
                  <li><a href="http://www.sqfi.com/trademark/SQF2000_Certification_Trademark.pdf" target="_blank" style="color: #000000; font-weight: bold">SQF 2000 Certification Trademark Rules</a> </li>
                </ul>
            </li>
            <li>Spray Record Regulation Information </li>
            <li>State
                <ul>
                  <li><a href="http://oregon.gov/ODA/PEST/purs_index.shtml" target="_blank" style="color: #000000; font-weight: bold">ODA Pesticides Division</a> </li>
                </ul>
            </li>
            <li>Federal              
              <ul>
                <li><a href="http://www.epa.gov/pesticides/regulating/" target="_blank" style="color: #000000; font-weight: bold">EPA - Regulating Pesticides</a></li>
              </ul>
            </li>
            </ul>
          <p style="font-weight: bold; color: #B33E44; font-size: 14px">Weather Sites </p>
          <ul>
            <li><a href="http://www.usbr.gov/pn/agrimet/wxdata.html" target="_blank" style="color: #000000; font-weight: bold">AgriMet</a> The Pacific Northwest Cooperative Agricultural Weather Network </li>
            <li><a href="http://www.clearwest.com/" target="_blank" style="font-weight: bold; color: #000000">Clearwest</a> An Agricultural Weather Forecast Company</li>
            <li> <a href="http://webpages.charter.net/hoodriverweather/weather.htm" style="font-weight: bold; color: #000000">Hood River Weather</a></li>
            <li><a href="http://www.weatherunderground.com/" style="font-weight: bold; color: #000000">Weather Underground</a> </li>
            </ul>
          <p style="color: #B33E44; font-weight: bold; font-size: 14px">Horiculture Sites</p>
          <ul>
            <li><a href="http://www.tfrec.wsu.edu/index.php" target="_blank" style="font-weight: bold; color: #000000">Washington State University Tree Fruit Research Center</a></li>
            <li><a href="http://www.cdms.net/manuf/manuf.asp" target="_blank" style="font-weight: bold; color: #000000">Crop Data Management Systems</a></li>
            <li><a href="http://www.ncw.wsu.edu/treefruit/fireblight/2000f.htm" target="_blank" style="font-weight: bold; color: #000000">Cougarbllight Model</a></li>
            <li><a href="http://www.goodfruit.com/" target="_blank" style="font-weight: bold; color: #000000">Good Fruit Grower Magazine</a></li>
            <li><a href="http://oregonstate.edu/dept/mcarec/" target="_blank" style="font-weight: bold; color: #000000">Hood River Experiment Station</a></li>
            <li><a href="http://extension.oregonstate.edu/wasco/" target="_blank" style="font-weight: bold; color: #000000">Wasco County Extension</a></li>
            <li><a href="http://www.ncw.wsu.edu/treefruit/" target="_blank" style="font-weight: bold; color: #000000">Washington State University North Central Extension </a></li>
          </ul>
          <p>&nbsp;</p>
        <!-- InstanceEndEditable --></td>
      </tr>
    </table>

<!--#include file="i_adminfooter.asp" -->
</body>
</html> 