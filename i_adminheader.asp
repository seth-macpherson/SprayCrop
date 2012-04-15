<!--#include file="include/i_data.asp"-->


<%
    dim isGrower,isPacker,isAdmin


    if session("adminid")>0 then 
        isAdmin=true
    elseif session("growerid")>0 then 
        isGrower=true
    elseif session("packerid")>0 then 
        isPacker=true
    end if
    
	dim rsSprayYear, sActiveSprayYear, l_sql
	set conn = Connect()
	set rsSprayYear = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT SprayYear FROM SprayYears WHERE Active = 1"
	set rsSprayYear = conn.execute(l_sql)
	if not rsSprayYear.EOF then
		sActiveSprayYear = rsSprayYear(0)
	else
		sActiveSprayYear = "NONE"
	end if
	rsSprayYear.Close
	set rsSprayYear = Nothing

    'JMS
    function selectoption(a,b)
        if a=b then 
            selectoption="selected"
        else
            selectoption=""
        end if
    end function
    function so(a,b)
        so=selectoption(a,b)
    end function
    
%>
 
<table width="1000" align="center" cellspacing="0" cellpadding="5" border="0" bgcolor=#ffffff>
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
		<td align="right" style="color:black;"><strong><%=sActiveSprayYear%></strong> |  Logged in as <strong><%=session("username")%> 
		    </strong> 
		    <%select case session("accessid") 
		        case 3
		            response.Write " (Grower)" 
    		        
    		        'if false then
                        'dim roleConn: set roleConn = Connect()
		                'dim rsRoles: set rsRoles = conn.execute("exec growerunit$bygrower " & session("growerid")) 
                       ' with response
                        
                        '.Write "<select name=growerrole onchange=document.roles.changerole.value=1;document.roles.submit();>"
                        'do until rsRoles.eof
                           
                       '     .Write "<option value="""&rsRoles("growerid")&"|"&rsRoles("growername")&""""
                       '     if cint(session("growerid"))=rsRoles("growerid") then .Write " selected"
                        '    .Write ">"
                        '    .Write rsRoles("growername")
                        '    .write "</option>"
                            
                        'rsRoles.movenext
                        'loop
                        '.Write "</select>"
                        
                        'end with
                   	    'EndConnect(roleConn)
        		        
    		        'end if
    		        
		        case 2: response.Write " (PSA)" 
		        case 1: response.Write " (Admin)" 
		        
		      end select
		    %>
		    | <a href="logout.asp">Log Out</a></td>
	</tr>
</table>

<table width="1000" align="center" cellspacing="0" cellpadding="0" border="0" bgcolor=013166>
	<tr>
		<td align="right" style="font-size:6pt;" height=8></td>
	</tr>
</table>

<table width="1000" height=600 align="center" cellspacing="0" cellpadding="0" border="0" bgcolor=beige>

    <tr bgcolor=#8DB33C><td height=25>

            <SCRIPT language="JavaScript" SRC="dropdown.js"></SCRIPT>

	        <%
	        if session("termsagreed") then
                if listContains("1", session("accessid")) then
                    server.Execute "include/incHtmNav-Admin.asp"
                elseif listContains("2", session("accessid")) then   
                    server.Execute "include/incHtmNav-Packer.asp"            
                elseif listContains("3", session("accessid")) then    
                    server.Execute "include/incHtmNav-Grower.asp"          
                end if
            else
                server.Execute "include/incHtmNav-Default.asp"              
            end if
            %>	
    
    </td>
    </tr>
	
	<tr valign="top" >
		
		<td width="100%">
        
        <br />