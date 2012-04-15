<%@ LANGUAGE = VBScript %>
<%
Option Explicit
%>

<!--#include file="include/i_data.asp"-->
<!--#include file="i_growers.asp"-->
<% 
Dim u, p, e
    u = lcase(request.form("username")) 
    p = lcase(request.form("password")) 
    e = lcase(request.form("email")) 
 
 	if e <> "" then
		'initialize the connection
		set conn = Connect()
		set rs = Server.CreateObject("ADODB.RecordSet")
		response.write e
		r = EmailPassword(e)
		set rs = nothing
		EndConnect(conn)
		response.redirect("loginForm.asp?r=" & (r+1))
	end if
 
    '--------------------------------------------------------- 
    '-- check to see that the form was completely filled out-- 
    '--------------------------------------------------------- 
    if u="" or p="" then 
        response.redirect("loginForm.asp") 
    end if 

	Dim conn,sql,rs,counter
	'initialize the connection
	set conn = Connect()
	set rs = Server.CreateObject("ADODB.RecordSet")
	
 	sql = "SELECT * FROM PackerUsers u INNER JOIN packers p ON u.packerid=p.packerid WHERE u.Username = '" & u & "' AND u.Password = '" & p & "'"
	set rs = conn.execute(sql)

    '--------------------------------------------------------- 
    '-- check for a match, this could be against a database!-- 
    '--------------------------------------------------------- 
    IF rs.eof THEN

	 	sql = "SELECT gu.*, g.* FROM GrowerUsers gu INNER JOIN growers g on gu.growerid = g.growerid WHERE gu.Username = '" & u & "' AND gu.Password = '" & p & "'"
		set rs = conn.execute(sql)
	    
	    IF rs.eof THEN
	    	    
 	        sql = "SELECT * FROM administrators WHERE Username = '" & u & "' AND Password = '" & p & "'"
	        set rs = conn.execute(sql)
	    
	        if not rs.eof then
	        
                session("login")=true 
                session("termsagreed") = 1
                session("username")=u 
                session("packerid")=0
                session("growerid")=0
		        session("accessid") = rs.Fields("AccessID") 
		        response.Redirect "sprayrecords_list.asp"
	        
	        else
    	           
	            'access denied 
	            session("login")=false 
	            session("username")=""
    	        response.redirect "logout.asp" 
		
		    end if
		
		END IF
	    
	    'WE FOUND A GROWER!
        session("login")=true 
        session("username")=u 
        session("growerid")=rs.Fields("GrowerID")
        session("growername")=rs.fields("Growername")
        
		session("accessid") = rs.Fields("AccessID") 
    	IF NOT rs.Fields("TermsAgreed") THEN
			session("termsagreed") = 0
			response.redirect("agreetoterms.asp")
		ELSE
			session("termsagreed") = 1
			response.redirect ("enterspraydata.asp") 
		END IF
		EndConnect(conn)
		set rs = nothing
		
    else 
 
        session("login")=true 
        session("username")=u 
        session("packerid")=rs.Fields("PackerID")
        session("fullrights")=rs.Fields("fullrights")
        if session("fullrights") and not isnull(rs.fields("logofileext")) then _
            session("logofile")=rs.Fields("packernumber")&rs.fields("logofileext")

		session("accessid") = rs.Fields("AccessID") 
    	IF NOT rs.Fields("TermsAgreed") THEN
        	session("termsagreed") = 0
			response.redirect("agreetoterms.asp")
		ELSE
			session("termsagreed") = 1
			response.redirect ("SPRAYRECORDS_LIST.asp") 
		END IF
		EndConnect(conn)
		set rs = nothing
		
    end if 
    
%> 
