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
	
 	sql = "SELECT * FROM Administrators WHERE username = '" & u & "' AND Password = '" & p & "'"
	set rs = conn.execute(sql)

    '--------------------------------------------------------- 
    '-- check for a match, this could be against a database!-- 
    '--------------------------------------------------------- 
    IF rs.eof THEN
		'try for a grower
'		IF (not isnumeric(u)) THEN
'			u = 0
'		END IF
	 	sql = "SELECT * FROM Growers WHERE GrowerNumber = '" & u & "' AND GrowerPassword = '" & p & "'"
		set rs = conn.execute(sql)
	    IF rs.eof THEN
	        'access denied 
	        session("login")=false 
	        session("username")=""
			EndConnect(conn)
			set rs = nothing
	        response.redirect ("loginForm.asp?er=1") 
		ELSE
			'WE FOUND A GROWER!
		END IF
        session("login")=true 
        session("username")=u 
        session("growerid")=rs.Fields("GrowerID")
        session("AdditionalNumbers")=rs.Fields("AdditionalGrowerNumbers")
		session("accessid") = 3
    	IF NOT rs.Fields("TermsAgreed") THEN
			session("termsagreed") = 0
			response.redirect("agreetoterms.asp")
		ELSE
			session("termsagreed") = 1
			response.redirect ("index.asp") 
		END IF
		EndConnect(conn)
		set rs = nothing
		
    else 
 
        ' let them in! 
		session("termsagreed") = 1
        session("login")=true 
        session("username")=u 
		session("accessid") = rs.Fields("AccessID") 
        session("growerid")=0
		EndConnect(conn)
		set rs = nothing
    	    
		response.redirect ("index.asp") 
    end if 
%> 
