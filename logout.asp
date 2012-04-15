<%
  session("login")=false 
  session("accessid") = 0
  session.Abandon()
  
  set session("SprayArray") = Nothing
  
  response.redirect ("default.asp") 
  
 %>
 