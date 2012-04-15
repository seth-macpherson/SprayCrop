
<%
' FILE: i_growerusers.asp 
' CREATED by www.LocusInteractive.net on 08/02/2005

' ******************************************************* 
' ************GetgrowerusersByID ******************************* 
' ******************************************************* 
function GetgrowerusersByID(groweruserID)
	sql = "SELECT * FROM growerusers WHERE groweruserID = " & groweruserID
	set GetgrowerusersByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllGrowerusers ******************************* 
' ******************************************************* 
function GetAllgrowerusers()
	sql = "SELECT gu.*,g.growername FROM growerusers gu inner join growers g on gu.growerid = g.growerid"
	set GetAllgrowerusers = conn.execute(sql)
end function

' ******************************************************* 
' ************GetPackerGrowerUsers ******************************* 
' ******************************************************* 
function GetPackerGrowerUsers()
	sql = "SELECT gu.*,g.growername FROM growerusers gu inner join growers g on gu.growerid = g.growerid"
	sql = sql & " inner join GrowerPacker gp on g.growerid = gp.growerid WHERE packerid = " & session("packerid")
	set GetPackerGrowerUsers = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllgrowers ******************************* 
' ******************************************************* 
function GetAllgrowers()
	sql = "SELECT * FROM growers WHERE (1=1) "
	IF session("growerid") <> 0 THEN
		sql = sql
	ELSEIF session("growerid") <> 0 THEN
		sql = sql & " AND growerid = " & session("growerid")  & ")"
	END IF
	sql = sql & " ORDER BY growername"
	set GetAllgrowers = conn.execute(sql)
end function

' ******************************************************* 
' ************Deletegrowerusers ******************************* 
' ******************************************************* 
function Deletegrowerusers(groweruserID)
	sql = "DELETE FROM growerusers WHERE groweruserID = " & groweruserID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************Insertgrowerusers ******************************* 
' ******************************************************* 
function Insertgrowerusers(Username,Password,growerID,AccessID)
REM old code that sucks, didn't work on new HRGS install... 1/30/2007 kmiers
'	sql = "INSERT INTO growerusers(Username,Password,AccessID) VALUES ("
'	sql = sql & "'" & EscapeQuotes(Username) & "'"
'	sql = sql & ",'" & EscapeQuotes(Password) & "'," & AccessID 
'	sql = sql & ")"
'response.write sql
'	conn.execute sql, , 129
REM

if true then

	sql = "INSERT INTO growerusers(Username,Password,growerID,AccessID) VALUES ("
	sql = sql & "'" & EscapeQuotes(Username) & "'"
	sql = sql & ",'" & EscapeQuotes(Password) & "'," & growerID & ", " & AccessID 
	sql = sql & ")"
    'response.write sql
	conn.execute sql, , 129

else
		dim oRS, newID

		Set oRS = Server.CreateObject("ADODB.recordset")

	    oRS.ActiveConnection = conn
	    oRS.CursorLocation = 2 'adUseServer
	    '1 = adOpenKeyset
	    'Const adLockOptimistic = 3
	    oRS.Open "SELECT * FROM growerusers", conn, 1, 3

		oRS.AddNew

		oRS.Fields("Username") = Username
		oRS.Fields("Password") = Password
		oRS.Fields("growerID") = growerID
		oRS.Fields("AccessID") = AccessID

		oRS.Update
		newID = oRS.Fields("groweruserID")

		'  Destruction
		'  ***********
		oRS.Close
		Set oRS = Nothing 

end if

REM
'the existing code which errors out.
'	sql = "SELECT MAX(groweruserID) AS insertid FROM growerusers"
'	set rs = conn.execute(sql)
'	newID = rs(0)

	Insertgrowerusers = newID
end Function

' ******************************************************* 
' ************Updategrowerusers ******************************* 
' ******************************************************* 
function X_Updategrowerusers(groweruserID,Username,Password,AccessID)
	sql = "UPDATE growerusers SET Username = '" & Username & "',Password ='" & Password & "',AccessID = " & AccessID & " WHERE groweruserID = " & groweruserID
response.write sql
	conn.execute sql, , 129
	Updategrowerusers = groweruserID
end Function

function Updategrowerusers(ID,Username,Password,growerID,AccessID)
		dim oRS

		Set oRS = Server.CreateObject("ADODB.recordset")

	    oRS.ActiveConnection = conn
	    oRS.CursorLocation = 2 'adUseServer
	    '1 = adOpenKeyset
	    'Const adLockOptimistic = 3
	    oRS.Open "SELECT * FROM growerusers WHERE groweruserID = " & ID, conn, 1, 3

		oRS.Fields("Username") = Username
		oRS.Fields("Password") = Password
		oRS.Fields("growerID") = growerID
		oRS.Fields("AccessID") = AccessID

		oRS.Update

		'  Destruction
		'  ***********
		oRS.Close
		Set oRS = Nothing 

	Updategrowerusers = ID
end Function
%>