
<%
' FILE: i_packerusers.asp 
' CREATED by www.LocusInteractive.net on 08/02/2005

' ******************************************************* 
' ************GetpackerusersByID ******************************* 
' ******************************************************* 
function GetpackerusersByID(packeruserID)
	sql = "SELECT * FROM packerusers WHERE packeruserID = " & packeruserID
	set GetpackerusersByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllpackeruserss ******************************* 
' ******************************************************* 
function GetAllpackerusers()
	sql = "SELECT * FROM packerusers pu inner join packers p on pu.packerid = p.packerid"
	if session("packerid")>0 then
	sql = sql & " WHERE p.packerid = " & session("packerid")
	end if 
	set GetAllpackerusers = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllPackers ******************************* 
' ******************************************************* 
function GetAllPackers()
	sql = "SELECT * FROM Packers WHERE (1=1) "
	IF session("packerid") <> 0 THEN
		sql = sql
	ELSEIF session("packerid") <> 0 THEN
		sql = sql & " AND packerid = " & session("packerid")  & ")"
	END IF
	sql = sql & " ORDER BY PackerName"
	set GetAllPackers = conn.execute(sql)
end function

' ******************************************************* 
' ************Deletepackerusers ******************************* 
' ******************************************************* 
function Deletepackerusers(packeruserID)
	sql = "DELETE FROM packerusers WHERE packeruserID = " & packeruserID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************Insertpackerusers ******************************* 
' ******************************************************* 
function Insertpackerusers(Username,Password,PackerID,AccessID)
REM old code that sucks, didn't work on new HRGS install... 1/30/2007 kmiers
'	sql = "INSERT INTO packerusers(Username,Password,AccessID) VALUES ("
'	sql = sql & "'" & EscapeQuotes(Username) & "'"
'	sql = sql & ",'" & EscapeQuotes(Password) & "'," & AccessID 
'	sql = sql & ")"
'response.write sql
'	conn.execute sql, , 129
REM

if true then

	sql = "INSERT INTO packerusers(Username,Password,PackerID,AccessID) VALUES ("
	sql = sql & "'" & EscapeQuotes(Username) & "'"
	sql = sql & ",'" & EscapeQuotes(Password) & "'," & PackerID & ", " & AccessID 
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
	    oRS.Open "SELECT * FROM packerusers", conn, 1, 3

		oRS.AddNew

		oRS.Fields("Username") = Username
		oRS.Fields("Password") = Password
		oRS.Fields("PackerID") = PackerID
		oRS.Fields("AccessID") = AccessID

		oRS.Update
		newID = oRS.Fields("packeruserID")

		'  Destruction
		'  ***********
		oRS.Close
		Set oRS = Nothing 

end if

REM
'the existing code which errors out.
'	sql = "SELECT MAX(packeruserID) AS insertid FROM packerusers"
'	set rs = conn.execute(sql)
'	newID = rs(0)

	Insertpackerusers = newID
end Function

' ******************************************************* 
' ************Updatepackerusers ******************************* 
' ******************************************************* 
function X_Updatepackerusers(packeruserID,Username,Password,AccessID)
	sql = "UPDATE packerusers SET Username = '" & Username & "',Password ='" & Password & "',AccessID = " & AccessID & " WHERE packeruserID = " & packeruserID
response.write sql
	conn.execute sql, , 129
	Updatepackerusers = packeruserID
end Function

function Updatepackerusers(ID,Username,Password,PackerID,AccessID)
		dim oRS

		Set oRS = Server.CreateObject("ADODB.recordset")

	    oRS.ActiveConnection = conn
	    oRS.CursorLocation = 2 'adUseServer
	    '1 = adOpenKeyset
	    'Const adLockOptimistic = 3
	    oRS.Open "SELECT * FROM packerusers WHERE packeruserID = " & ID, conn, 1, 3

		oRS.Fields("Username") = Username
		oRS.Fields("Password") = Password
		oRS.Fields("PackerID") = PackerID
		oRS.Fields("AccessID") = AccessID

		oRS.Update

		'  Destruction
		'  ***********
		oRS.Close
		Set oRS = Nothing 

	Updatepackerusers = ID
end Function
%>