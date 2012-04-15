
<%
' FILE: i_Administrators.asp 
' CREATED by www.LocusInteractive.net on 08/02/2005

' ******************************************************* 
' ************GetAdministratorsByID ******************************* 
' ******************************************************* 
function GetAdministratorsByID(AdministratorID)
	sql = "SELECT * FROM Administrators WHERE AdministratorID = " & AdministratorID
	set GetAdministratorsByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllAdministratorss ******************************* 
' ******************************************************* 
function GetAllAdministrators()
	sql = "SELECT * FROM Administrators"
	set GetAllAdministrators = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteAdministrators ******************************* 
' ******************************************************* 
function DeleteAdministrators(AdministratorID)
	sql = "DELETE FROM Administrators WHERE AdministratorID = " & AdministratorID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertAdministrators ******************************* 
' ******************************************************* 
function InsertAdministrators(Username,Password,AccessID)
REM old code that sucks, didn't work on new HRGS install... 1/30/2007 kmiers
'	sql = "INSERT INTO Administrators(Username,Password,AccessID) VALUES ("
'	sql = sql & "'" & EscapeQuotes(Username) & "'"
'	sql = sql & ",'" & EscapeQuotes(Password) & "'," & AccessID 
'	sql = sql & ")"
'response.write sql
'	conn.execute sql, , 129
REM
		dim oRS, newID

		Set oRS = Server.CreateObject("ADODB.recordset")

	    oRS.ActiveConnection = conn
	    oRS.CursorLocation = 2 'adUseServer
	    '1 = adOpenKeyset
	    'Const adLockOptimistic = 3
	    oRS.Open "SELECT * FROM Administrators", conn, 1, 3

		oRS.AddNew

		oRS.Fields("Username") = Username
		oRS.Fields("Password") = Password
		oRS.Fields("AccessID") = AccessID

		oRS.Update
		newID = oRS.Fields("AdministratorID")

		'  Destruction
		'  ***********
		oRS.Close
		Set oRS = Nothing 

REM
'the existing code which errors out.
'	sql = "SELECT MAX(AdministratorID) AS insertid FROM Administrators"
'	set rs = conn.execute(sql)
'	newID = rs(0)

	InsertAdministrators = newID
end Function

' ******************************************************* 
' ************UpdateAdministrators ******************************* 
' ******************************************************* 
function X_UpdateAdministrators(AdministratorID,Username,Password,AccessID)
	sql = "UPDATE Administrators SET Username = '" & Username & "',Password ='" & Password & "',AccessID = " & AccessID & " WHERE AdministratorID = " & AdministratorID
response.write sql
	conn.execute sql, , 129
	UpdateAdministrators = AdministratorID
end Function

function UpdateAdministrators(ID,Username,Password,AccessID)
		dim oRS

		Set oRS = Server.CreateObject("ADODB.recordset")

	    oRS.ActiveConnection = conn
	    oRS.CursorLocation = 2 'adUseServer
	    '1 = adOpenKeyset
	    'Const adLockOptimistic = 3
	    oRS.Open "SELECT * FROM Administrators WHERE AdministratorID = " & ID, conn, 1, 3

		oRS.Fields("Username") = Username
		oRS.Fields("Password") = Password
		oRS.Fields("AccessID") = AccessID

		oRS.Update

		'  Destruction
		'  ***********
		oRS.Close
		Set oRS = Nothing 

	UpdateAdministrators = ID
end Function
%>