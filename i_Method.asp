
<%
' FILE: i_Method.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetMethodByID ******************************* 
' ******************************************************* 
function GetMethodByID(ID)
	sql = "SELECT * FROM Methods WHERE MethodID in (" & ID & ")"
	set GetMethodByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllMethods ******************************* 
' ******************************************************* 
function GetAllMethod()
	sql = "SELECT * FROM Methods ORDER BY SortOrder"
	set GetAllMethod = conn.execute(sql)
end function

' ******************************************************* 
' ************GetActiveMethods ******************************* 
' ******************************************************* 
function GetActiveMethods()
	sql = "SELECT * FROM Methods WHERE Active= 1 ORDER BY SortOrder"
	set GetActiveMethods = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteMethod ******************************* 
' ******************************************************* 
function DeleteMethod(ID)
	sql = "DELETE FROM Methods WHERE MethodID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertMethod ******************************* 
' ******************************************************* 
function InsertMethod(Method)
Dim newSortOrder
sql = "SELECT MAX(SortOrder) AS maxSort FROM Methods"
set rs = conn.execute(sql)
newSortOrder = rs(0) + 1
sql = "INSERT INTO Methods(Method,SortOrder) VALUES ("
sql = sql & "'" & EscapeQuotes(Method) & "'," & newSortOrder
sql = sql & ")"
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(MethodID) AS insertid FROM Methods"
	set rs = conn.execute(sql)
	newID = rs(0)
	InsertMethod = newID
end Function
' ******************************************************* 
' ************UpdateMethod ******************************* 
' ******************************************************* 
function UpdateMethod(ID,Method)
	sql = "UPDATE Methods SET Method ='" & Method & "' WHERE MethodID = " & ID
	conn.execute sql, , 129
	UpdateMethod = ID
end Function
' ******************************************************* 
' ************ActivateMethod ******************************* 
' ******************************************************* 
function ActivateMethod(ID)
	sql = "Update Methods SET Active = 1 WHERE MethodID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateMethod ******************************* 
' ******************************************************* 
function DeActivateMethod(ID)
	sql = "Update Methods SET Active = 0 WHERE MethodID = " & ID
	conn.execute sql, , 129
end function
%>












