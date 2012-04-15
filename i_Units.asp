<%
' FILE: i_Units.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetUnitsByID ******************************* 
' ******************************************************* 
function GetUnitsByID(ID)
	sql = "SELECT * FROM Units WHERE UnitID in (" & ID & ")"
	set GetUnitsByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllUnitss ******************************* 
' ******************************************************* 
function GetAllUnits()
	sql = "SELECT * FROM Units ORDER BY SortOrder"
	set GetAllUnits = conn.execute(sql)
end function

' ******************************************************* 
' ************GetActiveUnits ******************************* 
' ******************************************************* 
function GetActiveUnits()
	sql = "SELECT * FROM Units WHERE Active = true ORDER BY SortOrder"
	set GetActiveUnits = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteUnits ******************************* 
' ******************************************************* 
function DeleteUnits(ID)
	sql = "DELETE FROM Units WHERE UnitID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertUnits ******************************* 
' ******************************************************* 
function InsertUnits(Unit, PURSUnit)
	DIM newSortOrder
	sql = "SELECT MAX(SortOrder) AS maxSort FROM Units"
	set rs = conn.execute(sql)
	if IsNull(rs(0)) then
		newSortOrder = 1
	else
		newSortOrder = rs(0) + 1
	end if

	sql = "INSERT INTO Units(Unit,PURSUnit,SortOrder) VALUES (" & _
		"'" & EscapeQuotes(Unit) & "','" & EscapeQuotes(PURSUnit) & "'," & newSortOrder & ")"
	response.write sql
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(UnitID) AS insertid FROM Units"
	set rs = conn.execute(sql)
	newID = rs(0)
	InsertUnits = newID
end Function

' ******************************************************* 
' ************UpdateUnits ******************************* 
' ******************************************************* 
function UpdateUnits(ID, Unit, PURSUnit)
	sql = "UPDATE Units SET Unit ='" & EscapeQuotes(Unit) & "', PURSUnit = '" & _
		EscapeQuotes(PURSUnit) & "' WHERE UnitID = " & ID
	conn.execute sql, , 129
	UpdateUnits = ID
end Function

' ******************************************************* 
' ************ActivateUnits ******************************* 
' ******************************************************* 
function ActivateUnits(ID)
	sql = "Update Units SET Active = 1 WHERE UnitID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateUnits ******************************* 
' ******************************************************* 
function DeActivateUnits(ID)
	sql = "Update Units SET Active = 0 WHERE UnitID = " & ID
	conn.execute sql, , 129
end function
%>