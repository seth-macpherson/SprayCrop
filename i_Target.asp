<%
' FILE: i_Target.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetTargetByID ******************************* 
' ******************************************************* 
function GetTargetByID(ID)
	sql = "SELECT * FROM Targets WHERE TargetID in (" & ID & ")"
	set GetTargetByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllTargets ******************************* 
' ******************************************************* 
function GetAllTarget()
	sql = "SELECT * FROM Targets ORDER BY SortOrder"
	set GetAllTarget = conn.execute(sql)
end function

' ******************************************************* 
' ************GetActiveTarget ******************************* 
' ******************************************************* 
function GetActiveTarget()
	sql = "SELECT * FROM Targets WHERE Active = 1 ORDER BY SortOrder"
	set GetActiveTarget = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteTarget ******************************* 
' ******************************************************* 
function DeleteTarget(ID)
	sql = "DELETE FROM Targets WHERE TargetID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertTarget ******************************* 
' ******************************************************* 
function InsertTarget(Target, PURSTarget)
	DIM newSortOrder
	sql = "SELECT MAX(SortOrder) AS maxSort FROM Targets"
	set rs = conn.execute(sql)
	newSortOrder = rs(0)+ 1					
	sql = "INSERT INTO Targets(Target,PURS_Target,SortOrder) VALUES ('" & _
		EscapeQuotes(Target) & "','" & PURSTarget & "'," & newSortOrder & ")"
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(TargetID) AS insertid FROM Targets"
	set rs = conn.execute(sql)
	newID = rs(0)
	InsertTarget = newID
end Function

' ******************************************************* 
' ************UpdateTarget ******************************* 
' ******************************************************* 
function UpdateTarget(ID,Target,PURSTarget)
	sql = "UPDATE Targets SET Target ='" & Target & "', PURS_Target = '" & PURSTarget & "' WHERE TargetID = " & ID
	conn.execute sql, , 129
	UpdateTarget = ID
end Function

' ******************************************************* 
' ************ActivateTarget ******************************* 
' ******************************************************* 
function ActivateTarget(ID)
	sql = "Update Targets SET Active = 1 WHERE TargetID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateTarget ******************************* 
' ******************************************************* 
function DeActivateTarget(ID)
	sql = "Update Targets SET Active = 0 WHERE TargetID = " & ID
	conn.execute sql, , 129
end function
%>












