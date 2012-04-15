
<%
' FILE: i_Stage.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetStageByID ******************************* 
' ******************************************************* 
function GetStageByID(ID)
	sql = "SELECT * FROM Stages WHERE StageID in (" & ID & ")"
	set GetStageByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllStages ******************************* 
' ******************************************************* 
function GetAllStage()
	sql = "SELECT * FROM Stages ORDER BY SortOrder"
	set GetAllStage = conn.execute(sql)
end function
' ******************************************************* 
' ************GetActiveStages ******************************* 
' ******************************************************* 
function GetActiveStages()
	sql = "SELECT * FROM Stages WHERE Active = 1 ORDER BY SortOrder"
	set GetActiveStages = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteStage ******************************* 
' ******************************************************* 
function DeleteStage(ID)
	sql = "DELETE FROM Stages WHERE StageID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertStage ******************************* 
' ******************************************************* 
function InsertStage(Stage)
	DIM newSortOrder
	sql = "SELECT MAX(SortOrder) AS maxSort FROM Stages"
	set rs = conn.execute(sql)
	newSortOrder = rs(0) +1								
	sql = "INSERT INTO Stages(Stage,SortOrder) VALUES ("
sql = sql & "'" & EscapeQuotes(Stage) & "'"
sql = sql & "," & newSortOrder 
sql = sql & ")"
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(StageID) AS insertid FROM Stages"
	set rs = conn.execute(sql)
	newID = rs(0)
	InsertStage = newID
end Function
' ******************************************************* 
' ************UpdateStage ******************************* 
' ******************************************************* 
function UpdateStage(ID,Stage)
	sql = "UPDATE Stages SET Stage ='" & Stage & "' WHERE StageID = " & ID
	conn.execute sql, , 129
	UpdateStage = ID
end Function
' ******************************************************* 
' ************ActivateStage ******************************* 
' ******************************************************* 
function ActivateStage(ID)
	sql = "Update Stages SET Active = 1 WHERE StageID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateStage ******************************* 
' ******************************************************* 
function DeActivateStage(ID)
	sql = "Update Stages SET Active = 0 WHERE StageID = " & ID
	conn.execute sql, , 129
end function
%>






