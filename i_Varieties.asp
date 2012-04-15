
<%
' FILE: i_Variety.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetVarietyByVarietyID ******************************* 
' ******************************************************* 
function GetVarietyByVarietyID(VarietyID)
	sql = "SELECT * FROM Varieties WHERE VarietyID in (" & VarietyID & ")"
	set GetVarietyByVarietyID = conn.execute(sql)
end function



' ******************************************************* 
' ************GetAllVarieties ******************************* 
' ***************************************************** 
function GetAllVarieties()
	sql = "SELECT * FROM Varieties ORDER BY SortOrder"
	set GetAllVarieties= conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllVarietiesByCropID ******************************* 
' ***************************************************** 
function GetAllVarietiesByCropID(CropID)
	sql = "SELECT * FROM Varieties WHERE CropID in (" & CropID & ")  ORDER BY SortOrder"
	set GetAllVarietiesByCropID= conn.execute(sql)
end function

' ******************************************************* 
' ************GetActiveVarieties ******************************* 
' ***************************************************** 
function GetActiveVarieties()
	sql = "SELECT * FROM Varieties WHERE Active = 1 ORDER BY SortOrder"
	set GetActiveVarieties= conn.execute(sql)
end function

' ******************************************************* 
' ************GetActiveVarietiesByCropID ******************************* 
' ******************************************************* 
function GetActiveVarietiesByCropID(CropID)
	sql = "SELECT * FROM Varieties WHERE Active = 1 AND CropID in (" & CropID & ")  ORDER BY SortOrder"
	set GetActiveVarietiesByCropID = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteVarieties******************************* 
' ******************************************************* 
function DeleteVariety(VarietyID)
	sql = "DELETE FROM Varieties WHERE VarietyID = " & VarietyID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertVarieties******************************* 
' ******************************************************* 
function InsertVariety(Variety,CropID)
	DIM newSortOrder
	sql = "SELECT count(VarietyID) as cnt FROM Varieties WHERE CropID=" & CropID
	set rs = conn.execute(sql)
	
	if rs(0) > 0 then
		sql = "SELECT MAX(SortOrder) AS maxSort FROM Varieties WHERE CropID=" & CropID
		set rs = conn.execute(sql)
		newSortOrder = rs(0) +1								
	else
		newSortOrder = 1
	end if
	sql = "INSERT INTO Varieties(Variety,CropID,SortOrder) VALUES ("
sql = sql & "'" & EscapeQuotes(Variety) & "'"
sql = sql & "," & CropID 
sql = sql & "," & newSortOrder 
sql = sql & ")"
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(VarietyID) AS insertid FROM Varieties"
	set rs = conn.execute(sql)
	newID = rs(0)
	InsertVariety = newID
end Function	

' ******************************************************* 
' ************UpdateVarieties******************************* 
' ******************************************************* 
function UpdateVariety(VarietyID,Variety,CropID)
	sql = "UPDATE Varieties SET Variety ='" & Variety & "',CropID=" & CropID & " WHERE VarietyID = " & VarietyID
	conn.execute sql, , 129
	UpdateVariety= VarietyID
end Function
' ******************************************************* 
' ************ActivateVarieties******************************* 
' ******************************************************* 
function ActivateVariety(VarietyID)
	sql = "Update Varieties SET Active = 1 WHERE VarietyID = " & VarietyID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateVarieties******************************* 
' ******************************************************* 
function DeActivateVariety(VarietyID)
	sql = "Update Varieties SET Active = 0 WHERE VarietyID = " & VarietyID
	conn.execute sql, , 129
end function
%>