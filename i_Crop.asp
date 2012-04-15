
<%
' FILE: i_Crop.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetCropByCropID ******************************* 
' ******************************************************* 
function GetCropByCropID(CropID)
	sql = "SELECT * FROM Crops WHERE CropID in (" & CropID & ")"
	set GetCropByCropID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllCrops ******************************* 
' ***************************************************** 
function GetAllCrops()
	sql = "SELECT * FROM Crops ORDER BY SortOrder"
	set GetAllCrops= conn.execute(sql)
end function

' ******************************************************* 
' ************GetActiveCrops ******************************* 
' ***************************************************** 
function GetActiveCrops()
if session("packerid")<>"" then
    	sql = "SELECT * FROM Crops WHERE Active = 1 ORDER BY SortOrder"
elseif session("growerid")<>"" then
   	sql = "SELECT * FROM Crops c INNER JOIN growercrop gc on c.cropid=gc.cropid WHERE Active = 1 AND growerid = " & session("growerid") & " ORDER BY SortOrder"
else
    	sql = "SELECT * FROM Crops WHERE Active = 1 ORDER BY SortOrder"
end if
	set GetActiveCrops= conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteCrops******************************* 
' ******************************************************* 
function DeleteCrop(CropID)
	sql = "DELETE FROM Crops WHERE CropID = " & CropID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertCrops******************************* 
' ******************************************************* 
function InsertCrop(ByVal Crop, ByVal PURS_SiteCategory, ByVal PURS_SpecificSite)
	DIM newSortOrder
	sql = "SELECT MAX(SortOrder) AS maxSort FROM Crops"
	set rs = conn.execute(sql)
	newSortOrder = rs(0) + 1
	sql = "INSERT INTO Crops(Crop,PURS_SiteCategory,PURS_SpecificSite,SortOrder) VALUES (" & _
			"'" & EscapeQuotes(Crop) & "'," & _
			"'" & EscapeQuotes(PURS_SiteCategory) & "'," & _
			"'" & EscapeQuotes(PURS_SpecificSite) & "'," & _
			newSortOrder & ")"
	response.write sql
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(CropID) AS insertid FROM Crops"
	set rs = conn.execute(sql)
	newID = rs(0)
	InsertCrop = newID
end Function	

' ******************************************************* 
' ************UpdateCrops******************************* 
' ******************************************************* 
function UpdateCrop(ByVal CropID, ByVal Crop, ByVal PURS_SiteCategory, ByVal PURS_SpecificSite)
	sql = "UPDATE Crops SET Crop='" & EscapeQuotes(Crop) & _
			"', PURS_SiteCategory = '" & EscapeQuotes(PURS_SiteCategory) & _
			"', PURS_SpecificSite = '" & EscapeQuotes(PURS_SpecificSite) & _
			"' WHERE CropID = " & CropID
	conn.execute sql, , 129
	UpdateCrop= CropID
end Function

' ******************************************************* 
' ************ActivateCrops******************************* 
' ******************************************************* 
function ActivateCrop(CropID)
	sql = "Update Crops SET Active = 1 WHERE CropID = " & CropID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateCrops******************************* 
' ******************************************************* 
function DeActivateCrop(CropID)
	sql = "Update Crops SET Active = 0 WHERE CropID = " & CropID
	conn.execute sql, , 129
end function
%>