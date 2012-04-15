<!--#include file="i_SprayList.asp"-->
<%
' FILE: i_SprayYears.asp 
' CREATED by www.LocusInteractive.net on 08/04/2005

' ******************************************************* 
' ************GetSprayYearsByID ******************************* 
' ******************************************************* 
function GetSprayYearsByID(SprayYearID)
	sql = "SELECT * FROM SprayYears WHERE SprayYearID = " & SprayYearID
	set GetSprayYearsByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllSprayYearss ******************************* 
' ******************************************************* 
function GetAllSprayYears()
	sql = "SELECT * FROM SprayYears where sprayyear<=year(getdate()) order by active desc, sprayyear desc"
	set GetAllSprayYears = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteSprayYears ******************************* 
' ******************************************************* 
function DeleteSprayYears(SprayYearID)
	response.write "Permission not granted"
	response.end

	sql = "DELETE FROM SprayYears WHERE SprayYearID = " & SprayYearID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertSprayYears ******************************* 
' ******************************************************* 
function InsertSprayYears(SprayYear)
	sql = "INSERT INTO SprayYears(SprayYear) VALUES (" &_
		"'" & EscapeQuotes(SprayYear) & "'" &_
		")"
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(SprayYearID) AS insertid FROM SprayYears"
	set rs = conn.execute(sql)
	newID = rs(0)
	
sql = "insert into SprayList " & _
	"(SprayYearID, Name, PURS_Name, PURS_EPA_Number, PURS_Report, UnitID, REI, PHI, MAXUseApp, MAXUseSeason, Active, ActiveInd, ReEntryIntervalHours, ReEntryIntervalDays, PreharvestInterval, PreviousSprayListID) " &_
	"SELECT " & newID & ", SprayList.Name, SprayList.PURS_Name, SprayList.PURS_EPA_Number, SprayList.PURS_Report, SprayList.UnitID, SprayList.REI, SprayList.PHI, SprayList.MAXUseApp, SprayList.MAXUseSeason, SprayList.Active, SprayList.ActiveInd, SprayList.ReEntryIntervalHours, SprayList.ReEntryIntervalDays, SprayList.PreharvestInterval, SprayList.SprayListID " &_
	"FROM SprayList INNER JOIN SprayYears ON SprayList.SprayYearID = SprayYears.SprayYearID " &_
	"where SprayYears.Active = true and SprayList.Active=true;"

conn.execute(sql)

REM now the newer CropsToSprays table which was not previously handled.
sql = "insert into CropsToSprays (CropID, SprayID) " &_
	"SELECT  CTS.CropID, SprayList_1.SprayListID " &_
	"FROM SprayList AS SprayList_1 INNER JOIN (CropsToSprays AS CTS INNER JOIN (SprayList INNER JOIN SprayYears ON SprayList.SprayYearID = SprayYears.SprayYearID) ON CTS.SprayID = SprayList.SprayListID) ON SprayList_1.PreviousSprayListID = SprayList.SprayListID " &_
	"WHERE (((SprayYears.Active)=True) AND ((SprayList.Active)=True));"

conn.execute(sql)

'	DIM rsSL,thisName,thisUnitID,thisREI,thisPHI,thisMAXUseApp,thisMAXUseSeason,slID,thisActiveInd,thisReEntryIntervalDays,thisReEntryIntervalHours,thisPreharvestInterval
'	set rsSL = Server.CreateObject("ADODB.RecordSet")
'	set rsSL = GetActiveSprayList()

'	IF not rsSL.EOF THEN
'		DO WHILE not rsSL.eof 
'			thisName = rsSL.Fields("Name")
'			thisUnitID = rsSL.Fields("UnitID")
'			thisREI = rsSL.Fields("REI")
'			thisPHI=rsSL.Fields("PHI")
'			thisMAXUseApp=rsSL.Fields("MAXUseApp")
'			thisMAXUseSeason=rsSL.Fields("MAXUseSeason")
'			thisActiveInd=rsSL.Fields("ActiveInd")
'			thisReEntryIntervalDays=rsSL.Fields("ReEntryIntervalDays")
'			thisReEntryIntervalHours=rsSL.Fields("ReEntryIntervalHours")
'			thisPreharvestInterval=rsSL.Fields("PreharvestInterval")
'			slID = InsertSprayListWithYear(newID,thisName,thisUnitID,thisREI,thisPHI,thisMAXUseApp,thisMAXUseSeason,thisActiveInd,thisReEntryIntervalDays,thisReEntryIntervalHours,thisPreharvestInterval)
'
'			rsSL.MoveNext
'		LOOP
'	END IF
	
	InsertSprayYears = newID
end Function
' ******************************************************* 
' ************UpdateSprayYears ******************************* 
' ******************************************************* 
function UpdateSprayYears(SprayYearID,SprayYear)
	sql = "UPDATE SprayYears SET SprayYear ='" & SprayYear & "' WHERE SprayYearID = " & SprayYearID
	conn.execute sql, , 129
	UpdateSprayYears = SprayYearID
end Function

' ******************************************************* 
' ************ActivateSprayList ******************************* 
' ******************************************************* 
function ActivateSprayYears(ID)
	sql = "Update SprayYears SET Active = 0 "
	conn.execute sql, , 129
	sql = "Update SprayYears SET Active = 1 WHERE SprayYearID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateMethod ******************************* 
' ******************************************************* 
function DeActivateSprayYears(ID)
	sql = "Update SprayYears SET Active = 0 WHERE SprayYearID = " & ID
	conn.execute sql, , 129
end function
%>