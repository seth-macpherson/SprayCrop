
<%
' FILE: i_SprayList.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005
' MODIFIED kalanmiers 12/4/2006
'	added PURS fields, general cleanup!

' ******************************************************* 
' ************GetSprayListByID ******************************* 
' ******************************************************* 
function GetSprayListByID(ID)
	sql = "SELECT SprayList.*, Units.Unit FROM SprayList LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE SprayListID in (" & ID & ")"
	set GetSprayListByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetCropsBySprayListID ******************************* 
' ******************************************************* 
function GetCropsBySprayListID(ID)
	sql = "SELECT Crops.* FROM Crops LEFT JOIN CropsToSprays ON Crops.CropID = CropsToSprays.CropID WHERE  CropsToSprays.SprayID = " & ID 
	set GetCropsBySprayListID = conn.execute(sql)
end function


' ******************************************************* 
' ************ GetAllSprayLists ************************* 
' ******************************************************* 
function GetAllSprayList()

	'# added 3-Apr-2011
	dim yr
	if request.form("SprayYear")<>"" then
		yr=request.form("SprayYear")
	else
		sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
		set rs = conn.execute(sql)
		yr = rs.collect(0)
	end if

	sql = "SELECT SprayList.*, Units.Unit FROM SprayList LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE SprayYearID = " & yr & " ORDER BY Name"
	
	'# debug
	'response.write sql
	'response.end
	
	set GetAllSprayList = conn.execute(sql)
end function

' ******************************************************* 
' ************ GetAllReportableSprayLists *************** 
' ******************************************************* 
function GetAllReportableSprayList()
	sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
	set rs = conn.execute(sql)

	sql = "SELECT SprayList.*, Units.Unit FROM SprayList LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE PURS_Report = 1 AND SprayYearID = " & rs(0) & " ORDER BY Name"
'response.write sql
	set GetAllReportableSprayList = conn.execute(sql)
end function

' ******************************************************* 
' ************ GetAllNONReportableSprayLists ************ 
' ******************************************************* 
function GetAllNONReportableSprayList()
	sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
	set rs = conn.execute(sql)

	sql = "SELECT SprayList.*, Units.Unit FROM SprayList LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE PURS_Report = 0 AND SprayYearID = " & rs(0) & " ORDER BY Name"
'response.write sql
	set GetAllNONReportableSprayList = conn.execute(sql)
end function


' ******************************************************* 
' ************GetActiveSprayList ******************************* 
' ******************************************************* 
function GetActiveSprayList()
	sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
	set rs = conn.execute(sql)
	
	sql = "SELECT SprayList.*, Units.Unit FROM SprayList LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE SprayList.Active= 1 AND SprayYearID = " & rs(0) & " ORDER BY Name"
	set GetActiveSprayList = conn.execute(sql)
end function


' ******************************************************* 
' ************GetActiveSprayListByCropID ******************************* 
' ******************************************************* 
function GetActiveSprayListByCropID(CropID)
	sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1 "
	set rs = conn.execute(sql)
	
	sql = "SELECT SprayList.*, Units.Unit FROM SprayList LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE SprayList.Active= 1 AND SprayYearID = " & rs(0) & " and SprayListID in (SELECT SprayID from CropsToSprays WHERE CropID = " & CropID & ")  ORDER BY Name"
'response.write sql
	set GetActiveSprayListByCropID = conn.execute(sql)
end function


' ******************************************************* 
' ************GetActiveSprayListByCropYear ******************************* 
' ******************************************************* 
function GetActiveSprayListByCropYear(CropID,YearID)
	dim yr
	yr=yearid
	if yr="" then
		sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
		set rs = conn.execute(sql)
		yr = rs.collect(0)
	end if

	sql = "SELECT SprayList.*, Units.Unit FROM SprayList LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE SprayList.Active= 1 AND SprayYearID = " & yr & " and SprayListID in (SELECT SprayID from CropsToSprays WHERE CropID = " & CropID & ")  ORDER BY Name"
	set GetActiveSprayListByCropYear = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteSprayList ******************************* 
' ******************************************************* 
function DeleteSprayList(ID)
	sql = "DELETE FROM SprayList WHERE SprayListID = " & ID
	conn.execute sql, , 129
end function


' ******************************************************* 
' ************InsertCropToSpray ******************************* 
' ******************************************************* 
function InsertCropToSpray(SprayID,CropID)
	sql = "INSERT INTO CropsToSprays (CropID,SprayID) VALUES (" & CropID & "," & SprayID & ")"
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeleteCropToSprayBySpray ******************************* 
' ******************************************************* 
function DeleteCropToSprayBySpray(SprayID)
	sql = "DELETE FROM CropsToSprays WHERE SprayID = " &  SprayID
	conn.execute sql, , 129
end function


' ******************************************************* 
' ************InsertSprayList ******************************* 
' ******************************************************* 
'function InsertSprayList(Name,UnitID,REI,PHI,MAXUseApp,MAXUseSeason,ActiveInd,ReEntryIntervalHours,ReEntryIntervalDays,PreharvestInterval)
function InsertSprayList(Name, PURS_Name, PURS_EPA_Number, PURS_Report, _
	UnitID, REI, PHI, MAXUseApp, MAXUseSeason, ActiveInd, _
	ReEntryIntervalHours, ReEntryIntervalDays, PreharvestInterval)

	if MaxUseApp = "" then
		MaxUseApp = "NULL"
	end if
	if MAXUseSeason = "" then
		MAXUseSeason = "NULL"
	end if
	if not IsNumeric(MaxUseApp) then
		MaxUseApp = "NULL"
	end if
	if not IsNumeric(MAXUseSeason) then
		MAXUseSeason = "NULL"
	end if
	if not IsNumeric(ReEntryIntervalDays) then
		ReEntryIntervalDays = "NULL"
	end if
	if not IsNumeric(ReEntryIntervalHours) then
		ReEntryIntervalHours = "NULL"
	end if
	if not IsNumeric(PreharvestInterval) then
		PreharvestInterval = "NULL"
	end if
	if UnitID = "" then
		UnitID = "NULL"
	end if
	if not IsNumeric(UnitID) then
		UnitID = "NULL"
	end if

	if PURS_Report = "" or PURS_Report <> "1" then
		PURS_Report = 0 'true
	end if
	'if VarType(PURS_Report) <> vbBoolean then
	'	PURS_Report = 0 'false
	'end if

	DIM thisSprayYearID
	sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
	set rs = conn.execute(sql)
	thisSprayYearID = rs(0)
	
	sql = "INSERT INTO SprayList(Name, PURS_Name, PURS_EPA_Number, PURS_Report, UnitID, REI, PHI, " & _
		"MAXUseApp, MAXUseSeason, SprayYearID, ActiveInd, " & _
		"ReEntryIntervalDays, ReEntryIntervalHours, PreharvestInterval) VALUES (" & _
		"'" & EscapeQuotes(Name) & "'" & _
		",'" & EscapeQuotes(PURS_Name) & "'" & _
		",'" & EscapeQuotes(PURS_EPA_Number) & "'" & _
		"," & EscapeQuotes(PURS_Report) & _
		"," & EscapeQuotes(UnitID) & _
		",'" & EscapeQuotes(REI) & "'" & _
		",'" & EscapeQuotes(PHI) & "'" & _
		"," & EscapeQuotes(MAXUseApp) & _
		"," & EscapeQuotes(MAXUseSeason) & "," & thisSprayYearID & _
		",'" & EscapeQuotes(ActiveInd) & "'" & _
		"," & EscapeQuotes(ReEntryIntervalDays) & _
		"," & EscapeQuotes(ReEntryIntervalHours) & _
		"," & EscapeQuotes(PreharvestInterval) & _
		")"
	'response.Write sql
	'response.End	
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(SprayListID) AS insertid FROM SprayList"
	set rs = conn.execute(sql)
	newID = rs(0)
	InsertSprayList = newID
end Function


' ******************************************************* 
' ************ InsertSprayListWithYear ******************
' ******************************************************* 
function InsertSprayListWithYear(InsertYearID,Name,UnitID,REI,PHI,MAXUseApp,MAXUseSeason,ActiveInd,ReEntryIntervalHours,ReEntryIntervalDays,PreharvestInterval)
	if MaxUseApp = "" then
		MaxUseApp = "NULL"
	end if
	if MAXUseSeason = "" then
		MAXUseSeason = "NULL"
	end if
	if not IsNumeric(MaxUseApp) then
		MaxUseApp = "NULL"
	end if
	if not IsNumeric(MAXUseSeason) then
		MAXUseSeason ="NULL"
	end if
	if not IsNumeric(ReEntryIntervalDays) then
		ReEntryIntervalDays = "NULL"
	end if
	if not IsNumeric(ReEntryIntervalHours) then
		ReEntryIntervalHours = "NULL"
	end if
	if not IsNumeric(PreharvestInterval) then
		PreharvestInterval = "NULL"
	end if
	if UnitID = "" then
		UnitID = "NULL"
	end if
	if not IsNumeric(UnitID) then
		UnitID = "NULL"
	end if
	
	sql = "INSERT INTO SprayList(Name, UnitID, REI, PHI, MAXUseApp, MAXUseSeason, SprayYearID," & _
		"ActiveInd, ReEntryIntervalDays, ReEntryIntervalHours, PreharvestInterval) VALUES (" & _
		"'" & EscapeQuotes(Name) & "'" & _
		"," & EscapeQuotes(UnitID) & _
		",'" & EscapeQuotes(REI) & "'" & _
		",'" & EscapeQuotes(PHI) & "'" & _
		"," & EscapeQuotes(MAXUseApp) & _
		"," & EscapeQuotes(MAXUseSeason) & "," & InsertYearID & _
		",'" & EscapeQuotes(ActiveInd) & "'" & _
		"," & ReEntryIntervalDays & _
		"," & ReEntryIntervalHours & _
		"," & EscapeQuotes(PreharvestInterval) & _
		")"
'response.write sql
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(SprayListID) AS insertid FROM SprayList"
	set rs = conn.execute(sql)
	newID = rs(0)
	response.write newID
	InsertSprayListWithYear = newID
end Function


' ******************************************************* 
' ************ UpdateSprayList ************************** 
' ******************************************************* 
'function UpdateSprayList(ID,Name,UnitID,REI,PHI,MAXUseApp,MAXUseSeason,ActiveInd,ReEntryIntervalHours,ReEntryIntervalDays,PreharvestInterval)
function UpdateSprayList(ID, Name, PURS_Name, PURS_EPA_Number, PURS_Report, UnitID, REI, PHI, _
	MAXUseApp, MAXUseSeason, ActiveInd, _
	ReEntryIntervalHours, ReEntryIntervalDays, PreharvestInterval)

	if MaxUseApp = "" then
		MaxUseApp = "NULL"
	end if
	if MAXUseSeason = "" then
		MAXUseSeason = "NULL"
	end if
	if not IsNumeric(MaxUseApp) then
		MaxUseApp = "NULL"
	end if
	if (not IsNumeric(ReEntryIntervalDays)) or ReEntryIntervalDays = "" then
		ReEntryIntervalDays = "NULL"
	end if
	if (not IsNumeric(ReEntryIntervalHours)) or ReEntryIntervalHours = "" then
		ReEntryIntervalHours = "NULL"
	end if
	if (not IsNumeric(PreharvestInterval)) or PreharvestInterval = "" then
		PreharvestInterval = "NULL"
	end if
	if not IsNumeric(MAXUseSeason) then
		MAXUseSeason = "NULL"
	end if
	if UnitID = "" then
		UnitID = "NULL"
	end if
	if not IsNumeric(UnitID) then
		UnitID = "NULL"
	end if

	'vbBoolean = 11
	if PURS_Report = "" or PURS_Report <> "1" then
		PURS_Report = 0
	end if
	'if VarType(PURS_Report) <> vb then
	'	PURS_Report = false
	'end if

'	sql = "UPDATE SprayList SET Name ='" & Name & "', ActiveInd ='" & ActiveInd & "',UnitID ='" & UnitID & "',REI ='" & REI & "',PHI ='" & PHI & "',MAXUseApp =" & MAXUseApp & ",MAXUseSeason =" & MAXUseSeason &  ",ReEntryIntervalDays =" & ReEntryIntervalDays & ",ReEntryIntervalHours =" & ReEntryIntervalHours & ",PreharvestInterval =" & PreharvestInterval & " WHERE SprayListID = " & ID

	sql = "UPDATE SprayList SET Name ='" & Name & _
		"', PURS_Name ='" & PURS_Name & _
		"', PURS_EPA_Number ='" & PURS_EPA_Number & _
		"', PURS_Report = " & PURS_Report & _
		", ActiveInd ='" & ActiveInd & "',UnitID ='" & UnitID & _
		"',REI ='" & REI & "',PHI ='" & PHI & "',MAXUseApp =" & MAXUseApp & ",MAXUseSeason =" & MAXUseSeason &  ",ReEntryIntervalDays =" & ReEntryIntervalDays & ",ReEntryIntervalHours =" & ReEntryIntervalHours & ",PreharvestInterval =" & PreharvestInterval & " WHERE SprayListID = " & ID
	conn.execute sql, , 129
	UpdateSprayList = ID
end Function


' ******************************************************* 
' ************ ActivateSprayList ************************ 
' ******************************************************* 
function ActivateSprayList(ID)
	sql = "Update SprayList SET Active = 1 WHERE SprayListID = " & ID
	conn.execute sql, , 129
end function


' ******************************************************* 
' ************ DeActivateSprayList ********************** 
' ******************************************************* 
function DeActivateSprayList(ID)
	sql = "Update SprayList SET Active = 0 WHERE SprayListID = " & ID
	conn.execute sql, , 129
end function
%>