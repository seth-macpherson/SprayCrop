
<%
' FILE: i_GrowerLocations.asp 
' CREATED by Kim Miers on 07/16/2006

' ******************************************************* 
' ************ GetGrowersLocationsByID ******************* 
' ******************************************************* 
function GetGrowersLocationsByID(pID)
	dim l_sql
	l_sql = "SELECT * FROM GrowersLocations WHERE GLoc_GrowerLocationID in (" & pID & ") AND GLoc_Active = 1"
	set GetGrowersLocationsByID = conn.execute(l_sql)
end function

' ******************************************************* 
' ************ GetGrowersLocationsByGrowerID ************ 
' ******************************************************* 
function GetGrowersLocationsByGrowerID(pGrowerID)

	IF pGrowerID <> "" THEN
		sql = "SELECT * FROM GrowersLocations WHERE GLoc_GrowerID in (" & pGrowerID & ") AND GLoc_Active = 1"
	ELSEIF session("packerid") <> 0 THEN
		sql = "SELECT * FROM GrowersLocations WHERE GLoc_GrowerID in (select distinct growerid from sprayrecord where packerid = " & session("packerid")  & ")"
	END IF
	set GetGrowersLocationsByGrowerID = conn.execute(sql)

end function

' ******************************************************* 
' ************ GetALLGrowersLocationsByGrowerID ************ 
' ******************************************************* 
function GetALLGrowersLocationsByGrowerID(pGrowerID)
	dim l_sql
	l_sql = "SELECT * FROM GrowersLocations WHERE GLoc_GrowerID in (" & pGrowerID & ")"
	set GetALLGrowersLocationsByGrowerID = conn.execute(l_sql)
end function

' ******************************************************* 
' ************ GetAllGrowerLocations ******************** 
' ******************************************************* 
function GetAllGrowersLocations()
	sql = "SELECT * FROM GrowersLocations ORDER BY GLoc_Location"
	set GetAllGrowersLocations = conn.execute(sql)
end function

' ******************************************************* 
' ************ GetActiveGrowersLocations ***************** 
' ******************************************************* 
function GetActiveGrowersLocations()
	sql = "SELECT * FROM GrowersLocations WHERE Active = 1 ORDER BY GLoc_Location"
	set GetActiveGrowersLocations = conn.execute(sql)
end function

' ******************************************************* 
' ************ DeleteGrowersLocation ********************* 
' ******************************************************* 
function DeleteGrowersLocation(pID)
	dim l_sql
	l_sql = "DELETE FROM GrowersLocations WHERE GLoc_GrowerLocationID = " & pID
	conn.execute l_sql, , 129
end function

' ******************************************************* 
' ************ InsertGrowersLocation ********************* 
' ******************************************************* 
function InsertGrowersLocation(pGrowerID, pLocation, pWatershed)
	dim rsInsert
	dim l_sql
	on error resume next
'Response.Write("<br>InsertGrowersLocation<br>GrowerID: " & pGrowerID)
'Response.flush
	l_sql = "INSERT INTO GrowersLocations(GLoc_GrowerID, GLoc_Location, GLoc_PURS_Watershed) VALUES (" & _
				pGrowerID & ",'" & pLocation & "','" & pWatershed & "')"
	conn.execute l_sql, , 129
	
REM
'    Dim objCmd, oRS

'    Set objCmd = Server.CreateObject("ADODB.Command")
'    objCmd.ActiveConnection = conn 'objStoreConn

'    objCmd.CommandText = "GrowersLocations"
'    objCmd.CommandType = adCmdTable
'    Set oRS = Server.CreateObject("ADODB.Recordset")
'    oRS.CursorType = adOpenStatic
'    oRS.CursorLocation = adUseClientBatch
'    oRS.LockType = adLockPessimistic
'    oRS.Open objCmd
'    oRS.AddNew

'    oRS.Fields("GLoc_GrowerID") = pGrowerID
'    oRS.Fields("GLoc_Location") = pLocation
    
'    on error resume next
'    oRS.Update

'    Dim iErr, iNumErrs, bReturn
'    iNumErrs = oRS.ActiveConnection.Errors.Count
'    If iNumErrs > 0 Then
'        bReturn = TRUE
'        For iErr = 0 to iNumErrs
'            Response.Write("<br><b>Number: " & CStr(iErr) & "</b><br>" & oRS.ActiveConnection.Errors(iErr).Number)
'            Response.Write("<br><b>Description</b><br>" & oRS.ActiveConnection.Errors(iErr).Description)
'            Response.Write("<br><b>NativeError</b><br>" & oRS.ActiveConnection.Errors(iErr).NativeError)
'            Response.Write("<br><b>Source</b><br>" & oRS.ActiveConnection.Errors(iErr).Source)
'            Response.Write("<br><b>SQLState</b><br>" & oRS.ActiveConnection.Errors(iErr).SQLState)
'            Response.Write("<br><b>HelpContext</b><br>" & oRS.ActiveConnection.Errors(iErr).HelpContext)
'            Response.Write("<br><b>HelpFile</b><br>" & oRS.ActiveConnection.Errors(iErr).HelpFile)

			REM SPECIFIC TO DATA PROVIDER!!!
'			If oRS.ActiveConnection.Errors(iErr).NativeError = -105121349 Then
'				InsertGrower = 0
'			End If
'        Next
'		If InsertGrower = -1 Then
'			Response.flush
'			Response.End
'		End If
'    Else
'		InsertGrower = oRS.Fields("GrowerID")
'    End If

'    oRS.Close
'    Set oRS = Nothing

'    Set objCmd = Nothing

REM	
	
	if err then
		newID = 0
	else
REM this is a lousy way of doing an insert. Kim Miers 7/16/2006
		DIM newID
		l_sql = "SELECT MAX(GLoc_GrowerLocationID) AS insertid FROM GrowersLocations"
		set rsInsert = conn.execute(l_sql)
		newID = rsInsert(0)
	end if
	InsertGrowersLocation = newID
end function

' ******************************************************* 
' ************ UpdateGrowersLocation ********************* 
' ******************************************************* 
function UpdateGrowersLocation(pID, pGrowerLocation, pWatershed)
	dim l_sql
	l_sql = "UPDATE GrowersLocations SET GLoc_Location ='" & _
		pGrowerLocation & "', GLoc_PURS_Watershed = '" & _
		pWatershed & "' WHERE GLoc_GrowerLocationID = " & pID
	conn.execute l_sql, , 129
	UpdateGrowersLocation = pID
end function

' ******************************************************* 
' ************ ActivateGrowersLocation ******************* 
' ******************************************************* 
function ActivateGrowersLocation(pID)
	dim l_sql
	l_sql = "Update GrowersLocations SET GLoc_Active = 1 WHERE GLoc_GrowerLocationID = " & pID
	conn.execute l_sql, , 129
end function

' ******************************************************* 
' ************ DeActivateGrowersLocation ****************** 
' ******************************************************* 
function DeActivateGrowersLocation(pID)
	dim l_sql
	l_sql = "Update GrowersLocations SET GLoc_Active = 0 WHERE GLoc_GrowerLocationID = " & pID
	conn.execute l_sql, , 129
end function
%>












