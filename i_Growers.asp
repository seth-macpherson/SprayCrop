<!-- #include file="Adovbs.asp" -->
<%
' FILE: i_Growers.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetAllPackers ******************************* 
' ******************************************************* 
function GetAllPackers()
	sql = "SELECT * FROM Packers WHERE (1=1) "
	IF session("growerid") <> 0 THEN
		sql = sql
	ELSEIF session("packerid") <> 0 THEN
		sql = sql & " AND packerid = " & session("packerid")  & ")"
	END IF
	sql = sql & " ORDER BY PackerNumber"
	set GetAllPackers = conn.execute(sql)
end function

' ******************************************************* 
' ************GetActivePackers ******************************* 
' ******************************************************* 
function GetActivePackers()
	sql = "SELECT * FROM Packers WHERE Active=1 "
	IF session("growerid") <> 0 THEN
		sql = sql
	ELSEIF session("packerid") <> 0 THEN
		sql = sql & " AND packerid = " & session("packerid")  & ")"
	END IF
	sql = sql & " ORDER BY PackerNumber"
	set GetActivePackers = conn.execute(sql)
end function

' ******************************************************* 
' ************ GetPackersByNameNumber ******************* 
' ******************************************************* 
function GetPackersByNameNumber(Name,Number)
	sql = "SELECT * FROM Packers WHERE PackerName like '%" & rTrim(lTrim(Name)) & "%' AND PAckerNumber like '%" & rTrim(lTrim(Number)) & "%'"
	IF session("growerid") <> 0 THEN
		sql = sql
	ELSEIF session("packerid") <> 0 THEN
		sql = sql & " AND packerid = " & session("packerid")  & ")"
	END IF
	sql = sql & "ORDER BY PackerName "
	set GetPackersByNameNumber = conn.execute(sql)
end function

' ******************************************************* 
' ************GetGrowersByID ******************************* 
' ******************************************************* 
function GetGrowersByID(ID)
	dim sql
	sql = "SELECT * FROM Growers WHERE GrowerID in (" & ID & ")"
	set GetGrowersByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllGrowerss ******************************* 
' ******************************************************* 
function GetAllGrowers()
REM fix additional numbers.
REM 7/14/2006 kim miers. changed grower number from numeric to text blew this up.
	dim cAddNums
	cAddNums = session("AdditionalNumbers")
	cAddNums = "'" + cAddNums + "'"
	cAddNums = replace(cAddNums,",","','")
	sql = "SELECT * FROM Growers WHERE (1=1) "
	IF session("growerid") <> 0 THEN
		sql = sql & " AND (Growers.GrowerID = " & session("growerid")&")"  '& " OR Growers.GrowerNumber IN (" & cAddNums & "))
	ELSEIF session("packerid") <> 0 THEN
		sql = sql & " AND Growers.GrowerID IN (select distinct growerid from sprayrecord where packerid = " & session("packerid")  & ")"
		'sql = sql & " AND Growers.GrowerID IN (select growerid from growerpacker where packerid = " & session("packerid")  & ")"
	END IF
	sql = sql & " ORDER BY GrowerName"
	set GetAllGrowers = conn.execute(sql)
end function

' ******************************************************* 
' ************GetActiveGrowers ******************************* 
' ******************************************************* 
function GetActiveGrowers()
REM fix additional numbers.
REM 7/14/2006 kim miers. changed grower number from numeric to text blew this up.
	dim cAddNums
	cAddNums = session("AdditionalNumbers")
	cAddNums = "'" + cAddNums + "'"
	cAddNums = replace(cAddNums,",","','")
	sql = "SELECT * FROM Growers WHERE Active = 1 "
	IF listContains("1", session("accessid")) then
		sql = sql
	ELSEIF session("growerid") <> 0 THEN
		sql = sql & " AND (Growers.GrowerID = " & session("growerid")  & ")" ' OR Growers.GrowerNumber IN (" & cAddNums & "))"
	ELSEIF session("packerid") <> 0 THEN
		sql = sql & " AND Growers.GrowerID IN (select distinct growerid from sprayrecord where packerid = " & session("packerid")  & ")"
		'sql = sql & " AND Growers.GrowerID IN (select growerid from growerpacker where packerid = " & session("packerid")  & ")"
	END IF
	sql = sql & " ORDER BY GrowerName"
'	response.write sql
	set GetActiveGrowers = conn.execute(sql)
end function

' ******************************************************* 
' ************ GetGrowersByNameNumber ******************* 
' ******************************************************* 
function GetGrowersByNameNumber(Name,Number)
	sql = "SELECT * FROM Growers WHERE GrowerName like '%" & rTrim(lTrim(Name)) & "%'"
	IF listContains("1", session("accessid")) then
		sql = sql
	ELSEIF session("growerid") <> 0 THEN
		sql = sql & " AND (Growers.GrowerID = " & session("growerid")  & ")" ' OR Growers.GrowerNumber IN (" & cAddNums & "))"
	ELSEIF session("packerid") <> 0 THEN
		sql = sql & " AND Growers.GrowerID IN (select distinct growerid from growerpacker where packerid = " & session("packerid")  & ")"
	END IF
	sql = sql & " ORDER BY GrowerName"
	
	set GetGrowersByNameNumber = conn.execute(sql)
end function

' ******************************************************* 
' ************ xGetGrowersLocationsByID ****************** 
' ******************************************************* 
function xGetGrowersLocationsByID(pGrowerID)
'	sql = "SELECT GrowersLocations.* FROM GrowersLocations WHERE GrowerID = " & pGrowerID & " ORDER BY GLoc_Location "
'	set GetGrowersLocationsByID = conn.execute(sql)

REM TEMP KILROY
    Dim objCmd, oRS

    Set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = conn

    objCmd.CommandText = "Q_GrowersLocationsByGrowerID"
    objCmd.CommandType = adCmdStoredProc
    
    Set objParam = objCmd.CreateParameter("pGrowerID", adNumeric, adParamInput, Len(pGrowerID), pGrowerID)
    objCmd.Parameters.Append objParam
    
    Set oRS = Server.CreateObject("ADODB.Recordset")
    oRS.CursorType = adOpenStatic
    oRS.CursorLocation = adUseClientBatch
    oRS.LockType = adLockPessimistic
    oRS.Open objCmd

	set xGetGrowersLocationsByID = oRS
'    oRS.Close
'    Set oRS = Nothing

'    Set objCmd = Nothing

REM TEMP KILROY

end function

' ******************************************************* 
' ************ DeleteGrowers **************************** 
' ******************************************************* 
function DeleteGrowers(ID)
	sql = "DELETE FROM Growers WHERE GrowerID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertGrowers ******************************* 
' ******************************************************* 
function InsertGrowers(AdditionalGrowerNumbers,GrowerNumber,GrowerName,Email,Username,Password,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,Cell,Fieldman,ApplicatorSupervisor,SupervisorLicense,Applicator,ApplicatorLicense,ChemicalSupplier,RecommendedBy,InternalNote)
if ZipCode = "" or NOT IsNumeric(ZipCode) THEN
	ZipCode = "NULL"
END IF

'sql = "INSERT INTO Growers(GrowerNumber,AdditionalGrowerNumbers,GrowerName,Email,GrowerPassword,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,Cell,Fieldman,ApplicatorSupervisor,SupervisorLicense,Applicator,ApplicatorLicense,ChemicalSupplier,RecommendedBy,InternalNote,Active) VALUES ("
sql = "INSERT INTO Growers(GrowerName,Email,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,Cell,Fieldman,ApplicatorSupervisor,SupervisorLicense,Applicator,ApplicatorLicense,ChemicalSupplier,RecommendedBy,InternalNote,Active) VALUES ("
'sql = sql & "'" & EscapeQuotes(GrowerNumber) & "'"
'sql = sql & ",'" & EscapeQuotes(AdditionalGrowerNumbers) & "'"
sql = sql & ",'" & EscapeQuotes(GrowerName) & "'"
sql = sql & ",'" & EscapeQuotes(Email) & "'"
'sql = sql & ",'" & EscapeQuotes(Password) & "'"
sql = sql & ",'" & EscapeQuotes(Address) & "'"
sql = sql & ",'" & EscapeQuotes(City) & "'"
sql = sql & ",'" & EscapeQuotes(State) & "'," & ZipCode  
sql = sql & ",'" & EscapeQuotes(Contact) & "'"
sql = sql & ",'" & EscapeQuotes(Telephone1) & "'"
sql = sql & ",'" & EscapeQuotes(Telephone2) & "'"
sql = sql & ",'" & EscapeQuotes(Fax) & "'"
sql = sql & ",'" & EscapeQuotes(Cell) & "'"
sql = sql & ",'" & EscapeQuotes(Fieldman) & "'"
sql = sql & ",'" & EscapeQuotes(ApplicatorSupervisor) & "'"
sql = sql & ",'" & EscapeQuotes(SupervisorLicense) & "'"
sql = sql & ",'" & EscapeQuotes(Applicator) & "'"
sql = sql & ",'" & EscapeQuotes(ApplicatorLicense) & "'"
sql = sql & ",'" & EscapeQuotes(ChemicalSupplier) & "'"
sql = sql & ",'" & EscapeQuotes(RecommendedBy) & "'"
sql = sql & ",'" & EscapeQuotes(InternalNote) & "'" 
sql = sql & ",1" 
sql = sql & ")"
response.write sql

	conn.execute sql, , 129

	DIM newID
	sql = "SELECT MAX(GrowerID) AS insertid FROM Growers"
	set rs = conn.execute(sql)
	newID = rs(0)

	IF session("packerid") <> 0 THEN
		sql = "INSERT growerpacker (growerid,packerid) values ("&newid&","&session("packerid")&")"
	END IF

	InsertGrowers = newID
end Function

'KILROY

FUNCTION InsertGrower(AdditionalGrowerNumbers,GrowerNumber,GrowerName,Email,Username,Password,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,Cell,Fieldman,ApplicatorSupervisor,SupervisorLicense,Applicator,ApplicatorLicense,ChemicalSupplier,RecommendedBy,InternalNote,fname,lname)
REM *** New f(x) to InsertGrower 7/13/2006
REM *** Kim Miers
REM *** provides additional check for uniqueness of Grower Number
	REM initially set return value, display if remains -1.
	InsertGrower = -1

	IF ZipCode = "" or NOT IsNumeric(ZipCode) THEN
		ZipCode = NULL
	END IF

    Dim objCmd, oRS

    Set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = conn 'objStoreConn

    objCmd.CommandText = "Growers"
    objCmd.CommandType = adCmdTable
    Set oRS = Server.CreateObject("ADODB.Recordset")
    oRS.CursorType = adOpenStatic
    oRS.CursorLocation = adUseClientBatch
    oRS.LockType = adLockPessimistic
    oRS.Open objCmd
    oRS.AddNew

    oRS.Fields("GrowerNumber") = GrowerNumber
    oRS.Fields("AdditionalGrowerNumbers") = AdditionalGrowerNumbers
    oRS.Fields("GrowerName") = GrowerName
    oRS.Fields("Email") = Email
    oRS.Fields("GrowerPassword") = Password
    oRS.Fields("Address") = Address
    oRS.Fields("City") = City
    oRS.Fields("State") = State
    oRS.Fields("ZipCode") = ZipCode
    oRS.Fields("Contact") = Contact
    oRS.Fields("Telephone1") = Telephone1
    oRS.Fields("Telephone2") = Telephone2
    oRS.Fields("Fax") = Fax
    oRS.Fields("Cell") = Cell
    oRS.Fields("Fieldman") = Fieldman
    oRS.Fields("ApplicatorSupervisor") = ApplicatorSupervisor
    oRS.Fields("SupervisorLicense") = SupervisorLicense
    oRS.Fields("Applicator") = Applicator
    oRS.Fields("ApplicatorLicense") = ApplicatorLicense
    oRS.Fields("ChemicalSupplier") = ChemicalSupplier
    oRS.Fields("RecommendedBy") = RecommendedBy
	oRS.Fields("InternalNote") = Trim(InternalNote)
    oRS.Fields("Active") = 1
    
    on error resume next
    oRS.Update

    Dim iErr, iNumErrs, bReturn
    iNumErrs = oRS.ActiveConnection.Errors.Count
    If iNumErrs > 0 Then
        bReturn = TRUE
        For iErr = 0 to iNumErrs
            Response.Write("<br><b>Number: " & CStr(iErr) & "</b><br>" & oRS.ActiveConnection.Errors(iErr).Number)
            Response.Write("<br><b>Description</b><br>" & oRS.ActiveConnection.Errors(iErr).Description)
            Response.Write("<br><b>NativeError</b><br>" & oRS.ActiveConnection.Errors(iErr).NativeError)
            Response.Write("<br><b>Source</b><br>" & oRS.ActiveConnection.Errors(iErr).Source)
            Response.Write("<br><b>SQLState</b><br>" & oRS.ActiveConnection.Errors(iErr).SQLState)
            Response.Write("<br><b>HelpContext</b><br>" & oRS.ActiveConnection.Errors(iErr).HelpContext)
            Response.Write("<br><b>HelpFile</b><br>" & oRS.ActiveConnection.Errors(iErr).HelpFile)

			REM SPECIFIC TO DATA PROVIDER!!!
			If oRS.ActiveConnection.Errors(iErr).NativeError = -105121349 Then
				InsertGrower = 0
			End If
        Next
		If InsertGrower = -1 Then
			Response.flush
			Response.End
		End If
    Else
    
    	IF session("packerid") <> 0 THEN
	    	sql = "INSERT growerpacker (growerid,packerid) values ("&ors.fields("growerid")&","&session("packerid")&")"
            conn.execute sql, , 129
	    END IF
	    
    	sql = "INSERT INTO growerusers(Username,Password,Fname,Lname,growerID,AccessID) "
    	sql = sql & "VALUES ('" & EscapeQuotes(fname)&EscapeQuotes(lname)&oRS.Fields("GrowerID") & "', "
    	sql = sql & "'" & EscapeQuotes(fname)&EscapeQuotes(lname)&oRS.Fields("GrowerID") & "', "
    	sql = sql & "'" & EscapeQuotes(fname) & "', '" & EscapeQuotes(lname) & "'," & ors.fields("growerid") & ", 3)"
        conn.execute sql, , 129
		InsertGrower = oRS.Fields("GrowerID")
		
    End If

    oRS.Close
    Set oRS = Nothing

    Set objCmd = Nothing

	
END FUNCTION 'InsertGrower
'KILROY

FUNCTION Local_ADO_RS_Error(ByRef pRS)
    Dim iErr, iNumErrs, bReturn
    bReturn = FALSE
    REM keep going if further errors
    on error resume next
    iNumErrs = pRS.ActiveConnection.Errors.Count
    If iNumErrs > 0 Then
        bReturn = TRUE
        For iErr = 0 to iNumErrs
            WriteLn("<br><b>Number: " & CStr(iErr) & "</b><br>" & pRS.ActiveConnection.Errors(iErr).Number)
            WriteLn("<br><b>Description</b><br>" & pRS.ActiveConnection.Errors(iErr).Description)
            WriteLn("<br><b>NativeError</b><br>" & pRS.ActiveConnection.Errors(iErr).NativeError)
            WriteLn("<br><b>Source</b><br>" & pRS.ActiveConnection.Errors(iErr).Source)
            WriteLn("<br><b>SQLState</b><br>" & pRS.ActiveConnection.Errors(iErr).SQLState)
            WriteLn("<br><b>HelpContext</b><br>" & pRS.ActiveConnection.Errors(iErr).HelpContext)
            WriteLn("<br><b>HelpFile</b><br>" & pRS.ActiveConnection.Errors(iErr).HelpFile)
        Next
    End If
    Local_ADO_RS_Error = bReturn
END FUNCTION



' ******************************************************* 
' ************UpdateGrowers ******************************* 
' ******************************************************* 
function UpdateGrowers(GrowerID,AdditionalGrowerNumbers,GrowerNumber,GrowerName,Email,Username,Password,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,Cell,Fieldman,ApplicatorSupervisor,SupervisorLicense,Applicator,ApplicatorLicense,ChemicalSupplier,RecommendedBy,InternalNote,Active,CreateDate,UnitID)
    'response.write("hi<br><br><br>")
	if ZipCode = "" or NOT IsNumeric(ZipCode) THEN
		ZipCode = "NULL"
	END IF
	if unitid=0 or not isnumeric(unitid) then
        unitid="null"        
	end if
	sql = "UPDATE Growers SET GrowerNumber ='" & GrowerNumber & "',AdditionalGrowerNumbers ='" & _
		AdditionalGrowerNumbers & "',GrowerName ='" & EscapeQuotes(GrowerName) &  "', Email ='" & _
		EscapeQuotes(Email) &  "', GrowerPassword ='" & EscapeQuotes(Password) &  "',  Address ='" & _
		EscapeQuotes(Address) & "', City ='" & City & "',State ='" & State & "',ZipCode = " & _
		ZipCode & ",Contact ='" & Contact & "',Telephone1 ='" & Telephone1 & "',Telephone2 ='" & _
		Telephone2 & "',Fax ='" & Fax & "',Cell ='" & _
		Cell & "',Fieldman ='" & Fieldman & "',ApplicatorSupervisor ='" & _
		ApplicatorSupervisor & "',SupervisorLicense ='" & SupervisorLicense & "',Applicator ='" & _
		Applicator & "',ApplicatorLicense ='" & ApplicatorLicense & "',ChemicalSupplier ='" & _
		ChemicalSupplier & "',RecommendedBy ='" & RecommendedBy & "',InternalNote ='" & _
		InternalNote & "', UnitID="&unitID&" WHERE GrowerID = " & GrowerID
'response.write sql
	conn.execute sql, , 129
	UpdateGrowers = GrowerID
end Function


' ******************************************************* 
' ************UpdateGrowersByGrower ******************************* 
' ******************************************************* 
function UpdateGrowersByGrower(GrowerID,Email,Username,Password,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,Cell,Fieldman,ApplicatorSupervisor,SupervisorLicense,Applicator,ApplicatorLicense,ChemicalSupplier,RecommendedBy)
response.write("hi<br><br><br>")
if ZipCode = "" or NOT IsNumeric(ZipCode) THEN
	ZipCode = "NULL"
END IF
	sql = "UPDATE Growers SET  Email ='" & EscapeQuotes(Email) &  "', GrowerPassword ='" & EscapeQuotes(Password) &  "',  Address ='" & Address & "', City ='" & City & "',State ='" & State & "',ZipCode = " & ZipCode & ",Contact ='" & Contact & "',Telephone1 ='" & Telephone1 & "',Telephone2 ='" & Telephone2 & "',Fax ='" & Fax & "',Cell ='" 
sql = sql & Cell & "',Fieldman ='" & Fieldman & "',ApplicatorSupervisor ='" & ApplicatorSupervisor & "',SupervisorLicense ='" & SupervisorLicense & "',Applicator ='" & Applicator & "',ApplicatorLicense ='" & ApplicatorLicense & "',ChemicalSupplier ='" & ChemicalSupplier & "',RecommendedBy ='" & RecommendedBy & "' WHERE GrowerID = " & GrowerID
response.write sql
	conn.execute sql, , 129
	UpdateGrowersByGrower = GrowerID
end Function

' ******************************************************* 
' ************AgreeToTerms ******************************* 
' ******************************************************* 
function AgreeToGrowerTerms(Name)
    sql = "UPDATE GrowerUsers SET  TermsAgreed = 1, AgreedBy ='" & EscapeQuotes(Name) &  "', AgreedDate = getdate() WHERE growerid = " & session("growerid") & " and Username = '" & session("username") & "'"
    'response.write sql
	conn.execute sql, , 129
	AgreeToGrowerTerms = session("growerid")
end Function


' ******************************************************* 
' ************AgreeToTerms ******************************* 
' ******************************************************* 
function AgreeToPackerTerms(Name)
    sql = "UPDATE PackerUsers SET  TermsAgreed = 1, AgreedBy ='" & EscapeQuotes(Name) &  "', AgreedDate = getdate() WHERE packerid = " & session("packerid") & " and Username = '" & session("username") & "'"
    'response.write sql
	conn.execute sql, , 129
	AgreeToPackerTerms = session("packerid")
end Function

' ******************************************************* 
' ************EmailPassword ******************************* 
' ******************************************************* 
function EmailPassword(email)
	sql = "SELECT count(GrowerPassword) FROM Growers WHERE Email = '" & EscapeQuotes(email) & "'"
	set rs = conn.execute(sql)
	IF (rs(0) > 0) THEN
		sql = "SELECT GrowerPassword FROM Growers WHERE Email = '" & EscapeQuotes(email) & "'"
		set rs = conn.execute(sql)
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			Mailer.RemoteHost = "mail.gorge.net"
			Mailer.FromAddress = Application("CONTACT_EMAIL")
			Mailer.AddRecipient "","email" 
			Mailer.Subject = Application("HOST_WEBSITE") & " Information"
			Mailer.BodyText = "This is your password reminder from " & Application("HOST_WEBSITE") & ".  It is: " & rs.Fields("GrowerPassword") & "\n\nFor more information please contact " & Application("CONTACT_NAME") & " at " & Application("CONTACT_EMAIL")
			Mailer.SendMail    
		Set Mailer = Nothing			
		EmailPassword = 1

	ELSE
		EmailPassword = 0
	END IF
end function

' ******************************************************* 
' ************ActivateGrowers ******************************* 
' ******************************************************* 
function ActivateGrowers(ID)
	sql = "Update Growers SET Active = 1 WHERE GrowerID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateGrowers ******************************* 
' ******************************************************* 
function DeActivateGrowers(ID)
	sql = "Update Growers SET Active = 0 WHERE GrowerID = " & ID
	conn.execute sql, , 129
end function
%>