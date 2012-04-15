<!-- #include file="Adovbs.asp" -->
<%
' FILE: i_packers.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetAllPackers ******************************* 
' ******************************************************* 
function GetAllPackers()
	sql = "SELECT * FROM Packers WHERE (1=1) "
	IF session("packerid") <> 0 THEN
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
	IF session("packerid") <> 0 THEN
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
	sql = "SELECT * FROM Packers WHERE PackerName like '%" & rTrim(lTrim(Name)) & "%' AND PAckerNumber like '%" & rTrim(lTrim(Number)) & "%' ORDER BY PackerName "
	set GetPackersByNameNumber = conn.execute(sql)
end function

' ******************************************************* 
' ************GetpackersByID ******************************* 
' ******************************************************* 
function GetpackersByID(ID)
	dim sql
	sql = "SELECT * FROM packers WHERE packerID in (" & ID & ")"
	set GetpackersByID = conn.execute(sql)
end function

' ******************************************************* 
' ************ Deletepackers **************************** 
' ******************************************************* 
function Deletepackers(ID)
	sql = "DELETE FROM packers WHERE packerID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************Insertpackers ******************************* 
' ******************************************************* 
function Insertpackers(packerName,Email,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax)
if ZipCode = "" or NOT IsNumeric(ZipCode) THEN
	ZipCode = "NULL"
END IF

    if growerlimit="" then growerlimit=0

	sql = "INSERT INTO packers(packerName,Email,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,Active) VALUES ("
sql = sql & ",'" & EscapeQuotes(packerName) & "'"
sql = sql & ",'" & EscapeQuotes(Email) & "'"
sql = sql & ",'" & EscapeQuotes(Address) & "'"
sql = sql & ",'" & EscapeQuotes(City) & "'"
sql = sql & ",'" & EscapeQuotes(State) & "'," & ZipCode  
sql = sql & ",'" & EscapeQuotes(Contact) & "'"
sql = sql & ",'" & EscapeQuotes(Telephone1) & "'"
sql = sql & ",'" & EscapeQuotes(Telephone2) & "'"
sql = sql & ",'" & EscapeQuotes(Fax) & "'"
sql = sql & ",1" 
sql = sql & ")"
response.write sql

	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(packerID) AS insertid FROM packers"
	set rs = conn.execute(sql)
	newID = rs(0)
	Insertpackers = newID
end Function

'KILROY

FUNCTION Insertpacker(packerName,Email,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,FullRights,GrowerLimit)
REM *** New f(x) to Insertpacker 7/13/2006
REM *** Kim Miers
REM *** provides additional check for uniqueness of packer Number
	REM initially set return value, display if remains -1.
	Insertpacker = -1

	IF ZipCode = "" or NOT IsNumeric(ZipCode) THEN
		ZipCode = NULL
	END IF

    Dim objCmd, oRS

    Set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = conn 'objStoreConn

    objCmd.CommandText = "packers"
    objCmd.CommandType = adCmdTable
    Set oRS = Server.CreateObject("ADODB.Recordset")
    oRS.CursorType = adOpenStatic
    oRS.CursorLocation = adUseClientBatch
    oRS.LockType = adLockPessimistic
    oRS.Open objCmd
    oRS.AddNew

    if growerlimit="" then growerlimit=0

    oRS.Fields("packerName") = packerName
    ors.Fields("fullrights") = cint(fullrights)
    ors.Fields("growerlimit")= cint(growerlimit)
    oRS.Fields("Email") = Email
    oRS.Fields("Address") = Address
    oRS.Fields("City") = City
    oRS.Fields("State") = State
    oRS.Fields("ZipCode") = ZipCode
    oRS.Fields("Contact") = Contact
    oRS.Fields("Telephone1") = Telephone1
    oRS.Fields("Telephone2") = Telephone2
    oRS.Fields("Fax") = Fax
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
				Insertpacker = 0
			End If
        Next
		If Insertpacker = -1 Then
			Response.flush
			Response.End
		End If
    Else
		Insertpacker = oRS.Fields("packerID")
    End If

    oRS.Close
    Set oRS = Nothing

    Set objCmd = Nothing

	
END FUNCTION 'Insertpacker
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
' ************Updatepackers ******************************* 
' ******************************************************* 
function Updatepackers(packerID,packerName,Email,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,FullRights,GrowerLimit)
'response.write("hi<br><br><br>")
	if ZipCode = "" or NOT IsNumeric(ZipCode) THEN
		ZipCode = "NULL"
	END IF
	
    if growerlimit="" then growerlimit=0
	sql = "UPDATE packers SET packerName ='" & EscapeQuotes(packerName) &  "', FullRights="&cint(fullrights)&", Growerlimit="&growerlimit&", Email ='" & _
		EscapeQuotes(Email) &  "',  Address ='" & _
		EscapeQuotes(Address) & "', City ='" & City & "',State ='" & State & "',ZipCode = " & _
		ZipCode & ",Contact ='" & Contact & "',Telephone1 ='" & Telephone1 & "',Telephone2 ='" & _
		Telephone2 & "',Fax ='" & Fax & "' WHERE packerID = " & packerID
'response.write sql
	conn.execute sql, , 129
	Updatepackers = packerID
end Function


' ******************************************************* 
' ************UpdatepackersBypacker ******************************* 
' ******************************************************* 
function UpdatepackersBypacker(packerID,Email,Username,Password,Address,City,State,ZipCode,Contact,Telephone1,Telephone2,Fax,Cell,Fieldman,ApplicatorSupervisor,SupervisorLicense,Applicator,ApplicatorLicense,ChemicalSupplier,RecommendedBy)
response.write("hi<br><br><br>")
if ZipCode = "" or NOT IsNumeric(ZipCode) THEN
	ZipCode = "NULL"
END IF
	sql = "UPDATE packers SET  Email ='" & EscapeQuotes(Email) &  "', packerPassword ='" & EscapeQuotes(Password) &  "',  Address ='" & Address & "', City ='" & City & "',State ='" & State & "',ZipCode = " & ZipCode & ",Contact ='" & Contact & "',Telephone1 ='" & Telephone1 & "',Telephone2 ='" & Telephone2 & "',Fax ='" & Fax & "',Cell ='" 
sql = sql & Cell & "',Fieldman ='" & Fieldman & "',ApplicatorSupervisor ='" & ApplicatorSupervisor & "',SupervisorLicense ='" & SupervisorLicense & "',Applicator ='" & Applicator & "',ApplicatorLicense ='" & ApplicatorLicense & "',ChemicalSupplier ='" & ChemicalSupplier & "',RecommendedBy ='" & RecommendedBy & "' WHERE packerID = " & packerID
response.write sql
	conn.execute sql, , 129
	UpdatepackersBypacker = packerID
end Function

' ******************************************************* 
' ************AgreeToTerms ******************************* 
' ******************************************************* 
function AgreeToTerms(Name)
	sql = "UPDATE packers SET  TermsAgreed = 1, AgreedBy ='" & EscapeQuotes(Name) &  "', AgreedDate = getdate() WHERE packerID = " & session("packerid")
response.write sql
	conn.execute sql, , 129
	AgreeToTerms = session("packerid")
end Function

' ******************************************************* 
' ************EmailPassword ******************************* 
' ******************************************************* 
function EmailPassword(email)
	sql = "SELECT count(packerPassword) FROM packers WHERE Email = '" & EscapeQuotes(email) & "'"
	set rs = conn.execute(sql)
	IF (rs(0) > 0) THEN
		sql = "SELECT packerPassword FROM packers WHERE Email = '" & EscapeQuotes(email) & "'"
		set rs = conn.execute(sql)
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			Mailer.RemoteHost = "mail.gorge.net"
			Mailer.FromAddress = Application("CONTACT_EMAIL")
			Mailer.AddRecipient "","email" 
			Mailer.Subject = Application("HOST_WEBSITE") & " Information"
			Mailer.BodyText = "This is your password reminder from " & Application("HOST_WEBSITE") & ".  It is: " & rs.Fields("packerPassword") & "\n\nFor more information please contact " & Application("CONTACT_NAME") & " at " & Application("CONTACT_EMAIL")
			Mailer.SendMail    
		Set Mailer = Nothing			
		EmailPassword = 1

	ELSE
		EmailPassword = 0
	END IF
end function

' ******************************************************* 
' ************Activatepackers ******************************* 
' ******************************************************* 
function Activatepackers(ID)
	sql = "Update packers SET Active = 1 WHERE packerID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivatepackers ******************************* 
' ******************************************************* 
function DeActivatepackers(ID)
	sql = "Update packers SET Active = 0 WHERE packerID = " & ID
	conn.execute sql, , 129
end function
%>