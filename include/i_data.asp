
<%


FUNCTION RemoveQuotes(txt)
	if txt <> "" then
		RemoveQuotes = Replace(txt,"'","")
	end if
END FUNCTION

FUNCTION EscapeQuotes(txt)
response.write("<br>" & txt)
	if txt <> "" then
		EscapeQuotes = Replace(txt,"'","''")
	end if
END FUNCTION

FUNCTION ValidateDatatype(thisfield,thistype,thisname,required)
	ValidateDatatype = TRUE
	if thisfield = "" and required then
		ValidateDatatype = FALSE
		tempErrorMessage = thisname + " is required."
	end if
	if thistype = "datetime" then
		if thisfield <> "" then
			if not IsDate(thisfield) then
				ValidateDatatype = FALSE
				tempErrorMessage = thisname + " must be a date in the format 'mm/dd/yyyy'."
			end if
		end if
	end if
	if thistype = "float" then
		if thisfield <> "" then
			if NOT IsNumeric(thisfield) then
				ValidateDatatype = FALSE
				tempErrorMessage = thisname + " must be a number."
			end if
		end if
	end if
END FUNCTION

FUNCTION Connect()
	DIM oConn,cst
    set oConn = CreateObject("ADODB.Connection") 
    'oConn.open cst 
	oConn.mode = 3 ' adModeReadWrite

'	Dim fso, root, fold
'	set fso = CreateObject("Scripting.FileSystemObject") 
'	root = Server.MapPath("\") 
'	root = mid(root,1,len(root)- 7)
'	set fold = fso.getFolder(root) 

	oConn.open (Application("ConnStr"))

'	Set fso = Nothing
	Set Connect = oConn
END FUNCTION

FUNCTION EndConnect(oConn)
		oConn.Close
		SET oConn = Nothing
END FUNCTION

FUNCTION ListContains(listname,strValue)
	Dim tArray,count
    ListContains = FALSE
	tArray = Split(listname, ",")
  	IF isarray(tArray) THEN
    	FOR count = LBound(tArray) to UBound(tArray)
    		IF trim(lcase(strValue)) = trim(lcase(tArray(Count))) THEN
		        ListContains = True
		    END IF
		NEXT
	END IF
END FUNCTION

FUNCTION ListAppend(listname,strValue)
	ListAppend = listname & "," & strValue
END FUNCTION

FUNCTION RemoveWhitespace(strValue)
	RemoveWhitespace = replace(strValue," ","")
END FUNCTION

    Function isEmailValid(email) 
		dim regEx
        Set regEx = New RegExp 
        regEx.Pattern = "^\w+([-+.]\w+)*@\w+" & _ 
            "([-.]\w+)*\.\w+([-.]\w+)*$" 
        isEmailValid = regEx.Test(trim(email)) 
    End Function 
	
Function ValidString(sInput)
 Dim iLen 'length of the email address
 Dim iCtr 'track # of characters in email address
 Dim sChar
 Dim bFieldIsOkay 'boolean value each field
 Dim sSymbol 'denotes the @ symbol
 Dim sExtension 'last four characters of email address
 Dim bExtension  'boolean value for email extension
 Dim searchSTring,searchchar,mypos,searchstring2,searchchar2,mypos2,sValid,t

 bFieldIsOkay = true
 bExtension = false

 sExtension = right(sInput, 4) 'get the last 4 characters in the emailaddress
 iLen = len(sInput)            'get the length of the email address
 searchstring = sInput         'search string will equal the email address
 searchchar = "@"

 mypos = instr(1, searchstring, searchchar, 1)     'determines the positionof the @ in the email address
 sSymbol = mid(searchchar, 1, iLen)                'search for the @ in theemail address
 searchstring2 = searchstring                      'search string will equalthe email address
 searchchar2 = right(sInput, 4)                    'use this as reference 2for search
 mypos2 = instr(1, searchstring2, searchchar2, 1)  'determines the startposition of the extension

 'list of valid characters
 sValid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-."

 if iLen < 7 then
  bFieldIsOkay = false
  exit function
 else
  if searchchar <> sSymbol then
   bFieldIsOkay = false
   exit function
  else
   for t = 1 to iLen
    if instr(searchchar, LCase(mid(sInput, t, 1))) <> 0 then
     counter = counter + 1
    end if
   next
   if counter <> 1 then
    bFieldIsOkay = false
    exit function
   else
    Select Case sExtension
     case ".com", ".net", ".org", ".gov", ".edu"
     bExtension = true
    end select
   end if

   if bExtension = false then
    bFieldIsOkay = false
    exit function
   else
    for iCtr = 1 to mypos
     sChar = mid(sInput, iCtr, 1)
    next
    if iCtr <1 then
     bFieldIsOkay = false
     exit function
    else
     for iCtr = 1 to mypos -1
      sChar = mid(sInput, iCtr, 1)
      if not instr(sValid, sChar) <> 0 then
       bFieldIsOkay = false
      end if
     next
    end if
    for iCtr = mypos to mypos2
     sChar = mid(sInput, iCtr, 1)
    next
    if iCtr < 1 then
     bFieldIsOkay = false
     exit function
    else
     for iCtr = mypos + 1 to mypos2 - 1
      sChar = mid(sInput, iCtr, 1)
      if not instr(sValid, sChar) <> 0 then
       bFieldIsOkay = false
      end if
     next
    end if
   end if
  end if
 end if

 'If bFieldIsOkay = True Then Response.Redirect "response.asp"
 ValidString = bFieldIsOkay
end Function
	
Function FormDataDump(bolShowOutput, bolEndPageExecution)
  Dim sItem

  'What linebreak character do we need to use?
  Dim strLineBreak
  If bolShowOutput then
    'We are showing the output, so set the line break character
    'to the HTML line breaking character
    strLineBreak = "<br>"
  Else
    'We are nesting the data dump in an HTML comment block, so
    'use the carraige return instead of <br>
    'Also start the HTML comment block
    strLineBreak = vbCrLf
    Response.Write("<!--" & strLineBreak)
  End If
  

  'Display the Request.Form collection
  Response.Write("DISPLAYING REQUEST.FORM COLLECTION" & strLineBreak)
  For Each sItem In Request.Form
    Response.Write(sItem)
    Response.Write(" - [" & Request.Form(sItem) & "]" & strLineBreak)
  Next
  
  
  'Display the Request.QueryString collection
  Response.Write(strLineBreak & strLineBreak)
  Response.Write("DISPLAYING REQUEST.QUERYSTRING COLLECTION" & strLineBreak)
  For Each sItem In Request.QueryString
    Response.Write(sItem)
    Response.Write(" - [" & Request.QueryString(sItem) & "]" & strLineBreak)
  Next

  
  'If we are wanting to hide the output, display the closing
  'HTML comment tag
  If Not bolShowOutput then Response.Write(strLineBreak & "-->")

  'End page execution if needed
  If bolEndPageExecution then Response.End
End Function




%>