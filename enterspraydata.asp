<%if not session("login") or not listContains("1,2,3", session("accessid")) then
	response.redirect("index.asp")
end if
%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayRecord.asp"-->
<!--#include file="i_Crop.asp"-->
<!--#include file="i_Growers.asp"-->
<!--#include file="i_GrowerLocations.asp"-->
<!--#include file="i_Method.asp"-->
<!--#include file="i_SprayList.asp"-->
<!--#include file="i_Stage.asp"-->
<!--#include file="i_Target.asp"-->
<!--#include file="i_Units.asp"-->
<!--#include file="i_Weather.asp" -->
<!--#include file="i_Varieties.asp"-->
<%
'CREATED by LocusInteractive on 08/02/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlSprayRecordID,formSprayRecordID,urlAdding,formAdding,arraySprayList,formApplicator,formApplicatorLicense
Dim conn,sql,rs,rsSelect,rsSelectTarget,counter,rsSelect2,j,thisTotalMaterialApplied

'Initialize variables
errorFound = FALSE
formError = FALSE
errorMessage = "The following errors have occurred:"
formSprayRecordID=Request.Form.Item("SprayRecordID")
urlSprayRecordID=Request.QueryString("SprayRecordID")

'See if ID was passed through URL or FORM
IF urlSprayRecordID = "" THEN urlSprayRecordID = 0 END IF
IF formSprayRecordID = "" THEN formSprayRecordID = urlSprayRecordID End IF
urlSprayRecordID = formSprayRecordID

formAdding=Request.Form.Item("Adding")
urlAdding=Request.QueryString("Adding")

'See if ID was passed through URL or FORM
IF urlAdding = "" THEN urlAdding = 0 END IF
IF formAdding = "" THEN formAdding = urlAdding End IF
urlAdding = formAdding

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Response.Redirect("enterspraydata.asp")
END IF

'Initialize Form Fields
DIM formPackerID,formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formBartlet,formStageID,formLocation,formMethodID,formWeatherID,formAcresTreated,formRateAcre,formTotalMaterialApplied,formSprayListID,formUnitsOfProduct,formIFPRating,formTargetID,formTargetID1,formTargetID2,formTargetID3,formTargetID4,formTargetID5,formHarvestDate,formComments,formLocationOption,formLocationText,arrayGrower,urlAcresTreated,formNewWeather,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy
DIM urlSprayStartDate,urlTimeFinishedSpraying,urlGrowerID,urlStageID,urlMethodID,urlWeatherID,urlLocation,urlSprayEndDate,urlCropID,urlApplicator,urlApplicatorLicense,urlAdministrator,urlSupervisor,urlLicenseNumber,urlChemicalSupplier,urlRecommendedBy

DIM formWeather,formTemp,formWindSpd,formWindDir
dim thisAcresTreated

'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")
set rsSelect2 = Server.CreateObject("ADODB.RecordSet")
set rsSelectTarget = Server.CreateObject("ADODB.RecordSet")

formPackerID = request.Form("PackerID")
if formPackerID = "" THEN formPackerID = "NULL"

     if request.servervariables("request_method")="POST" and request.Form("changerole")>"" then

        dim g: g=split(request.Form("growerrole"),"|")

        if isarray(g) then
            session("growerid")=g(0)
            session("growername")=g(1)
        end if

    end if

	'arrayGrower = Split(Request.Form.Item("GrowerID"),"|")
	'IF IsArray(arrayGrower)  THEN
	'	IF Ubound(arrayGrower) >= 0 THEN
	'		formGrowerID = arrayGrower(0)
	'	END IF
	'END IF
	'urlGrowerID=Request.QueryString("GrowerID")
	'if formGrowerID = "" THEN formGrowerID = urlGrowerID END IF
	'if formGrowerID = "" AND session("growerid") <> 0 then
		formGrowerID = session("growerID")
	'end if

formAdministrator = Request.Form.Item("Administrator")
urlAdministrator=Request.QueryString("Administrator")
if formAdministrator = "" then formAdministrator = urlAdministrator end if

formSprayStartDate = Request.Form.Item("SprayStartDate")
urlSprayStartDate=Request.QueryString("SprayStartDate")
if formSprayStartDate = "" then formSprayStartDate = urlSprayStartDate end if

formTimeFinishedSpraying = Request.Form.Item("TimeFinishedSpraying")
urlTimeFinishedSpraying=Request.QueryString("TimeFinishedSpraying")
if formTimeFinishedSpraying = "" then formTimeFinishedSpraying = TimeFinishedSpraying end if

formSprayEndDate = Request.Form.Item("SprayEndDate")
urlSprayEndDate=Request.QueryString("SprayEndDate")
if formSprayEndDate = "" then formSprayEndDate = urlSprayEndDate end if
if formSprayEndDate = "" then formSprayEndDate = formSprayStartDate end if

formCropID = Request.Form.Item("CropID")
urlCropID=Request.QueryString("CropID")
if formCropID = "" then formCropID = urlCropID end if

formVarietyID = Request.Form.Item("VarietyID")
urlVarietyID=Request.QueryString("VarietyID")
if formVarietyID = "" then formVarietyID = urlVarietyID end if

formBartlet = Request.Form.Item("Bartlet")

formStageID = Request.Form.Item("StageID")
urlStageID=Request.QueryString("StageID")
if formStageID = "" then formStageID = urlStageID end if

formLocationOption = Request.Form.Item("LocationOptions")
urlLocation=Request.QueryString("Location")
if formLocationOption = "" then formLocationOption = urlLocation end if

formLocationText = Request.Form.Item("LocationText")
if trim(formLocationOption) <> "" THEN
	formLocation = formLocationOption
	formLocationText = ""
ELSE
	formLocation = formLocationText
END IF

formSupervisorOption = Request.Form.Item("SupervisorOptions")
urlSupervisor=Request.QueryString("Supervisor")
if formSupervisorOption = "" then formSupervisorOption = urlSupervisor end if

formSupervisorText = Request.Form.Item("SupervisorText")
if trim(formSupervisorOption) <> "" THEN
    'response.Write Request.Form("SupervisorOptions")
    'response.End
    if instr(formSupervisorOption,"|")>0 then
        formSupervisor = split(formSupervisorOption,"|")(0)
	    formLicenseNumber = split(formSupervisorOption,"|")(1)
    end if
	formSupervisorText = ""
ELSE
	formSupervisor = formSupervisorText
END IF

'formLicenseNumberOption = Request.Form.Item("LicenseNumberOptions")
'urlLicenseNumber=Request.QueryString("LicenseNumber")
'if formLicenseNumberOption = "" then formLicenseNumberOption = urlLicenseNumber end if

'formLicenseNumberText = Request.Form.Item("LicenseNumberText")
'if trim(formLicenseNumberOption) <> "" THEN
'	formLicenseNumber = formLicenseNumberOption
'	formLicenseNumberText = ""
'ELSE
'	formLicenseNumber = formLicenseNumberText
'END IF

formChemicalSupplierOption = Request.Form.Item("ChemicalSupplierOptions")
urlChemicalSupplier=Request.QueryString("ChemicalSupplier")
if formChemicalSupplierOption = "" then formChemicalSupplierOption = urlChemicalSupplier end if

formChemicalSupplierText = Request.Form.Item("ChemicalSupplierText")
if trim(formChemicalSupplierOption) <> "" THEN
	formChemicalSupplier = formChemicalSupplierOption
	formChemicalSupplierText = ""
ELSE
	formChemicalSupplier = formChemicalSupplierText
END IF

formRecommendedByOption = Request.Form.Item("RecommendedByOptions")
urlRecommendedBy=Request.QueryString("RecommendedBy")
if formRecommendedByOption = "" then formRecommendedByOption = urlRecommendedBy end if

formRecommendedByText = Request.Form.Item("RecommendedByText")
if trim(formRecommendedByOption) <> "" THEN
	formRecommendedBy = formRecommendedByOption
	formRecommendedByText = ""
ELSE
	formRecommendedBy = formRecommendedByText
END IF

formApplicatorOption = Request.Form.Item("ApplicatorOptions")
urlApplicator=Request.QueryString("Applicator")
if formApplicatorOption = "" then formApplicatorOption = urlApplicator end if

formApplicatorText = Request.Form.Item("ApplicatorText")
if trim(formApplicatorOption) <> "" THEN
	formApplicators = split(formApplicatorOption,",")
	Dim sprayer
    For Each sprayer In formApplicators
		If InStr(sprayer,"|") > 0 Then
	    formApplicator = formApplicator + split(sprayer,"|")(0) + ";"
	    formApplicatorLicense = split(sprayer,"|")(1)
	    End If
    Next
ELSE
	formApplicator = formApplicatorText
END IF


' Call FormDataDump("True", False)

'formApplicatorLicenseOption = Request.Form.Item("ApplicatorLicenseOptions")
'urlApplicatorLicense=Request.QueryString("ApplicatorLicense")
'if formApplicatorLicenseOption = "" then formApplicatorLicenseOption = urlApplicatorLicense end if

'formApplicatorLicenseText = Request.Form.Item("ApplicatorLicenseText")
'if trim(formApplicatorLicenseOption) <> "" THEN
'	formApplicatorLicense = formApplicatorLicenseOption
'	formApplicatorLicenseText = ""
'ELSE
'	formApplicatorLicense = formApplicatorLicenseText
'END IF

formMethodID = Request.Form.Item("MethodID")
urlMethodID=Request.QueryString("MethodID")
if formMethodID = "" then formMethodID = urlMethodID end if

rem 7/18/06 making weather freeform, will require db modification to SprayRecord
'formWeatherID = Request.Form.Item("WeatherID")
'urlWeatherID=Request.QueryString("WeatherID")
'if formWeatherID = "" then formWeatherID = urlWeatherID end if

'formNewWeather = Request.Form.Item("NewWeather")
'if formWeatherID = "" AND formNewWeather <> "" THEN
'	formWeatherID = InsertWeather(formNewWeather)
'	formNewWeather = ""
'END IF

'formWeather = Request.Form.Item("Weather")
if Request.Form.Item("WeatherTemp")>"" then formTemp = Request.Form.Item("WeatherTemp")
if Request.Form.Item("WeatherWindSpd") >"" then formWindSpd = Request.Form.Item("WeatherWindSpd")
if Request.Form.Item("WeatherWindDir") > "" then formWindDir = Request.Form.Item("WeatherWindDir")
formWeather = formTemp & "F " & formWindSpd & "mph " & formWindDir

REM when adding spray data w/ more to follow only the first spray product is getting acres treated prefilled.
REM is this an issue??? kmiers 7/18/06
formAcresTreated = Request.Form.Item("AcresTreated")
urlAcresTreated=Request.QueryString("AcresTreated")
if formAcresTreated = "" then formAcresTreated = urlAcresTreated end if

formRateAcre = Request.Form.Item("RateAcre")

formAcresTreated1 = Request.Form.Item("AcresTreated1")
formRateAcre1 = Request.Form.Item("RateAcre1")
formIFPRating1 = Request.Form.Item("IFPRating1")
formAcresTreated2 = Request.Form.Item("AcresTreated2")
formRateAcre2 = Request.Form.Item("RateAcre2")
formIFPRating2 = Request.Form.Item("IFPRating2")
formAcresTreated3 = Request.Form.Item("AcresTreated3")
formRateAcre3 = Request.Form.Item("RateAcre3")
formIFPRating3 = Request.Form.Item("IFPRating3")
formAcresTreated4 = Request.Form.Item("AcresTreated4")
formRateAcre4 = Request.Form.Item("RateAcre4")
formIFPRating4 = Request.Form.Item("IFPRating4")
formAcresTreated5 = Request.Form.Item("AcresTreated5")
formRateAcre5 = Request.Form.Item("RateAcre5")
formIFPRating5 = Request.Form.Item("IFPRating5")
formTotalMaterialApplied = Request.Form.Item("TotalMaterialApplied")
formTotalMaterialApplied1 = Request.Form.Item("TotalMaterialApplied1")
formTotalMaterialApplied2 = Request.Form.Item("TotalMaterialApplied2")
formTotalMaterialApplied3 = Request.Form.Item("TotalMaterialApplied3")
formTotalMaterialApplied4 = Request.Form.Item("TotalMaterialApplied4")
formTotalMaterialApplied5 = Request.Form.Item("TotalMaterialApplied5")

arraySprayList = Split(Request.Form.Item("SprayListID"),"|")
IF IsArray(arraySprayList) THEN
	IF Ubound(arraySprayList) >= 0 THEN
		formSprayListID = arraySprayList(0)
	END IF
END IF

arraySprayList = Split(Request.Form.Item("SprayListID1"),"|")
IF IsArray(arraySprayList) THEN
	IF Ubound(arraySprayList) >= 0 THEN
		formSprayListID1 = arraySprayList(0)
	END IF
END IF

arraySprayList = Split(Request.Form.Item("SprayListID2"),"|")
IF IsArray(arraySprayList) THEN
	IF Ubound(arraySprayList) >= 0 THEN
		formSprayListID2 = arraySprayList(0)
	END IF
END IF

arraySprayList = Split(Request.Form.Item("SprayListID3"),"|")
IF IsArray(arraySprayList) THEN
	IF Ubound(arraySprayList) >= 0 THEN
		formSprayListID3 = arraySprayList(0)
	END IF
END IF

arraySprayList = Split(Request.Form.Item("SprayListID4"),"|")
IF IsArray(arraySprayList) THEN
	IF Ubound(arraySprayList) >= 0 THEN
		formSprayListID4 = arraySprayList(0)
	END IF
END IF

arraySprayList = Split(Request.Form.Item("SprayListID5"),"|")
IF IsArray(arraySprayList) THEN
	IF Ubound(arraySprayList) >= 0 THEN
		formSprayListID5 = arraySprayList(0)
	END IF
END IF


formIFPRating = Request.Form.Item("IFPRating")
formTargetID = Request.Form.Item("TargetID")
formTargetID1 = Request.Form.Item("TargetID1")
formTargetID2 = Request.Form.Item("TargetID2")
formTargetID3 = Request.Form.Item("TargetID3")
formTargetID4 = Request.Form.Item("TargetID4")
formTargetID5 = Request.Form.Item("TargetID5")
formHarvestDate = Request.Form.Item("HarvestDate")
formComments = Request.Form.Item("Comments")

if Request.Form.Item("NewProdID") <> 0 THEN
	formSprayListID = Request.Form.Item("NewProdID")
end IF

IF Request.QuerySTring("task") = "d" and urlSprayRecordID <> "" THEN
	DeleteSprayRecord(urlSprayRecordID)
	Response.Redirect("enterspraydata.asp")
END IF
IF Request.QuerySTring("task") = "activate" and urlSprayRecordID <> "" THEN
	ActivateSprayRecord(urlSprayRecordID)
	Response.Redirect("enterspraydata.asp")
END IF
IF Request.QuerySTring("task") = "deactivate" and urlSprayRecordID <> "" THEN
	DeActivateSprayRecord(urlSprayRecordID)
	Response.Redirect("enterspraydata.asp")
END IF

'Response.Write("<br>Are we saving spray data?")
'Response.Flush

DIM lSaveSprayData
lSaveSprayData = FALSE
'Form Was Submitted
'IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN
IF Request.Form.Item("insert") <> "" OR _
	Request.Form.Item("insert_n_save_spray") <> "" OR _
	Request.Form.Item("update") <> "" THEN

	IF Request.Form.Item("insert_n_save_spray") <> "" THEN
		lSaveSprayData = TRUE
	END IF

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("GrowerID"), "char","Grower", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("SprayStartDate"), "datetime","SprayStartDate", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("SprayEndDate"), "datetime","SprayEndDate", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("HarvestDate"), "datetime","Harvest Date", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("CropID"), "varchar","Crop", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF

	if Request.Form.Item("CropID") = 1 then 'bartletts.
'		Response.Write("<br>bartlett: " & Request.Form.Item("Bartlet"))
		if trim(Request.Form.Item("Bartlet")) = "" then
			errorFound = TRUE
			errorMessage = errorMessage + "<br>Bartlett selection (YES/NO) is required."
		end if
	end if

	IF NOT ValidateDatatype(Request.Form.Item("Bartlet"), "bit","Bartlet", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("StageID"), "int","Stage", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(formApplicator, "varchar","Applicator", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(formApplicatorLicense, "varchar","Applicator License", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(formSupervisor, "varchar","Supervisor", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(formLicenseNumber, "varchar","Supervisor License", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(formRecommendedBy, "varchar","Recommended By", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(formChemicalSupplier, "varchar","Chemical Supplier", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF


rem check to see if spray product has been selected.  will need to be modified to allow this to be not selected
rem as long as one of the other lines has selected product
	IF (NOT ValidateDatatype(formSprayListID, "int","Product", TRUE)) THEN
		if (NOT ValidateDatatype(formSprayListID1, "int","Product", TRUE)) then
		if (NOT ValidateDatatype(formSprayListID2, "int","Product", TRUE)) then
	    if (NOT ValidateDatatype(formSprayListID3, "int","Product", TRUE)) then
	    if (NOT ValidateDatatype(formSprayListID4, "int","Product", TRUE)) then
	    if (NOT ValidateDatatype(formSprayListID5, "int","Product", TRUE)) then
		errorFound = TRUE
		errorMessage = errorMessage + "<br>No Spray Products have been selected."
		end if
		end if
		end if
		end if
		end if
	END IF

	rem making weather freeform
'	IF NOT ValidateDatatype(Request.Form.Item("WeatherID"), "int","Weather", FALSE) THEN
'		errorFound = TRUE
'		errorMessage = errorMessage + "<br>" + tempErrorMessage
'	END IF
	IF NOT ValidateDatatype(formLocation, "nvarchar","Location", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("MethodID"), "int","Method", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("TargetID"), "int","Target", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("AcresTreated"), "float","Acres Treated", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("RateAcre"), "float","Rate/Acre", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("IFPRating"), "nvarchar","IFP Rating", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF

	IF	Request.Form.Item("insert") <> "" OR Request.Form.Item("insert_n_save_spray") <> "" THEN

		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID1 <> "" THEN
			IF NOT ValidateDatatype(Request.Form.Item("AcresTreated1"), "float","Acres Treated #2", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("RateAcre1"), "float","Rate/Acre #2", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("TargetID1"), "int","Target #2", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF

		END IF
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID2 <> "" THEN
			IF NOT ValidateDatatype(Request.Form.Item("AcresTreated2"), "float","Acres Treated #3", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("RateAcre2"), "float","Rate/Acre #3", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("TargetID2"), "int","Target #3", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
		END IF
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID3 <> "" THEN
			IF NOT ValidateDatatype(Request.Form.Item("AcresTreated3"), "float","Acres Treated #4", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("RateAcre3"), "float","Rate/Acre #4", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("TargetID3"), "int","Target #4", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
		END IF
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID4 <> "" THEN
			IF NOT ValidateDatatype(Request.Form.Item("AcresTreated4"), "float","Acres Treated #5", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("RateAcre4"), "float","Rate/Acre #5", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("TargetID4"), "int","Target #5", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
		END IF
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID5 <> "" THEN
			IF NOT ValidateDatatype(Request.Form.Item("AcresTreated5"), "float","Acres Treated #6", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("RateAcre5"), "float","Rate/Acre #6", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
			IF NOT ValidateDatatype(Request.Form.Item("TargetID5"), "int","Target #6", TRUE) THEN
				errorFound = TRUE
				errorMessage = errorMessage + "<br>" + tempErrorMessage
			END IF
		END IF
	END IF


	IF NOT ValidateDatatype(Request.Form.Item("Comments"), "nvarchar","Comments", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN
		urlSprayRecordID = UPDATESprayRecord(formSprayRecordID,formPackerID,formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated,formRateAcre,formSprayListID,formIFPRating,formTargetID,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)

        'new update spray record, weather is text (KILROY)
        '		urlSprayRecordID = UPDATESprayRecord2(formSprayRecordID,formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated,formRateAcre,formSprayListID,formIFPRating,formTargetID,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)

        '		Response.Redirect("enterspraydata.asp?success=1&AcresTreated=" &  formAcresTreated  & "&GrowerID=" & formGrowerID & "&SprayStartDate="& formSprayStartDate & "&StageID=" & formStageID & "&MethodID=" & formMethodID & "&Location=" & Server.URLEncode(formLocation) & "&SprayEndDate=" & formSprayEndDate  & "&CropID=" & formCropID & "&VarietyID=" & formVarietyID &  "&Applicator=" &  Server.URLEncode(formApplicator) & "&ApplicatorLicense=" &  formApplicatorLicense & "&Administrator=" &  formAdministrator & "&Supervisor=" &  formSupervisor & "&LicenseNumber=" &  formLicenseNumber& "&ChemicalSupplier=" &  formChemicalSupplier & "&RecommendedBy=" &  formRecommendedBy)
        '11/6/2006 dont save spray data: date, stage, location
		'Response.Redirect("enterspraydata.asp?success=1&AcresTreated=" &  formAcresTreated  & "&GrowerID=" & formGrowerID & "&MethodID=" & formMethodID & "&CropID=" & formCropID & "&VarietyID=" & formVarietyID &  "&Applicator=" &  Server.URLEncode(formApplicator) & "&ApplicatorLicense=" &  formApplicatorLicense & "&Administrator=" &  formAdministrator & "&Supervisor=" &  formSupervisor & "&LicenseNumber=" &  formLicenseNumber& "&ChemicalSupplier=" &  formChemicalSupplier & "&RecommendedBy=" &  formRecommendedBy)
		Response.Redirect("sprayrecords_list.asp#searchresults")

		'END UPDATE
	END IF

	'INSERT
	IF NOT errorFound AND (Request.Form.Item("insert") <> "" OR Request.Form.Item("insert_n_save_spray") <> "") THEN

		If lSaveSprayData Then
			'setup session array to keep spray info.

			Dim SprayArray(5,4), iRow, iColumn
'rem what is the for loop???
			SprayArray(0,0) = ""
			SprayArray(1,0) = ""
			SprayArray(2,0) = ""
			SprayArray(3,0) = ""
			SprayArray(4,0) = ""
			SprayArray(5,0) = ""

			i_counter = 0
			DO WHILE i_counter <= 5
				SprayArray(i_counter,1) = ""
				SprayArray(i_counter,2) = ""
				SprayArray(i_counter,3) = ""
				SprayArray(i_counter,4) = ""
				i_counter = i_counter + 1
			LOOP

			IF formSprayListID <> "" THEN
				SprayArray(0,0) = formSprayListID
				SprayArray(0,1) = formAcresTreated
				SprayArray(0,2) = formRateAcre
				SprayArray(0,3) = formIFPRating
				SprayArray(0,4) = formTargetID
			END IF

			IF formSprayListID1 <> "" THEN
				SprayArray(1,0) = formSprayListID1
				SprayArray(1,1) = formAcresTreated1
				SprayArray(1,2) = formRateAcre1
				SprayArray(1,3) = formIFPRating1
				SprayArray(1,4) = formTargetID1
			END IF
			IF formSprayListID2 <> "" THEN
				SprayArray(2,0) = formSprayListID2
				SprayArray(2,1) = formAcresTreated2
				SprayArray(2,2) = formRateAcre2
				SprayArray(2,3) = formIFPRating2
				SprayArray(2,4) = formTargetID2
			END IF
			IF formSprayListID3 <> "" THEN
				SprayArray(3,0) = formSprayListID3
				SprayArray(3,1) = formAcresTreated3
				SprayArray(3,2) = formRateAcre3
				SprayArray(3,3) = formIFPRating3
				SprayArray(3,4) = formTargetID3
			END IF
			IF formSprayListID4 <> "" THEN
				SprayArray(4,0) = formSprayListID4
				SprayArray(4,1) = formAcresTreated4
				SprayArray(4,2) = formRateAcre4
				SprayArray(4,3) = formIFPRating4
				SprayArray(4,4) = formTargetID4
			END IF
			IF formSprayListID5 <> "" THEN
				SprayArray(5,0) = formSprayListID5
				SprayArray(5,1) = formAcresTreated5
				SprayArray(5,2) = formRateAcre5
				SprayArray(5,3) = formIFPRating5
				SprayArray(5,4) = formTargetID5
			END IF
			session("SprayArray") = SprayArray
		Else
			'clear session array
			session("SprayArray") = ""
		End If

		Dim selectedTarget,formTargets,formTargetList,tmpElement

		'formTargetList(5) = {'formTargetID','formTargetID1',formTargetID2'}

		For Each tmpElement In Request.Form
			Response.Write("<br />ELEMENT: " + tmpElement + " <br>")
			If InStr(tmpElement,"TargetID") > 0 Then
				formTargets = Split(Request.Form(tmpElement),",")
				If UBound(formTargets) > 0 Then
					elementPosition = Right(tmpElement,1)
					If Not IsNumeric(elementPosition) Then
						elementPosition = ""
					End If
					Response.Write("Element Pos: " + elementPosition)
					formAcresTreated = Request.Form("AcresTreated" + elementPosition)
					formRateAcre = Request.Form("RateAcre" + elementPosition)
					tmpSprayList = Split(Request.Form("SprayListID" + elementPosition),"|")
					If UBound(tmpSprayList) > 0 Then
						formSprayListID = tmpSprayList(0)
					Else
						formSprayListID = "ABC"
					End If
					formIFPRating = Request.Form("FIFPRating" + elementPosition)

					urlSprayRecordID = InsertSprayRecord2(formPackerID,formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated,formRateAcre,formSprayListID,formIFPRating,formTargets,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)
					'For Each selectedTarget In formTargets
					'Next
				End If
			End If
		Next




'		Response.Redirect("enterspraydata.asp?success=1&AcresTreated=" &  formAcresTreated  & "&GrowerID=" & formGrowerID & "&SprayStartDate="& formSprayStartDate & "&StageID=" & formStageID & "&MethodID=" & formMethodID & "&Location=" & replace(formLocation,"&","%26") & "&SprayEndDate=" & formSprayEndDate  & "&CropID=" & formCropID& "&VarietyID=" & formVarietyID &  "&Applicator=" &  formApplicator & "&ApplicatorLicense=" &  formApplicatorLicense & "&Administrator=" &  formAdministrator & "&Supervisor=" &  formSupervisor & "&LicenseNumber=" &  formLicenseNumber& "&ChemicalSupplier=" &  formChemicalSupplier & "&RecommendedBy=" &  formRecommendedBy)
rem on insert & save spray data, don't carry over the AcresTreated
		If lSaveSprayData Then
			Response.Redirect("enterspraydata.asp?success=1&GrowerID=" & formGrowerID & "&SprayStartDate="& formSprayStartDate & "&StageID=" & formStageID & "&MethodID=" & formMethodID & "&Location=" & replace(formLocation,"&","%26") & "&SprayEndDate=" & formSprayEndDate  & "&CropID=" & formCropID& "&VarietyID=" & formVarietyID &  "&Applicator=" &  formApplicator & "&ApplicatorLicense=" &  formApplicatorLicense & "&Administrator=" &  formAdministrator & "&Supervisor=" &  formSupervisor & "&LicenseNumber=" &  formLicenseNumber& "&ChemicalSupplier=" &  formChemicalSupplier & "&RecommendedBy=" &  formRecommendedBy)
		Else
			Response.Redirect("enterspraydata.asp?success=1&GrowerID=" & formGrowerID & "&MethodID=" & formMethodID & "&CropID=" & formCropID& "&VarietyID=" & formVarietyID &  "&Applicator=" &  formApplicator & "&ApplicatorLicense=" &  formApplicatorLicense & "&Administrator=" &  formAdministrator & "&Supervisor=" &  formSupervisor & "&LicenseNumber=" &  formLicenseNumber& "&ChemicalSupplier=" &  formChemicalSupplier & "&RecommendedBy=" &  formRecommendedBy)
		End If

	END IF 'insert
ELSEIF formSprayRecordID = 0 THEN
	'formSprayDate = month(now()) & "/" & day(now()) & "/" & year(now())
	'formHarvestDate = month(now()) & "/" & day(now()) & "/" & year(now())
END IF 'form submitted

IF formSprayRecordID  <>  0 and not errorFound and formCropID = 0 THEN
	set rs = GetSprayRecordByID(formSprayRecordID)
	IF NOT rs.eof THEN

		formPackerID = rs.Fields("PackerID")
		formGrowerID = rs.Fields("GrowerID")
		session("growerid")=formGrowerID

		formSprayStartDate = rs.Fields("SprayStartDate")
		formTimeFinishedSpraying = rs.Fields("TimeFinishedSpraying")
		formSprayEndDate = rs.Fields("SprayEndDate")
		formCropID = rs.Fields("CropID")
		formVarietyID = rs.Fields("VarietyID1")
		formVarietyID = listAppend(formVarietyID,rs.Fields("VarietyID2"))
		formVarietyID = listAppend(formVarietyID,rs.Fields("VarietyID3"))
		formVarietyID = listAppend(formVarietyID,rs.Fields("VarietyID4"))
		formBartlet = rs.Fields("Bartlet")
		formStageID = rs.Fields("StageID")
REM KILROY, any action needed here???
		formLocation = rs.Fields("Location")
		formMethodID = rs.Fields("MethodID")
'		formWeatherID = rs.Fields("WeatherID")
'		formWeather = rs.Fields("Weather") '(KILROY)
		formApplicator = rs.Fields("Applicator")
		formApplicatorLicense = rs.Fields("ApplicatorLicense")
		formAcresTreated = rs.Fields("AcresTreated")
		formRateAcre = rs.Fields("RateAcre")
		formTotalMaterialApplied = rs.Fields("TotalMaterialApplied")
		formSprayListID = rs.Fields("ProductID")
		formIFPRating = rs.Fields("IFPRating")
		formTargetID = rs.Fields("TargetID")
		formHarvestDate = rs.Fields("HarvestDate")
		formComments = rs.Fields("Comments")
		formAdministrator = rs.Fields("Administrator")
		formSupervisor = rs.Fields("Supervisor")
		formLicenseNumber = rs.Fields("LicenseNumber")
		formChemicalSupplier = rs.Fields("ChemicalSupplier")
		formRecommendedBy = rs.Fields("RecommendedBy")
		thisYearID = rs.Fields("SprayYearID")

		'response.write thisSprayYear
		'response.end


        formWeather = rs.Fields("Weather")
        'response.Write formweather
        'response.End
        if ubound(split(formWeather," ")) >= 0 then
            formTemp = replace(split(formWeather," ")(0),"F","")
            formWindSpd = replace(split(formWeather," ")(1),"mph","")
            formWindDir = split(formWeather," ")(2)
        end if

		formTotalMaterialApplied = Round(rs.Fields("RateAcre") * rs.Fields("AcresTreated"),2)
	END IF
END IF%>
<html>
<head>
	<title><%=Application("BusinessName")%>&nbsp;-&nbsp;<%=Application("ProgramName")%>&nbsp;-&nbsp;Enter Spray Data</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
	<script language="JavaScript" src="datepicker.js"></script>
	<script type="text/javascript">
		function PegarEnter(e) {
			//alert("PegarEnter");
			//above alert does get called, but i don't see any of the following alerts.  what's up???
			//where if ever is this called???
			if(document.all) {
				var Tecla = event.keyCode;
				}
			else if(document.layers) {
				var Tecla = e.which;
				}
			if(Tecla == 13){
				alert("13 return false");
				return false;
				}
			}
			document.onkeypress = PegarEnter;
	</script>

	<!--#include file="SprayDataCalcs.js"-->

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><img src="images/spacer.gif" height="4" width="1" border="0"><br>
<h1>> Enter Spray Data</h1><br><img src="images/spacer.gif" height="4" width="1" border="0"></td></tr></table><br />

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="ffffff" class="bodytext"><br>
<% IF Request.QueryString("success") = 1 THEN%>

<br>
<font color="990000"><strong>Success! Please see record below.</strong></font>
<%END IF %>
<%'if not formAdding then%>
<!--- <a href="enterspraydata.asp?adding=1"><strong>Add SprayRecord</strong></a><br><br> ---></td>
<%'end if%>
</tr>
<tr>
<td colspan="2" class="bodytext">
<%'if formAdding then %>
<form action="enterspraydata.asp" method="post" name="frmadd">
<input type="hidden" name="RefreshData" value="0">
<input type="hidden" name="NewProdID" value="0">
<input type="hidden" name="Adding" value="1">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td>
<table width="100%" border="1" cellpadding="2" cellspacing="0">
<tr bgcolor="#cccccc">
<td align="left" class="bodytext" colspan="3"><strong><%if urlSprayRecordID <> 0 then%>
EDITING RECORD #<%=formSprayRecordID%>
<%else%>
ADDING RECORD
<%end if%></strong>
</td>
</tr>
<tr bgcolor=#eeeeee>
<td align="left" class="bodytext"><strong>* indicates required field</strong><br>
</td><td colspan="2"><strong>dates must be mm/dd/yyyy</strong><br>
</td>
</tr>
<% if errorFound then%>
<tr>

<td class="bodytext" colspan="2"><font color="red"><% =errorMessage%></font></td><td>&nbsp;</td>
</tr>
<% End If %>

<input type="hidden" value="<% =urlSprayRecordID%>" name="SprayRecordID">
<input type="hidden" value="0" name="changedGrower">

<tr valign="top">


<td>

    <table cellpadding=0 cellspacing=10 border=0><tr valign=top>


<td>
    <span class="subtitle"><label for="GrowerName">*Grower</label>:</span><br><span class="bodytext">
   <%if false then %>

        <%
	        set rsSelect = GetActiveGrowers()
        %>
        <SELECT name="GrowerID" onchange="javascript:form.submit();" style="background-color:beige">
        <%IF session("growerid") = 0 THEN%>
        <option value="">---SELECT A GROWER---</option>
        <%
        END IF
        IF not rsSelect.EOF THEN
        DO WHILE not rsSelect.eof
        %>
        <option value="<%	response.write(rsSelect.Fields("GrowerID"))%>"
        <%if trim(formGrowerID) = trim(rsSelect.Fields("GrowerID")) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("GrowerName")%></option>
        <%
        rsSelect.MoveNext
        LOOP
        END IF
        %>
        </select></span>
    <%
    elseif true then

        dim roleConn: set roleConn = Connect()
        dim rsRoles
	if listContains("1", session("accessid")) then
		set rsRoles = GetActiveGrowers()
	else
		set rsRoles = conn.execute("exec growerunit$bygrower " & session("growerid"))
	end if

        with response

        .Write "<input type=hidden name=changerole />"
        .Write "<select name=growerrole onchange=form.changerole.value=1;form.submit(); style=background-color:beige;>"
        do until rsRoles.eof

            .Write "<option value="""&rsRoles("growerid")&"|"&rsRoles("growername")&""""
            if cint(session("growerid"))=rsRoles("growerid") then .Write " selected"
            .Write ">"
            .Write rsRoles("growername")
            .write "</option>"

        rsRoles.movenext
        loop
        .Write "</select>"

        end with
       	EndConnect(roleConn)

    %>
    <%end if 'else%>
        <input type=hidden name=GrowerID value="<%=session("GrowerID")%>" />
    <%'end if %>

    </td>
<td>
    <span class="subtitle"><label for="CropID">*Crop</label>:</span><br>
    <Select name="CropID" onchange="javascript:form.submit();"><option value=0 style="background-color:beige"></option>
    <%

        dim rsdefcrop: set rsdefcrop = conn.execute("select cropid from growercrop where growerid="&session("growerid")&" and isdefault=1")
        dim defcrop: if not rsdefcrop.eof then defcrop = rsdefcrop.collect(0)

	    set rsSelect = GetActiveCrops()
	    thisCropID = 0
	    i = 0
	    IF not rsSelect.EOF THEN
		    'thisCropID = rsSelect.Fields("CropID")
		    DO WHILE not rsSelect.eof
			    i = i + 1
    %>
    <option value="<%=rsSelect.Fields("CropID")%>"
    <%
    if formCropID="" and rsSelect.Fields("CropID")=defcrop then
        thisCropID = rsSelect.Fields("CropID")
        formCropID = rsSelect.Fields("CropID")
        response.write("selected")
    elseif listContains(formCropID, trim(rsSelect.Fields("CropID"))) then
        thisCropID = rsSelect.Fields("CropID")
        response.write("selected")
    end if
    %> style="background-color:beige"><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Crop")%></option>
    <%
			    rsSelect.MoveNext
		    LOOP
	    END IF
    %>
    </SELECT><br>

    </td>
    <td>

    <span class="subtitle"><label for="PackerName">Packer</label>:</span><br><span class="bodytext">
    <%

        if formcropid="" then formcropid=0
        dim rsdefpack: set rsdefpack = conn.execute("select packerid from growercrop where growerid="&session("growerid")&" and cropid="&formcropid)
        dim defpack: if not rsdefpack.eof then defpack = rsdefpack.collect(0)

	    set rsSelect = GetActivePackers()
    %>
    <SELECT name="PackerID" style="background-color:beige;width:150px;">
    <%IF session("packerid") = 0 THEN%>
    <option value="">Other/Unspecified</option>
    <%
    END IF
    IF not rsSelect.EOF THEN
    DO WHILE not rsSelect.eof
    %>
    <option value="<%	response.write(rsSelect.Fields("PackerID"))%>"
    <%
    if rsSelect.Fields("PackerID")=defpack then
        response.write("selected")
    elseif trim(formPackerID) = trim(rsSelect.Fields("PackerID")) then
        response.write("selected")
    end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=string(6-len(rsSelect.Fields("PackerNumber")),"0")&rsSelect.Fields("PackerNumber")%></option>
    <%
    rsSelect.MoveNext
    LOOP
    END IF
    %>
    </select></span>

    </td>

    </tr>

    <%if thisCropID = 1 then%>
    <tr><td colspan=3><span class="subtitle"><label for="Bartlet">Bartlett:</label></span> <span class="bodytext"><span class="bodytext"><input type="radio" value="1" name="Bartlet" <% if ListContains(formBartlet,"True") OR ListContains(formBartlet,"1")  THEN %>Checked<% END IF %> style="background-color:beige">YES <input type="radio" value="0" name="Bartlet" <% if ListContains(formBartlet,"False") OR ListContains(formBartlet,"0")  THEN %>Checked<% END IF %> style="background-color:beige">NO</span></td></tr>
    <%end if%>

    </table>

</td>

<td>

    <table border=0 cellspacing=0 cellpadding=10>
    <tr valign=top>
    <td>
        <table border=0 cellspacing=0 cellpadding=0><tr><td><span class="subtitle"><label for="HarvestDate">Harvest Date</label>:</span><br /><input type="text" value="<%=formHarvestDate%>" name="HarvestDate" size="15" maxlength="21" class="bodytext"></td><td><br /><a href="javascript:show_calendar('frmadd.HarvestDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table>
    </td>

    </tr>
    </table>

</td>
</tr>

<%
	set rsSelect = GetActiveVarietiesByCropID(thisCropID)
    i = 0
    IF not rsSelect.EOF THEN%>
    <tr><td colspan=2>
    <table><tr><td class="smalltext" colspan="2"><span class="subtitle">Choose up to 4 varieties:</span></td></tr>
    <tr><td valign=top>
    <%
    DO WHILE not rsSelect.eof
    i = i + 1
    %>
    <input type="checkbox" name="VarietyID" value="<%=rsSelect.Fields("VarietyID")%>"<%if listContains(formVarietyID, trim(rsSelect.Fields("VarietyID"))) then response.write("checked") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Variety")%></option>
    <%
    'if (i = 7) then
	'    response.write ("</td><td valign=top>")
    'end if
    rsSelect.MoveNext
    LOOP%>
    </td></tr></table>

    </td></tr>
    <%
    end if
    %>


<tr>

<td valign="top" nowrap>

    <table width="100%" border="0" cellpadding="3" cellspacing="0" border=0>
    <tr valign="top">
    <td width="50%">
        <span class="subtitle"><label for="SprayStartDate">* Spray Start Date</label>:</span><br /><i style="font-size:7pt;"> </i>

        <table border=0 cellspacing=0 cellpadding=0>
        <tr><td><input type="text" value="<%=formSprayStartDate%>" name="SprayStartDate" size="15" maxlength="21" class="bodytext" onfocus="frmadd.SprayEndDate.value=this.value;" style="background-color:beige"></td>
        <td><a href="javascript:show_calendar('frmadd.SprayStartDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr>
        </table>

        <br />

        <span class="subtitle"><label for="SprayEndDate">Spray End Date</label>:</span><br /><i style="font-size:7pt;">(If different than start date.)</i><br>
        <table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=formSprayEndDate%>" name="SprayEndDate" size="15" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmadd.SprayEndDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td>
        </tr></table></span>

    </td>
    <td>

        <span class="subtitle"><label for="Weather">Weather</label>:</span><br>
        <span class="bodytext">

        <table>
        <tr><td>Temperature:</td><td><input type="text" value="<%=RemoveQuotes(formTemp)%>" name="WeatherTemp"  class="bodytext" size="2" maxlength="3"> &deg;F</td></tr>
        <tr><td>Wind Speed:</td><td><input type="text" value="<%=RemoveQuotes(formWindSpd)%>" name="WeatherWindSpd"  class="bodytext" size="2" maxlength="3"> mph</td></tr>
        <tr><td>Wind Dir:</td><td><select name=WeatherWindDir style="font-size:9pt;">
        <option></option>
        <option <%=so("N",formWindDir)%>>N</option><option <%=so("NNE",formWindDir)%>>NNE</option><option <%=so("ENE",formWindDir)%>>ENE</option>
        <option <%=so("E",formWindDir)%>>E</option><option <%=so("ESE",formWindDir)%>>ESE</option><option <%=so("SSE",formWindDir)%>>SSE</option>
        <option <%=so("S",formWindDir)%>>S</option><option <%=so("WSW",formWindDir)%>>WSW</option><option <%=so("SSW",formWindDir)%>>SSW</option>
        <option <%=so("W",formWindDir)%>>W</option><option <%=so("WNW",formWindDir)%>>WNW</option><option <%=so("NNW",formWindDir)%>>NNW</option>
        </select></td></tr>
        </table>

        </span>


    </td>
    </tr>

    <tr><td colspan=2>

        <br><span class="subtitle">Time of Application </span> <!-- hh:mm [AM/PM] -->

        <br><i style="font-size:7pt;">(Time Required only if Field or Central Posting Report desired.)</i><br>

<input type="hidden" value="<%=formTimeFinishedSpraying%>" name="TimeFinishedSpraying" /> <!-- size="15" maxlength="21" class="bodytext" -->

<select name=SelectTimeFinishedSpraying onchange="this.form.TimeFinishedSpraying.value=this[this.selectedIndex].text;//alert(this.form.TimeFinishedSpraying.value);"><option></option>
<%

'# if formTimeFinishedSpraying="" then formTimeFinishedSpraying="12:30 PM"

dim time: for time=1 to 24

response.write "<option style=text-align:right;>"
if time<=12 then
	response.write time & ":00"
else
	response.write time-12 & ":00"
end if
if time<12 or time=24 then
	response.write " AM"
else
	response.write " PM"
end if
response.write "</option>"


response.write "<option style=text-align:right;>"
if time<=12 then
	response.write time & ":30"
else
	response.write time-12 & ":30"
end if
if time<12 or time=24 then
	response.write " AM"
else
	response.write " PM"
end if
response.write "</option>"

next
%>

</select>


    </td></tr>

   <!--
    <tr><td>
    <br><span class="subtitle">
    <span class="subtitle"><label for="HarvestDate">Harvest Date</label>:</span><table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=formHarvestDate%>" name="HarvestDate" size="15" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmadd.HarvestDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table>

    </td>
    <td></td>
   </tr>
   -->

   </table>

</td>

<td valign="top">

            <table border="0" cellpadding="5" cellspacing="0">
            <tr>
            <td valign="top" colspan=2>
            <span class="subtitle">* Supervisor</span>:<br>
            <select name="SupervisorOptions" class="bodytext" style="width:300px;background-color:beige">
            <option value=""></option>
            <%
            set rsSelect2 = conn.execute("exec growersupervisor$list " & session("growerid"))
	            IF not rsSelect2.EOF THEN
		            DO WHILE not rsSelect2.eof
			            if rsselect2("active") then
			            response.write("<option value='" & rsSelect2.Fields("Supervisor") & "|" & rsSelect2.Fields("LicenseNo") & "'")
			            IF rsSelect2.Fields("Supervisor") = RemoveQuotes(formSupervisor) THEN
				            response.write("selected")
				            formSupervisor = ""
			            END if
			            response.write (">" & rsSelect2.Fields("Supervisor") & " (" & rsSelect2.Fields("LicenseNo") & ")</option>")
                        end if
			            rsSelect2.movenext
		            LOOP
	            END IF
            %>
            </select>
            </td>
            </tr>

            <tr>
            <td valign="top" colspan=2><br><span class="subtitle">* Applicator</span>:<br>
			<select name="ApplicatorOptions" multiple="multiple" size="4" class="bodytext" style="width:300px;background-color:beige">
            <%
            set rsSelect2 = conn.execute("exec growerapplicator$list " & session("growerid"))
	            IF not rsSelect2.EOF THEN
		            DO WHILE not rsSelect2.eof
		            if rsselect2("active") then
			            response.write("<option value='" & rsSelect2.Fields("Applicator") & "|" & rsSelect2.Fields("LicenseNo") & "'")
			            IF rsSelect2.Fields("Applicator") = RemoveQuotes(formApplicator) THEN
				            response.write("selected")
				            formApplicator = ""
			            END if
			            response.write (">" & rsSelect2.Fields("Applicator") & " (" & rsSelect2.Fields("LicenseNo") & ")</option>")
			        end if
			            rsSelect2.movenext
		            LOOP
	            END IF
            %>
            </select>
            </td>
            </tr>

            <tr>
            <td>&nbsp;&nbsp;<br>
            <span class="subtitle">* Chemical Supplier</span>:<br><select name="ChemicalSupplierOptions" class="bodytext" style="width:150px;background-color:beige">
            <option value=""></option>
            <%
            set rsSelect2 = conn.execute("exec growersupplier$list " & session("growerid"))
	            IF not rsSelect2.EOF THEN
		            DO WHILE not rsSelect2.eof
		            if rsselect2("active") then
			            response.write("<option value='" & rsSelect2.Fields("Supplier") & "'")
			            IF rsSelect2.Fields("Supplier") = RemoveQuotes(formChemicalSupplier) THEN
				            response.write("selected")
				            formChemicalSupplier = ""
			            END if
			            response.write (">" & rsSelect2.Fields("Supplier") & "</option>")
                    end if
			            rsSelect2.movenext
		            LOOP
	            END IF
            %>
            </select></td>
            <td><br>
            <span class="subtitle">* Recommended By</span>:<br><select name="RecommendedByOptions" class="bodytext" style="width:150px;background-color:beige">
            <option value=""></option>
            <%
            set rsSelect2 = conn.execute("exec growerreferrer$list " & session("growerid"))
	            IF not rsSelect2.EOF THEN
		            DO WHILE not rsSelect2.eof
		            if rsselect2("active") then
			            response.write("<option value='" & rsSelect2.Fields("referrer") & "'")
			            IF rsSelect2.Fields("referrer") = RemoveQuotes(formRecommendedBy) THEN
				            response.write("selected")
				            formRecommendedBy = ""
			            END if
			            response.write (">" & rsSelect2.Fields("referrer") & "</option>")
			        end if
			            rsSelect2.movenext
		            LOOP
	            END IF
            %>
            </select>
            </td>
            </tr>
            </table>

</td>

</tr>

<tr>

<td valign="top" colspan="3">


</td>

</tr>

<tr>
<td valign="top" colspan="3">
<table align=center width="80%" cellpadding=3 cellspacing=0>
<tr valign=top align=center>

<td><span class="subtitle"><label for="LocationOptions">*Location</label>:</span><br>

<span class="bodytext"><select name="LocationOptions" class="bodytext" style="background-color:beige" size=5 multiple>
<option value=""></option>
<%

	set rsSelect2 = GetGrowersLocationsByGrowerID(formGrowerID)
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("<option value='" & rsSelect2.Fields("GLoc_Location") & "'")
			IF rsSelect2.Fields("GLoc_Location") = RemoveQuotes(formLocation) THEN
				response.write("selected")
				formLocation = ""
			END if
			response.write (">" & rsSelect2.Fields("GLoc_Location") & "</option>")
			rsSelect2.movenext
		LOOP
	END IF

%>
</select>
<br>&nbsp;<br>&nbsp;<!--
require them to have locations setup in advance... kim miers 7/16/2006.
<span class="subtitle">Enter New Location:</span><br>
<input type="text" value="<%=RemoveQuotes(formLocation)%>" name="LocationText"  class="bodytext" size="15" maxlength="50"></span>
-->
</td>
<td valign="top">
<span class="subtitle"><label for="MethodID">*Method</label>:</span><br>
<%
	set rsSelect = GetActiveMethods()
%>
<SELECT name="MethodID" style="background-color:beige">
<option value=""></option>
<%
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof
%>
<option value="<%=rsSelect.Fields("MethodID")%>"<%if trim(formMethodID) = trim(rsSelect.Fields("MethodID")) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Method")%></option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select>
</td>
<td valign="top"><span class="subtitle"><label for="StageID">*Stage</label>:</span><br>
<%
	set rsSelect = GetActiveStages()
%>
<SELECT name="StageID" style="background-color:beige">
<option value=""></option>
<%
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof
%>
<option value="<%=rsSelect.Fields("StageID")%>"<%if trim(formStageID) = trim(rsSelect.Fields("StageID")) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Stage")%></option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select>
</td>
</tr>
</table>
</td>
</tr>

<tr><td valign="top" colspan="3">
<%
	Dim formMaxAppUse,formMaxAppSeason,formUnits
	'set rsSelect = GetActiveSprayListByCropID(thisCropID)
	set rsSelect = GetActiveSprayListByCropYear(thisCropID,thisYearID)

%>

<table width="100%" border="1" cellpadding="3" cellspacing="0">
<tr>
<td>&nbsp;</td>
<td>

<% IF session("growerid") = 0 THEN
rem window.open() was pu_spraylist_list.asp
%>
<a href="javascript:void(0);" onclick="window.open('spraylist_list.asp?Action=AddOnly','add','width=550,height=550,scrollbars=no,resizable=yes');" class="bodytext"><strong>Add New Product</strong></a>
<br>
<% END IF %><span class="subtitle"><label for="SprayListID">Product Name and Formulation</label>:</span>
</td>
<td bgcolor="cccccc" valign="bottom"><strong>Unit</strong></td>
<td bgcolor="cccccc" valign="bottom"><strong>Max Rate Use</strong></td>
<td bgcolor="cccccc" valign="bottom"><strong>Max Season</strong></td>
<td valign="bottom"><strong>Acres Treat</strong><br># only</td>
<td valign="bottom"><strong>Rate Acre</strong><br># only</td>
<td valign="bottom"><strong>Total Applied</strong><br># only</td>
<td bgcolor="cccccc" valign="bottom"><strong>Target</strong></td>
<td bgcolor="cccccc" valign="bottom"><strong>Recalc</strong></td>
</tr>

<tr>
<td class="bodytext">1)</td>
<td>*<SELECT name="SprayListID" style="background-color:beige" onchange="javascript: displaySprayListData();">
<option value="|||">---SELECT A PRODUCT---</option>
<%

	REM setup vars from session SprayArray
	IF IsArray(session("SprayArray")) THEN
		localSprayArray = session("SprayArray")
		local_SprayID0 = localSprayArray(0,0)
		local_SprayID1 = localSprayArray(1,0)
		local_SprayID2 = localSprayArray(2,0)
		local_SprayID3 = localSprayArray(3,0)
		local_SprayID4 = localSprayArray(4,0)
		local_SprayID5 = localSprayArray(5,0)
	END IF

dim l_TargetsHaveArray
l_TargetsHaveArray = false
If IsArray(localSprayArray) Then
	cKilroy = "-"
	l_TargetsHaveArray = true
End If
dim l_match
l_match = false


IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID) = trim(rsSelect.Fields("SprayListID")) or local_SprayID0 = trim(rsSelect.Fields("SprayListID")) then
	response.write("selected")
	formMaxAppUse =  rsSelect.Fields("MaxUseApp")
	formMaxAppSeason = rsSelect.Fields("MaxUseSeason")
	formUnits = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF

if l_TargetsHaveArray then
	if localSprayArray(0,2) > 0 then
		formRateAcre = trim(CStr(localSprayArray(0,2)))
	end if
end if
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units" size="2" readonly value="<%=formUnits%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse" size="2" readonly value="<%=formMaxAppUse%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason" value="<%=formMaxAppSeason%>" size="2" readonly style="border: 0px;"></td>
<td><span class="bodytext"><input type="text" style="background-color:beige" name="AcresTreated" value="<%=formAcresTreated%>" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal();prefilAcres();"></span></td>
<td><span class="bodytext"><input type="text" value="<%=formRateAcre%>" style="background-color:beige" name="RateAcre" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal();"></span></td>
<td nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied%>" style="background-color:beige" name="TotalMaterialApplied" size="5" maxlength="11" class="bodytext" onchange="javascript:document.frmadd.RateAcre.value=0;calculateTotal();"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating%>" name="IFPRating"  class="bodytext" size="8" maxlength="50"> --->

<%
	set rsSelectTarget = GetActiveTarget()
%>
<select name="TargetID" multiple="multiple" size="5" style="background-color:beige;">
<!-- <option value="">---TARGET---</option> -->
<%
IF not rsSelectTarget.EOF THEN
l_match = false
DO WHILE not rsSelectTarget.eof
	if l_TargetsHaveArray then
		if trim(CStr(rsSelectTarget.Fields("TargetID"))) = trim(CStr(localSprayArray(0,4))) or trim(formTargetID) = trim(rsSelectTarget.Fields("TargetID")) then
			l_match = true
		else
			l_match = false
		end if
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if l_match then response.write(" selected ") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
	else
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if trim(formTargetID) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target") %></option>
<%
	end if
	rsSelectTarget.MoveNext
LOOP
END IF
%>
</select>
</td>
<td><input type=button value=Recalc onclick="calculateTotal();" /></td>
</tr>


<!--- ABILITY TO ENTER 5 MORE RECORDS ON INSERT --->
<%IF urlSprayRecordID=0 and thisCropID>0 THEN%>

<!--- ADDITIONAL RECORD 1 --->
<tr>
<td class="bodytext">2)</td>

<td><SELECT name="SprayListID1" style="background-color:beige" onchange="javascript: displaySprayListData1();">
<option value="|||">---SELECT A PRODUCT---</option>
<%

IF rsSelect.EOF THEN
rsSelect.MoveFirst
DO WHILE not rsSelect.eof
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID1) = trim(rsSelect.Fields("SprayListID")) or local_SprayID1 = trim(rsSelect.Fields("SprayListID")) then
	response.write("selected")
	formMaxAppUse1 =  rsSelect.Fields("MaxUseApp")
	formMaxAppSeason1 = rsSelect.Fields("MaxUseSeason")
	formUnits1 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF

if l_TargetsHaveArray then
	if IsNumeric(localSprayArray(1,2)) then
		formRateAcre1 = trim(CStr(localSprayArray(1,2)))
	end if
end if
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units1" size="2" readonly value="<%=formUnits1%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse1" size="2" readonly value="<%=formMaxAppUse1%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason1" value="<%=formMaxAppSeason1%>" size="2" readonly style="border: 0px;"></td>
<td ><span class="bodytext"><input type="text" style="background-color:beige" name="AcresTreated1" value="<%=formAcresTreated1%>" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal1();"></span></td>
<td ><span class="bodytext"><input type="text" value="<%=formRateAcre1%>" style="background-color:beige" name="RateAcre1" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal1();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied1%>" style="background-color:beige" name="TotalMaterialApplied1" size="5" maxlength="11" class="bodytext" onchange="javascript:document.frmadd.RateAcre1.value=0;calculateTotal1();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating1%>" name="IFPRating1"  class="bodytext" size="8" maxlength="50"> --->
<select name="TargetID1" multiple="multiple" size="5" style="background-color:beige;">
<!-- <option value="">---TARGET---</option> -->
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
l_match = false
DO WHILE not rsSelectTarget.eof
	if l_TargetsHaveArray then
		if trim(CStr(rsSelectTarget.Fields("TargetID"))) = trim(CStr(localSprayArray(1,4))) or trim(formTargetID1) = trim(rsSelectTarget.Fields("TargetID")) then
			l_match = true
		else
			l_match = false
		end if
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if l_match then response.write(" selected ") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
	else
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if trim(formTargetID1) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target") %></option>
<%
	end if
	rsSelectTarget.MoveNext
LOOP
END IF
%>
</select></td>
<td><input type=button value=Recalc onclick="calculateTotal();" /></td>
</tr>


<!--- ADDITIONAL RECORD 2 --->
<tr>
<td class="bodytext">3)</td>

<td><SELECT name="SprayListID2" style="background-color:beige" onchange="javascript: displaySprayListData2();">
<option value="|||">---SELECT A PRODUCT---</option>
<%

IF rsSelect.EOF THEN
rsSelect.MoveFirst
DO WHILE not rsSelect.eof
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID2) = trim(rsSelect.Fields("SprayListID")) or local_SprayID2 = trim(rsSelect.Fields("SprayListID")) then
	response.write("selected")
	formMaxAppUse2 =  rsSelect.Fields("MaxUseApp")
	formMaxAppSeason2 = rsSelect.Fields("MaxUseSeason")
	formUnits2 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF

if l_TargetsHaveArray then
	if IsNumeric(localSprayArray(2,2)) then
		formRateAcre2 = trim(CStr(localSprayArray(2,2)))
	end if
end if
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units2" size="2" readonly value="<%=formUnits2%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse2" size="2" readonly value="<%=formMaxAppUse2%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason2" value="<%=formMaxAppSeason2%>" size="2" readonly style="border: 0px;"></td>
<td ><span class="bodytext"><input type="text" style="background-color:beige" name="AcresTreated2" value="<%=formAcresTreated2%>" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal2();"></span></td>
<td ><span class="bodytext"><input type="text" value="<%=formRateAcre2%>" style="background-color:beige" name="RateAcre2" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal2();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied2%>" style="background-color:beige" name="TotalMaterialApplied2" size="5" maxlength="11" class="bodytext" onchange="javascript:document.frmadd.RateAcre2.value=0;calculateTotal2();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating2%>" name="IFPRating2"  class="bodytext" size="8" maxlength="50"> --->
<select name="TargetID2" multiple="multiple" size="5" style="background-color:beige;">
<!-- <option value="">---TARGET---</option> -->
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
l_match = false
DO WHILE not rsSelectTarget.eof
	if l_TargetsHaveArray then
		if trim(CStr(rsSelectTarget.Fields("TargetID"))) = trim(CStr(localSprayArray(2,4))) or trim(formTargetID2) = trim(rsSelectTarget.Fields("TargetID")) then
			l_match = true
		else
			l_match = false
		end if
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if l_match then response.write(" selected ") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
	else
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if trim(formTargetID2) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target") %></option>
<%
	end if
	rsSelectTarget.MoveNext
LOOP
END IF
%>
</select></td>
<td><input type=button value=Recalc onclick="calculateTotal();" /></td>
</tr>


<!--- ADDITIONAL RECORD 3 --->
<tr>
<td class="bodytext">4)</td>

<td><SELECT name="SprayListID3" style="background-color:beige" onchange="javascript: displaySprayListData3();">
<option value="|||">---SELECT A PRODUCT---</option>
<%
IF rsSelect.EOF THEN
rsSelect.MoveFirst
DO WHILE not rsSelect.eof
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID3) = trim(rsSelect.Fields("SprayListID")) or local_SprayID3 = trim(rsSelect.Fields("SprayListID")) then
	response.write("selected")
	formMaxAppUse3 =  rsSelect.Fields("MaxUseApp")
	formMaxAppSeason3 = rsSelect.Fields("MaxUseSeason")
	formUnits3 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF

if l_TargetsHaveArray then
	if IsNumeric(localSprayArray(3,2)) then
		formRateAcre3 = trim(CStr(localSprayArray(3,2)))
	end if
end if
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units3" size="2" readonly value="<%=formUnits3%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse3" size="2" readonly value="<%=formMaxAppUse3%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason3" value="<%=formMaxAppSeason3%>" size="2" readonly style="border: 0px;"></td>
<td ><span class="bodytext"><input type="text" style="background-color:beige" name="AcresTreated3" value="<%=formAcresTreated3%>" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal3();"></span></td>
<td ><span class="bodytext"><input type="text" value="<%=formRateAcre3%>" style="background-color:beige" name="RateAcre3" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal3();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied3%>" style="background-color:beige" name="TotalMaterialApplied3" size="5" maxlength="11" class="bodytext" onchange="javascript:document.frmadd.RateAcre3.value=0;calculateTotal3();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating3%>" name="IFPRating3"  class="bodytext" size="8" maxlength="50"> --->
<select name="TargetID3" multiple="multiple" size="5" style="background-color:beige;">
<!-- <option value="">---TARGET---</option> -->
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
l_match = false
DO WHILE not rsSelectTarget.eof
	if l_TargetsHaveArray then
		if trim(CStr(rsSelectTarget.Fields("TargetID"))) = trim(CStr(localSprayArray(3,4))) or trim(formTargetID3) = trim(rsSelectTarget.Fields("TargetID")) then
			l_match = true
		else
			l_match = false
		end if
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if l_match then response.write(" selected ") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
	else
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if trim(formTargetID3) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target") %></option>
<%
	end if
	rsSelectTarget.MoveNext
LOOP
END IF
%>
</select></td>
<td><input type=button value=Recalc onclick="calculateTotal();" /></td>
</tr>


<!--- ADDITIONAL RECORD 4 --->
<tr>
<td class="bodytext">5)</td>

<td><SELECT name="SprayListID4" style="background-color:beige" onchange="javascript: displaySprayListData4();">
<option value="|||">---SELECT A PRODUCT---</option>
<%
IF rsSelect.EOF THEN
rsSelect.MoveFirst
DO WHILE not rsSelect.eof
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID4) = trim(rsSelect.Fields("SprayListID")) or local_SprayID4 = trim(rsSelect.Fields("SprayListID")) then
	response.write("selected")
	formMaxAppUse4 =  rsSelect.Fields("MaxUseApp")
	formMaxAppSeason4 = rsSelect.Fields("MaxUseSeason")
	formUnits4 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF

if l_TargetsHaveArray then
	if IsNumeric(localSprayArray(4,2)) then
		formRateAcre4 = trim(CStr(localSprayArray(4,2)))
	end if
end if
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units4" size="2" readonly value="<%=formUnits4%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse4" size="2" readonly value="<%=formMaxAppUse4%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason4" value="<%=formMaxAppSeason4%>" size="2" readonly style="border: 0px;"></td>
<td ><span class="bodytext"><input type="text" style="background-color:beige" name="AcresTreated4" value="<%=formAcresTreated4%>" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal4();"></span></td>
<td ><span class="bodytext"><input type="text" value="<%=formRateAcre4%>" style="background-color:beige" name="RateAcre4" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal4();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied4%>" style="background-color:beige" name="TotalMaterialApplied4" size="5" maxlength="11" class="bodytext" onchange="javascript:document.frmadd.RateAcre4.value=0;calculateTotal4();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating4%>" name="IFPRating4"  class="bodytext" size="8" maxlength="50"> --->
<select name="TargetID4" multiple="multiple" size="5" style="background-color:beige;">
<!-- <option value="">---TARGET---</option> -->
<%
rsSelectTarget.MoveFirst

IF not rsSelectTarget.EOF THEN
l_match = false
DO WHILE not rsSelectTarget.eof
	if l_TargetsHaveArray then
		if trim(CStr(rsSelectTarget.Fields("TargetID"))) = trim(CStr(localSprayArray(4,4))) or trim(formTargetID4) = trim(rsSelectTarget.Fields("TargetID")) then
			l_match = true
		else
			l_match = false
		end if
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if l_match then response.write(" selected ") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
	else
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if trim(formTargetID4) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target") %></option>
<%
	end if
	rsSelectTarget.MoveNext
LOOP
END IF
%>
</select></td>
<td><input type=button value=Recalc onclick="calculateTotal();" /></td>
</tr>


<!--- ADDITIONAL RECORD 5 --->
<tr>
<td class="bodytext">6)</td>

<td><SELECT name="SprayListID5" style="background-color:beige" onchange="javascript: displaySprayListData5();">
<option value="|||">---SELECT A PRODUCT---</option>
<%
IF rsSelect.EOF THEN
rsSelect.MoveFirst
DO WHILE not rsSelect.eof
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID5) = trim(rsSelect.Fields("SprayListID")) or local_SprayID5 = trim(rsSelect.Fields("SprayListID")) then
	response.write("selected")
	formMaxAppUse5 =  rsSelect.Fields("MaxUseApp")
	formMaxAppSeason5 = rsSelect.Fields("MaxUseSeason")
	formUnits5 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF

if l_TargetsHaveArray then
	if IsNumeric(localSprayArray(5,2)) then
		formRateAcre5 = trim(CStr(localSprayArray(5,2)))
	end if
end if
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units5" size="2" readonly value="<%=formUnits5%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse5" size="2" readonly value="<%=formMaxAppUse5%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason5" value="<%=formMaxAppSeason5%>" size="2" readonly style="border: 0px;"></td>
<td ><span class="bodytext"><input type="text" style="background-color:beige" name="AcresTreated5" value="<%=formAcresTreated5%>" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal5();"></span></td>
<td ><span class="bodytext"><input type="text" value="<%=formRateAcre5%>" style="background-color:beige" name="RateAcre5" size="5" maxlength="38" class="bodytext" onchange="javascript:calculateTotal5();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied5%>" style="background-color:beige" name="TotalMaterialApplied5" size="5" maxlength="11" class="bodytext" onchange="javascript:document.frmadd.RateAcre5.value=0;calculateTotal5();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating5%>" name="IFPRating5"  class="bodytext" size="8" maxlength="50"> --->
<select name="TargetID5" multiple="multiple" size="5" style="background-color:beige;">
<!-- <option value="">---TARGET---</option> -->
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
l_match = false
DO WHILE not rsSelectTarget.eof
	if l_TargetsHaveArray then
		if trim(CStr(rsSelectTarget.Fields("TargetID"))) = trim(CStr(localSprayArray(5,4))) or trim(formTargetID5) = trim(rsSelectTarget.Fields("TargetID")) then
			l_match = true
		else
			l_match = false
		end if
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if l_match then response.write(" selected ") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
	else
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>" <%if trim(formTargetID5) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target") %></option>
<%
	end if
	rsSelectTarget.MoveNext
LOOP
END IF
%>
</select>
</td>
<td><input type=button value=Recalc onclick="calculateTotal();" /></td>
</tr>




<%END IF%>
</table>


</td></tr>
<tr><td ></td>
<td ></td>
</tr>

<table>
<tr><td ><br />
<span class="subtitle"><label for="Comments">Comments</label>:</span></td><td><br /><span class="bodytext"><textarea cols="110" rows="6" name="Comments" class="bodytext"><%=formComments%></textarea></span></td>
</tr>
</table>
</td></tr>
<tr>
<td align=center><% IF  urlSprayRecordID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext"><% ELSE %><input type="submit" name="insert" value="Add Spray Record" class="bodytext">&nbsp;&nbsp;<input type="submit" name="insert_n_save_spray" value="Add Spray Record and Save Spray Data" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>
</table></form>

<%
IF TRUE THEN
%>

Up to last 5 entered records in your login.

<table width="90%" border="1" cellpadding="2" cellspacing="0">
<% if  delerror then%>
<tr>
<td colspan="27" class="bodytext"><font color="red"><%= delerrormessage %></font></td>
</tr>
<% end if %>
<%
set rs = GetSprayRecordsByLogin()
Dim i,maxMaterialApplied,maxSeasonApplied
i = 0

IF not rs.EOF THEN
DO WHILE not rs.eof
thisReentryIntervalDays = rs.Fields("ReentryIntervalDays")
thisReentryIntervalHours = rs.Fields("ReentryIntervalHours")
thisPreharvestInterval = rs.Fields("Preharvestinterval")
i = i + 1
if i < 5 then%>
<tr <%if i mod 2 = 1 then %>bgcolor="cccccc"<%end if%>>
<td rowspan="3"><%=i%></td>
<td colspan="3"><strong><%=rs.Fields("GrowerName")%></strong>&nbsp;&nbsp;<a href="enterSpraydata.asp?adding=1&SprayRecordID=<%=rs.Fields("SprayRecordID")%>#edit" class="bodytext">Edit</a>&nbsp;&nbsp;<a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="enterSpraydata.asp?SprayRecordID=<%=rs.Fields("SprayRecordID")%>&task=d" class="bodytext">Delete</a>&nbsp;&nbsp;&nbsp;&nbsp;  Rec.#<%=rs.Fields("SprayRecordID")%></td>
</tr>
<tr <%if i mod 2 = 1 then %>bgcolor="cccccc"<%end if%>>
	<td class="bodytext" nowrap>

	<%if rs.Fields("PackerNUmber")<>"" then %>Packer: <%=rs.Fields("PackerNumber")%><br><%end if %>
	Harvest Date: <%=rs.Fields("HarvestDate")%><br>
Spray Date: <%=rs.Fields("SprayStartDate")%>-<%=rs.Fields("SprayEndDate")%><br>
Time Finished: <%=rs.Fields("TimeFinishedSpraying")%><br>Crop: <%=rs.Fields("Crop")%><br>
Varieties:
<%if rs.Fields("Variety") <> "" then
	response.write( rs.Fields("Variety"))
end if%>
<%if rs.Fields("Variety2") <> "" then
	response.write("," & rs.Fields("Variety2"))
end if%>
<%if rs.Fields("Variety3") <> "" then
	response.write("," & rs.Fields("Variety3"))
end if%>
<%if rs.Fields("Variety4") <> "" then
	response.write("," & rs.Fields("Variety4"))
end if%>
&nbsp;<br>
	Bartlett: <% if  rs.Fields("Bartlet") then response.write("Yes") else response.write("No") end if%> <br>
	Stage: <%=rs.Fields("Stage")%><br>
	Location: <%=rs.Fields("Location")%><br>
	Weather: <%=rs.Fields("Weather")%></td>
	<td class="bodytext" nowrap><%=rs.Fields("Method")%>&nbsp;<br>
	Acres: <%=rs.Fields("AcresTreated")%>&nbsp;<br>

	Rate/Acre: <%=rs.Fields("RateAcre")%>&nbsp;| MAX App: <%=rs.Fields("MaxUseApp")%>| MAX Sea: <%=rs.Fields("MaxUseSeason")%><br><br>
	<%
		if rs.Fields("RateAcre") = "" THEN
			rs.Fileds("RateAcre") = 0
		END IF
		if rs.Fields("AcresTreated") = "" THEN
			rs.Fileds("AcresTreated") = 0
		END IF
		thisTotalMaterialApplied = rs.Fields("RateAcre") * rs.Fields("AcresTreated")
		maxMaterialApplied =  rs.Fields("MaxUseApp") * rs.Fields("AcresTreated")
		set rsSelect = GetSeasonQty(rs.Fields("SprayStartDate"),rs.Fields("SprayListID"),rs.Fields("GrowerID"),rs.Fields("Location"),rs.Fields("SprayYearID"))
		maxSeasonApplied =  rsSelect(1) * rs.Fields("MaxUseSeason")

		if thisReentryIntervalDays = "" then
			thisReentryIntervalDays = 0
		end if
		if thisReentryIntervalHours = "" then
			thisReentryIntervalHours = 0
		end if
		if thisPreharvestInterval = "" then
			thisPreharvestInterval = 0
		end if
	%>
	Total Material: <%=thisTotalMaterialApplied%> | MAX: <%=maxMaterialApplied%><br>
	Total This Season: <%=rsSelect(0)%> | MAX: <%=maxSeasonApplied%><br>
	Preharvest Interval: <%=thisPreharvestInterval%> Days</td>
	<td class="bodytext" nowrap><strong><%=rs.Fields("ProductNameAndFormulation")%></strong><br>
	Units: <%=rs.Fields("Unit")%><br>
	Target:
	<%
	Dim tmprs
	Set tmprs = conn.execute("SELECT t.target FROM targets t INNER JOIN SprayRecordTargets srt ON srt.TargetID = t.TargetID AND srt.sprayRecordID = " & rs.Fields("SprayRecordID"))
	tmprs.MoveFirst
	Do While Not tmprs.EOF
    	Response.Write tmprs.Fields("Target")
		tmprs.MoveNext
		If tmprs.EOF = False Then
		Response.Write ", "
		End If
	Loop
	%>
	<br>
	Applicator: <%=rs.Fields("Applicator")%><br>
	License: <%=rs.Fields("ApplicatorLicense")%><br>
	Supervisor: <%=rs.Fields("Supervisor")%><br>	Supervisor License: <%=rs.Fields("LicenseNumber")%><br>
	Chemical Supplier: <%=rs.Fields("ChemicalSupplier")%><br>
	Recommended By: <%=rs.Fields("RecommendedBy")%><br>
	</td></tr>
	<tr  <%if i mod 2 = 1 then %>bgcolor="cccccc"<%end if%>>
	<td colspan="1" style="border-bottom: thin solid Black;">Last Update: <%
	IF isdate(rs.Fields("UpdateDate")) then
		response.write(FormatDateTime(rs.Fields("UpdateDate")))
	else
		response.write(rs.Fields("UpdateDate"))
	end if%><br>
By: <%=rs.Fields("Administrator")%></td>
	<td colspan="3"style="border-bottom: thin solid Black;"><strong>Comments</strong>:<%=rs.Fields("Comments")%>&nbsp;<br>
	<%
		if  thisTotalMaterialApplied > maxMaterialApplied AND rs.Fields("MaxUseApp") <> 0 	then
			response.write("<font color=red><strong>Over Application: Yes</strong></font>")
		else
			response.write("<strong>Over Application: No</strong>")
		end if%>
				<br>
<%

		if rsSelect(0) > maxSeasonApplied AND rs.Fields("MaxUseSeason") <> 0 then
			response.write("<font color=red><strong>Over Season: Yes</strong></font>")
		else
			response.write("<strong>Over Season: No</strong>")
		end if

		%>
</td>
</tr>
<%
end if
rs.MoveNext
LOOP

Else
%>
<tr><td class="bodytext" colspan="27">No Records Selected</td></tr>
<%	end if %>
</table>

<%END IF%>

</td></tr>

</table>

<!--#include file="i_adminfooter.asp" -->


<%
	set rs = nothing
	set rsSelect = nothing
	set rsSelect2 = nothing
	set rsSelectTarget = nothing
	EndConnect(conn)

	function saveWeatherCode()
'	<span class="subtitle"><label for="WeatherID">Weather</label>:</span><br>
'<.
'	set rsSelect = GetActiveWeather()
'.>
'<SELECT name="WeatherID">
'<option value="">NEW WEATHER</option>
'<.
'IF not rsSelect.EOF THEN
'DO WHILE not rsSelect.eof
'.>
'<option value="<.=rsSelect.Fields("WeatherID").>"<.if trim(formWeatherID) = trim(rsSelect.Fields("WeatherID")) then response.write("selected") end if.>><.=rsSelect.Fields("Weather").></option>
'<.
'rsSelect.MoveNext
'LOOP
'END IF
'.>
'</select>
'<br><span class="subtitle">Enter New Weather:</span><br>
'<input type="text" value="<.=RemoveQuotes(formNewWeather).>" name="NewWeather"  class="bodytext" size="15" maxlength="20"></span>

	end function

%>


