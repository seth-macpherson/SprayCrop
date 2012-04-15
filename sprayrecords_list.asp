<%
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "CACHE-CONTROL", "NO-CACHE"
Response.Buffer = true
%>
<%if not session("login") or not listContains("1,2,3", session("accessid")) then
	response.redirect("index.asp")
end if%>
<!--#include file="include/i_data.asp"-->
<!--#include file="i_SprayRecord.asp"-->
<!--#include file="i_Crop.asp"-->
<!--#include file="i_Varieties.asp"-->
<!--#include file="i_Growers.asp"-->
<!--#include file="i_GrowerLocations.asp"-->
<!--#include file="i_Method.asp"-->
<!--#include file="i_SprayList.asp"-->
<!--#include file="i_SprayYears.asp"-->
<!--#include file="i_Stage.asp"-->
<!--#include file="i_Target.asp"-->
<!--#include file="i_Units.asp"-->
<!--#include file="i_Method.asp"-->
<!--#include file="i_Weather.asp"-->
<%
'CREATED by LocusInteractive on 08/02/2005
Dim errorFound,formError,errorMessage,tempErrorMessage,crossList,intCount,delError
Dim urlSprayRecordID,formSprayRecordID,urlAdding,formAdding,urlSearch,formSearch
Dim conn,sql,rs,rsSelect,counter,searchQueryString,page,searching,rsSelect2
dim searchSprayYear

'# 3-Apr-2011

searchSprayYear=request.form("SprayYear")
if request.querystring("searchsprayyear")<>"" then
	searchSprayYear=request.querystring("searchSprayYear")
end if

if listContains("1", session("accessid")) OR listContains("2", session("accessid")) then session("growerid")=0
'#

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

formAdding = Request.Form.Item("Adding")
urlAdding = Request.QueryString("Adding")
'IF urlAdding = "" THEN urlAdding = 0 END IF
IF urlAdding = "" THEN urlAdding = false END IF
IF formAdding = "" THEN formAdding = urlAdding End IF
urlAdding = formAdding

formSearch=Request.Form.Item("Search")
urlSearch=Request.QueryString("Search")
IF urlSearch = "" THEN urlSearch = 0 END IF
IF formSearch = "" THEN formSearch = urlSearch End IF
urlSearch = formSearch

dim searchGrower,searchCrop,searchVariety,searchBartlet,searchStage,searchLocation,searchMethod,searchProduct,searchTarget,searchUpdateBy,searchOverSeason,searchOverApplication,searchHighSprayDate,searchHighHarvestDate,searchLowSprayDate,searchLowHarvestDate

searching=Request.Form.Item("searching")
IF Request.QueryString("searching") <> "" AND searching = "" THEN 
	searching = Request.QueryString("searching")
END IF
urlSearch = 1 'wtf?
formSearch = 1
searchGrower=Request.Form.Item("searchGrower")
'Response.Write("<br>searchGrower (1): " & searchGrower)
IF Request.QueryString("searchGrower") <> "" AND searchGrower = "" THEN 
	searchGrower = Request.QueryString("searchGrower")
END IF
IF searchGrower = "" THEN
	tmpGrowerID = Request.Form.Item("GrowerID")
	arrayGrower = Split(tmpGrowerID,"|")
	IF IsArray(arrayGrower)  THEN
		IF Ubound(arrayGrower) >= 0 THEN
			searchGrower = arrayGrower(0)
		END IF
	ELSE
		searchGrower = tmpGrowerID
	END IF
END IF
'Response.Write("<br>searchGrower (2): " & searchGrower)

     if request.servervariables("request_method")="POST" and request.Form("changerole")>"" then
    
        dim g: g=split(request.Form("growerrole"),"|")
        
        if isarray(g) then 
            session("growerid")=g(0)
            session("growername")=g(1)
        end if
    
    end if

searchCrop=Request.Form.Item("searchCrop")
IF Request.QueryString("searchCrop") <> "" AND searchCrop = "" THEN 
	searchCrop = Request.QueryString("searchCrop")
END IF

searchBartlet=Request.Form.Item("searchBartlet")
if Request.QueryString("searchBartlet") <> "" AND searchBartlet = "" THEN 
	searchBartlet = Request.QueryString("searchBartlet")
END IF

searchStage=Request.Form.Item("searchStage")
if Request.QueryString("searchStage") <> "" AND searchStage = "" THEN 
	searchStage = Request.QueryString("searchStage")
END IF

searchLocation=Request.Form.Item("searchLocation")
if Request.QueryString("searchLocation") <> "" AND searchLocation = "" THEN 
	searchLocation = Request.QueryString("searchLocation")
END IF

searchMethod=Request.Form.Item("searchMethod")
if Request.QueryString("searchMethod") <> "" AND searchMethod = "" THEN 
	searchMethod = Request.QueryString("searchMethod")
END IF

searchProduct=Request.Form.Item("searchProduct")
'Response.Write("<br>searchProduct (1): " & searchProduct)
if Request.QueryString("searchProduct") <> "" AND searchProduct = "" THEN 
	searchProduct = Request.QueryString("searchProduct")
END IF
'Response.Write("<br>searchProduct (2): " & searchProduct)
'REM formSprayListID???

searchTarget=Request.Form.Item("searchTarget")
if Request.QueryString("searchTarget") <> "" AND searchTarget = "" THEN 
	searchTarget = Request.QueryString("searchTarget")
END IF

searchUpdateBy=Request.Form.Item("searchUpdateBy")
if Request.QueryString("searchUpdateBy") <> "" AND searchUpdateBy = "" THEN 
	searchUpdateBy = Request.QueryString("searchUpdateBy")
END IF

searchOverSeason=Request.Form.Item("searchOverSeason")
if Request.QueryString("searchOverSeason") <> "" AND searchOverSeason = "" THEN 
	searchOverSeason = Request.QueryString("searchOverSeason")
END IF

searchOverApplication=Request.Form.Item("searchOverApplication")
if Request.QueryString("searchOverApplication") <> "" AND searchOverApplication = "" THEN 
	searchOverApplication = Request.QueryString("searchOverApplication")
END IF

searchHighSprayDate=Request.Form.Item("searchHighSprayDate")
if Request.QueryString("searchHighSprayDate") <> "" AND searchHighSprayDate = "" THEN 
	searchHighSprayDate = Request.QueryString("searchHighSprayDate")
END IF
if not isDate(searchHighSprayDate) then searchHighSprayDate = "" end if

searchLowSprayDate=Request.Form.Item("searchLowSprayDate")
if Request.QueryString("searchLowSprayDate") <> "" AND searchLowSprayDate = "" THEN 
	searchLowSprayDate = Request.QueryString("searchLowSprayDate")
END IF
if not isDate(searchLowSprayDate) then searchLowSprayDate = "" end if

searchHighHarvestDate=Request.Form.Item("searchHighHarvestDate")
if Request.QueryString("searchHighHarvestDate") <> "" AND searchHighHarvestDate = "" THEN 
	searchHighHarvestDate = Request.QueryString("searchHighHarvestDate")
END IF
if not isDate(searchHighHarvestDate) then searchHighHarvestDate = "" end if

searchLowHarvestDate=Request.Form.Item("searchLowHarvestDate")
if Request.QueryString("searchLowHarvestDate") <> "" AND searchLowHarvestDate = "" THEN 
	searchLowHarvestDate = Request.QueryString("searchLowHarvestDate")
END IF
if not isDate(searchLowHarvestDate) then searchLowHarvestDate = "" end if

page=Request.Form.Item("page")
if Request.QueryString("page") <> "" AND page = "" THEN 
	page = Request.QueryString("page")
END IF
if Request.QueryString("newPage") <> "" THEN 
	page = Request.QueryString("newPage")
END IF
if page = "" then
	page = 1
end if

searchQueryString = "searchsprayyear="&searchsprayyear&"&searchGrower=" & RemoveWhitespace(searchGrower) & "&searchCrop=" & RemoveWhitespace(searchCrop) & "&searchBartlet=" & RemoveWhitespace(searchBartlet) & "&searchStage=" & RemoveWhitespace(searchStage) & "&searchLocation=" & RemoveWhitespace(searchLocation) & "&searchMethod=" & RemoveWhitespace(searchMethod) & "&searchProduct=" & RemoveWhitespace(searchProduct) & "&searchTarget=" & RemoveWhitespace(searchTarget) & "&searchUpdateBy=" & RemoveWhitespace(searchUpdateBy) & "&searchOverSeason=" & RemoveWhitespace(searchOverSeason) & "&searchOverApplication=" & RemoveWhitespace(searchOverApplication) & "&searchHighSprayDate=" & RemoveWhitespace(searchHighSprayDate) & "&searchHighHarvestDate=" & RemoveWhitespace(searchHighHarvestDate) & "&searchLowSprayDate=" & RemoveWhitespace(searchLowSprayDate) & "&searchLowHarvestDate=" & RemoveWhitespace(searchLowHarvestDate) & "&page=" & page & "&searching=" & searching

'Cancel button hit
IF Request.Form.Item("cancel") <> "" THEN
	Response.Redirect("sprayrecords_list.asp?" & searchQueryString)
END IF

'Initialize Form Fields
DIM formGrowerID,formSprayStartDate,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated,formRateAcre,formSprayListID,formUnitsOfProduct,formIFPRating,formTargetID,formHarvestDate,formComments,formUnitID,arrayGrower,arraySprayList,formLocationOption,formLocationText,formTotalMaterialApplied,urlLocation,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy,urlApplicatorLicense,urlAdministrator,urlSupervisor,urlLicenseNumber,urlChemicalSupplier,urlRecommendedBy,formAdministrator,formTimeFinishedSpraying

arrayGrower = Split(Request.Form.Item("GrowerID"),"|")
IF IsArray(arrayGrower)  THEN
	IF Ubound(arrayGrower) >= 0 THEN
		formGrowerID = arrayGrower(0)
	END IF
END IF

formSprayStartDate = Request.Form.Item("SprayStartDate")
formSprayEndDate = Request.Form.Item("SprayEndDate")
formCropID = Request.Form.Item("CropID")
formVarietyID = Request.Form.Item("VarietyID")
formBartlet = Request.Form.Item("Bartlet")
formStageID = Request.Form.Item("StageID")
formTimeFinishedSpraying = Request.Form.Item("TimeFinishedSpraying")
formLocationOption = Request.Form.Item("LocationOptions")
urlLocation=Request.QueryString("Location")
if formLocationOption = "" then formLocationOption = urlLocation end if

formLocationText = Request.Form.Item("LocationText")
if trim(formLocationOption) <> "" THEN
	formLocation = formLocationOption
ELSE
	formLocation = formLocationText
END IF
formMethodID = Request.Form.Item("MethodID")

formWeather = Request.Form.Item("Weather")
urlWeather=Request.QueryString("Weather")
if formWeather = "" then formWeather = urlWeather end if

formAcresTreated = Request.Form.Item("AcresTreated")
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

formSupervisorOption = Request.Form.Item("SupervisorOptions")
urlSupervisor=Request.QueryString("Supervisor")
'response.write(urlLocation)
if formSupervisorOption = "" then formSupervisorOption = urlSupervisor end if

formSupervisorText = Request.Form.Item("SupervisorText")
if trim(formSupervisorOption) <> "" THEN
	formSupervisor = formSupervisorOption
	formSupervisorText = ""
ELSE
	formSupervisor = formSupervisorText
END IF

formLicenseNumberOption = Request.Form.Item("LicenseNumberOptions")
urlLicenseNumber=Request.QueryString("LicenseNumber")
'response.write(urlLocation)
if formLicenseNumberOption = "" then formLicenseNumberOption = urlLicenseNumber end if

formLicenseNumberText = Request.Form.Item("LicenseNumberText")
if trim(formLicenseNumberOption) <> "" THEN
	formLicenseNumber = formLicenseNumberOption
	formLicenseNumberText = ""
ELSE
	formLicenseNumber = formLicenseNumberText
END IF

formChemicalSupplierOption = Request.Form.Item("ChemicalSupplierOptions")
urlChemicalSupplier=Request.QueryString("ChemicalSupplier")
'response.write(urlLocation)
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
'response.write(urlLocation)
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
'response.write(urlLocation)
if formApplicatorOption = "" then formApplicatorOption = urlApplicator end if

formApplicatorText = Request.Form.Item("ApplicatorText")
if trim(formApplicatorOption) <> "" THEN
	formApplicator = formApplicatorOption
	formApplicatorText = ""
ELSE
	formApplicator = formApplicatorText
END IF


formApplicatorLicenseOption = Request.Form.Item("ApplicatorLicenseOptions")
urlApplicatorLicense=Request.QueryString("ApplicatorLicense")
'response.write(urlLocation)
if formApplicatorLicenseOption = "" then formApplicatorLicenseOption = urlApplicatorLicense end if

formApplicatorLicenseText = Request.Form.Item("ApplicatorLicenseText")
if trim(formApplicatorLicenseOption) <> "" THEN
	formApplicatorLicense = formApplicatorLicenseOption
	formApplicatorLicenseText = ""
ELSE
	formApplicatorLicense = formApplicatorLicenseText
END IF
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
'initialize the connection
set conn = Connect()
set rs = Server.CreateObject("ADODB.RecordSet")
set rsSelect = Server.CreateObject("ADODB.RecordSet")
set rsSelect2 = Server.CreateObject("ADODB.RecordSet")

IF Request.QuerySTring("task") = "d" and urlSprayRecordID <> "" THEN
	DeleteSprayRecord(urlSprayRecordID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("sprayrecords_list.asp?" & searchQueryString)
END IF 
IF Request.QuerySTring("task") = "activate" and urlSprayRecordID <> "" THEN
	ActivateSprayRecord(urlSprayRecordID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("sprayrecords_list.asp?" & searchQueryString)
END IF 
IF Request.QuerySTring("task") = "deactivate" and urlSprayRecordID <> "" THEN
	DeActivateSprayRecord(urlSprayRecordID)
	set rs = nothing
	EndConnect(conn)
	Response.Redirect("sprayrecords_list.asp?" & searchQueryString)
END IF 


'Form Was Submitted
IF Request.Form.Item("insert") <> "" OR Request.Form.Item("update") <> "" THEN

	'Form Validation
	IF NOT ValidateDatatype(Request.Form.Item("GrowerID"), "char","Grower", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("SprayStartDate"), "datetime","SprayStartDate", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("SprayEndDate"), "datetime","SprayEndDate", TRUE) THEN
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
	IF NOT ValidateDatatype(formSprayListID, "int","Product", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(formLocation, "nvarchar","Location", TRUE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF
	IF NOT ValidateDatatype(Request.Form.Item("Weather"), "nvarchar","Weather", FALSE) THEN
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
	
	IF	Request.Form.Item("insert") <> "" THEN
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
		END IF
	END IF
	
	
	IF NOT ValidateDatatype(Request.Form.Item("Comments"), "nvarchar","Comments", FALSE) THEN
		errorFound = TRUE
		errorMessage = errorMessage + "<br>" + tempErrorMessage
	END IF

	'Update record
	IF NOT errorFound AND Request.Form.Item("update") <> "" THEN 
		urlSprayRecordID = UpdateSprayRecord(formSprayRecordID,formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated,formRateAcre,formSprayListID,formIFPRating,formTargetID,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)

'Response.Write("<br><b>QS: " & searchQueryString & "</b>")
'	Response.Redirect("sprayrecords_list.asp?" & searchQueryString)
		'END UPDATE
	END IF 
	'INSERT
	IF NOT errorFound AND Request.Form.Item("insert") <> "" THEN
'	response.write("growerId: " & CStr(formGrowerID))
		urlSprayRecordID = InsertSprayRecord(formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated,formRateAcre,formSprayListID,formIFPRating,formTargetID,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)


IF	Request.Form.Item("insert") <> "" THEN
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID1 <> "" THEN
			urlSprayRecordID = InsertSprayRecord(formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated1,formRateAcre1,formSprayListID1,formIFPRating1,formTargetID1,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)
		END IF
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID2 <> "" THEN
			urlSprayRecordID = InsertSprayRecord(formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated2,formRateAcre2,formSprayListID2,formIFPRating2,formTargetID2,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)
		END IF
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID3 <> "" THEN
			urlSprayRecordID = InsertSprayRecord(formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated3,formRateAcre3,formSprayListID3,formIFPRating3,formTargetID3,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)
		END IF
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID4 <> "" THEN
			urlSprayRecordID = InsertSprayRecord(formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated4,formRateAcre4,formSprayListID4,formIFPRating4,formTargetID4,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)
		END IF
		'CHECK ADDITIONAL SPRAY RECORDS
		IF formSprayListID5 <> "" THEN
			urlSprayRecordID = InsertSprayRecord(formGrowerID,formSprayStartDate,formTimeFinishedSpraying,formSprayEndDate,formCropID,formVarietyID,formBartlet,formStageID,formLocation,formMethodID,formAcresTreated5,formRateAcre5,formSprayListID5,formIFPRating5,formTargetID5,formHarvestDate,formComments,formWeather,formApplicator,formApplicatorLicense,formAdministrator,formSupervisor,formLicenseNumber,formChemicalSupplier,formRecommendedBy)
		END IF
	END IF



	Response.Redirect("sprayrecords_list.asp?" & searchQueryString)
	END IF 'insert	
ELSEIF formSprayRecordID = 0 THEN
	'formSprayDate = month(now()) & "/" & day(now()) & "/" & year(now())
	'formHarvestDate = month(now()) & "/" & day(now()) & "/" & year(now())
END IF 'form submitted 

IF formSprayRecordID <> 0 and not errorFound and formCropID = 0 THEN
	set rs = GetSprayRecordByID(formSprayRecordID)
	IF NOT rs.eof THEN

		formGrowerID = rs.Fields("GrowerID")
		formSprayStartDate = rs.Fields("SprayStartDate")
		formSprayEndDate = rs.Fields("SprayEndDate")
		formCropID = rs.Fields("CropID")
		formVarietyID = rs.Fields("VarietyID1")
		formVarietyID = listAppend(formVarietyID,rs.Fields("VarietyID2"))
		formVarietyID = listAppend(formVarietyID,rs.Fields("VarietyID3"))
		formVarietyID = listAppend(formVarietyID,rs.Fields("VarietyID4"))
		formBartlet = rs.Fields("Bartlet")
		formStageID = rs.Fields("StageID")
		formLocation = rs.Fields("Location")
		formMethodID = rs.Fields("MethodID")
		formAcresTreated = rs.Fields("AcresTreated")
		formRateAcre = rs.Fields("RateAcre")
		formSprayListID = rs.Fields("ProductID")
		formIFPRating = rs.Fields("IFPRating")
		formTargetID = rs.Fields("TargetID")
		formHarvestDate = rs.Fields("HarvestDate")
		formComments = rs.Fields("Comments")
		formWeather = rs.Fields("Weather")
		formApplicator = rs.Fields("Applicator")
		formApplicatorLicense = rs.Fields("ApplicatorLicense")		
		formAdministrator = rs.Fields("Administrator")
		formSupervisor = rs.Fields("Supervisor")
		formLicenseNumber = rs.Fields("LicenseNumber")
		formChemicalSupplier = rs.Fields("ChemicalSupplier")
		formRecommendedBy = rs.Fields("RecommendedBy")		
		formTimeFinishedSpraying = rs.Fields("TimeFinishedSpraying")		
		if rs.Fields("RateAcre") = "" THEN
			rs.Fileds("RateAcre") = 0
		END IF
		if rs.Fields("AcresTreated") = "" THEN
			rs.Fileds("AcresTreated") = 0
		END IF
		formTotalMaterialApplied = Round(rs.Fields("RateAcre") * rs.Fields("AcresTreated"),2)
	END IF
END IF%>
<html>
<head>
	<title><%=Application("BusinessName")%>&nbsp;-&nbsp;<%=Application("ProgramName")%>&nbsp;-&nbsp;Spray Record List</title>
    <link rel=stylesheet type="text/css" href="li_admin.css">
	<script language="JavaScript" src="datepicker.js"></script>
	
	<!--#include file="SprayDataCalcs.js"-->
<%
if formAdding and urlSprayRecordID <> 0 then
%>
	<script language"JavaScript">
		window.name = 'SprayRecordEdit'
	</script>
<%
else
%>
	<script language"JavaScript">
		window.name = 'SprayRecordsList';
	</script>
<%
end if
%>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" rightmargin="10" bottommargin="0" marginwidth="0" marginheight="0">
<!--#include file="i_adminheader.asp" -->

<table width="95%" border="0" cellspacing="0" cellpadding="0" align=center><tr><td><img src="images/spacer.gif" height="4" width="1" border="0"><br>
<h1>> Review Spray Data</h1><br><img src="images/spacer.gif" height="4" width="1" border="0"></td></tr></table><br />

<table width="95%" border="0" bgcolor="FFFFFF" align="center">
<tr><td bgcolor="ffffff" class="bodytext"><br>
Edit/delete spray records.<br><br>
<%if not formAdding then%>
<!--<a href="sprayrecords_list.asp?adding=true&<%=searchQuerySTring%>"><strong>Add SprayRecord</strong></a>
<a href="enterspraydata.asp"><strong>Add SprayRecord</strong></a><br><br>-->
<%end if%>


</td>
</tr>
<tr>
<td colspan="2" class="bodytext">
<%if formAdding then %>
<form action="sprayrecords_list.asp" method="post" name="frmadd">
<input type="hidden" name="RefreshData" value="0">
<input type="hidden" name="NewProdID" value="0">
<input type="hidden" name="Adding" value="true">
<table width="500" border="1" cellpadding="0" cellspacing="0">
<tr><td>
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr bgcolor="#cccccc">
<td align="left" class="bodytext" colspan="2"><strong><%if urlSprayRecordID <> 0 then%>
EDITING RECORD #<%=formSprayRecordID%>
<%else%>
ADDING RECORD
<%end if%></strong>
</td><td>&nbsp;</td>
</tr>
<tr>
<td align="left" class="bodytext"><strong>* indicates required field (blue)</strong><br>
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
<tr><td valign="top"><span class="subtitle"><label for="GrowerName">*Grower</label>:</span><br><span class="bodytext">
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
<option value="<%	response.write(rsSelect.Fields("GrowerID"))
	set rsSelect2 = GetGrowersLocations(rsSelect.Fields("GrowerID"))
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("|" & rsSelect2.Fields("Location"))
			rsSelect2.movenext
		LOOP
	END IF
%>" 

<%if trim(formGrowerID) = trim(rsSelect.Fields("GrowerID")) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("GrowerNumber")%>&nbsp;&nbsp;&nbsp;<%=rsSelect.Fields("GrowerName")%></option>
<%
	rsSelect.MoveNext
	LOOP
END IF
%>
</select></span><br><br><br>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
<td valign="top">&nbsp;&nbsp;



<span class="subtitle">*Supervisor</span>:<br>
<select name="SupervisorOptions" class="bodytext"style="background-color:beige">
<option value="">NEW SUPERVISOR</option>
<%
set rsSelect2 = GetGrowersSupervisors(formGrowerID)
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("<option value='" & rsSelect2.Fields("Supervisor") & "'")
			IF rsSelect2.Fields("Supervisor") = RemoveQuotes(formSupervisor) THEN
				response.write("selected")
				formSupervisor = ""
			END if
			response.write (">" & rsSelect2.Fields("Supervisor") & "</option>")
			rsSelect2.movenext
		LOOP
	END IF
%>
</select>
<br>
<input type="text" value="<%=formSupervisor%>" name="SupervisorText" size="20" maxlength="50" class="bodytext" style="background-color:beige">&nbsp;&nbsp;&nbsp;
<td valign="top">
<span class="subtitle">*Supervisor License</span>:<br>
<select name="LicenseNumberOptions" class="bodytext"style="background-color:beige">
<option value="">NEW SUPERVISOR LIC</option>
<%
set rsSelect2 = GetGrowersSupervisorLicenses(formGrowerID)
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("<option value='" & rsSelect2.Fields("LicenseNumber") & "'")
			IF rsSelect2.Fields("LicenseNumber") = RemoveQuotes(formLicenseNumber) THEN
				response.write("selected")
				formLicenseNumber = ""
			END if
			response.write (">" & rsSelect2.Fields("LicenseNumber") & "</option>")
			rsSelect2.movenext
		LOOP
	END IF
%>
</select><br><input type="text" value="<%=RemoveQuotes(formLicenseNumber)%>" name="LicenseNumberText" size="20" maxlength="50" class="bodytext" style="background-color:beige">&nbsp;&nbsp;&nbsp;
</td>
</tr>
<tr>
<td valign="top"><br>&nbsp;&nbsp;
<span class="subtitle">*Applicator(s)</span>:<br><select name="ApplicatorOptions" class="bodytext"style="background-color:beige">
<option value="">NEW APPLICATOR</option>
<%
set rsSelect2 = GetGrowersApplicators(formGrowerID)
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("<option value='" & rsSelect2.Fields("Applicator") & "'")
			IF rsSelect2.Fields("Applicator") = RemoveQuotes(formApplicator) THEN
				response.write("selected")
				formApplicator = ""
			END if
			response.write (">" & rsSelect2.Fields("Applicator") & "</option>")
			rsSelect2.movenext
		LOOP
	END IF
%>
</select><br><input type="text" value="<%=RemoveQuotes(formApplicator)%>" name="ApplicatorText" size="20" maxlength="50" class="bodytext" style="background-color:beige"> &nbsp;&nbsp;&nbsp;</td>
<td valign="top"><br>
<span class="subtitle">Applicator License</span>:<br><select name="ApplicatorLicenseOptions" class="bodytext">
<option value="">NEW APPLICATOR LIC</option>
<%
set rsSelect2 = GetGrowersApplicatorLicenses(formGrowerID)
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("<option value='" & rsSelect2.Fields("ApplicatorLicense") & "'")
			IF rsSelect2.Fields("ApplicatorLicense") = RemoveQuotes(formApplicatorLicense) THEN
				response.write("selected")
				formApplicatorLicense = ""
			END if
			response.write (">" & rsSelect2.Fields("ApplicatorLicense") & "</option>")
			rsSelect2.movenext
		LOOP
	END IF
%>
</select><br><input type="text" value="<%=RemoveQuotes(formApplicatorLicense)%>" name="ApplicatorLicenseText" size="20" maxlength="50" class="bodytext">&nbsp;&nbsp;&nbsp;
</td>


</tr>


<tr>
<td>&nbsp;&nbsp;<br>
<span class="subtitle">*Chemical Supplier</span>:<br><select name="ChemicalSupplierOptions" class="bodytext" style="background-color:beige">
<option value="">NEW CHEMICAL SUPL.</option>
<%
set rsSelect2 = GetGrowersChemicalSuppliers(formGrowerID)
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("<option value='" & rsSelect2.Fields("ChemicalSupplier") & "'")
			IF rsSelect2.Fields("ChemicalSupplier") = RemoveQuotes(formChemicalSupplier) THEN
				response.write("selected")
				formChemicalSupplier = ""
			END if
			response.write (">" & rsSelect2.Fields("ChemicalSupplier") & "</option>")
			rsSelect2.movenext
		LOOP
	END IF
%>
</select><br><input type="text" value="<%=RemoveQuotes(formChemicalSupplier)%>" name="ChemicalSupplierText" size="20" maxlength="50" class="bodytext" style="background-color:beige">&nbsp;&nbsp;&nbsp;
<td><br>
<span class="subtitle">*Recommended By</span>:<br><select name="RecommendedByOptions" class="bodytext"style="background-color:beige">
<option value="">NEW RECOMMENDED</option>
<%
set rsSelect2 = GetGrowersRecommendedBy(formGrowerID)
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("<option value='" & rsSelect2.Fields("RecommendedBy") & "'")
			IF rsSelect2.Fields("RecommendedBy") = RemoveQuotes(formRecommendedBy) THEN
				response.write("selected")
				formRecommendedBy = ""
			END if
			response.write (">" & rsSelect2.Fields("RecommendedBy") & "</option>")
			rsSelect2.movenext
		LOOP
	END IF
%>
</select><br><input type="text" value="<%=RemoveQuotes(formRecommendedBy)%>" name="RecommendedByText" size="20" maxlength="50" class="bodytext" style="background-color:beige">&nbsp;&nbsp;&nbsp;
</td>

</tr>

</table></td>


<td valign="top" nowrap colspan="3">
<table  border="0" cellpadding="0" cellspacing="0" ><tr>
<td valign="top">
<span class="subtitle"><label for="SprayStartDate">* Spray Start Date</label>:</span><br><span class="bodytext">

<table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=formSprayStartDate%>" name="SprayStartDate" size="20" maxlength="21" class="bodytext" style="background-color:beige"></td><td><a href="javascript:show_calendar('frmadd.SprayStartDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table><br><span class="subtitle">Time of Application </span><br>hh:mm [AM/PM]<br><input type="text" value="<%=formTimeFinishedSpraying%>" name="TimeFinishedSpraying" size="20" maxlength="21" class="bodytext"><br><em>Time Required only if Field or Central Posting Report desired.</em></span><br><br><span class="subtitle"><label for="SprayEndDate">* Spray End Date</label>:</span><br><span class="bodytext"><table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=formSprayEndDate%>" name="SprayEndDate" size="20" maxlength="21" class="bodytext" style="background-color:beige"></td><td><a href="javascript:show_calendar('frmadd.SprayEndDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td>
</tr></table></span>

<br><span class="subtitle">
<span class="subtitle"><label for="HarvestDate">Harvest Date</label>:</span><table border=0 cellspacing=0 cellpadding=0><tr><td><input type="text" value="<%=formHarvestDate%>" name="HarvestDate" size="20" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmadd.HarvestDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table>

</td>
<td>&nbsp;&nbsp;&nbsp;</td>
<td nowrap valign="top" class="bodytext" width="150"><span class="subtitle"><label for="CropID">*Crop</label>:</span><br>
<Select name="CropID" onchange="javascript:form.submit();">
<%
	set rsSelect = GetActiveCrops()
	thisCropID = 0
i = 0
IF not rsSelect.EOF THEN
	thisCropID = rsSelect.Fields("CropID")
DO WHILE not rsSelect.eof 
i = i + 1
%>
<option value="<%=rsSelect.Fields("CropID")%>"<%if listContains(formCropID, trim(rsSelect.Fields("CropID"))) then 
thisCropID = rsSelect.Fields("CropID") 
response.write("selected") 
end if%> style="background-color:beige"><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Crop")%></option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</SELECT><br>

<%
	set rsSelect = GetActiveVarietiesByCropID(thisCropID)
i = 0
IF not rsSelect.EOF THEN%>
<table><tr><td class="smalltext" colspan="2">Choose up to 4 varieties.</td></tr>
<tr><td valign=top>
<%
DO WHILE not rsSelect.eof 
i = i + 1
%>
<input type="checkbox" name="VarietyID" value="<%=rsSelect.Fields("VarietyID")%>"<%if listContains(formVarietyID, trim(rsSelect.Fields("VarietyID"))) then response.write("checked") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Variety")%></option>
<%
if (i = 7) then
	response.write ("</td><td valign=top>")
end if
rsSelect.MoveNext
LOOP%>
</td></tr></table>
<%
END IF
%>
&nbsp;&nbsp;
<%if thisCropID = 1 then%>
<br><span class="subtitle"><label for="Bartlet">Bartlett:</label></span> <span class="bodytext"><span class="bodytext"><input type="radio" value="1" name="Bartlet" <% if ListContains(formBartlet,"True") OR ListContains(formBartlet,"1")  THEN %>Checked<% END IF %> style="background-color:beige">YES <input type="radio" value="" name="Bartlet" <% if ListContains(formBartlet,"False") OR ListContains(formBartlet,"0")  THEN %>Checked<% END IF %> style="background-color:beige">NO</span>
<%end if%></td>

</tr></table>

<tr>

<td valign="top" colspan="3">




</td>

</tr>

<tr>
<td valign="top" colspan="3">
<table width="100%">
<tr>

<td><span class="subtitle"><label for="Weather">Weather</label>:</span><br>
<span class="bodytext"><input type="text" value="<%=RemoveQuotes(formWeather)%>" name="Weather"  class="bodytext" size="15" maxlength="50"></span>
<br>&nbsp;<br>&nbsp;
</td><td><span class="subtitle"><label for="LocationOptions">*Location</label>:</span><br>

<span class="bodytext"><select name="LocationOptions" class="bodytext"style="background-color:beige">
<option value="">SELECT LOCATION</option>
<%
'set rsSelect2 = GetGrowersLocations(formGrowerID)
'	IF not rsSelect2.EOF THEN
'		DO WHILE not rsSelect2.eof
'			response.write("<option value='" & rsSelect2.Fields("Location") & "'")
'			IF rsSelect2.Fields("Location") = RemoveQuotes(formLocation) THEN
'				response.write("selected")
'				formLocation = ""
'			END if
'			response.write (">" & rsSelect2.Fields("Location") & "</option>")
'			rsSelect2.movenext
'		LOOP
'	END IF
	set rsSelect2 = GetGrowersLocationsByGrowerID(CStr(formGrowerID))
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
<br><span class="subtitle">Enter New Location:</span><br>
<input type="text" value="<%=RemoveQuotes(formLocation)%>" name="LocationText"  class="bodytext" size="15" maxlength="50"></span>
-->
</td>
<td valign="top">
<span class="subtitle"><label for="MethodID">*Method</label>:</span><br>
<%
	set rsSelect = GetActiveMethods()
%>
<SELECT name="MethodID" style="background-color:beige">
<option value="">---SELECT A METHOD---</option>
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
<option value="">---SELECT A STAGE---</option>
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



<tr><td colspan="3">
<table width="100%" cellspacing="0" cellpadding="0" border="0">
<tr>


</tr>
</table>
</td>
</tr>

<tr><td valign="top" colspan="3">
<%
	Dim formMaxAppUse,formMaxAppSeason,formUnits
	set rsSelect = GetActiveSprayListByCropID(thisCropID)
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
<td  valign="bottom"><strong>Acres Treat</strong><br># only</td>
<td  valign="bottom"><strong>Rate Acre</strong><br># only</td>
<td  valign="bottom"><strong>Total Applied</strong><br># only</td>
<td bgcolor="cccccc" valign="bottom"><strong>Target</strong></td>
</tr>

<tr>
<td class="bodytext">1)</td>
<td valign="top">*<SELECT name="SprayListID"onchange="javascript: displaySprayListData();" style="background-color:beige">
<option value="">---SELECT A PRODUCT---</option>
<%

IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID) = trim(rsSelect.Fields("SprayListID")) then 
	response.write("selected") 
	formMaxAppUse =  rsSelect.Fields("MaxUseApp") 
	formMaxAppSeason = rsSelect.Fields("MaxUseSeason") 
	formUnits = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units" size="2" readonly value="<%=formUnits%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse" size="2" readonly value="<%=formMaxAppUse%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason" value="<%=formMaxAppSeason%>" size="2" readonly style="border: 0px;"></td>
<td  valign="top"><span class="bodytext"><input type="text" name="AcresTreated" value="<%=formAcresTreated%>" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal();prefilAcres();" style="background-color:beige"></span></td>
<td  valign="top"><span class="bodytext"><input type="text" value="<%=formRateAcre%>" name="RateAcre" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal();" style="background-color:beige"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied%>" name="TotalMaterialApplied" size="5" maxlength="11" class="bodytext" onChange="javascript:document.frmadd.RateAcre.value=0;calculateTotal();" style="background-color:beige"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating%>" name="IFPRating"  class="bodytext" size="8" maxlength="50"> --->

<%
	set rsSelectTarget = GetActiveTarget()
%>
<SELECT name="TargetID" style="background-color:beige">
<option value="">---TARGET---</option>
<%
IF not rsSelectTarget.EOF THEN
DO WHILE not rsSelectTarget.eof 
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>"<%if trim(formTargetID) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
rsSelectTarget.MoveNext
LOOP
END IF
%>
</select>
</td>
</tr>


<!--- ABILITY TO ENTER 5 MORE RECORDS ON INSERT --->
<%IF urlSprayRecordID = 0 THEN %>


<!--- ADDITIONAL RECORD 1 --->
<tr>
<td class="bodytext">2)</td>

<td valign="top"><SELECT name="SprayListID1"onchange="javascript: displaySprayListData1();">
<option value="">---SELECT A PRODUCT---</option>
<%
rsSelect.MoveFirst
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID1) = trim(rsSelect.Fields("SprayListID")) then 
	response.write("selected") 
	formMaxAppUse1 =  rsSelect.Fields("MaxUseApp") 
	formMaxAppSeason1 = rsSelect.Fields("MaxUseSeason") 
	formUnits1 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units1" size="2" readonly value="<%=formUnits1%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse1" size="2" readonly value="<%=formMaxAppUse1%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason1" value="<%=formMaxAppSeason1%>" size="2" readonly style="border: 0px;"></td>
<td  valign="top"><span class="bodytext"><input type="text" name="AcresTreated1" value="<%=formAcresTreated1%>" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal1();"></span></td>
<td  valign="top"><span class="bodytext"><input type="text" value="<%=formRateAcre1%>" name="RateAcre1" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal1();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied1%>" name="TotalMaterialApplied1" size="5" maxlength="11" class="bodytext" onChange="javascript:document.frmadd.RateAcre1.value=0;calculateTotal1();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating1%>" name="IFPRating1"  class="bodytext" size="8" maxlength="50"> --->
<SELECT name="TargetID1">
<option value="">---TARGET---</option>
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
DO WHILE not rsSelectTarget.eof 
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>"<%if trim(formTargetID1) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
rsSelectTarget.MoveNext
LOOP
END IF
%>
</select></td>
</tr>

<!--- ADDITIONAL RECORD 2 --->
<tr>
<td class="bodytext">3)</td>

<td valign="top"><SELECT name="SprayListID2"onchange="javascript: displaySprayListData2();">
<option value="">---SELECT A PRODUCT---</option>
<%
rsSelect.MoveFirst
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID2) = trim(rsSelect.Fields("SprayListID")) then 
	response.write("selected") 
	formMaxAppUse2 =  rsSelect.Fields("MaxUseApp") 
	formMaxAppSeason2 = rsSelect.Fields("MaxUseSeason") 
	formUnits2 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units2" size="2" readonly value="<%=formUnits2%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse2" size="2" readonly value="<%=formMaxAppUse2%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason2" value="<%=formMaxAppSeason2%>" size="2" readonly style="border: 0px;"></td>
<td  valign="top"><span class="bodytext"><input type="text" name="AcresTreated2" value="<%=formAcresTreated2%>" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal2();"></span></td>
<td  valign="top"><span class="bodytext"><input type="text" value="<%=formRateAcre2%>" name="RateAcre2" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal2();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied2%>" name="TotalMaterialApplied2" size="5" maxlength="11" class="bodytext" onChange="javascript:document.frmadd.RateAcre2.value=0;calculateTotal2();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating2%>" name="IFPRating2"  class="bodytext" size="8" maxlength="50"> --->
<SELECT name="TargetID2">
<option value="">---TARGET---</option>
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
DO WHILE not rsSelectTarget.eof 
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>"<%if trim(formTargetID2) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
rsSelectTarget.MoveNext
LOOP
END IF
%>
</select></td>
</tr>

<!--- ADDITIONAL RECORD 3 --->
<tr>
<td class="bodytext">4)</td>

<td valign="top"><SELECT name="SprayListID3" onchange="javascript: displaySprayListData3();">
<option value="">---SELECT A PRODUCT---</option>
<%
rsSelect.MoveFirst
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID3) = trim(rsSelect.Fields("SprayListID")) then 
	response.write("selected") 
	formMaxAppUse3 =  rsSelect.Fields("MaxUseApp") 
	formMaxAppSeason3 = rsSelect.Fields("MaxUseSeason") 
	formUnits3 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units3" size="2" readonly value="<%=formUnits3%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse3" size="2" readonly value="<%=formMaxAppUse3%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason3" value="<%=formMaxAppSeason3%>" size="2" readonly style="border: 0px;"></td>
<td  valign="top"><span class="bodytext"><input type="text" name="AcresTreated3" value="<%=formAcresTreated3%>" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal3();"></span></td>
<td  valign="top"><span class="bodytext"><input type="text" value="<%=formRateAcre3%>" name="RateAcre3" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal3();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied3%>" name="TotalMaterialApplied3" size="5" maxlength="11" class="bodytext" onChange="javascript:document.frmadd.RateAcre3.value=0;calculateTotal3();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating3%>" name="IFPRating3"  class="bodytext" size="8" maxlength="50"> --->
<SELECT name="TargetID3" >
<option value="">---TARGET---</option>
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
DO WHILE not rsSelectTarget.eof 
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>"<%if trim(formTargetID3) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
rsSelectTarget.MoveNext
LOOP
END IF
%>
</select></td>
</tr>

<!--- ADDITIONAL RECORD 4 --->
<tr>
<td class="bodytext">5)</td>

<td valign="top"><SELECT name="SprayListID4" onchange="javascript: displaySprayListData4();">
<option value="">---SELECT A PRODUCT---</option>
<%
rsSelect.MoveFirst
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID4) = trim(rsSelect.Fields("SprayListID")) then 
	response.write("selected") 
	formMaxAppUse4 =  rsSelect.Fields("MaxUseApp") 
	formMaxAppSeason4 = rsSelect.Fields("MaxUseSeason") 
	formUnits4 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units4" size="2" readonly value="<%=formUnits4%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse4" size="2" readonly value="<%=formMaxAppUse4%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason4" value="<%=formMaxAppSeason4%>" size="2" readonly style="border: 0px;"></td>
<td  valign="top"><span class="bodytext"><input type="text" name="AcresTreated4" value="<%=formAcresTreated4%>" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal4();"></span></td>
<td  valign="top"><span class="bodytext"><input type="text" value="<%=formRateAcre4%>" name="RateAcre4" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal4();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied4%>" name="TotalMaterialApplied4" size="5" maxlength="11" class="bodytext" onChange="javascript:document.frmadd.RateAcre4.value=0;calculateTotal4();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating4%>" name="IFPRating4"  class="bodytext" size="8" maxlength="50"> --->
<SELECT name="TargetID4">
<option value="">---TARGET---</option>
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
DO WHILE not rsSelectTarget.eof 
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>"<%if trim(formTargetID4) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
rsSelectTarget.MoveNext
LOOP
END IF
%>
</select></td>
</tr>

<!--- ADDITIONAL RECORD 5 --->
<tr>
<td class="bodytext">6)</td>

<td valign="top"><SELECT name="SprayListID5" onchange="javascript: displaySprayListData5();">
<option value="">---SELECT A PRODUCT---</option>
<%
rsSelect.MoveFirst
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("SprayListID")%>|<%=rsSelect.Fields("MaxUseApp")%>|<%=rsSelect.Fields("MaxUseSeason")%>|<%=rsSelect.Fields("Unit")%>"<%if trim(formSprayListID5) = trim(rsSelect.Fields("SprayListID")) then 
	response.write("selected") 
	formMaxAppUse5 =  rsSelect.Fields("MaxUseApp") 
	formMaxAppSeason5 = rsSelect.Fields("MaxUseSeason") 
	formUnits5 = rsSelect.Fields("Unit")
end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%>&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="Units5" size="2" readonly value="<%=formUnits5%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppUse5" size="2" readonly value="<%=formMaxAppUse%>" style="border: 0px;"></td>
<td bgcolor="cccccc" nowrap class="bodytext"><input type="text" name="MaxAppSeason5" value="<%=formMaxAppSeason5%>" size="2" readonly style="border: 0px;"></td>
<td  valign="top"><span class="bodytext"><input type="text" name="AcresTreated5" value="<%=formAcresTreated5%>" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal5();"></span></td>
<td  valign="top"><span class="bodytext"><input type="text" value="<%=formRateAcre5%>" name="RateAcre5" size="5" maxlength="38" class="bodytext" onChange="javascript:calculateTotal5();"></span></td>
<td  nowrap class="bodytext"><input type="text" value="<%=formTotalMaterialApplied5%>" name="TotalMaterialApplied5" size="5" maxlength="11" class="bodytext" onChange="javascript:document.frmadd.RateAcre5.value=0;calculateTotal5();" ></td>
<td bgcolor="cccccc" nowrap class="bodytext"><!--- <input type="text" value="<%=formIFPRating5%>" name="IFPRating5"  class="bodytext" size="8" maxlength="50"> --->
<SELECT name="TargetID5" >
<option value="">---TARGET---</option>
<%
rsSelectTarget.MoveFirst
IF not rsSelectTarget.EOF THEN
DO WHILE not rsSelectTarget.eof 
%>
<option value="<%=rsSelectTarget.Fields("TargetID")%>"<%if trim(formTargetID5) = trim(rsSelectTarget.Fields("TargetID")) then response.write("selected") end if%>><%if not rsSelectTarget.Fields("Active") then %>*<%end if%><%=rsSelectTarget.Fields("Target")%></option>
<%
rsSelectTarget.MoveNext
LOOP
END IF
%>
</select>
</td>
</tr>



<%END IF%>
</table>






</td></tr>
<tr><td valign="top" ></td>
<td valign="top" ></td>
</tr>

<table>
<tr><td valign="top" >
<span class="subtitle"><label for="Comments">Comments</label>:</span></td><td><span class="bodytext"><textarea cols="50" rows="4" name="Comments" class="bodytext"><%=formComments%></textarea></span></td>
</tr>
</table>
</td></tr>
<tr>
<td><% IF  urlSprayRecordID <> 0 THEN%><input type="submit" name="update" value="Update" class="bodytext" OnClick="alert('Update in Progress... \r\n The Edit Spray Record window will close when you press OK below. \r\n The Spray Records Search Page will automatically refresh. \r\n \r\n However, the browser may popup a dialog box stating that the information will be resent in order to refresh the page.  This is okay!  With Internet Explorer click the Retry button.  Then your search criteria will remain and the resulting list will show the editted spray record.');window.opener.location.reload(true);window.close()"><% ELSE %><input type="submit" name="insert" value="Add Spray Record" class="bodytext"><%END IF%>&nbsp;&nbsp;<input type="submit" name="cancel" value="Cancel" class="bodytext"></td>
</tr>
</table></form>
<% end if 'formadding %>


<!---SEARCH--->

<%if formSearch and not formAdding then %>

<form action="sprayrecords_list.asp?searching=1" method="post" name="frmsearch">
<table width="500" border="1" cellpadding="0" cellspacing="0">
<tr><td>
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr bgcolor="#cccccc">
<td align="left" class="bodytext" colspan="2"><strong>
SELECT SEARCH CRITERIA</strong> &nbsp;&nbsp; &nbsp;&nbsp;<a href="sprayrecords_list.asp?search=1">Reset Search</a><br>
Hold down the "CTRL" key to multiple select.<br>
* indicates in-active data</td><td>&nbsp;</td>
</tr>
<tr valign=top>

<td>

<br><span class="subtitle"><label for="GrowerName">Year</label>:</span> 

<%
	set rsSelect = GetAllSprayYears()
%>
<select name="SprayYear" size="1" onchange=this.form.submit() >
<%
IF not rsSelect.EOF THEN
DO WHILE not rsSelect.eof 
%>
<option value="<%=trim(rsSelect.Fields("SprayYearID"))%>" <%if listContains(trim(searchSprayYear), trim(rsSelect.Fields("SprayYearID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("SprayYear")%> </option>
<%
rsSelect.MoveNext
LOOP
END IF
%>
</select>
<br><bR>

<br><span class="subtitle"><label for="GrowerName">Grower</label>:</span>
<%if session("growerid")=0 then %>
    <%
	    set rsSelect = GetAllGrowers()
    %><br />
    <select name="SearchGrower" size="12" multiple>
    <%
    IF not rsSelect.EOF THEN
    DO WHILE not rsSelect.eof 
    %>
    <option value="<%=trim(rsSelect.Fields("GrowerID"))%>" <%if listContains(trim(SearchGrower), trim(rsSelect.Fields("GrowerID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%'=rsSelect.Fields("GrowerNumber")&"&nbsp;&nbsp;&nbsp;"%><%=rsSelect.Fields("GrowerName")%> </option>
    <%
    rsSelect.MoveNext
    LOOP
    END IF
    %>
    </select>
<%else%>


<%
    
        dim roleConn: set roleConn = Connect()
        dim rsRoles: set rsRoles = conn.execute("exec growerunit$bygrower " & session("growerid")) 
        with response
        
        .Write "<br><input type=hidden name=changerole />"
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

    <%'=session("growername")%>

<%end if%>
</td>

<td valign="top" nowrap colspan="2" ><span class="subtitle"><label for="SprayListID">Product Name and Formulation</label>:</span><br><%
'Response.Write("<br>searchProduct (4): " & searchProduct)
	set rsSelect = GetAllSprayList()
%>

<SELECT name="SearchProduct" size="12" multiple>
<option value="" <%if searchProduct="" then response.write "selected"%>>---SELECT ALL---</option>
<%
IF not rsSelect.EOF THEN
	DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("SprayListID")%>"<%if ListContains(trim(searchProduct), trim(rsSelect.Fields("SprayListID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Name")%></option>
<%
		rsSelect.MoveNext
	LOOP
END IF
%>
</select> 
</td>

</tr>
<tr>

<td valign="top" colspan="3">

<table>

<tr>
<td><span class="subtitle"><label for="CropID">Crop</label>:</span><br>
<%
	set rsSelect = GetAllCrops()
%>
<SELECT name="SearchCrop" size="5" multiple>
<option value="" <%if searchCrop="" then response.write "selected"%>>---SELECT ALL---</option>
<%
IF not rsSelect.EOF THEN
	DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("CropID")%>"<%if ListContains(trim(SearchCrop),trim(rsSelect.Fields("CropID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Crop")%></option>
<%
		rsSelect.MoveNext
	LOOP
END IF
%>
</select>
</td>

<td valign="top" ><span class="subtitle"><label for="SearchBartlet">Bartlet:</label></span><br><span class="bodytext"><span class="bodytext"><input type="radio" value="1" name="SearchBartlet" <% if ListContains(SearchBartlet,"True") OR ListContains(SearchBartlet,"1")  THEN %>Checked<% END IF %>>YES&nbsp;&nbsp;<input type="radio" value="" name="SearchBartlet" <% if ListContains(SearchBartlet,"False") OR ListContains(SearchBartlet,"0")  THEN %>Checked<% END IF %>>NO</span></td>
<td valign="top"><span class="subtitle"><label for="SearchStageID">Stage</label>:</span><br>
<%
	set rsSelect = GetAllStage()
%>
<SELECT name="searchStage" size="5" multiple>
<option value="" <%if searchStage="" then response.write "selected"%>>---SELECT ALL---</option>
<%
IF not rsSelect.EOF THEN
	DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("StageID")%>"<%if ListContains(trim(searchStage),trim(rsSelect.Fields("StageID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Stage")%></option>
<%
		rsSelect.MoveNext
	LOOP
END IF
%>
</select>
</td>
<td valign="top"><span class="subtitle"><label for="Location">Location</label>:</span><br>
<span class="bodytext">

<!--
<input type="text" value="<%=searchLocation%>" name="locationtext"  class="bodytext" size="15" maxlength="50">
-->

<!---->
<select name="searchlocation" class="bodytext" style="background-color:beige" size=1>
<option value=""> - ALL - </option>
<%
	
	set rsSelect2 = GetGrowersLocationsByGrowerID(session("GrowerID"))
	IF not rsSelect2.EOF THEN
		DO WHILE not rsSelect2.eof
			response.write("<option value='" & rsSelect2.Fields("GLoc_Location") & "' ")
			IF rsSelect2.Fields("GLoc_Location") = RemoveQuotes(searchlocation) THEN
				response.write("selected")
				formLocation = ""
			END if
			response.write (">" & rsSelect2.Fields("GLoc_Location") & "</option>")
			rsSelect2.movenext
		LOOP
	END IF
	
%>
</select>

</span>
</td>
</tr>


<tr>
<td valign="top">
<span class="subtitle"><label for="SearchHighSprayDate">SprayDate</label>:</span><br>Low: <table border=0 cellpadding="0"  cellspacing="0"><tr><td><input type="text" value="<%=SearchLowSprayDate%>" name="SearchLowSprayDate" size="10" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmsearch.SearchLowSprayDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table><br><span class="bodytext">High: <table border=0 cellpadding="0"  cellspacing="0"><tr><td><input type="text" value="<%=SearchHighSprayDate%>" name="SearchHighSprayDate" size="10" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmsearch.SearchHighSprayDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table></span></td>
<td valign="top">
<span class="subtitle"><label for="SearchHighHarvestDate">HarvestDate</label>:</span><br>
Low: <span class="bodytext"><table border=0 cellpadding="0"  cellspacing="0"><tr><td><input type="text" value="<%=SearchLowHarvestDate%>" name="SearchLowHarvestDate" size="10" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmsearch.SearchLowHarvestDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table></span><br>High: <span class="bodytext"><table border=0 cellpadding="0"  cellspacing="0"><tr><td><input type="text" value="<%=SearchHighHarvestDate%>" name="SearchHighHarvestDate" size="10" maxlength="21" class="bodytext"></td><td><a href="javascript:show_calendar('frmsearch.SearchHighHarvestDate');" onmouseover="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"><img src="images/calendar.gif" width=24 height=22 border=0></a></td></tr></table></span></td>
<td valign="top">
<span class="subtitle"><label for="SearchMethod">Method</label>:</span><br>
<%
	set rsSelect = GetAllMethod()
%>
<SELECT name="searchMethod" size="5" multiple>
<option value="" <%if searchMethod="" then response.write "selected"%>>---SELECT ALL---</option>
<%
IF not rsSelect.EOF THEN
	DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("MethodID")%>"<%if ListContains(trim(searchMethod),trim(rsSelect.Fields("MethodID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Method")%></option>
<%
		rsSelect.MoveNext
	LOOP
END IF
%>
</select>
</td>
<td valign="top"><span class="subtitle"><label for="searchTargetID">Target</label>:</span><br>
<%
	set rsSelect = GetAllTarget()
%>
<SELECT name="searchTarget" size="5" multiple>
<option value="" <%if searchTarget="" then response.write "selected"%>>---SELECT ALL---</option>
<%
IF not rsSelect.EOF THEN
	DO WHILE not rsSelect.eof 
%>
<option value="<%=rsSelect.Fields("TargetID")%>"<%if ListContains(trim(searchTarget),trim(rsSelect.Fields("TargetID"))) then response.write("selected") end if%>><%if not rsSelect.Fields("Active") then %>*<%end if%><%=rsSelect.Fields("Target")%></option>
<%
		rsSelect.MoveNext
	LOOP
END IF
%>
</select>
</td>
</tr>

</table>
<input type="submit" name="Go" value="--- SEARCH NOW---" class="bodytext">
</form>
</td>
</tr>
</table>
</td>
</tr>
</table>
<br><br>

<%elseif not formAdding then %>
<table width="500" border="1" cellpadding="0" cellspacing="0">
<tr><td>
<table width="500" border="0" cellpadding="2" cellspacing="0">
<tr bgcolor="#cccccc">
<strong>CURRENT SEARCH</strong> <a href="sprayrecords_list.asp?search=1&<%=searchQuerySTring%>">Edit Search</a>&nbsp;&nbsp;<a href="sprayrecords_list.asp?search=0">Reset Search</a><br>
</td></tr>
<tr><td valign="top">
<%
'Response.Write("<br>searchGrower (6): " & searchGrower)

if searchGrower <> "" then
	set rs = GetGrowersByID(searchGrower)
	Response.Write("<strong>Growers:</strong><br>" & vbCrLf)

	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			Response.Write(rs.Fields("GrowerName") & "<br>")
			rs.MoveNext
		LOOP
	END IF
end if

if searchCrop <> "" then
	set rs = GetCropByCropID(searchCrop)
	Response.Write("<br><strong>Crops:</strong><br>" & vbCrLf)

	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			Response.Write(rs.Fields("Crop") & "<br>")
			rs.MoveNext
		LOOP
	END IF
end if
if searchBartlet <> "" then
	Response.Write("<br> <strong>Bartlet:</strong> ")
	if searchBartlet = 1 then
		Response.Write("true")
	else
		Response.Write("false")
	end if
	Response.Write("<br>")
end if
if searchStage <> "" then
	set rs = GetStageByID(searchStage)
	Response.Write("<br> <strong>Stage:</strong><br>" & vbCrLf)
	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			Response.Write(rs.Fields("Stage") & "<br>")
			rs.MoveNext
		LOOP
	END IF
end if
if searchLocation <> "" then
	Response.Write("<br> <strong>Location:</strong> " & searchLocation & "<br>" & vbCrLf)
end if
if searchMethod <> "" then
	set rs = GetMethodByID(searchMethod)
	Response.Write("<br> <strong>Method: </strong><br>")
	IF not rs.EOF THEN
	DO WHILE not rs.eof 
		Response.Write(rs.Fields("Method") & "<br>")
		rs.MoveNext
	LOOP
	END IF
end if
%>
</td><td valign="top">
<%
	if searchHighSprayDAte <> "" then %>
 <strong>High Spray Date:</strong> <%=searchHighSprayDAte%><br>
<%	end if
	if searchLowSprayDAte <> "" then %>
<strong>Low Spray Date:</strong> <%=searchLowSprayDAte%><br>
<%	end if
	if searchHighHarvestDAte <> "" then %>
<strong>High Harvest Date:</strong> <%=searchHighHarvestDAte%><br>
<%	end if
	if searchHighHarvestDAte <> "" then %>
<strong>Low SHarvest Date:</strong> <%=searchLowHarvestDAte%><br>
<%	end if
if searchProduct <> "" then
	set rs = GetSprayListByID(searchProduct)
	Response.Write("<br> <strong>Product: </strong><br>")
	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			Response.Write(rs.Fields("Name") & "<br>")
			rs.MoveNext
		LOOP
	END IF
end if
if searchTarget <> "" then
	set rs = GetTargetByID(searchTarget)
	Response.Write("<br> <strong>Target:</strong><br> ")
	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			Response.Write(rs.Fields("Target") & "<br>")
			rs.MoveNext
		LOOP
	END IF
end if
if searchUpdateBy <> "" then
	Response.Write("<br> <strong>Updated By:</strong> " & searchUpdateBy & "<br>")
end if
if searchOverApplication <> "" then
	Response.Write("<br> <strong>Over Application:</strong> " & searchOverApplication & "<br>")
end if
if searchOverSeason <> "" then
	Response.Write("<br> <strong>Over Season:</strong> " & searchOverSeason & "<br>")
end if
%>


</td></tr></table>
</td></tr></table>
<%end if 'end form search%>



<%
'IF request.servervariables("request_method")="POST" then 'searching THEN
if searching then '# Jun-2010 need GET for pagination
'Response.Write("<br>searchGrower (5): " & searchGrower)
'response.write searchlocation

	if searching then

		set rs = GetCountSprayRecordsBySearch(searchGrower,searchCrop,"",searchBartlet,searchStage,searchLocation,searchMethod,searchProduct,searchTarget,searchUpdateBy,searchOverApplication,searchOverSeason,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate)		

	else
	
		set rs = GetLastSpraySearchCount()
		

	end if

	dim thisCount
	thisCount = rs(0)

	if searching then
	
		set rs = GetSprayRecordsBySearch(searchGrower,searchCrop,"",searchBartlet,searchStage,searchLocation,searchMethod,searchProduct,searchTarget,searchUpdateBy,searchOverApplication,searchOverSeason,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate)		

	else

		set rs = GetLastSpraySearch()


	end if

	Dim i,maxMaterialApplied,maxSeasonApplied
	i = 0

%><br><br>

<a name=searchresults>
<strong><%=thisCount%> records returned.</strong>
</a><br>

<%
	Dim startrow,endrow,recordsPerPage,nextPage
	recordsPerPage = 20
	if thisCount > recordsPerPage then
		startrow = (page - 1) * recordsPerPage + 1
		endrow = startrow + recordsPerPage
		if startrow > thisCount then
			startrow = 1
			endrow = thisCount + 1
		end if
		if endrow > thisCount then
			endrow = thisCount + 1
		end if
	else
		startrow = 1
		endrow = thisCount + 1
	end if 
	Response.Write(recordsPerPage & " records per page.")
	For i = 0 to thisCount Step recordsPerPage
		nextpage = (i + recordsPerPage)/recordsPerpage
		if trim(page) = trim(nextpage) then
			Response.Write("  &nbsp;&nbsp;<strong>" & nextpage & "</strong>")
		else
			Response.Write(" &nbsp;&nbsp;<a href=sprayrecords_list.asp?newPage=" & nextpage & "&"  & searchQuerySTring &  ">" & nextpage & "</a> ")
		end if
	Next

	'# debug: Response.Write("<br>"& startrow & ":" & endrow)
%>

<table width="90%" border="1" cellpadding="2" cellspacing="0">
<%
	if  delerror then%>
<tr>
<td colspan="9" class="bodytext" valign="top"><font color="red"><strong><%= delerrormessage %></strong></font></td>
</tr>
<%
	end if %>
<!--- <% if  errorFound then%>
<tr>
<td colspan="27" class="bodytext" valign="top"><a href="#edit"><font color="red">AN ERROR HAS OCCURRED, PLEASE SEE MESSAGE BELOW</font></a></td>
</tr>
<% end if %> --->


<tr>
<td  valign="top">&nbsp;</td>
	<td  valign="top"><h2>Crop</h2></td>
	<td  valign="top"><h2>Method</h2></td>
	<td  valign="top"><h2>Material</h2></td>
</tr>

<%
	dim j,thisTotalMaterialApplied
	i = 0
	j = 0
	IF not rs.EOF THEN
		DO WHILE not rs.eof 
			j = j + 1
			if j < endrow and j >= startrow THEN
				i = i + 1
%>

<tr <%if i mod 2 = 1 then %>bgcolor="cccccc"<%end if%>>
<td rowspan="3" valign="top"><%=i%></td>
<td colspan="3" valign="top"><strong><%=rs.Fields("GrowerName")%>&nbsp;#<%=rs.Fields("GrowerNumber")%></strong>&nbsp;&nbsp;
<a href="enterSpraydata.asp?adding=1&SprayRecordID=<%=rs.Fields("SprayRecordID")%>" class="bodytext">Edit</a>
&nbsp;&nbsp;<a  onclick="javascript:return confirm('Are you sure you want to delete this record?');"  href="enterSpraydata.asp?SprayRecordID=<%=rs.Fields("SprayRecordID")%>&task=d" class="bodytext">Delete</a>&nbsp;&nbsp;&nbsp;&nbsp;  Rec.#<%=rs.Fields("SprayRecordID")%></td>
</tr>
<tr <%if i mod 2 = 1 then %>bgcolor="cccccc"<%end if%>>
	<td class="bodytext" valign="top" nowrap>
	
	<%if rs.Fields("PackerNUmber")<>"" then %>Packer: <%=rs.Fields("PackerNumber")%><br><%end if %>
	Harvest Date: <%=rs.Fields("HarvestDate")%><br>
Spray Date: <%=rs.Fields("SprayStartDate")%>-<%=rs.Fields("SprayEndDate")%><br>
Time of Application: <%=rs.Fields("TimeFinishedSpraying")%><br>Crop: <%=rs.Fields("Crop")%><br>
Varieties:
<%if rs.Fields("Variety") <> "" then
	response.write( rs.Fields("Variety"))
end if
if rs.Fields("Variety2") <> "" then
	response.write("," & rs.Fields("Variety2"))
end if
if rs.Fields("Variety3") <> "" then
	response.write("," & rs.Fields("Variety3"))
end if
if rs.Fields("Variety4") <> "" then
	response.write("," & rs.Fields("Variety4"))
end if%>
&nbsp;<br>
	Bartlett: <% if  rs.Fields("Bartlet") then response.write("Yes") else response.write("No") end if%> <br>
	Stage: <%=rs.Fields("Stage")%><br>
	Location: <%=rs.Fields("Location")%><br>
	Weather: <%=rs.Fields("Weather")%></td>
	<td class="bodytext" valign="top" nowrap><%=rs.Fields("Method")%>&nbsp;<br>
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
	Preharvest Interval: <%=thisPreharvestInterval%> Days<br><br>
	
<%
	if rs.Fields("PURS_Reported") then
'		Response.Write("<font bgcolor=""#00FF33"">Reported to PURS</font>")
		Response.Write("<i>Reported to PURS</i>")
	else
'		Response.Write("<font bgcolor=""#FFFFCC"">NOT Reported to PURS</font>")
		Response.Write("<b>NOT Reported to PURS</b>")
	end if
%>
	
	</td>
	<td class="bodytext" valign="top" nowrap><strong><%=rs.Fields("ProductNameAndFormulation")%></strong><br>
	Units: <%=rs.Fields("Unit")%><br>
	Target: <%=rs.Fields("Target")%><br><br>
	Applicator: <%=rs.Fields("Applicator")%><br>
	License: <%=rs.Fields("ApplicatorLicense")%><br>
	Supervisor: <%=rs.Fields("Supervisor")%><br>
	Supervisor License: <%=rs.Fields("LicenseNumber")%><br>
	Chemical Supplier: <%=rs.Fields("ChemicalSupplier")%><br>
	Recommended By: <%=rs.Fields("RecommendedBy")%><br>
	</td></tr>
	<tr  <%if i mod 2 = 1 then %>bgcolor="cccccc"<%end if%>>
	<td valign="top" colspan="1" style="border-bottom: thin solid Black;">Last Update: <%
	IF isdate(rs.Fields("UpdateDate")) then 
		response.write(FormatDateTime(rs.Fields("UpdateDate"))) 
	else 
		response.write(rs.Fields("UpdateDate")) 
	end if%><br>
By: <%=rs.Fields("Administrator")%></td>
	<td colspan="3" valign="top"style="border-bottom: thin solid Black;"><strong>Comments</strong>:<%=rs.Fields("Comments")%>&nbsp;<br>
	<% 
		if  thisTotalMaterialApplied > maxMaterialApplied AND rs.Fields("MaxUseApp") <> 0 	then  
			response.write("<font color=red><strong>Over Appication: Yes</strong></font>") 
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
END IF
rs.MoveNext
LOOP

Else
%>
<tr><td class="bodytext" colspan="27">No Records Selected</td></tr>
<%	end if %>
</table>
<%	end if %>


</td></tr>

</table>
<%
	set rs = nothing
	EndConnect(conn)
%>
<!--#include file="i_adminfooter.asp" -->
</body>
</html>

