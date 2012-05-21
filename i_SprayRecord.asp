<%
' FILE: i_SprayRecord.asp
' CREATED by www.LocusInteractive.net on 08/02/2005
' MODIFIED kalanmiers 12/5/2006 cleanup!!!

' *******************************************************
' ************ GetSprayRecordByID ***********************
' *******************************************************
function GetSprayRecordByID(SprayRecordID)
	sql = "SELECT sr.*,s.sprayyearid FROM SprayRecord sr inner join spraylist s on sr.productid=s.spraylistid WHERE SprayRecordID = " & SprayRecordID
	set GetSprayRecordByID = conn.execute(sql)
end function

' *******************************************************
' ************ GetAllSprayRecords ***********************
' *******************************************************
function GetAllSprayRecord()
	sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
	set rs = conn.execute(sql)

	sql = "SELECT TOP 20 SprayRecord.SprayRecordID,SprayList.SprayYearID,SprayRecord.UpdateDAte,SprayRecord.CreateDate,SprayRecord.Administrator, SprayRecord.OverApplicationFlag,SprayRecord.GrowerID,SprayRecord.SprayStartDate,SprayRecord.TimeFinishedSpraying,SprayRecord.SprayEndDate,SprayRecord.CropID,VarietyID1,VarietyID2,VarietyID3,VarietyID4,SprayRecord.Bartlet,SprayRecord.StageID,SprayRecord.Location,SprayRecord.OverSeasonFlag,SprayRecord.MethodID,SprayRecord.AcresTreated,SprayRecord.RateAcre,SprayRecord.TotalMaterialApplied,SprayRecord.Applicator,SprayRecord.ApplicatorLicense,SprayRecord.GrowerID,SprayList.ReentryIntervalDays,SprayList.ReentryIntervalHours,SprayList.PreharvestInterval,SprayRecord.Weather,SprayList.SprayListID,SprayRecord.UnitID,SprayRecord.IFPRating,SprayRecord.TargetID,SprayRecord.HarvestDate,SprayRecord.Comments,Crops.Crop,Varieties.Variety,Varieties2.Variety as Variety2,Varieties3.Variety as Variety3, Varieties4.Variety as Variety4,Growers.GrowerName,Growers.GrowerNumber,Methods.Method,SprayList.Name,Stages.Stage,Targets.Target,Units.Unit,SprayList.Name as ProductNameAndFormulation,Spraylist.ActiveInd,SprayRecord.Administrator,SprayRecord.Supervisor,SprayRecord.LicenseNumber,SprayRecord.ChemicalSupplier,SprayRecord.RecommendedBy FROM (((((((((((SprayRecord LEFT JOIN Crops ON SprayRecord.CropID = Crops.CropID)LEFT JOIN Varieties ON SprayRecord.VarietyID1 = Varieties.VarietyID) LEFT JOIN Varieties Varieties2 ON SprayRecord.VarietyID2 = Crops2.VarietyID)LEFT JOIN Varieties Varieties3 ON SprayRecord.VarietyID3 = Varieties3.VarietyID)LEFT JOIN Varieties Varieties4 ON SprayRecord.VarietyID4 = Varieties4.VarietyID) LEFT JOIN Growers ON SprayRecord.GrowerID = Growers.GrowerID) LEFT JOIN Methods ON SprayRecord.MethodID = Methods.MethodID) LEFT JOIN SprayList ON SprayRecord.ProductID = SprayList.SprayListID"
	sql = sql & ") LEFT JOIN Stages ON SprayRecord.StageID = Stages.StageID) LEFT JOIN Targets ON SprayRecord.TargetID = Targets.TargetID) LEFT JOIN Units ON SprayRecord.UnitID = Units.UnitID) ORDER BY SprayRecord.CREATEDATE DESC, SprayRecord.SprayRecordID WHERE SprayYearID = " & rs(0)
'response.write sql
	set GetAllSprayRecord = conn.execute(sql)
end function

function GetSprayRecordsByLogin()
	dim cAddNums
	sql = 		"SELECT TOP 20 packers.packernumber,                    "
    sql = sql & "      sprayrecord.administrator,                       "
    sql = sql & "      sprayrecord.supervisor,                          "
    sql = sql & "      sprayrecord.updatedate,                          "
    sql = sql & "      sprayrecord.licensenumber,                       "
    sql = sql & "      sprayrecord.chemicalsupplier,                    "
    sql = sql & "      sprayrecord.recommendedby,                       "
    sql = sql & "      sprayrecord.applicator,                          "
    sql = sql & "      sprayrecord.sprayrecordid,                       "
    sql = sql & "      spraylist.sprayyearid,                           "
    sql = sql & "      sprayrecord.overapplicationflag,                 "
    sql = sql & "      sprayrecord.growerid,                            "
    sql = sql & "      sprayrecord.spraystartdate,                      "
    sql = sql & "      sprayrecord.timefinishedspraying,                "
    sql = sql & "      sprayrecord.sprayenddate,                        "
    sql = sql & "      sprayrecord.cropid,                              "
    sql = sql & "      varietyid1,                                      "
    sql = sql & "      varietyid2,                                      "
    sql = sql & "      varietyid3,                                      "
    sql = sql & "      varietyid4,                                      "
    sql = sql & "      sprayrecord.bartlet,                             "
    sql = sql & "      sprayrecord.stageid,                             "
    sql = sql & "      sprayrecord.location,                            "
    sql = sql & "      sprayrecord.overseasonflag,                      "
    sql = sql & "      sprayrecord.methodid,                            "
    sql = sql & "      sprayrecord.acrestreated,                        "
    sql = sql & "      sprayrecord.rateacre,                            "
    sql = sql & "      sprayrecord.totalmaterialapplied,                "
    sql = sql & "      sprayrecord.applicatorlicense,                   "
    sql = sql & "      spraylist.reentryintervaldays,                   "
    sql = sql & "      spraylist.reentryintervalhours,                  "
    sql = sql & "      spraylist.preharvestinterval,                    "
    sql = sql & "      sprayrecord.weather,                             "
    sql = sql & "      spraylist.spraylistid,                           "
    sql = sql & "      spraylist.unitid,                                "
    sql = sql & "      spraylist.maxuseapp,                             "
    sql = sql & "      spraylist.maxuseseason,                          "
    sql = sql & "      sprayrecord.ifprating,                           "
    sql = sql & "      sprayrecord.harvestdate,                         "
    sql = sql & "      sprayrecord.comments,                            "
    sql = sql & "      crops.crop,                                      "
    sql = sql & "      varieties.variety,                               "
    sql = sql & "      varieties2.variety AS variety2,                  "
    sql = sql & "      varieties3.variety AS variety3,                  "
    sql = sql & "      varieties4.variety AS variety4,                  "
    sql = sql & "      growers.growernumber,                            "
    sql = sql & "      growers.additionalgrowernumbers,                 "
    sql = sql & "      growers.growername,                              "
    sql = sql & "      growers.email,                                   "
    sql = sql & "      growers.growerpassword,                          "
    sql = sql & "      growers.address,                                 "
    sql = sql & "      growers.city,                                    "
    sql = sql & "      growers.state,                                   "
    sql = sql & "      growers.zipcode,                                 "
    sql = sql & "      growers.contact,                                 "
    sql = sql & "      growers.telephone1,                              "
    sql = sql & "      growers.telephone2,                              "
    sql = sql & "      growers.fax,                                     "
    sql = sql & "      growers.cell,                                    "
    sql = sql & "      growers.fieldman,                                "
    sql = sql & "      methods.method,                                  "
    sql = sql & "      spraylist.name,                                  "
    sql = sql & "      stages.stage,                                    "
    sql = sql & "      units.unit,                                      "
    sql = sql & "      spraylist.name     AS productnameandformulation, "
    sql = sql & "      spraylist.activeind                              "
	sql = sql & " 	FROM   sprayrecord                                            "
	sql = sql & "        LEFT JOIN packers                                        "
	sql = sql & "          ON sprayrecord.packerid = packers.packerid             "
	sql = sql & "        LEFT JOIN crops                                          "
	sql = sql & "          ON sprayrecord.cropid = crops.cropid                   "
	sql = sql & "        LEFT JOIN varieties                                      "
	sql = sql & "          ON sprayrecord.varietyid1 = varieties.varietyid        "
	sql = sql & "        LEFT JOIN varieties varieties2                           "
	sql = sql & "          ON sprayrecord.varietyid2 = varieties2.varietyid       "
	sql = sql & "        LEFT JOIN varieties varieties3                           "
	sql = sql & "          ON sprayrecord.varietyid3 = varieties3.varietyid       "
	sql = sql & "        LEFT JOIN varieties varieties4                           "
	sql = sql & "          ON sprayrecord.varietyid4 = varieties4.varietyid       "
	sql = sql & "        LEFT JOIN growers                                        "
	sql = sql & "          ON sprayrecord.growerid = growers.growerid             "
	sql = sql & "        LEFT JOIN methods                                        "
	sql = sql & "          ON sprayrecord.methodid = methods.methodid             "
	sql = sql & "        INNER JOIN spraylist                                     "
	sql = sql & "          ON sprayrecord.productid = spraylist.spraylistid       "
	sql = sql & "        INNER JOIN (SELECT sprayyearid                           "
	sql = sql & "                    FROM   sprayyears                            "
	sql = sql & "                    WHERE  active = 1) s_years                   "
	sql = sql & "          ON s_years.sprayyearid = spraylist.sprayyearid         "
	sql = sql & "        LEFT JOIN stages                                         "
	sql = sql & "          ON sprayrecord.stageid = stages.stageid                "
	sql = sql & "        LEFT JOIN units                                          "
	sql = sql & "          ON spraylist.unitid = units.unitid                     "
	sql = sql & " 	WHERE 														  "
IF session("growerid") = 0 THEN
	sql = sql & "1=0" '" SprayRecord.Administrator = '" & session("username") & "'"
ELSE
	cAddNums = session("AdditionalNumbers")
	cAddNums = "'" + cAddNums + "'"
	cAddNums = replace(cAddNums,",","','")

	sql = sql & "  (Growers.GrowerID = " & session("growerid")  & ")" '" OR Growers.GrowerNumber IN (" & cAddNums & "))"
'	sql = sql & "  (Growers.GrowerID = " & session("growerid")  & " OR Growers.GrowerNumber IN (0" & session("AdditionalNumbers") & "))"
END IF
sql = sql & " ORDER BY SprayRecord.CREATEDATE DESC, SprayRecord.SprayRecordID"
'response.write sql
	set GetSprayRecordsByLogin = conn.execute(sql)
end function


' *******************************************************
' ************ GetSprayRecordsByGrower ******************
' *******************************************************
function GetSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate,searchSprayYear,searchCropID,searchVarietyID,searchBartlet)
	sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
	set rs = conn.execute(sql)
	sql = 		"SELECT packers.packernumber,												"
	sql = sql & "       sprayrecord.administrator,											"
	sql = sql & "       sprayrecord.supervisor,												"
	sql = sql & "       sprayrecord.updatedate,												"
	sql = sql & "       sprayrecord.licensenumber,											"
	sql = sql & "       sprayrecord.chemicalsupplier,										"
	sql = sql & "       sprayrecord.recommendedby,											"
	sql = sql & "       sprayrecord.applicator,												"
	sql = sql & "       sprayrecord.sprayrecordid,											"
	sql = sql & "       spraylist.sprayyearid,												"
	sql = sql & "       sprayrecord.overapplicationflag,									"
	sql = sql & "       sprayrecord.growerid,												"
	sql = sql & "       sprayrecord.spraystartdate,											"
	sql = sql & "       sprayrecord.timefinishedspraying,									"
	sql = sql & "       sprayrecord.sprayenddate,											"
	sql = sql & "       sprayrecord.cropid,													"
	sql = sql & "       varietyid1,															"
	sql = sql & "       varietyid2,															"
	sql = sql & "       varietyid3,															"
	sql = sql & "       varietyid4,															"
	sql = sql & "       sprayrecord.bartlet,												"
	sql = sql & "       sprayrecord.stageid,												"
	sql = sql & "       sprayrecord.location,												"
	sql = sql & "       sprayrecord.overseasonflag,											"
	sql = sql & "       sprayrecord.methodid,												"
	sql = sql & "       sprayrecord.acrestreated,											"
	sql = sql & "       sprayrecord.rateacre,												"
	sql = sql & "       sprayrecord.totalmaterialapplied,									"
	sql = sql & "       sprayrecord.applicatorlicense,										"
	sql = sql & "       spraylist.reentryintervaldays,										"
	sql = sql & "       spraylist.reentryintervalhours,										"
	sql = sql & "       spraylist.preharvestinterval,										"
	sql = sql & "       sprayrecord.weather,												"
	sql = sql & "       spraylist.spraylistid,												"
	sql = sql & "       spraylist.unitid,													"
	sql = sql & "       spraylist.maxuseapp,												"
	sql = sql & "       spraylist.maxuseseason,												"
	sql = sql & "       sprayrecord.ifprating,												"
	sql = sql & "       sprayrecord.harvestdate,											"
	sql = sql & "       sprayrecord.comments,												"
	sql = sql & "       crops.crop,															"
	sql = sql & "       varieties.variety,													"
	sql = sql & "       varieties2.variety AS variety2,										"
	sql = sql & "       varieties3.variety AS variety3,										"
	sql = sql & "       varieties4.variety AS variety4,										"
	sql = sql & "       growers.growernumber,												"
	sql = sql & "       growers.additionalgrowernumbers,									"
	sql = sql & "       growers.growername,													"
	sql = sql & "       growers.email,														"
	sql = sql & "       growers.growerpassword,												"
	sql = sql & "       growers.address,													"
	sql = sql & "       growers.city,														"
	sql = sql & "       growers.state,														"
	sql = sql & "       growers.zipcode,													"
	sql = sql & "       growers.contact,													"
	sql = sql & "       growers.telephone1,													"
	sql = sql & "       growers.telephone2,													"
	sql = sql & "       growers.fax,														"
	sql = sql & "       growers.cell,														"
	sql = sql & "       growers.fieldman,													"
	sql = sql & "       methods.method,														"
	sql = sql & "       spraylist.name,														"
	sql = sql & "       stages.stage,														"
	sql = sql & "       units.unit,															"
	sql = sql & "       spraylist.name     AS productnameandformulation,					"
	sql = sql & "       spraylist.activeind,												"
	sql = sql & "       spraylist.phi,														"
	sql = sql & "       spraylist.rei,														"
	sql = sql & "       NULL as Targets														"
	sql = sql & "FROM   sprayrecord															"
	sql = sql & "       LEFT JOIN packers													"
	sql = sql & "         ON sprayrecord.packerid = packers.packerid						"
	sql = sql & "       LEFT JOIN crops														"
	sql = sql & "         ON sprayrecord.cropid = crops.cropid								"
	sql = sql & "       LEFT JOIN varieties													"
	sql = sql & "         ON sprayrecord.varietyid1 = varieties.varietyid					"
	sql = sql & "       LEFT JOIN varieties varieties2										"
	sql = sql & "         ON sprayrecord.varietyid2 = varieties2.varietyid					"
	sql = sql & "       LEFT JOIN varieties varieties3										"
	sql = sql & "         ON sprayrecord.varietyid3 = varieties3.varietyid					"
	sql = sql & "       LEFT JOIN varieties varieties4										"
	sql = sql & "         ON sprayrecord.varietyid4 = varieties4.varietyid					"
	sql = sql & "       LEFT JOIN growers													"
	sql = sql & "         ON sprayrecord.growerid = growers.growerid						"
	sql = sql & "       LEFT JOIN methods													"
	sql = sql & "         ON sprayrecord.methodid = methods.methodid						"
	sql = sql & "       LEFT JOIN spraylist													"
	sql = sql & "         ON sprayrecord.productid = spraylist.spraylistid					"
	sql = sql & "       LEFT JOIN stages													"
	sql = sql & "         ON sprayrecord.stageid = stages.stageid							"
	sql = sql & "       LEFT JOIN units														"
	sql = sql & "         ON spraylist.unitid = units.unitid								"
	sql = sql & "WHERE  NOT sprayrecord.growerid IS NULL									"
	sql = sql & "       AND NOT growers.growerid IS NULL									"
	sql = sql & "        AND sprayyearid = "  &  searchSprayYear
	if searchGrower <> "0" and searchGrower <> "" then
		sql = sql & " AND SprayRecord.GrowerID in (" & searchGrower  & ")"
	end if
	if searchHighSprayDate <> "" then
		sql = sql & " AND SprayRecord.SprayStartDate <=  '" & CDate(searchHighSprayDate) & "'"
	end if
	if searchLowSprayDate <> "" then
		sql = sql & " AND SprayRecord.SprayStartDate >= '" & CDate(searchLowSprayDate) & "'"
	end if
	if searchHighHarvestDate <> "" then
		sql = sql & " AND SprayRecord.HarvestDate < '" & CDate(searchHighHarvestDate) & "'"
	end if
	if searchLowHarvestDate <> "" then
		sql = sql & " AND SprayRecord.HarvestDate > '" & CDate(searchLowHarvestDate) & "'"
	end if
	if searchCropID <> "" then
		sql = sql & " AND (SprayRecord.CropID in (" & searchCropID & ") )"
	end if
	if searchBartlet <> "" then
		if searchBartlet = 1 then
			sql = sql & " AND (SprayRecord.Bartlet = 1)"
		else
			sql = sql & " AND (SprayRecord.Bartlet = 0)"
		end if
	end if
	if searchVarietyID <> "" and searchVarietyID <> "0" then
		sql = sql & " AND (SprayRecord.VarietyID1 in (" & searchVarietyID & ") OR SprayRecord.VarietyID2 in (" & searchVarietyID & ") OR SprayRecord.VarietyID3 in (" & searchVarietyID & ") OR SprayRecord.VarietyID4 in (" & searchVarietyID & "))"
	end if
	if session("packerid") <> 0 then
		sql = sql & " AND SprayRecord.packerid = " & session("packerid")
	end if
	sql = sql &  " ORDER BY SprayRecord.GrowerID,SprayRecord.SprayStartDate ASC, SprayRecord.SprayRecordID"

'response.write sql
'response.end

	set GetSprayRecordsByGrower = conn.execute(sql)
end function


' *******************************************************
' ************ GetCountSprayRecordsByGrower *************
' *******************************************************
function GetCountSprayRecordsByGrower(searchGrower,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate,searchSprayYear,searchCropID,searchVarietyID,searchBartlet)

	sql = "SELECT Count(SprayRecord.SprayRecordID) FROM ((((((SprayRecord LEFT JOIN Crops ON SprayRecord.CropID = Crops.CropID) LEFT JOIN Growers ON SprayRecord.GrowerID = Growers.GrowerID) LEFT JOIN Methods ON SprayRecord.MethodID = Methods.MethodID) LEFT JOIN SprayList ON SprayRecord.ProductID = SprayList.SprayListID"
	sql = sql & ") LEFT JOIN Stages ON SprayRecord.StageID = Stages.StageID) LEFT JOIN Targets ON SprayRecord.TargetID = Targets.TargetID) LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE (not SprayRecord.GrowerID is null AND not Growers.GrowerID is null)  AND SprayYearID = " &  searchSprayYear
if searchGrower <> "0" and searchGrower <> "" then
	sql = sql & " AND SprayRecord.GrowerID in (" & searchGrower  & ")"
end if
if searchHighSprayDate <> "" then
	sql = sql & " AND SprayRecord.SprayStartDate <=  '" & DateValue(searchHighSprayDate) & "'"
end if
if searchLowSprayDate <> "" then
	sql = sql & " AND SprayRecord.SprayStartDate >= '" & DateValue(searchLowSprayDate) & "'"
end if
if searchHighHarvestDate <> "" then
	sql = sql & " AND SprayRecord.HarvestDate <= '" & DateValue(searchHighHarvestDate) & "'"
end if
if searchLowHarvestDate <> "" then
	sql = sql & " AND SprayRecord.HarvestDate >= '" & DateValue(searchLowHarvestDate) & "'"
end if
if searchCropID <> "" then
	sql = sql & " AND SprayRecord.CropID in (" & searchCropID & ") "
end if
if searchBartlet <> "" then
	if searchBartlet = 1 then
		sql = sql & " AND (SprayRecord.Bartlet = 1)"
	else
		sql = sql & " AND (SprayRecord.Bartlet = 0)"
	end if
end if

if searchVarietyID <> "" and searchVarietyID <> "0" then
	sql = sql & " AND (SprayRecord.VarietyID1 in (" & searchVarietyID & ") OR SprayRecord.VarietyID2 in (" & searchVarietyID & ") OR SprayRecord.VarietyID3 in (" & searchVarietyID & ") OR SprayRecord.VarietyID4 in (" & searchVarietyID & "))"
end if
if session("packerid") <> 0 then
	sql = sql & " AND SprayRecord.packerid = " & session("packerid")
end if

'response.write sql
'response.end

	set GetCountSprayRecordsByGrower = conn.execute(sql)
end function

' *******************************************************
' ************ GetCountSprayRecordsBySearch *************
' *******************************************************
function GetCountSprayRecordsBySearch(searchGrower,searchCrop,searchVariety,searchBartlet,searchStage,searchLocation,searchMethod,searchProduct,searchTarget,searchUpdateBy,searchOverApplication,searchOverSeason,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate)


	'# added 3-Apr-2011
	dim yr
	if request.form("SprayYear")<>"" then
		yr=request.form("SprayYear")
	elseif request.querystring("searchSprayYear")<>"" then
		yr=request.querystring("searchSprayYear")
	else
		sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
		set rs = conn.execute(sql)
		yr = rs.collect(0)
	end if

	sql = "SELECT Count( SprayRecord.SprayRecordID) as recordCount FROM ((((((SprayRecord LEFT JOIN Crops ON SprayRecord.CropID = Crops.CropID) LEFT JOIN Growers ON SprayRecord.GrowerID = Growers.GrowerID) LEFT JOIN Methods ON SprayRecord.MethodID = Methods.MethodID) LEFT JOIN SprayList ON SprayRecord.ProductID = SprayList.SprayListID"
	sql = sql & ") LEFT JOIN Stages ON SprayRecord.StageID = Stages.StageID) LEFT JOIN Targets ON SprayRecord.TargetID = Targets.TargetID) LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE (1=1) AND SprayYearID = " & yr

if searchGrower <> "" then
	sql = sql & " AND SprayRecord.GrowerID in (" & searchGrower & ")"
end if
if searchCrop <> "" then
	sql = sql & " AND (SprayRecord.CropID in (" & searchCrop & ") )"
end if
if searchVariety <> "" then
	sql = sql & " AND (SprayRecord.VarietyID1 in (" & searchVariety & ") OR SprayRecord.VarietyID2 in (" & searchVariety & ") OR SprayRecord.VarietyID3 in (" & searchVariety & ") OR SprayRecord.VarietyID4 in (" & searchVariety & "))"
end if
if searchBartlet <> "" then
	if searchBartlet = 1 then
		sql = sql & " AND SprayRecord.Bartlet = 1"
	else
		sql = sql & " AND SprayRecord.Bartlett = 0"
	end if
end if
if searchStage <> "" then
	sql = sql & " AND SprayRecord.StageID  in (" & searchStage & ")"
end if
if searchLocation <> "" then
	sql = sql & " AND SprayRecord.Location like '%" & searchLocation & "%'"
end if
if searchMethod <> "" then
	sql = sql & " AND SprayRecord.MethodID in (" & searchMethod & ")"
end if
if searchProduct <> "" then
	sql = sql & " AND SprayRecord.ProductID in (" & searchProduct & ")"
end if
if searchTarget <> "" then
	sql = sql & " AND SprayRecord.TargetID in (" & searchTarget & ")"
end if
if searchUpdateBy <> "" then
	sql = sql & " AND SprayRecord.Administrator = '" & searchUpdateBy & "'"
end if
if searchOverApplication <> "" then
	sql = sql & " AND SprayRecord.OverApplicationFlag = " & searchOverApplication
end if
if searchOverSeason <> "" then
	sql = sql & " AND SprayRecord.OverSeasonFlag = " & searchOverSeason
end if
if searchHighSprayDate <> "" then
	sql = sql & " AND SprayRecord.SprayStartDate <=  '" & DateValue(searchHighSprayDate) & "'"
end if
if searchLowSprayDate <> "" then
	sql = sql & " AND SprayRecord.SprayStartDate >= '" & DateValue(searchLowSprayDate) & "'"
end if
if searchHighHarvestDate <> "" then
	sql = sql & " AND SprayRecord.HarvestDate <= '" & DateValue(searchHighHarvestDate) & "'"
end if
if searchLowHarvestDate <> "" then
	sql = sql & " AND SprayRecord.HarvestDate >= '" & DateValue(searchLowHarvestDate) & "'"
end if
IF session("growerid") <> 0 THEN
	cAddNums = session("AdditionalNumbers")
	cAddNums = "'" + cAddNums + "'"
	cAddNums = replace(cAddNums,",","','")

'	sql = sql & " AND (Growers.GrowerID = " & session("growerid")  & " OR Growers.GrowerNumber IN (0" & session("AdditionalNumbers") & "))"
	sql = sql & " AND (Growers.GrowerID = " & session("growerid")  & ")" 'OR Growers.GrowerNumber IN (" & cAddNums & "))"

END IF
if session("packerid") <> 0 then
	sql = sql & " AND SprayRecord.packerid = " & session("packerid")
end if

	session("LastSpraySearchCount")=sql
	set GetCountSprayRecordsBySearch = conn.execute(sql)
end function

function GetLastSpraySearchCount()

	if session("LastSpraySearchCount")>"" then
		set GetLastSpraySearchCount= conn.execute(session("LastSpraySearchCount"))
	else
		set GetLastSpraySearchCount = nothing
	end if


end function

function GetLastSpraySearch()

	if session("LastSpraySearch")>"" then
		set GetLastSpraySearch= conn.execute(session("LastSpraySearch"))
	else
		set GetLastSpraySearch = nothing
	end if


end function

' *******************************************************
' ************ GetSprayRecordsBySearch ******************
' *******************************************************
function GetSprayRecordsBySearch(searchGrower,searchCrop,searchVariety,searchBartlet,searchStage,searchLocation,searchMethod,searchProduct,searchTarget,searchUpdateBy,searchOverApplication,searchOverSeason,searchHighSprayDate,searchLowSprayDate,searchHighHarvestDate,searchLowHarvestDate)

	'# added 3-Apr-2011
	dim yr
	if request.form("SprayYear")<>"" then
		yr=request.form("SprayYear")
	elseif request.querystring("searchSprayYear")<>"" then
		yr=request.querystring("searchSprayYear")
	else
		sql = "SELECT SprayYearID FROM SprayYears WHERE Active = 1"
		set rs = conn.execute(sql)
		yr = rs.collect(0)
	end if

'	sql = "SELECT  SprayRecord.Administrator, SprayRecord.Supervisor, SprayRecord.UpdateDAte,SprayRecord.LicenseNumber,SprayRecord.ChemicalSupplier,SprayRecord.RecommendedBy,SprayRecord.Applicator,SprayRecord.SprayRecordID,SprayList.SprayYearID,SprayRecord.OverApplicationFlag,SprayRecord.GrowerID,SprayRecord.SprayStartDate,SprayRecord.TimeFinishedSpraying,SprayRecord.SprayEndDate,SprayRecord.CropID,VarietyID1,VarietyID2,VarietyID3,VarietyID4,SprayRecord.Bartlet,SprayRecord.StageID,SprayRecord.Location,SprayRecord.OverSeasonFlag,SprayRecord.MethodID,SprayRecord.AcresTreated,SprayRecord.RateAcre,SprayRecord.TotalMaterialApplied,SprayRecord.ApplicatorLicense,SprayList.ReentryIntervalDays,SprayList.ReentryIntervalHours,SprayList.PreharvestInterval,SprayRecord.Weather,SprayList.SprayListID,SprayList.UnitID,SprayList.MaxUseApp,SprayList.MaxUseSeason,SprayRecord.IFPRating,SprayRecord.TargetID,SprayRecord.HarvestDate,SprayRecord.Comments,Crops.Crop,Varieties.Variety,Varieties2.Variety as Variety2,Varieties3.Variety as Variety3, Varieties4.Variety as Variety4,Growers.GrowerNumber,Growers.AdditionalGrowerNumbers,Growers.GrowerName,Growers.Email,Growers.GrowerPassword,Growers.Address,Growers.City,Growers.State,Growers.ZipCode,Growers.Contact,Growers.Telephone1,Growers.Telephone2,Growers.Fax,Growers.Cell,Growers.Fieldman,Methods.Method,SprayList.Name,Stages.Stage,Targets.Target,Units.Unit,SprayList.Name as ProductNameAndFormulation ,Spraylist.ActiveInd"
'	sql = sql & " FROM ((((((((((SprayRecord LEFT JOIN Crops ON SprayRecord.CropID = Crops.CropID) LEFT JOIN Varieties ON SprayRecord.VarietyID1 = Varieties.VarietyID) LEFT JOIN Varieties Varieties2 ON SprayRecord.VarietyID2 = Varieties2.VarietyID) LEFT JOIN Varieties Varieties3 ON SprayRecord.VarietyID3 = Varieties3.VarietyID) LEFT JOIN Varieties Varieties4 ON SprayRecord.VarietyID4 = Varieties4.VarietyID)  LEFT JOIN Growers ON SprayRecord.GrowerID = Growers.GrowerID) LEFT JOIN Methods ON SprayRecord.MethodID = Methods.MethodID) LEFT JOIN SprayList ON SprayRecord.ProductID = SprayList.SprayListID"
'	sql = sql & ") LEFT JOIN Stages ON SprayRecord.StageID = Stages.StageID) LEFT JOIN Targets ON SprayRecord.TargetID = Targets.TargetID) LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE (not IsNull(SprayRecord.GrowerID) and not IsNull(Growers.GrowerID))  AND SprayYearID = "  &   rs(0)

	sql = "SELECT  Packers.PackerNumber,SR.Administrator, SR.Supervisor, SR.UpdateDAte, SR.LicenseNumber, " & _
		"SR.ChemicalSupplier, SR.RecommendedBy, SR.Applicator, SR.SprayRecordID, " & _
		"SprayList.SprayYearID, SR.OverApplicationFlag, SR.GrowerID, SR.SprayStartDate, " & _
		"SR.TimeFinishedSpraying, SR.SprayEndDate, SR.CropID, VarietyID1, VarietyID2, VarietyID3, VarietyID4, " & _
		"SR.Bartlet, SR.StageID, SR.Location, SR.OverSeasonFlag, SR.MethodID, SR.AcresTreated, " & _
		"SR.RateAcre, SR.TotalMaterialApplied, SR.ApplicatorLicense, SR.PURS_Reported, " & _
		"SprayList.ReentryIntervalDays,SprayList.ReentryIntervalHours,SprayList.PreharvestInterval,SR.Weather,SprayList.SprayListID,SprayList.UnitID,SprayList.MaxUseApp,SprayList.MaxUseSeason,SR.IFPRating,SR.TargetID,SR.HarvestDate,SR.Comments,Crops.Crop,Varieties.Variety,Varieties2.Variety as Variety2,Varieties3.Variety as Variety3, Varieties4.Variety as Variety4, " & _
		"Growers.GrowerNumber,Growers.AdditionalGrowerNumbers,Growers.GrowerName,Growers.Email,Growers.GrowerPassword,Growers.Address,Growers.City,Growers.State,Growers.ZipCode,Growers.Contact,Growers.Telephone1,Growers.Telephone2,Growers.Fax,Growers.Cell,Growers.Fieldman,Methods.Method,SprayList.Name,Stages.Stage,Targets.Target,Units.Unit,SprayList.Name as ProductNameAndFormulation ,Spraylist.ActiveInd"
	sql = sql & " FROM (((((((((((SprayRecord SR LEFT JOIN Packers ON SR.PackerID = Packers.PackerID) LEFT JOIN Crops ON SR.CropID = Crops.CropID) LEFT JOIN Varieties ON SR.VarietyID1 = Varieties.VarietyID) LEFT JOIN Varieties Varieties2 ON SR.VarietyID2 = Varieties2.VarietyID) LEFT JOIN Varieties Varieties3 ON SR.VarietyID3 = Varieties3.VarietyID) LEFT JOIN Varieties Varieties4 ON SR.VarietyID4 = Varieties4.VarietyID)  LEFT JOIN Growers ON SR.GrowerID = Growers.GrowerID) LEFT JOIN Methods ON SR.MethodID = Methods.MethodID) LEFT JOIN SprayList ON SR.ProductID = SprayList.SprayListID"
	sql = sql & ") LEFT JOIN Stages ON SR.StageID = Stages.StageID) LEFT JOIN Targets ON SR.TargetID = Targets.TargetID) LEFT JOIN Units ON SprayList.UnitID = Units.UnitID WHERE (not SR.GrowerID is null and not Growers.GrowerID is null)  AND SprayYearID = "  & yr

'if searchGrower <> "" then
'	sql = sql & " AND SprayRecord.GrowerID in (" & searchGrower & ")"
'end if
'if searchCrop <> "" then
'	sql = sql & " AND (SprayRecord.CropID in (" & searchCrop & ") )"
'end if
'if searchVariety <> "" then
'	sql = sql & " AND (SprayRecord.VarietyID1 in (" & searchVariety & ") OR SprayRecord.VarietyID2 in (" & searchVariety & ") OR SprayRecord.VarietyID3 in (" & searchVariety & ") OR SprayRecord.VarietyID4 in (" & searchVariety & "))"
'end if
'if searchBartlet <> "" then
'	if searchBartlet = 1 then
'		sql = sql & " AND SprayRecord.Bartlet = " & true
'	else
'		sql = sql & " AND SprayRecord.Bartlett = " & false
'	end if
'end if
'if searchStage <> "" then
'	sql = sql & " AND SprayRecord.StageID  in (" & searchStage & ")"
'end if
'if searchLocation <> "" then
'	sql = sql & " AND SprayRecord.Location like '*" & searchLocation & "*'"
'end if
'if searchMethod <> "" then
'	sql = sql & " AND SprayRecord.MethodID in (" & searchMethod & ")"
'end if
'if searchProduct <> "" then
'	sql = sql & " AND SprayRecord.ProductID in (" & searchProduct & ")"
'end if
'if searchTarget <> "" then
'	sql = sql & " AND SprayRecord.TargetID in (" & searchTarget & ")"
'end if
'if searchUpdateBy <> "" then
'	sql = sql & " AND SprayRecord.Administrator = '" & searchUpdateBy & "'"
'end if
'if searchOverApplication <> "" then
'	sql = sql & " AND SprayRecord.OverApplicationFlag = " & searchOverApplication
'end if
'if searchOverSeason <> "" then
'	sql = sql & " AND SprayRecord.OverSeasonFlag = " & searchOverSeason
'end if
'if searchHighSprayDate <> "" then
'	sql = sql & " AND SprayRecord.SprayStartDate <=  '" & DateValue(searchHighSprayDate) & "'"
'end if
'if searchLowSprayDate <> "" then
'	sql = sql & " AND SprayRecord.SprayStartDate >= '" & DateValue(searchLowSprayDate) & "'"
'end if
'if searchHighHarvestDate <> "" then
'	sql = sql & " AND SprayRecord.HarvestDate <= '" & DateValue(searchHighHarvestDate) & "'"
'end if
'if searchLowHarvestDate <> "" then
'	sql = sql & " AND SprayRecord.HarvestDate >= '" & DateValue(searchLowHarvestDate) & "'"
'end if
'IF session("growerid") <> 0 THEN
'	cAddNums = session("AdditionalNumbers")
'	cAddNums = "'" + cAddNums + "'"
'	cAddNums = replace(cAddNums,",","','")

''	sql = sql & " AND  (Growers.GrowerID = " & session("growerid")  & " OR Growers.GrowerNumber IN (0" & session("AdditionalNumbers") & "))"
'	sql = sql & " AND  (Growers.GrowerID = " & session("growerid")  & " OR Growers.GrowerNumber IN (" & cAddNums & "))"
'END IF
'sql = sql  & " ORDER BY SprayRecord.CREATEDATE DESC, SprayRecord.SprayRecordID"

if searchGrower <> "" then
	sql = sql & " AND SR.GrowerID in (" & searchGrower & ")"
end if
if searchCrop <> "" then
	sql = sql & " AND (SR.CropID in (" & searchCrop & ") )"
end if
if searchVariety <> "" then
	sql = sql & " AND (SR.VarietyID1 in (" & searchVariety & ") OR SR.VarietyID2 in (" & searchVariety & ") OR SR.VarietyID3 in (" & searchVariety & ") OR SR.VarietyID4 in (" & searchVariety & "))"
end if
if searchBartlet <> "" then
	if searchBartlet = 1 then
		sql = sql & " AND SR.Bartlet = 1"
	else
		sql = sql & " AND SR.Bartlett = 0"
	end if
end if
if searchStage <> "" then
	sql = sql & " AND SR.StageID  in (" & searchStage & ")"
end if
if searchLocation <> "" then
	sql = sql & " AND SR.Location like '%" & searchLocation & "%'"
end if
if searchMethod <> "" then
	sql = sql & " AND SR.MethodID in (" & searchMethod & ")"
end if
if searchProduct <> "" then
	sql = sql & " AND SR.ProductID in (" & searchProduct & ")"
end if
if searchTarget <> "" then
	sql = sql & " AND SR.TargetID in (" & searchTarget & ")"
end if
if searchUpdateBy <> "" then
	sql = sql & " AND SR.Administrator = '" & searchUpdateBy & "'"
end if
if searchOverApplication <> "" then
	sql = sql & " AND SR.OverApplicationFlag = " & searchOverApplication
end if
if searchOverSeason <> "" then
	sql = sql & " AND SR.OverSeasonFlag = " & searchOverSeason
end if
if searchHighSprayDate <> "" then
	sql = sql & " AND SR.SprayStartDate <=  '" & DateValue(searchHighSprayDate) & "'"
end if
if searchLowSprayDate <> "" then
	sql = sql & " AND SR.SprayStartDate >= '" & DateValue(searchLowSprayDate) & "'"
end if
if searchHighHarvestDate <> "" then
	sql = sql & " AND SR.HarvestDate <= '" & DateValue(searchHighHarvestDate) & "'"
end if
if searchLowHarvestDate <> "" then
	sql = sql & " AND SR.HarvestDate >= '" & DateValue(searchLowHarvestDate) & "'"
end if
IF session("growerid") <> 0 THEN
	cAddNums = session("AdditionalNumbers")
	cAddNums = "'" + cAddNums + "'"
	cAddNums = replace(cAddNums,",","','")

'	sql = sql & " AND  (Growers.GrowerID = " & session("growerid")  & " OR Growers.GrowerNumber IN (0" & session("AdditionalNumbers") & "))"
	sql = sql & " AND  (Growers.GrowerID = " & session("growerid")  & ")" 'OR Growers.GrowerNumber IN (" & cAddNums & "))"
END IF
if session("packerid") <> 0 then
	sql = sql & " AND sr.packerid = " & session("packerid")
end if
sql = sql  & " ORDER BY SR.CREATEDATE DESC, SR.SprayRecordID"

	'response.write sql
	'response.end

	session("LastSpraySearch")=sql
	set GetSprayRecordsBySearch = conn.execute(sql)
end function




' *******************************************************
' ************ GetGrowersLocations **********************
' *******************************************************
function GetGrowersLocations(GrowerID)
	IF (GrowerID = "") THEN
		GrowerID = 0
	END IF
	sql = "SELECT Distinct Location FROM SprayRecord WHERE Location <> '' AND GrowerID = " & GrowerID
	set GetGrowersLocations = conn.execute(sql)
end function

' *******************************************************
' ************ GetGrowersSupervisors ********************
' *******************************************************
function GetGrowersSupervisors(GrowerID)
	IF (GrowerID = "") THEN
		GrowerID = 0
	END IF
	sql = "SELECT Distinct Supervisor FROM SprayRecord WHERE Supervisor <> '' AND GrowerID = " & GrowerID
	set GetGrowersSupervisors = conn.execute(sql)
end function

' *******************************************************
' ************ GetGrowersSupervisorLicenses *************
' *******************************************************
function GetGrowersSupervisorLicenses(GrowerID)
	IF (GrowerID = "") THEN
		GrowerID = 0
	END IF
	sql = "SELECT Distinct LicenseNumber FROM SprayRecord WHERE LicenseNumber <> '' AND GrowerID = " & GrowerID
	set GetGrowersSupervisorLicenses = conn.execute(sql)
end function

' *******************************************************
' ************ GetGrowersApplicators ********************
' *******************************************************
function GetGrowersApplicators(GrowerID)
	IF (GrowerID = "") THEN
		GrowerID = 0
	END IF
	sql = "SELECT Distinct Applicator FROM SprayRecord WHERE Applicator <> '' AND GrowerID = " & GrowerID
	set GetGrowersApplicators = conn.execute(sql)
end function

' *******************************************************
' ************ GetGrowersApplicatorLicenses *************
' *******************************************************
function GetGrowersApplicatorLicenses(GrowerID)
	IF (GrowerID = "") THEN
		GrowerID = 0
	END IF
	sql = "SELECT Distinct ApplicatorLicense FROM SprayRecord WHERE ApplicatorLicense <> '' AND GrowerID = " & GrowerID
	set GetGrowersApplicatorLicenses = conn.execute(sql)
end function

' *******************************************************
' ************ GetGrowersChemicalSuppliers **************
' *******************************************************
function GetGrowersChemicalSuppliers(GrowerID)
	IF (GrowerID = "") THEN
		GrowerID = 0
	END IF
	sql = "SELECT Distinct ChemicalSupplier FROM SprayRecord WHERE ChemicalSupplier <> '' AND GrowerID = " & GrowerID
	set GetGrowersChemicalSuppliers = conn.execute(sql)
end function

' *******************************************************
' ************ GetGrowersRecommendedBy ******************
' *******************************************************
function GetGrowersRecommendedBy(GrowerID)
	IF (GrowerID = "") THEN
		GrowerID = 0
	END IF
	sql = "SELECT Distinct RecommendedBy FROM SprayRecord WHERE RecommendedBy <> '' AND GrowerID = " & GrowerID
	set GetGrowersRecommendedBy = conn.execute(sql)
end function

' *******************************************************
' ************ GetSeasonQty *****************************
' *******************************************************
function GetSeasonQty(SprayDate,SprayListID,GrowerID,Location,SprayYearID)
	if not IsDate(SprayDate) THEN
		SprayDate = now()
	END IF
	sql = "SELECT Sum(RateAcre*AcresTreated),Sum(AcresTreated) AS ApplicationSeason FROM SprayRecord LEFT JOIN SprayList ON SprayRecord.ProductID = SprayList.SprayListID WHERE SprayRecord.GrowerID= " & GrowerID & " AND SprayRecord.Location = '" & EscapeQuotes(Location) & "' AND  SprayRecord.ProductID= " & SprayListID & " AND SprayStartDate <= '" & DateValue(SprayDate) & "' and SprayYearID = " & SprayYearID
'response.write sql
	set GetSeasonQty = conn.execute(sql)
end function

' *******************************************************
' ************ GetRecordCountBySprayListID **************
' *******************************************************
function GetRecordCountBySprayListID(SprayListID)
	sql = "SELECT Count(SprayRecordID) AS reccount FROM SprayRecord WHERE SprayRecord.ProductID= " & SprayListID
'response.write sql
	set GetRecordCountBySprayListID = conn.execute(sql)
end function

' *******************************************************
' ************ DeleteSprayRecord ************************
' *******************************************************
function DeleteSprayRecord(SprayRecordID)
	sql = "DELETE FROM SprayRecord WHERE SprayRecordID = " & SprayRecordID
	conn.execute sql, , 129
end function

' *******************************************************
' ************ InsertSprayRecord ************************
' *******************************************************
function InsertSprayRecord(GrowerID,SprayStartDate,TimeFinishedSpraying,SprayEndDate,CropID,VarietyID,Bartlet,StageID,Location,MethodID,AcresTreated,RateAcre,SprayListID,IFPRating,TargetID,HarvestDate,Comments,Weather,Applicator,ApplicatorLicense,Administrator,Supervisor,LicenseNumber,ChemicalSupplier,RecommendedBy)

	arrayVarieties = Split(VarietyID,",")
response.write("<br>--" & VarietyID & "<br>")
	VarietyID1 = 0
	VarietyID2 = 0
	VarietyID3 = 0
	VarietyID4 = 0
	IF IsArray(arrayVarieties) THEN
		IF Ubound(arrayVarieties) >= 0 THEN
			VarietyID1 = arrayVarieties(0)
		END IF
		IF Ubound(arrayVarieties) >= 1 THEN
			VarietyID2 = arrayVarieties(1)
		END IF
		IF Ubound(arrayVarieties) >= 2 THEN
			VarietyID3 = arrayVarieties(2)
		END IF
		IF Ubound(arrayVarieties) >= 3 THEN
			VarietyID4 = arrayVarieties(3)
		END IF
	END IF



	IF Bartlet = "" THEN
		Bartlet = 0
	END IF
	sql = "INSERT INTO SprayRecord(GrowerID,SprayStartDate,TimeFinishedSpraying"
	if SprayEndDate <> "" then
		sql = sql & ",SprayEndDate"
	end if
	sql = sql & ",CropID,VarietyID1,VarietyID2,VarietyID3,VarietyID4,Bartlet,StageID,Location,MethodID,AcresTreated,RateAcre,ProductID,IFPRating,TargetID"
	if HarvestDate <> "" then
		sql = sql & ",HarvestDate"
	end if
	sql = sql & ",Comments,Administrator,Weather,Applicator,ApplicatorLicense,Supervisor,LicenseNumber,ChemicalSupplier,RecommendedBy) VALUES ("
sql = sql  & GrowerID
sql = sql & ",'" & CDate(SprayStartDate) & "','"  & TimeFinishedSpraying & "'"
if SprayEndDate <> "" then
	sql = sql & ",'" & CDate(SprayEndDate) & "'"
end if
sql = sql &  "," & CropID &  "," & VarietyID1 &  "," & VarietyID2 &  "," & VarietyID3 &  "," & VarietyID4 & "," & Bartlet & "," & StageID
sql = sql & ",'" & RemoveQuotes(Location) & "'," & MethodID & "," & AcresTreated & "," & RateAcre &  "," & SprayListID
sql = sql & ",'" & EscapeQuotes(IFPRating) & "'," & TargetID
if HarvestDate <> "" then
	sql = sql & ",'" & DateValue(HarvestDate) & " '"
end if
sql = sql & ",'" & EscapeQuotes(Comments) & "'"
sql = sql & ",'" & Session("username") & "'"
sql = sql & ",'" & RemoveQuotes(Weather) & "'"
sql = sql & ",'" & RemoveQuotes(Applicator) & "'"
sql = sql & ",'" & RemoveQuotes(ApplicatorLicense) & "'"
sql = sql & ",'" & RemoveQuotes(Supervisor) & "'"
sql = sql & ",'" & RemoveQuotes(LicenseNumber) & "'"
sql = sql & ",'" & RemoveQuotes(ChemicalSupplier) & "'"
sql = sql & ",'" & RemoveQuotes(RecommendedBy) & "'"
sql = sql & ")"
response.write sql
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(SprayRecordID) AS insertid FROM SprayRecord"
	set rs = conn.execute(sql)
	newID = rs(0)

	sql = "SELECT Sum(RateAcre) AS OverAppSeason FROM SprayRecord LEFT JOIN SprayList ON SprayRecord.ProductID = SprayList.SprayListID WHERE SprayRecord.GrowerID= " & GrowerID & " AND SprayRecord.Location = '" & EscapeQuotes(Location) & "' AND  SprayRecord.ProductID= " & SprayListID

	InsertSprayRecord = newID
end Function

' *******************************************************
' ************ InsertSprayRecord2 ***********************
' rem weather text stored in SprayRecord now.
' *******************************************************
function InsertSprayRecord2(PackerID,GrowerID,SprayStartDate,TimeFinishedSpraying,SprayEndDate,CropID,VarietyID,Bartlet,StageID,Location,MethodID,AcresTreated,RateAcre,SprayListID,IFPRating,ArrayTargetIDs,HarvestDate,Comments,Weather,Applicator,ApplicatorLicense,Administrator,Supervisor,LicenseNumber,ChemicalSupplier,RecommendedBy)

	arrayVarieties = Split(VarietyID,",")
response.write("<br>--" & VarietyID & "<br>")
	VarietyID1 = 0
	VarietyID2 = 0
	VarietyID3 = 0
	VarietyID4 = 0
	IF IsArray(arrayVarieties) THEN
		IF Ubound(arrayVarieties) >= 0 THEN
			VarietyID1 = arrayVarieties(0)
		END IF
		IF Ubound(arrayVarieties) >= 1 THEN
			VarietyID2 = arrayVarieties(1)
		END IF
		IF Ubound(arrayVarieties) >= 2 THEN
			VarietyID3 = arrayVarieties(2)
		END IF
		IF Ubound(arrayVarieties) >= 3 THEN
			VarietyID4 = arrayVarieties(3)
		END IF
	END IF

	IF Bartlet = "" THEN
		Bartlet = 0
	END IF
	sql = "SET NOCOUNT ON INSERT INTO SprayRecord(PackerID,GrowerID,SprayStartDate,TimeFinishedSpraying"
	if SprayEndDate <> "" then
		sql = sql & ",SprayEndDate"
	end if
	sql = sql & ",CropID,VarietyID1,VarietyID2,VarietyID3,VarietyID4,Bartlet,StageID,Location,MethodID,AcresTreated,RateAcre,ProductID,IFPRating"
	if HarvestDate <> "" then
		sql = sql & ",HarvestDate"
	end if
	sql = sql & ",Comments,Administrator,Weather,Applicator,ApplicatorLicense,Supervisor,LicenseNumber,ChemicalSupplier,RecommendedBy) VALUES ("
	sql = sql & PackerID & "," & GrowerID
	sql = sql & ",'" & CDate(SprayStartDate) & "','"  & TimeFinishedSpraying & "'"
	if SprayEndDate <> "" then
		sql = sql & ",'" & CDate(SprayEndDate) & "'"
	end if
	sql = sql &  "," & CropID &  "," & VarietyID1 &  "," & VarietyID2 &  "," & VarietyID3 &  "," & VarietyID4 & "," & Bartlet & "," & StageID
	sql = sql & ",'" & RemoveQuotes(Location) & "'," & MethodID & "," & AcresTreated & "," & RateAcre &  "," & SprayListID
	sql = sql & ",'" & EscapeQuotes(IFPRating) & "'"
	'sql = sql & ",'" & EscapeQuotes(IFPRating) & "'," & TargetID
	if HarvestDate <> "" then
		sql = sql & ",'" & DateValue(HarvestDate) & " '"
	end if
	sql = sql & ",'" & EscapeQuotes(Comments) & "'"
	sql = sql & ",'" & Session("username") & "'"
	sql = sql & ",'" & RemoveQuotes(Weather) & "'"
	sql = sql & ",'" & RemoveQuotes(Applicator) & "'"
	sql = sql & ",'" & RemoveQuotes(ApplicatorLicense) & "'"
	sql = sql & ",'" & RemoveQuotes(Supervisor) & "'"
	sql = sql & ",'" & RemoveQuotes(LicenseNumber) & "'"
	sql = sql & ",'" & RemoveQuotes(ChemicalSupplier) & "'"
	sql = sql & ",'" & RemoveQuotes(RecommendedBy) & "'"
	sql = sql & ")"
	sql = sql & " SELECT insertid = SCOPE_IDENTITY()"

	DIM newID
  '  Response.Write(sql)
  'response.end
	Set rs = conn.execute(sql)
	newID = rs(0)
	sql = ""
	For Each targetID In ArrayTargetIDs
		sql = sql & " INSERT INTO dbo.SprayRecordTargets (SprayRecordID, TargetID) VALUES (" & newID & "," & targetID & ") "
	Next
  'Response.Write(sql)
	Set rs = conn.execute(sql)
	'sql = "SELECT Sum(RateAcre) AS OverAppSeason FROM SprayRecord LEFT JOIN SprayList ON SprayRecord.ProductID = SprayList.SprayListID WHERE SprayRecord.GrowerID= " & GrowerID & " AND SprayRecord.Location = '" & EscapeQuotes(Location) & "' AND  SprayRecord.ProductID= " & SprayListID

	InsertSprayRecord2 = newID
end Function

' *******************************************************
' ************ UpdateSprayRecord ************************
' *******************************************************
function UpdateSprayRecord(SprayRecordID,PackerID,GrowerID,SprayStartDate,TimeFinishedSpraying,SprayEndDate,CropID,VarietyID,Bartlet,StageID,Location,MethodID,AcresTreated,RateAcre,SprayListID,IFPRating,TargetID,HarvestDate,Comments,Weather,Applicator,ApplicatorLicense,Administrator,Supervisor,LicenseNumber,ChemicalSupplier,RecommendedBy)
	IF Bartlet = "" THEN
		Bartlet = 0
	END IF
		arrayVarieties = Split(VarietyID,",")
'response.write("<br>--" & VarietyID & "<br>")
	VarietyID1 = 0
	VarietyID2 = 0
	VarietyID3 = 0
	VarietyID4 = 0
	IF IsArray(arrayVarieties) THEN
		IF Ubound(arrayVarieties) >= 0 THEN
			VarietyID1 = arrayVarieties(0)
		END IF
		IF Ubound(arrayVarieties) >= 1 THEN
			VarietyID2 = arrayVarieties(1)
		END IF
		IF Ubound(arrayVarieties) >= 2 THEN
			VarietyID3 = arrayVarieties(2)
		END IF
		IF Ubound(arrayVarieties) >= 3 THEN
			VarietyID4 = arrayVarieties(3)
		END IF
	END IF

	sql = "UPDATE SprayRecord SET PackerID=" & PackerID &", GrowerID =" & GrowerID & ",SprayStartDate = '" & SprayStartDate & "'"
	if trim(SprayEndDate) <> "" then
		sql = sql & ",SprayEndDate = '" & SprayEndDate & "'"
	end if
	sql = sql & ",CropID = " & CropID & ",Bartlet = " & Bartlet & ",StageID = " & StageID & ",Location ='" & RemoveQuotes(Location) & "',MethodID = " & MethodID & ",AcresTreated = " & AcresTreated & ",RateAcre = " & RateAcre & ",ProductID = " & SprayListID & ",IFPRating ='" & IFPRating & "',TargetID = " & TargetID
	if trim(HarvestDate) <> "" then
		sql = sql  & ",HarvestDate = '" & HarvestDate & "'"
	end if
	sql = sql  & ",TimeFinishedSpraying = '" & TimeFinishedSpraying & "'"
	sql = sql  & ",Weather = '" & RemoveQuotes(Weather) & "'"
	sql = sql  & ",VarietyID1 = " & VarietyID1
	sql = sql  & ",VarietyID2 = " & VarietyID2
	sql = sql  & ",VarietyID3 = " & VarietyID3
	sql = sql  & ",VarietyID4 = " & VarietyID4
	sql = sql  & ",Applicator = '" & RemoveQuotes(Applicator) & "'"
	sql = sql  & ",ApplicatorLicense = '" & RemoveQuotes(ApplicatorLicense) & "'"
	sql = sql  & ",Supervisor = '" & RemoveQuotes(Supervisor) & "'"
	sql = sql  & ",LicenseNumber = '" & RemoveQuotes(LicenseNumber) & "'"
	sql = sql  & ",ChemicalSupplier = '" & RemoveQuotes(ChemicalSupplier) & "'"
	sql = sql  & ",RecommendedBy = '" & RemoveQuotes(RecommendedBy) & "'"
	sql = sql & ",Comments ='" & Comments & "',Administrator ='" & Session("username") & "',UpdateDate =  getdate()    WHERE SprayRecordID = " & SprayRecordID

'	response.write sql
	conn.execute sql, , 129
	UpdateSprayRecord = SprayRecordID
end Function

' *******************************************************
' ************ UpdateSprayRecord ************************
' rem weather text stored in SprayRecord now.
' *******************************************************
function UpdateSprayRecord2(SprayRecordID,PackerID,GrowerID,SprayStartDate,TimeFinishedSpraying,SprayEndDate,CropID,VarietyID,Bartlet,StageID,Location,MethodID,AcresTreated,RateAcre,SprayListID,IFPRating,TargetID,HarvestDate,Comments,Weather,Applicator,ApplicatorLicense,Administrator,Supervisor,LicenseNumber,ChemicalSupplier,RecommendedBy)
	IF Bartlet = "" THEN
		Bartlet = 0
	END IF
		arrayVarieties = Split(VarietyID,",")
response.write("<br>--" & VarietyID & "<br>")
	VarietyID1 = 0
	VarietyID2 = 0
	VarietyID3 = 0
	VarietyID4 = 0
	IF IsArray(arrayVarieties) THEN
		IF Ubound(arrayVarieties) >= 0 THEN
			VarietyID1 = arrayVarieties(0)
		END IF
		IF Ubound(arrayVarieties) >= 1 THEN
			VarietyID2 = arrayVarieties(1)
		END IF
		IF Ubound(arrayVarieties) >= 2 THEN
			VarietyID3 = arrayVarieties(2)
		END IF
		IF Ubound(arrayVarieties) >= 3 THEN
			VarietyID4 = arrayVarieties(3)
		END IF
	END IF

	sql = "UPDATE SprayRecord SET PackerID=" & PackerID & ",GrowerID =" & GrowerID & ",SprayStartDate = '" & SprayStartDate & "'"
	if trim(SprayEndDate) <> "" then
		sql = sql & ",SprayEndDate = '" & SprayEndDate & "'"
	end if
	sql = sql & ",CropID = " & CropID & ",Bartlet = " & Bartlet & ",StageID = " & StageID & ",Location ='" & RemoveQuotes(Location) & "',MethodID = " & MethodID & ",AcresTreated = " & AcresTreated & ",RateAcre = " & RateAcre & ",ProductID = " & SprayListID & ",IFPRating ='" & IFPRating & "',TargetID = " & TargetID
	if trim(HarvestDate) <> "" then
		sql = sql  & ",HarvestDate = '" & HarvestDate & "'"
	end if
	sql = sql  & ",TimeFinishedSpraying = '" & TimeFinishedSpraying & "'"
	sql = sql  & ",Weather = '" & RemoveQuotes(Weather) & "'"
	sql = sql  & ",VarietyID1 = " & VarietyID1
	sql = sql  & ",VarietyID2 = " & VarietyID2
	sql = sql  & ",VarietyID3 = " & VarietyID3
	sql = sql  & ",VarietyID4 = " & VarietyID4
	sql = sql  & ",Applicator = '" & RemoveQuotes(Applicator) & "'"
	sql = sql  & ",ApplicatorLicense = '" & RemoveQuotes(ApplicatorLicense) & "'"
	sql = sql  & ",Supervisor = '" & RemoveQuotes(Supervisor) & "'"
	sql = sql  & ",LicenseNumber = '" & RemoveQuotes(LicenseNumber) & "'"
	sql = sql  & ",ChemicalSupplier = '" & RemoveQuotes(ChemicalSupplier) & "'"
	sql = sql  & ",RecommendedBy = '" & RemoveQuotes(RecommendedBy) & "'"
	sql = sql & ",Comments ='" & Comments & "',Administrator ='" & Session("username") & "',UpdateDate =  Now()    WHERE SprayRecordID = " & SprayRecordID

	response.write sql
	conn.execute sql, , 129
	UpdateSprayRecord = SprayRecordID
end Function
%>