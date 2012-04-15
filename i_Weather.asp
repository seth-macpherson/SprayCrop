
<%
' FILE: i_Weather.asp 
' CREATED by www.LocusInteractive.net on 07/21/2005

' ******************************************************* 
' ************GetWeatherByID ******************************* 
' ******************************************************* 
function GetWeatherByID(ID)
	sql = "SELECT * FROM Weather WHERE WeatherID in (" & ID & ")"
	set GetWeatherByID = conn.execute(sql)
end function

' ******************************************************* 
' ************GetAllWeather ******************************* 
' ******************************************************* 
function GetAllWeather()
	sql = "SELECT * FROM Weather ORDER BY SortOrder"
	set GetAllWeather = conn.execute(sql)
end function

' ******************************************************* 
' ************GetActiveWeather ******************************* 
' ******************************************************* 
function GetActiveWeather()
	sql = "SELECT * FROM Weather WHERE Active= true ORDER BY SortOrder"
	set GetActiveWeather = conn.execute(sql)
end function

' ******************************************************* 
' ************DeleteWeather ******************************* 
' ******************************************************* 
function DeleteWeather(ID)
	sql = "DELETE FROM Weather WHERE WeatherID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************InsertWeather ******************************* 
' ******************************************************* 
function InsertWeather(Weather)
Dim newSortOrder
sql = "SELECT Count(SortOrder) AS cntSort FROM Weather"
set rs = conn.execute(sql)
if rs(0) = 0 THEN
	newSortOrder = 1
else
	sql = "SELECT MAX(SortOrder) AS maxSort FROM Weather"
	set rs = conn.execute(sql)
	newSortOrder = rs(0) + 1
end if

sql = "INSERT INTO Weather(Weather,SortOrder) VALUES ("
sql = sql & "'" & EscapeQuotes(Weather) & "'," & newSortOrder
sql = sql & ")"
'response.write sql
	conn.execute sql, , 129
	DIM newID
	sql = "SELECT MAX(WeatherID) AS insertid FROM Weather"
	set rs = conn.execute(sql)
	newID = rs(0)
	InsertWeather = newID
end Function
' ******************************************************* 
' ************UpdateWeather ******************************* 
' ******************************************************* 
function UpdateWeather(ID,Weather)
	sql = "UPDATE Weather SET Weather ='" & Weather & "' WHERE WeatherID = " & ID
	conn.execute sql, , 129
	UpdateWeather = ID
end Function
' ******************************************************* 
' ************ActivateWeather ******************************* 
' ******************************************************* 
function ActivateWeather(ID)
	sql = "Update Weather SET Active = 1 WHERE WeatherID = " & ID
	conn.execute sql, , 129
end function

' ******************************************************* 
' ************DeActivateWeather ******************************* 
' ******************************************************* 
function DeActivateWeather(ID)
	sql = "Update Weather SET Active = 0 WHERE WeatherID = " & ID
	conn.execute sql, , 129
end function
%>












