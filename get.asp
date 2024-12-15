<% language="VBScript"
'* openSug.js Private outcome source (ASP)
'* Robin lee <openSug@qq.com>
'* Thu, 21 Dec 2006 14:08:27 GMT
'* kw Query String
'* cb CallBack Function

' Define constants for database name and HTTP status message
CONST AccessMDB			= "database"
CONST StatusMSG			= "404 Not Found"

' Set server script timeout to 30 seconds
SERVER.SCRIPTTIMEOUT	= 30

' Set response expiration to -1 to prevent caching
RESPONSE.EXPIRES		= -1

' Set response charset to UTF-8
RESPONSE.CHARSET		= "UTF-8"

' Set response codepage to 65001 (UTF-8)
RESPONSE.CODEPAGE		= 65001

' Handle errors by resuming next statement
ON ERROR RESUME NEXT

' Declare variables for query string parameters
DIM WORD, FUNC

' Get the value of the 'kw' query string parameter
WORD	= REQUEST.QUERYSTRING("kw")

' Get the value of the 'cb' query string parameter
FUNC	= REQUEST.QUERYSTRING("cb")

' Check if 'kw' or 'cb' parameters are empty, request method is not GET, or referer length is less than 10
IF WORD = "" OR FUNC = "" OR REQUEST.SERVERVARIABLES("REQUEST_METHOD") <> "GET" OR LEN(REQUEST.SERVERVARIABLES("HTTP_REFERER")) < 10 THEN
    ' Set response status to 404 Not Found and end the response
    RESPONSE.STATUS = StatusMSG
    RESPONSE.END()
END IF

' Declare variables for regular expression and match collection
DIM REG, MAT

' Create a new RegExp object
SET REG		= New RegExp

' Set the pattern to match expected callback function format
REG.PATTERN	= "^BaiduSuggestion\.res\.__\d+$"

' Execute the regular expression on the 'cb' parameter and get matches
SET MAT		= REG.EXECUTE(FUNC)

' Release the RegExp object
SET REG		= NOTHING

' If no matches found, set response status to 404 Not Found and end the response
IF MAT.COUNT = 0 THEN
    RESPONSE.STATUS = StatusMSG
    RESPONSE.END()
END IF

' Set response content type to text/javascript
RESPONSE.CONTENTTYPE	= "text/javascript"

' Declare variables for database connection, recordset, and string
DIM DBC, RST, STR

' Create an ADODB.Connection object
SET DBC			= SERVER.CREATEOBJECT("ADODB.CONNECTION")

' Set the mode to read/write
DBC.MODE		= 3

' Set the provider to Microsoft Jet OLEDB 4.0
DBC.PROVIDER	= "Microsoft.Jet.OLEDB.4.0"

' Open the database connection using the mapped path of the Access database file
DBC.OPEN(SERVER.MAPPATH(AccessMDB &".mdb"))

' Create an ADODB.Recordset object
SET RST = SERVER.CREATEOBJECT("ADODB.RECORDSET")

' Open a recordset to check if the word already exists in the database
RST.OPEN "SELECT word FROM words WHERE word = '" & REPLACE(WORD, "'", "''") & "';", DBC, 1, 1

' If the word does not exist, insert it into the database along with IP address and timestamp
IF RST.EOF THEN
    DBC.EXECUTE "INSERT INTO words(word, ip, times) VALUES('" & REPLACE(WORD, "'", "''") & "', '"& REQUEST.SERVERVARIABLES("REMOTE_ADDR") &"', '"& DATEDIFF("s", #1/1/1970#, NOW()) &"');"
END IF

' Close and release the recordset object
SET RST = NOTHING

' Reopen the recordset to retrieve up to 10 suggestions based on the input word
SET RST = SERVER.CREATEOBJECT("ADODB.RECORDSET")
RST.OPEN "SELECT TOP 10 word FROM words WHERE word LIKE '%" & REPLACE(WORD, "'", "''") & "%' ORDER BY times;", DBC, 1, 1

' Loop through the recordset and build a JSON-like string of suggestions
IF NOT RST.EOF THEN
    DO WHILE NOT RST.EOF
        STR = STR & """" & REPLACE(RST("word"), """", "\""") & ""","
        RST.MOVENEXT
    LOOP
END IF

' Close and release the recordset object
RST.CLOSE
SET RST = NOTHING

' Close and release the database connection object
DBC.CLOSE
SET DBC = NOTHING

' Remove the trailing comma from the suggestions string if it exists
IF LEN(STR) > 0 THEN
    STR = LEFT(STR, LEN(STR) - 1)
END IF

' Write the JavaScript code to the response to call the callback function with the suggestions
RESPONSE.WRITE("'use strict';(function(w){'use strict';if(typeof w.BaiduSuggestion==='object'&&typeof w.BaiduSuggestion.res==='object'&&typeof w."& FUNC & "==='function')w."& FUNC & "({s:[" & STR & "]});" &"}(window));")

' End the response
RESPONSE.END()
%>