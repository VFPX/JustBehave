* Program:     NameValue.PRG
* Description: Save and Retreive Values by name from various os sources.
* Created:     2/1/2005
* Developer:   Gregory L Reichert - GregReichert@GLRsoftware.com - http://www.GLRsoftware.com
*------------------------------------------------------------
* Id Date        By         Description
*  1 2/1/2005  Gregory L Reichert Initial Creation
*
*------------------------------------------------------------
*!*	This routine allows the coder to save and retreive data as a Name / Value. 
*!*	Several mediums can be address to hold the data. Manage INI, Registry, FoxUser,
*!*	Configuration Tables, XML files, Collection or Arrays with the same rotuine, and vary 
*!*	similar syntax.

*!*	Syntax:
*!*	[cValue =] NameValue( ["GET" | "PUT",] <path> [,vValue] )

*!*
*!*	Examples:
*!*
*!*	*-- INI file
*!*	? NameValue("Put","c:\test.ini/test1/Test Name","Hello World")
*!*	? NameValue("Get","c:\test.ini/test1/Test Name")
*!*
*!*	*-- Registry
*!*	? NameValue("Put","HKEY_CURRENT_USER\Software\test1\Test Name","Hello World")
*!*	? NameValue("Get","HKEY_CURRENT_USER\Software\test1\Test Name")
*!*
*!*	*-- FoxUser Resource File
*!*	? NameValue("Put","resource/Test/test1/Test Name","Hello World")
*!*	? NameValue("Get","resource/Test/test1/Test Name")
*!*
*!*	*-- Configuration table with specific field names.
*!*	? NameValue("Put","c:\test.dbf/test1/Test Name","Hello World")
*!*	? NameValue("Get","c:\test.dbf/test1/Test Name")
*!*
*!*	*-- XML file with specific format
*!*	? NameValue("Put","c:\test.xml/test1/Test Name","Hello World")
*!*	? NameValue("Get","c:\test.xml/test1/Test Name")
*!*	 
*!*	*-- Collection or Array
*!*	? NameValue("Put","collection/test1/Test Name","Hello World")
*!*	? NameValue("Get","collection/test1/Test Name")



LPARAMETERS tcAction, tcPath, tvValue

#if VERSION(5) >= 800
TRY
#endif

LOCAL pcnt
pcnt = PARAMETERS()
DO CASE
*!*		CASE EMPTY(tcPath)
*!*			RETURN ""
	CASE pcnt=0
		RETURN ""
	CASE pcnt=1 ;
			AND VARTYPE(tcAction)="C" AND NOT EMPTY(tcAction)

		tcPath = tcAction
		tcAction = "GET"

	CASE pcnt=2 ;
			AND VARTYPE(tcAction)="C" AND NOT EMPTY(tcAction)

		tvValue = tcPath
		tcPath = tcAction
		tcAction = "PUT"

	CASE pcnt=3 ;
			AND VARTYPE(tcAction)="C" AND NOT EMPTY(tcAction) ;
			AND VARTYPE(tcPath)="C" AND NOT EMPTY(tcPath)

	OTHERWISE
		*-- Invalid parameter
		RETURN ""

ENDCASE
tcAction = ALLTRIM(UPPER(tcAction))
tcPath   = ALLTRIM(tcPath)

LOCAL lcResult
lcResult = ""

DO CASE
	CASE LEFT(UPPER(tcPath),5)=="HKEY_"			&& registry

		DO CASE
			CASE tcAction == "GET"
				lcResult = Registry("GET", tcPath)
			CASE tcAction == "PUT"
				IF Registry("PUT", tcPath, TRANSFORM(tvValue))
					lcResult = Registry("GET", tcPath)
				ENDIF
		ENDCASE

	CASE ATC(".INI/",tcPath)>0 					&& INI file

		LOCAL lcPath, lcFilename, lcSection, lcEntry, lcData

		lcPath = UPPER(ALLTRIM(tcPath)) +"///"
		lcFilename = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcSection   = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcEntry = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))

		lcData = TRANSFORM(tvValue)

		DO CASE
			CASE tcAction == "GET"
				lcResult = ALLTRIM(GetProfileString( lcFilename,lcSection,lcEntry, 255))
			CASE tcAction == "PUT"
				IF WriteProfileString( lcFilename,lcSection,lcEntry, lcData )
					lcResult = ALLTRIM(GetProfileString( lcFilename,lcSection,lcEntry, 255))
				ENDIF
		ENDCASE


	CASE LEFT(UPPER(tcPath),9)=="RESOURCE/"		&& FoxUser resource file.

		LOCAL lcPath, lcType, lcID, lcName, lcData

		lcPath = UPPER(ALLTRIM(tcPath))
		lcPath = STRTRAN(lcPath, "RESOURCE/","") +"///"
		lcType = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcID   = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcName = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))

		lcData = TRANSFORM(tvValue)

		DO CASE
			CASE tcAction == "GET"
				lcResult = Resource_Load( lcType, lcID, lcName )
			CASE tcAction == "PUT"
				IF Resource_Save( lcType, lcID, lcName, lcData )
					lcResult = Resource_Load( lcType, lcID, lcName )
				ENDIF
		ENDCASE

	CASE ATC(".DBF/",tcPath)>0 					&& Table

		LOCAL lcPath, lcFilename, lcSection, lcEntry, lcData

		lcPath = UPPER(ALLTRIM(tcPath)) +"///"
		lcFilename = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcSection   = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcEntry = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))

		lcData = TRANSFORM(tvValue)

		DO CASE
			CASE tcAction == "GET"
				lcResult = ALLTRIM(Table_Get( lcFilename,lcSection,lcEntry))
			CASE tcAction == "PUT"
				IF Table_Put( lcFilename,lcSection,lcEntry, lcData )
					lcResult = ALLTRIM(Table_Get( lcFilename,lcSection,lcEntry))
				ENDIF
		ENDCASE

	CASE ATC(".XML/",tcPath)>0 					&& XML File

		LOCAL lcPath, lcFilename, lcSection, lcEntry, lcData

		lcPath = UPPER(ALLTRIM(tcPath)) +"///"
		lcFilename = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcSection   = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcEntry = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))

		lcData = TRANSFORM(tvValue)

		DO CASE
			CASE tcAction == "GET"
				lcResult = ALLTRIM(XML_Get( lcFilename,lcSection,lcEntry))
			CASE tcAction == "PUT"
				IF XML_Put( lcFilename,lcSection,lcEntry, lcData )
					lcResult = ALLTRIM(XML_Get( lcFilename,lcSection,lcEntry))
				ENDIF
		ENDCASE

	CASE LEFT(UPPER(tcPath),11)=="COLLECTION/" OR LEFT(UPPER(tcPath),6)=="ARRAY/"	&& Collection / Array

		LOCAL lcPath, lcData, __lcExact
		__lcExact = SET("Exact")
		SET EXACT ON

		lcPath = UPPER(ALLTRIM(tcPath))
		lcPath = STRTRAN(lcPath, "COLLECTION/","") 
		lcPath = STRTRAN(lcPath, "ARRAY/","") 
		lcPath = lcPath + "///"
		lcCollection = ALLTRIM(LEFT(lcPath,AT("/",lcPath)-1))
		lcPath = SUBSTR(lcPath, AT("/",lcPath)+1 )
		lcName   = ALLTRIM(LEFT(lcPath,AT("/",lcPath)+1))

		lcData = TRANSFORM(tvValue)

		DO CASE
			CASE tcAction == "GET"
				lcResult = Collection_Get( lcCollection, lcName )
			CASE tcAction == "PUT"
				IF Collection_Put( lcCollection, lcName, lcData )
					lcResult = Collection_Get( lcCollection, lcName )
				ENDIF
		ENDCASE
		SET EXACT &__lcExact

ENDCASE

#if VERSION(5) >= 800
CATCH
	lcResult = ""
ENDTRY
#endif 


RETURN lcResult

*******************************************************************************
*******************************************************************************


FUNCTION GetProfileString
	************************************************************************
	* wwAPI :: GetProfileString
	***************************
	***  Modified: 09/26/95
	***  Function: Read Profile String information from a given
	***            text file using Windows INI formatting conventions
	***      Pass: pcFileName   -    Name of INI file
	***            pcSection    -    [Section] in the INI file ("Drivers")
	***            pcEntry      -    Entry to retrieve ("Wave")
	***                              If this value is a null string
	***                              all values for the section are
	***                              retrieved seperated by CHR(13)s
	***    Return: Value(s) or .NULL. if not found
	************************************************************************
	LPARAMETERS pcFilename,pcSection,pcEntry, pnBufferSize
	LOCAL lcIniValue, lnResult

	*** Initialize buffer for result
	lcIniValue=SPACE(IIF( TYPE("pnBufferSize")="N",pnBufferSize,MAX_INI_BUFFERSIZE) )

	DECLARE INTEGER GetPrivateProfileString ;
		IN WIN32API ;
		STRING cSection,;
		STRING cEntry,;
		STRING cDefault,;
		STRING @cRetVal,;
		INTEGER nSize,;
		STRING cFileName

	lnResult=GetPrivateProfileString(pcSection,pcEntry,"*None*",;
		@lcIniValue,LEN(lcIniValue),pcFilename)

	CLEAR DLLS GetPrivateProfileString

	*** Strip out Nulls
	IF TYPE("pcEntry")="N" AND pcEntry=0
		*** 0 was passed to get all entry labels
		*** Seperate all of the values with a Carriage Return
		lcIniValue=TRIM(CHRTRAN(lcIniValue,CHR(0),CHR(13)) )
	ELSE
		*** Individual Entry
		lcIniValue=SUBSTR(lcIniValue,1,lnResult)
	ENDIF

	*** On error the result contains "*None"
	IF lcIniValue="*None*"
		lcIniValue=.NULL.
	ENDIF

	RETURN lcIniValue
ENDFUNC
* GetProfileString

************************************************************************
* wwAPI :: WriteProfileString
*********************************
***  Function: Writes a value back to an INI file
***      Pass: pcFileName    -   Name of the file to write to
***            pcSection     -   Profile Section
***            pcKey         -   The key to write to
***            pcValue       -   The value to write
***    Return: .T. or .F.
************************************************************************
FUNCTION WriteProfileString
	LPARAMETERS pcFilename,pcSection,pcEntry,pcValue

	DECLARE INTEGER WritePrivateProfileString ;
		IN WIN32API ;
		STRING cSection,STRING cEntry,STRING cEntry,;
		STRING cFileName

	lnRetVal=WritePrivateProfileString(pcSection,pcEntry,pcValue,pcFilename)

	CLEAR DLLS WritePrivateProfileString

	IF lnRetVal=1
		RETURN .T.
	ENDIF

	RETURN .F.
ENDFUNC
* WriteProfileString

**************************************************************
**************************************************************

* Program:     REGISTRY.PRG
* Created:     10/19/2004
* Developer:   Greg Reichert
* Copyright:   public domain 2004 GLR software
*------------------------------------------------------------
* Description: Read and write to registry in Name / Value pair method.
* Parameters:  tKey, req, d="", Registry key.
*			   tValue, opt, d="", value to write to key.
* Return:      Value from key.
* Syntax:
* 	Set Value:
*		Registry( "HKEY_CURRENT_USER\SOFTWARE\Microsoft\VisualFoxPro\MyKey", "myValue" )
*
*	Get Value:
*		lcValue = Registry( "HKEY_CURRENT_USER\SOFTWARE\Microsoft\VisualFoxPro\MyKey" )
*------------------------------------------------------------
* Id Date        By         Description
*  1 08/27/2002  Greg Reichert    Initial Creation
*
*------------------------------------------------------------
*------------------------------------------------------------
* Description: Get / Set Registry
* Parameters:  <para>, <req/opt>, D=<def>, <desc>
* Return:
* Use:
*------------------------------------------------------------
* Id Date        By         Description
*  1 04/12/2002  Greg Reichert    Initial Creation
*
*------------------------------------------------------------
PROCEDURE Registry
	LPARAMETERS tAction, tKey, tValue
	* Registry roots
	HKEY_CLASSES_ROOT           = -2147483648  && BITSET(0,31)
	HKEY_CURRENT_USER           = -2147483647  && BITSET(0,31)+1
	HKEY_LOCAL_MACHINE          = -2147483646  && BITSET(0,31)+2
	HKEY_USERS                  = -2147483645  && BITSET(0,31)+3

	#DEFINE REG_SZ 				1	&& Data string
	#DEFINE REG_BINARY 			3	&& Binary data in any form.
	#DEFINE REG_DWORD 			4	&& A 32-bit number.

	DECLARE INTEGER RegCreateKey IN Win32API ;
		INTEGER nHKey, STRING @cSubKey, INTEGER @nResult

	DECLARE INTEGER RegOpenKey IN Win32API ;
		INTEGER nHKey, STRING @cSubKey, INTEGER @nResult

	DECLARE INTEGER RegCloseKey IN Win32API ;
		INTEGER nHKey

	DECLARE INTEGER RegQueryValueEx IN Win32API ;
		INTEGER nHKey, STRING lpszValueName, INTEGER dwReserved,;
		INTEGER @lpdwType, STRING @lpbData, INTEGER @lpcbData

	DECLARE INTEGER RegSetValueEx IN Win32API ;
		INTEGER hKey, STRING lpszValueName, INTEGER dwReserved,;
		INTEGER fdwType, STRING lpbData, INTEGER cbData

	*--------------------------------------------
	*  parse cKey
	*--------------------------------------------
	LOCAL hKey, cKey, cData
	hKey	= LEFT( tKey, AT("\", tKey)-1)
	tKey 	= SUBSTR( tKey, AT("\",tKey)+1)
	cKey	= LEFT( tKey, RAT("\",tKey)-1)
	tKey 	= SUBSTR( tKey, RAT("\",tKey)+1)
	cData	= tKey

	LOCAL nCurrentKey,nErrCode
	nCurrentKey = 0
	nErrCode = RegOpenKey(&hKey,cKey,@nCurrentKey)
	IF nErrCode>0
		nErrCode = RegCreateKey(&hKey,cKey,@nCurrentKey)
	ENDIF

	LOCAL lpdwReserved,lpdwType,lpbData,lpcbData,nErrCode
	STORE 0 TO lpdwReserved,lpdwType
	STORE SPACE(256) TO lpbData
	STORE LEN(m.lpbData) TO m.lpcbData

	DO CASE
		CASE UPPER(ALLTRIM(tAction))="GET"

			m.nErrCode=RegQueryValueEx(nCurrentKey,cData, m.lpdwReserved,@lpdwType,@lpbData,@lpcbData)
			m.cKeyValue = LEFT(m.lpbData,m.lpcbData-1)

		CASE UPPER(ALLTRIM(tAction))="PUT"
			* Make sure we null terminate this guy
			lpcbData = TRANSFORM(m.tValue)+CHR(0)
			nValueSize = LEN(m.lpcbData)

			* Set the key value here
			m.nErrCode = RegSetValueEx(nCurrentKey,m.cData,0,;
				REG_SZ,m.lpcbData ,m.nValueSize)

			cKeyValue = m.nErrCode=0
	ENDCASE

	=RegCloseKey(nCurrentKey)

	#IF VERSION(5)>=700
		CLEAR DLLS RegCreateKey
		CLEAR DLLS RegOpenKey
		CLEAR DLLS RegCloseKey
		CLEAR DLLS RegQueryValueEx
		CLEAR DLLS RegSetValueEx
	#ENDIF

	RETURN cKeyValue
ENDPROC
***********************************************
***********************************************


*------------------------------------------------------------
* Description:
* Parameters:  <para>, <req/opt>, D=<def>, <desc>
* Return:
* Use:
*------------------------------------------------------------
* Id Date        By         Description
*  1 02/01/2005  gregory.re Initial Creation
*
*------------------------------------------------------------
PROCEDURE Resource_Save
	LPARAMETERS tcType, tcID, tcName, tcData

	LOCAL ls
	LOCAL lcFoxUser
	LOCAL lcResult

	lcFoxUser = SYS(2005)
	ls = SELECT()

	SELECT 0
	USE (lcFoxUser) AGAIN SHARED

	UPDATE (lcFoxUser) ;
		SET DATA = m.tcData, ;
		ckval = VAL(SYS(2007,m.tcData)) ;
		WHERE TYPE == m.tcType ;
		AND ID == m.tcID ;
		AND ALLTRIM(NAME) == m.tcName

	IF _TALLY=0

		INSERT INTO (lcFoxUser) (TYPE, ID, NAME, DATA, ckval) VALUES (m.tcType, m.tcID, m.tcName, m.tcData, VAL(SYS(2007,m.tcData)) )

	ENDIF

	lcResult = (_TALLY>0)

	USE IN (SELECT("FoxUser"))
	SELECT (ls)

	RETURN lcResult

ENDPROC


*------------------------------------------------------------
* Description:
* Parameters:  <para>, <req/opt>, D=<def>, <desc>
* Return:
* Use:
*------------------------------------------------------------
* Id Date        By         Description
*  1 02/01/2005  gregory.re Initial Creation
*
*------------------------------------------------------------
PROCEDURE Resource_Load
	LPARAMETERS tcType, tcID, tcName

	LOCAL ARRAY laPosition[1,1]
	LOCAL ls
	LOCAL lcResult
	LOCAL lcFoxUser

	lcFoxUser = SYS(2005)
	lcResult = ""
	ls = SELECT()

	IF EMPTY(lcFoxUser)
		SELECT 0
		CREATE TABLE (THIS.cFoxUser) (TYPE C(12), ID C(12), NAME M, READONLY L, ckval N(6), DATA M, UPDATED D)
		USE
		SET RESOURCE TO (THIS.cFoxUser)
	ENDIF

	SELECT 0
	USE (lcFoxUser) AGAIN SHARED

	SELECT DATA FROM (lcFoxUser) ;
		WHERE TYPE == m.tcType ;
		AND ID == m.tcID ;
		AND ALLTRIM(NAME) == m.tcName ;
		INTO ARRAY laPosition

	IF _TALLY>0
		lcResult = laPosition[1,1]
	ENDIF

	USE IN (SELECT("FoxUser"))
	SELECT (ls)

	RETURN lcResult

ENDPROC



*******************************************
**********************************************

*------------------------------------------------------------
* Description:
* Parameters:  <para>, <req/opt>, D=<def>, <desc>
* Return:
* Use:
*------------------------------------------------------------
* Id Date        By         Description
*  1 02/01/2005  gregory.re Initial Creation
*
*------------------------------------------------------------
PROCEDURE Table_Put
	LPARAMETERS pcFilename,pcSection,pcEntry,pcValue

	LOCAL ls, lcResult
	lcResult = ""
	ls = SELECT()
	SELECT 0

	*--------------------------------------------
	*  If not found, create
	*--------------------------------------------
	IF NOT FILE(FULLPATH(pcFilename))
		CREATE TABLE (FULLPATH(pcFilename)) ;
			(fcSection C(20), ;
			fcEntry C(20), ;
			fcData M)
	ENDIF

	*--------------------------------------------
	*  Update existing record
	*--------------------------------------------
	UPDATE (FULLPATH(pcFilename)) ;
		SET fcData = m.pcValue ;
		WHERE ALLTRIM(fcSection) == m.pcSection ;
		AND ALLTRIM(fcEntry) == m.pcEntry ;
		AND NOT DELETED()

	lcResult = (_TALLY>0)
	IF _TALLY=0
		*--------------------------------------------
		*  If update did not occur, add it as new.
		*--------------------------------------------
		INSERT INTO (FULLPATH(pcFilename)) ;
			(fcSection, fcEntry, fcData ) ;
			VALUES ;
			(m.pcSection, m.pcEntry, m.pcValue )
		lcResult = .T.

	ENDIF

	USE IN (SELECT(JUSTSTEM(m.pcFilename)))
	SELECT (ls)

	RETURN lcResult

ENDPROC

*---------------

*------------------------------------------------------------
* Description:
* Parameters:  <para>, <req/opt>, D=<def>, <desc>
* Return:
* Use:
*------------------------------------------------------------
* Id Date        By         Description
*  1 02/01/2005  gregory.re Initial Creation
*
*------------------------------------------------------------
PROCEDURE Table_Get
	LPARAMETERS pcFilename,pcSection,pcEntry

	LOCAL ls, lcResult
	lcResult = ""
	ls = SELECT()
	SELECT 0

	*--------------------------------------------
	*  If not found, create
	*--------------------------------------------
	IF NOT FILE(FULLPATH(pcFilename))
		CREATE TABLE (FULLPATH(pcFilename)) ;
			(fcSection C(20), ;
			fcEntry C(20), ;
			fcData M)
	ENDIF

	*--------------------------------------------
	*  Get last position and size.
	*--------------------------------------------
	LOCAL ARRAY laData[1,1]

	SELECT fcData FROM (FULLPATH(pcFilename)) ;
		WHERE ALLTRIM(fcSection) == m.pcSection ;
		AND ALLTRIM(fcEntry) == m.pcEntry ;
		AND NOT DELETED() ;
		INTO ARRAY laData

	*--------------------------------------------
	*  If found, update form.
	*--------------------------------------------
	IF _TALLY>0
		lcResult = laData[1,1]
	ENDIF

	USE IN (SELECT(JUSTSTEM(m.pcFilename)))
	SELECT (ls)

	RETURN lcResult

ENDPROC

********************************************
********************************************

*------------------------------------------------------------
* Description:
* Parameters:  <para>, <req/opt>, D=<def>, <desc>
* Return:
* Use:
*------------------------------------------------------------
* Id Date        By         Description
*  1 02/01/2005  gregory.re Initial Creation
*
*------------------------------------------------------------
PROCEDURE XML_Get
	LPARAMETERS pcFilename,pcSection,pcEntry

	LOCAL ls, lcResult
	ls = SELECT()
	lcResult = ""

	SELECT 0
	CREATE CURSOR tmpXMLfile ;
		(fcSection C(20), ;
		fcEntry C(20), ;
		fcData M)

	IF FILE(pcFilename)
		XMLTOCURSOR(pcFilename,"tmpXMLfile",4+512+8192)
	ENDIF

	*--------------------------------------------
	*  Get last position and size.
	*--------------------------------------------
	LOCAL ARRAY laData[1,1]

	SELECT fcData FROM tmpXMLfile ;
		WHERE ALLTRIM(fcSection) == m.pcSection ;
		AND ALLTRIM(fcEntry) == m.pcEntry ;
		AND NOT DELETED() ;
		INTO ARRAY laData

	*--------------------------------------------
	*  If found, update form.
	*--------------------------------------------
	IF _TALLY>0
		lcResult = laData[1,1]
	ENDIF

	USE IN (SELECT("tmpXMLfile"))
	SELECT (ls)

	RETURN lcResult

ENDPROC

*------------------------------------------------------------
* Description:
* Parameters:  <para>, <req/opt>, D=<def>, <desc>
* Return:
* Use:
*------------------------------------------------------------
* Id Date        By         Description
*  1 02/01/2005  gregory.re Initial Creation
*
*------------------------------------------------------------
PROCEDURE XML_Put
	LPARAMETERS pcFilename,pcSection,pcEntry, pcValue

	LOCAL ls, lcResult
	ls = SELECT()
	lcResult = .F.

	SELECT 0
	CREATE CURSOR tmpXMLfile ;
		(fcSection C(20), ;
		fcEntry C(20), ;
		fcData M)

	IF FILE(pcFilename)
		XMLTOCURSOR(pcFilename,"tmpXMLfile",4+512+8192)
	ENDIF

	*--------------------------------------------
	*  Update existing record
	*--------------------------------------------
	UPDATE tmpXMLfile ;
		SET fcData = m.pcValue ;
		WHERE ALLTRIM(fcSection) == m.pcSection ;
		AND ALLTRIM(fcEntry) == m.pcEntry ;
		AND NOT DELETED()

	lcResult = (_TALLY>0)
	IF _TALLY=0
		*--------------------------------------------
		*  If update did not occur, add it as new.
		*--------------------------------------------
		INSERT INTO tmpXMLfile ;
			(fcSection, fcEntry, fcData ) ;
			VALUES ;
			(m.pcSection, m.pcEntry, m.pcValue )
		lcResult = .T.

	ENDIF

	lcResult = (CURSORTOXML("tmpXMLfile",pcFilename,1,2+4+8+512,0,"")>0)


	USE IN (SELECT("tmpXMLfile"))
	SELECT (ls)

	RETURN lcResult
ENDPROC

*****************************************************************
*****************************************************************
PROCEDURE Collection_Get
	LPARAMETERS tcCollection, tcName

	IF TYPE(tcCollection)="U"
		RELEASE &tcCollection
		PUBLIC &tcCollection.[1,2]
	ENDIF

	LOCAL i
	i = ASCAN(&tcCollection ,tcName,1, ALEN(&lcCollection,0))
	IF i>0
		i = ASUBSCRIPT(&tcCollection ,i,1)
		lcResult = &tcCollection.[i,2]
	ENDIF

	RETURN lcResult

ENDPROC

*-----------

PROCEDURE Collection_Put
	LPARAMETERS tcCollection, tcName, tcValue


	IF TYPE(tcCollection)="U"
		RELEASE &tcCollection
		PUBLIC ARRAY &tcCollection.[1,2]
	ENDIF

	LOCAL i
	i = ASCAN(&tcCollection ,tcName,1, ALEN(&lcCollection,0))
	IF i>0
		i = ASUBSCRIPT(&tcCollection ,i,1)
	ELSE
		i = 1
		IF VARTYPE( &tcCollection.[1,1] )<>"L"
			i = ALEN(&tcCollection,1)+1
		ENDIF
		PUBLIC ARRAY &tcCollection.[i,2]
	ENDIF

	&tcCollection.[i,1] = tcName
	&tcCollection.[i,2] = tcValue

	lcResult = .T.

	RETURN lcResult

ENDPROC


* Eof NAMEVALUE.PRG


