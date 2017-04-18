********************************************************************
*** Name.....: ERRORHANDLER.PRG
*** Author...: Andy Kramek & Marcia Akins
*** Date.....: 08/28/2005
*** Notice...: Copyright (c) 2005 Tightline Computers, Inc
*** Compiler.: Visual FoxPro 09.00.0000.2412 for Windows 
*** Function.: Class Definition for the TLC Generic Error Handler
*** Returns..: VOID
********************************************************************
*** System Definitions
#DEFINE DEF_EMAILTO    "glennpd@aol.com"
#define DEF_CPHONE     "2312768043"
#DEFINE DEF_CCMAILTO   ""
#DEFINE DEF_COMPANY   "Glenn Domeracki"
#DEFINE DEF_EMSUBJECT  "Application Error"

*** Basic Definitions
#DEFINE CRLF CHR(13) + CHR(10) 
#DEFINE DIVLINE REPLICATE( "-", 79 )

*** Standard Error Message Definitions:
#DEFINE MSG_MinorTitle "A minor application error occurred"
#DEFINE MSG_StdTitle   "Application Error"
#DEFINE MSG_PressOK    "Press OK to continue"
#DEFINE MSG_PromEMail  "Select 'Yes' to send an E-Mail to: " 
#DEFINE MSG_VFPError   "You should note exactly what you were doing and report the error to IT Support"
#DEFINE MSG_AppGoOn    "The application can attempt to carry on despite this error, however," + CRLF ;
                     + "you may prefer to cancel and re-start, in which case choose 'Cancel'" 
#DEFINE MSG_NoEMail    "(Note: This will also prevent the E-Mail being sent)" 
#DEFINE MSG_WillClose  "In order to ensure the safety of your data the application will be closed down" 
#DEFINE MSG_RollBack   "Any data which has not already been committed to the database will" + CRLF ;
                     + "be lost in the close-down. When you re-start you should carefully check your " + CRLF ;
                     + "data to ensure that everything is correct." 

*** Instantiate the handler and run the process
PRIVATE oHandler, lcToDo
*** These are create as PRIVATE because VFP does not actually create the variables
*** declared as private until they are assigned a value - unlike LOCAL or PUBLIC which
*** both get initialized to .F. as soon as they are declared
oHandler = CREATEOBJECT( "generrhandler", SET( "Datasession" ))
lcToDo = oHandler.cToDo
RELEASE oHandler
IF NOT EMPTY( lcToDo )
  *** Take the required action 
  DO CASE
    CASE LOWER( lcToDo ) =  'quit'
      *** Release all public variables
      CANCEL
      CLEAR EVENTS
      CLEAR ALL
      RELEASE ALL EXTENDED
      *** And quit
      QUIT
    CASE LOWER( lcToDo ) =  'cancel'
      *** Release all public variables
      CANCEL
      CLEAR EVENTS
      CLEAR ALL
      RELEASE ALL EXTENDED
    OTHERWISE
      *** We are going to Retry the command
      RETRY
  ENDCASE        
ENDIF

DEFINE CLASS generrhandler AS SESSION
  PROTECTED cUserLogin, cAddr, cCompany, cEMail, cPhone, cAppID 
  PROTECTED cErrHandler, cFileName, cProgName, cStatus, nErrID
  PROTECTED bSMTP
  *** Datasession created by the error handler itself
  nDSid = 1
  *** User's log in details
  cUserLogin = SYS(0)
  *** Property to define the E-Mail Address to send to
  cAddr    = DEF_EMAILTO
  *** Properties for customer information 
  cCompany = DEF_COMPANY
  cEMail   = DEF_EMAILTO
  cPhone   = DEF_CPHONE
  *** Property to store the old Error Handler
  cErrHandler = ""
  *** Property for the name of the temporary file used for status info
  cFileName   = ""
  *** Array property for AERROR result
  DIMENSION aErrs[1]
  *** Property for the Name of the program in which error occurred
  cProgName   = ""
  *** Property for the Error Handler Status Report Text
  cStatus = ""
  *** Property for the ID of the error (if it is logged successfully)
  nErrID = 0
  *** Property for the action to take on completion
  cToDo = ""
     
  ********************************************************************
  *** [E] INIT(): Standard Initialization method
  ********************************************************************
  FUNCTION Init( tnDSID )
    PRIVATE lcToDo, lcErr, lnErr
    WITH This

      *** Store Error Information to property
      lnErr = AERROR( This.aErrs )
      IF lnErr < 1
        *** There WAS no recorded error
        RETURN
      ENDIF

      *** Store the current error handler and disable further error handling
      .cErrHandler = ON( 'ERROR' )
      ON ERROR *

      *** Open error log without buffering
      IF NOT USED( 'tlc_error_log' )
        .OpenLog()
      ENDIF
      
      *** Get into the right DataSession
      IF VARTYPE( tnDSID ) = "N" AND NOT EMPTY( tnDSID )
        *** Save current datasession
        This.nDSid = This.DataSessionID
        *** And switch to the datasessionn in which the error occurred
        This.DataSessionId = tnDSID
      ENDIF
      
      *** Call Error Handling routines
      .cToDo = .HandleError()

      *** Reset error handling
      IF NOT EMPTY( .cErrHandler )
        *** Restore original error handler if there was one
        lcErr = This.cErrHandler
        ON ERROR &lcErr
      ELSE
        *** Restore the default VFP error handler
        ON ERROR
      ENDIF
      
      *** Return 
      RETURN
    ENDWITH
  ENDFUNC

  ********************************************************************
  *** [P] HANDLEERROR(): Control method for error handling object
  ********************************************************************
  PROTECTED FUNCTION HandleError
    WITH This
      *** Get the contents of memory first before we pollute things!
      .GetMemory()
      
      *** Now go ahead and get the rest of the information that we need
      LOCAL lcDetails, loAction, llStatus, llSendEMail, lcNextCmd, lcMailTxt
      
      *** Now determine how we are going to handle the error
      loAction = .AssessError()

      *** Get the memory contents into a character string
      lcDetails = CHRTRAN( FILETOSTR( This.cFileName ), CHR(26), "" )
	  
      *** Then dump status information to the temporary file
      .GetStatus()

      *** And add it to the Details String
      lcDetails = lcDetails + CRLF + CHRTRAN( FILETOSTR( This.cFileName ), CHR(26), "" )
	  
      *** Next we need the state of any open tables/views
      lcDetails = lcDetails + CRLF + .GetDataDets()
     
      *** Next the calling Stack
      lcDetails = lcDetails + CRLF + .GetStack()

      *** Finally prepend the Header information
      lcMailTxt = .GenHeader( loAction.cMessage ) 
      lcDetails = lcMailTxt + CRLF + lcDetails

      *** Do we want to log this one?
      IF loAction.lLogErr
        *** Status will be written to the status property
        This.WriteLog( lcDetails )
        IF NOT EMPTY( .cStatus )
          *** Add on the status
          lcMailTxt = lcMailTxt + CRLF + ALLTRIM( .cStatus ) 
        ENDIF
      ELSE
        *** No need to log this error
        .cStatus = "This error does not need to be logged" + CRLF
      ENDIF

      *** Can we even offer to send E-Mail
      llSendEMail = ( NOT EMPTY( .cAddr ))

      *** Return the necessary action     
      lcNextCmd = LOWER( ALLTRIM( .ExecAction( loAction.cAction )))
      
      *** Do we need to show the user a message? 
      IF NOT EMPTY( loAction.cMessage )
        *** Call the display message method, and pass the object
        *** Note that if they return No or Cancel we need to suppress 
        *** sending the E-Mail, so store the return value
        llSendEMail = llSendEMail AND .ShowMessage( loAction, llSendEMail )
      ENDIF      
    
      IF NOT( lcNextCmd == 'retry' )
        *** If all OK here, send the E-Mail with the contents of the error dump string
        IF llSendEMail AND NOT EMPTY( .cAddr )
          .SendEMail( lcMailTxt )
        ENDIF
      
        *** Clean up the temporary file
        IF FILE( .cFileName )
          DELETE FILE (.cFileName)
        ENDIF
      ENDIF
            
      RETURN lcNextCmd

    ENDWITH
  ENDFUNC

  ********************************************************************
  *** [P] ASSESSERROR(): Check the error number and determine what to do about it
  ********************************************************************
  PROTECTED FUNCTION AssessError
  LOCAL lnError, lcStr, lcAction, llLogError
    *** Set Default return values
    lcStr = ""
    lcAction  = 'Quit'  && Default action is to quit app
    llQuitApp = .T.     && Should app be terminated
    llLogError = .T.    && Should Error be logged
    ***********************************************************************
    *** Set the action (based on the error number)
    ***********************************************************************
    lnError = IIF( VARTYPE( This.aErrs[1, 1] ) = "N", This.aErrs[1, 1], 0 )
    DO CASE
      ******<Application Specific (i.e. Custom) Errors>*******************************
      CASE lnError = 1098
        lcStr = 'The Application has reported an internal error that is either' + CRLF +;
           'related to a problem in the data, or the availability of a required' + CRLF +;
           'file or data set. Please re-try after re-starting the application' + CRLF +;
           'and if the problem persists contact your IT Support Person'
      
      ******<I/O Errors>**************************************************************
      CASE INLIST( lnError,1,3,4,5,6,7,15,17,19,41,50,54,55,114,1002,1166,1167,1551,1552,1553)  
        lcStr = 'The Application has reported a File or Input/Output Error.  This may'+ CRLF +;
          'be caused by corrupt data or be due to an error at the Windows operating '+ CRLF +;
          'system level.  Try re-booting your PC after exiting Windows, but if that '+ CRLF +;
          'does not cure the problem you should contact your IT support person'

      ******<File Access Errors>**************************************************************
      CASE INLIST( lnError,1102,1103,1104,1105,1106,1111,1112,1113,1115,1152,1153,1157,1169,1178 )
        lcStr = 'The Application has reported a problem while attempting to read from, or'+ CRLF +;
          'write to a file. There are many possible causes of this problem some of which' + CRLF +;
          'may require changes to your setup. So if the error persists after re-starting' + CRLF +;
          'the application you should contact your IT support person'
      
      ******<Lock Errors>**************************************************************
      CASE INLIST( lnError, 101,102,108,109,110,111,130 )
        lcStr = 'The Application has reported a File Lock or Access Error.  This may'+ CRLF +;
          'be caused by someone else accessing the same record that you are trying'+ CRLF +;
          'to use. If you are sure that this is not the case, and re-starting the  '+ CRLF +;
          'application does not cure the problem, you should contact your IT support person'

      ******<Memory Problems>**************************************************************
      CASE INLIST( lnError,22,43,1149,1150,1151,1308,1600,1881,1986)
        lcStr = 'The Application has run low on memory.  Quit one or more'+ CRLF + ;
            'applications or restart Windows.  You may also benefit'+ CRLF + ;
            'from re-booting your PC after exiting Windows'
        lcAction = 'QuitRB'

      ******<Diskspace Problems>***********************************************************
      CASE INLIST( lnError, 56 )  && Not enough disk space for "<file name>"
        lcStr = 'The Application has run low on disk space.  Try increasing'+ CRLF + ;
          'the space on the hard-disk by removing unused files and'+ CRLF + ;
          're-start the application.  Contact your IT support person if'+ CRLF + ;
          'you are unsure of what to do next.'
        lcAction = 'QuitRB'

      ******<Printer Not Ready>*******************************************
      CASE INLIST( lnError,125,1643,1644 )  && Printer is not ready
        lcStr = 'The selected printer is reporting "NOT READY".  Please'+ CRLF + ;
          'ensure printer is on-line and ready to print before clearing'+ CRLF + ;
          'this message and allowing the system to try again.'+ CRLF + CRLF + ;
          'If the printer appears to be ready to print and this error' + CRLF +;
          'is continually appearing, you may be experiencing network' + CRLF +;
          'problems and you should contact your IT support person in this case.'
        llLogError = .F.
        lcAction = 'Retry'

      ******<Feature Not Available>***************************************
      CASE INLIST( lnError,1001,1443,1444,1445,1999)  && Feature is not available
        lcStr = 'An attempt has been made to access a feature not available'+ CRLF +;
          'in this version of the software.'

      ******<API Errors>**************************************************
      CASE INLIST( lnError, 1711, 1726,2027,2028)
        lcStr = 'Windows has reported an error from one of its system files'+ CRLF +;
          'This is nothing to do with the application, and if the problem'+ CRLF +;
          'persists after re-starting your PC, you should seek advice from your'+ CRLF +;
          'IT Support Person'

      ******<OLE Errors>***************************************************
      CASE INLIST( lnError,1420,1421,1422,1423,1424,1426,1427,1428,1429,1431,1434,1436,1437,1438, 1440)
        lcStr = 'The Application has reported an error while trying to access an'+ CRLF +;
          'external program or control. The most probable cause is that the target' + CRLF +;
          'is not available at this time. Please re-try after re-starting the application' + CRLF +;
          'and if the problem persists contact your IT Support Person'

      ******<Disk Access Error>***************************************************
      CASE INLIST( lnError,1410,1961,1962,1963 )  
        lcStr = 'The Application was unable to write temporary work files.'+ CRLF + ;
          'This may be due to a full or missing temporary directory or'+ CRLF + ;
          'insufficient priviliges on a network drive. If you are in doubt' + CRLF + ;
          'you should contact your IT support person.'
        llLogError = .F.
        lcAction = 'QuitRB'

      ********************************************************************
      CASE INLIST(lnError,1462,1522, 1941)  && Bad unknown error, VFP will exit
        lcStr = 'The Application has experienced an internal error which' + CRLF +;
          'cannot be identified.  To absolutely ensure data integrity the ' + CRLF +;
          'Application will be safely closed down.'
        lcAction = 'QuitRB'

      ******<Connect/SPT Errors>***************************************
      CASE INLIST( lnError,1463,1465,1466,1472,1474,1475,1476,1477,1479,1488,1491,1492,1493)  
        lcStr = 'The application has reported a problem trying to connect'+ CRLF +;
          'to the database. Please re-try after re-starting the application'

      ********************************************************************
      CASE INLIST( lnError,1589)  && MULTILOCKS OFF -- set on and return
        *** No need to log this one - it's fixable at run time
        SET MULTILOCKS ON
        llLogError = .F.
        llQuitApp = .F.
        lcAction = 'Retry'

      ********************************************************************
      CASE INLIST( lnError,1545 )  && uncommited changes - revert
        *** But log the error - we need to know about it!
        *** Get table alias for bad table
        LOCAL lnStt, lnEnd, lcTable, lcMessage
        lcMessage = This.aErrs[2]
        lnStt   = AT( '"', lcMessage ) + 1
        lnEnd   = AT( '"', lcMessage, 2 ) - 1
        lcTable = SUBSTR( lcMessage, lnStt, lnEnd - lnStt + 1 )
        IF USED( lcTable ) AND CURSORGETPROP("Buffering", lcTable ) > 1
          TABLEREVERT( .T., lcTable )
        ENDIF
        llQuitApp = .F.
        lcAction  = 'Retry'
        
      ********************************************************************
      CASE INLIST( lnError,1736)  && Cannot instantiate, may be low memory
        lcStr = 'The Application was unable to create an application object.  This may'+ CRLF +;
          'be due to low memory or an invalid DLL library.  Quit one or more'+ CRLF +;
          'applications or try restarting Windows.  You may also benefit'+ CRLF +;
          'from re-booting your PC after exiting Windows'
        lcAction = 'QuitRB'

      ********************************************************************
      CASE INLIST( lnError,1527,1528,1753, 1754)  && Invalid .DLL -- may be corrupt
        lcStr = 'Windows reports an error in one of its essential system files.'+ CRLF + ;
            'This has nothing to do with the Application. You can try re-booting '+ CRLF +;
            'your PC after exiting Windows, to see if Windows can resolve the '+ CRLF +;
            'problem, but the nature of this error may mean that you will need '+ CRLF +;
            'to have the installation of Windows checked for errors.'
        lcAction = 'QuitRB'

      ********************************************************************
      CASE INLIST( lnError,1907)  && Invalid drive -- check you are logged-on
        lcStr = 'The Application has referenced an invalid drive.  Please ensure'+ CRLF +;
          'that you have logged-on to the network successfully.'+ CRLF +;
          'Please try restarting Windows and you may also benefit'+ CRLF +;
          'from re-booting your PC after exiting Windows'
        lcAction = 'QuitRB'

      ********************************************************************
      CASE INLIST( lnError,1955)  && WIN.INI/Registry corrupt -- try re-booting
        lcStr = 'Windows reports an error in its .INI File and/or Registry.'+ CRLF + ;
            'This has nothing to do with the application. You can try re-booting '+ CRLF +;
            'your PC after exiting Windows, to see if Windows can resolve the '+ CRLF +;
            'problem, but the nature of this error may mean that you will need '+ CRLF +;
            'to re-install the Windows Operating System.'
        lcAction = 'QuitRB'

      ********************************************************************
      CASE INLIST( lnError,1956,1957,1958)  && Unable to access printer
        lcStr = 'The Application cannot access the specified printer, although '+ CRLF + ;
          'the printer itself appears to be ready to print.'+ CRLF + CRLF + ;
          'There may, therefore be a system problem and should you contact' + CRLF +;
          'your IT support person immediately.'

      ********************************************************************
      CASE INLIST( lnError,1526)  && ODBC Connection Error
        *** Check for deadlock errors here - we can re-try them without logging them
        LOCAL ARRAY laErr[1]
        AERROR( laErr )
        *** Some other connection error
        lcStr = 'The database is currently unavailable. This may be due to a database problem ' + CRLF ;
              + 'but there may also be a system problem and if this error occurs repeatedly' + CRLF ;
              + 'you should check with your IT support person.'

      ********* OTHER FOXPRO ERROR **********************************
      OTHERWISE
        lcStr = 'The Application has reported an internal error identified as ' + CRLF ;
              + 'Error #: ' + TRANSFORM( lnError )
    ENDCASE
    *** Return a parameter object
    loPObj = CREATEOBJECT( 'empty' )
    WITH loPObj
      AddProperty( loPObj, 'cMessage', lcStr )
      AddProperty( loPObj, 'cAction', lcAction )
      AddProperty( loPObj, 'lLogErr', llLogError )
      AddProperty( loPObj, 'lQuitApp', llQuitApp )
    ENDWITH
    RETURN loPObj
  ENDFUNC

  ********************************************************************
  *** [P] CLOSEAPP(): Close Down the application. Rollback Transactions, 
  *** Revert Pending changes, Close forms. NB: Error handling is disabled!
  ********************************************************************
  PROTECTED FUNCTION CloseApp()
    LOCAL ARRAY laSess[1], laUsed[1]
    LOCAL lnSess, lnCnt, lnCntUsed, lnSessCnt
    *** Check all Data Sessions
    lnSess = ASESSIONS( laSess )
    FOR lnCnt = 1 TO lnSess
      SET DATASESSION TO ( laSess[ lnCnt ] )
      *** Roll Back Any Transactions
      IF TXNLEVEL() > 0
        DO WHILE TXNLEVEL() > 0
          ROLLBACK
        ENDDO
      ENDIF
      *** Revert Tables and Close Them
      lnCntUsed = AUSED(laUsed)
      FOR lnSessCnt = 1 TO lnCntUsed
        SELECT ( laUsed[ lnSessCnt, 2 ] )
        IF CURSORGETPROP('Buffering') > 1
          TABLEREVERT( .T. )
        ENDIF
        USE
      NEXT
    NEXT

    *** Now Close any open forms
    IF _Screen.FormCount > 0
      FOR lnCnt = _Screen.FormCount TO 1 STEP -1
        loObj = _Screen.Forms[ lnCnt ]
        loObj.Release()
      NEXT
    ENDIF
 
    *** Issue a Clear Events 
    CLEAR EVENTS

    *** Restore System Menu if in DevMode
    IF VERSION(2) # 0
      SET SYSMENU TO DEFAULT
    ENDIF
    
    *** And Exit here
    RETURN
  ENDFUNC

  ********************************************************************
  *** [P] EXECACTION(): Call CloseApp if not going to Retry
  ********************************************************************
  PROTECTED FUNCTION ExecAction( tcToDo )
  LOCAL lcToDo
    WITH This
      *** Default to QUIT if we are wrong here
      lcToDo = IIF( VARTYPE( tcToDo ) = "C" AND NOT EMPTY( tcToDo ), ALLTRIM( LOWER( tcToDo )), "quit" )
      IF 'quit' $ lcToDo
        *** If in Runtime, command = QUIT, otherwise CANCEL
        lcToDo = IIF( VERSION(2) = 0, "quit", "cancel" )
        *** But close down as best we can anyway
        .CloseApp()
      ELSE
        *** We are just going to re-try the last command anyway
        lcToDo = 'retry'
      ENDIF
      RETURN lcToDo 
    ENDWITH
  ENDFUNC

  ********************************************************************
  *** [P] GENHEADER(): Generate the header information for the Error log
  ********************************************************************
  PROTECTED FUNCTION GenHeader( tcUserMsg )
  LOCAL lcMerge, lcRetVal, lcLocation
    *** We need Textmerge here
    lcMerge = SET( "Textmerge" )
    IF lcMerge # "ON"
      SET TEXTMERGE ON
    ENDIF
    IF NOT EMPTY( This.cCompany )
      lcLocation = "At " + This.cCompany
      lcLocation = lcLocation + IIF( NOT EMPTY( This.cEMail ), CRLF + "(Email: " + This.cEMail + ")", "")
      lcLocation = lcLocation + IIF( NOT EMPTY( This.cPhone ), CRLF + "(Phone: " + This.cPhone + ")", "")
      lcLocation = lcLocation + CRLF
    ELSE
      lcLocation = ""
    ENDIF
    *** Create the header as a text block
    TEXT TO lcRetVal NOSHOW
------------------------------------------------------------------------------
<<DEF_EMSUBJECT>> occurred at <<DATETIME()>> 
<<lcLocation>>To the user logged in as <<This.cUserLogin>>
------------------------------------------------------------------------------
The error occurred in: <<This.cProgName>>
The VFP Error Number was:   <<This.aErrs[1]>>
And the VFP Error Text was: <<This.aErrs[2]>>
------------------------------------------------------------------------------ 
The User Message displayed was: <<tcUserMsg>>
    ENDTEXT
    *** Restore a priori status
    IF lcMerge = "OFF"
      SET TEXTMERGE OFF
    ENDIF
    *** And return the text block
    RETURN lcRetVal
  ENDFUNC

  ********************************************************************
  *** [P] GETDATADETS(): Get the state of the data tables that are open
  ********************************************************************
  PROTECTED FUNCTION GetDataDets()
  LOCAL ARRAY laUsed[1]
  LOCAL lnSelect, lnCnt, lcStr, lcStat, lnFld, lcFVal, lnBMode, lcBMode
    lnSelect = SELECT()
    lcStr  = CRLF + DIVLINE ;
           + CRLF + "*** Open Table Information ***" ;
           + CRLF + DIVLINE + CRLF
    FOR lnCnt = 1 TO AUSED(laUsed)
      *** Select the work area
      SELECT ( laUsed[ lnCnt, 2 ] )
      lnBMode = CURSORGETPROP( "BUFFERING" )
      lcBMode = IIF( lnBMode = 1, "NONE", ;
                  IIF( lnBMode = 2, "Pessimistic Row", ;
                  IIF( lnBMode = 3, "Optimistic Row", ;
                  IIF( lnBMode = 4, "Pessimistic Table", "Optimistic Table"))))
      *** Get the Alias name and Record Number
      lcStr = lcStr + 'TABLE: ' + ALIAS() ;
            + ' (Buffering: ' + lcBMode + ')' + CRLF
      lcStr = lcStr + SPACE( 5 ) + 'Record #' + TRANSFORM( RECNO() )
      *** Check the record status
      IF CURSORGETPROP( 'Buffering' ) > 1 
        lcStat = NVL( GETFLDSTATE( -1 ), '0' )
        IF '3' $ lcStat OR '4' $ lcStat
          lcStr = lcStr + SPACE(5) + '  ** APPENDED **'
        ENDIF
        *** And Deletion Status
        IF GETFLDSTATE(0)%2 = 0
          IF DELETED()
            lcStr = lcStr + SPACE(5) + '  ** DELETED **'
          ELSE
            lcStr = lcStr + SPACE(5) + '  ** UNDELETED **'
          ENDIF
        ENDIF
      ENDIF
      lcStr = lcStr + CRLF
      *** Get the Details of the current record too
      FOR lnFld = 1 TO FCOUNT()
        lcFVal = ALLTRIM( TRANSFORM( EVAL( FIELD( lnFld ))))
        IF CURSORGETPROP('Buffering')>1
          *** Flag if Edited
          lcStr = lcStr + SPACE(5) ;
            + IIF( INLIST( GETFLDSTATE( lnFld ), 2, 4), ' *', '  ') + ' ' ;
            + TYPE( FIELD( lnFld )) + ' ' ;
            + PADR( FIELD( lnFld ) + ' ', 25, '.' ) + ' ' ;
            + lcFVal + CRLF
        ELSE
          *** Can't tell if it is edited or not
          lcStr = lcStr + SPACE(5) ;
            + '   ' + TYPE( FIELD( lnFld )) + ' ' ;
            + PADR( FIELD( lnFld ) + ' ', 25, '.' ) + ' ' ;
            + lcFVal + CRLF
        ENDIF
      NEXT
      lcStr = lcStr + CRLF
    NEXT
    RETURN lcStr
  ENDFUNC

  ********************************************************************
  *** [P] GETMEMORY(): Get the contents of memory into a string for Error Logging
  ********************************************************************
  PROTECTED FUNCTION GetMemory()
    *** We need to use a temporary file for this
    IF EMPTY( This.cFileName )
      This.cFileName = SUBSTR(SYS(2015), 3) + ".txt"
    ENDIF
    SET CONSOLE OFF
    SET ALTERNATE TO ( This.cFileName )
    SET ALTERNATE ON
    *** Dump the memory
    ? DIVLINE
    ? "*** Contents of memory ***"
    ? DIVLINE
    LIST MEMORY LIKE *
    SET ALTERNATE OFF
    SET ALTERNATE TO
    SET CONSOLE ON
  ENDFUNC

  ********************************************************************
  *** [P] GETSTACK()(): Retrieve calling stack information
  ********************************************************************
  PROTECTED FUNCTION GetStack()
    LOCAL ARRAY laStax[1]
    LOCAL lcStr, lnStax, lnCnt
    WITH This
      lnStax = ASTACKINFO( laStax )
      lcStr  = CRLF + DIVLINE ;
             + CRLF + "*** Calling Stack ***" ;
             + CRLF + DIVLINE + CRLF
      FOR lnCnt = lnStax TO 1 STEP -1
        *** Ignore anything in the error handler
        IF "errorhandler" $ LOWER( laStax[ lnCnt, 4 ] )
          LOOP
        ENDIF
        lcStr = lcStr + "Called From Line " ;
              + TRANSFORM( laStax[lnCnt, 5] ) + " of " ;
              + laStax[ lnCnt, 3] + " in " + laStax[ lnCnt, 2] ;
              + IIF( NOT EMPTY( laStax[ lnCnt, 6] ), CRLF + "[" + laStax[ lnCnt, 6] + "]", "") ;
              + CRLF
        IF EMPTY( .cProgName )
          *** Add the calling program name (first item we process)
          .cProgName = laStax[ lnCnt, 2]
        ENDIF
      NEXT
      RETURN lcStr
    ENDWITH
  ENDFUNC

  ********************************************************************
  *** [P] GETSTATUS(): Get the current status into a string for Error Logging
  ********************************************************************
  PROTECTED FUNCTION GetStatus()
    *** We need to use a temporary file for this
    IF EMPTY( This.cFileName )
      This.cFileName = SUBSTR(SYS(2015), 3) + ".txt"
    ENDIF
    SET CONSOLE OFF
    SET ALTERNATE TO ( This.cFileName )
    SET ALTERNATE ON
    *** Dump the memory
    ? DIVLINE
    ? "*** System Status ***"
    ? DIVLINE
    LIST STATUS
    SET ALTERNATE OFF
    SET ALTERNATE TO
    SET CONSOLE ON
  ENDFUNC

  ********************************************************************
  *** [P] HANDLEINTERNALERROR(): Retrieve error information about an 
  *** that occurred while inside the error handler itself! The information 
  *** is written to the cStatus Property so as to be available for display 
  *** to the end user if needed
  ********************************************************************
  PROTECTED FUNCTION HandleInternalError()
    LOCAL ARRAY laErr[1]
    AERROR( laErr )
    lcTxt = "An error occurred while the error handler was processing a system error" + CRLF
    lcTxt = lcTxt + "VFP Error #" + TRANSFORM( laErr[1] )
    *** Deal 
    IF laErr[1] = 1526 
      *** It was an ODBC Connectivity Error. 
      *** The Source error number is in Element 5, and the Error Text in Element 3 
      lcTxt = lcTxt + "Details: " + TRANSFORM( laErr[5] ) + ": " + ALLTRIM( laErr[3] )
    ELSE
      IF INLIST( laErr[1], 1427, 1429 )
        *** This is an OLE Error
        *** Application name in Element 4, Error Text in Element 3
        lcTxt = lcTxt + "Details: " + ALLTRIM( laErr[4] ) + ": " + ALLTRIM( laErr[3] )
      ELSE
        *** This was a standard FoxPro Error
        *** The text is in Element 2 of the Error Array
        lcTxt = lcTxt + "Details: " + ALLTRIM( laErr[2] )
     ENDIF
    ENDIF
    *** And update the status property accordingly
    .cStatus = lcTxt + CRLF
  ENDFUNC
  
  ********************************************************************
  *** [P] OPENLOG: Open, or create if missing, the error log table
  ********************************************************************
  PROTECTED FUNCTION OpenLog()
    WITH This
      IF NOT FILE( 'tlc_error_log.dbf' )
        CREATE TABLE tlc_error_log FREE ( ;
          ierr_pk   I AUTOINC NEXTVALUE 1 STEP 1, ;
          terrdt    T (   8 ), ;
          cerruser  C (  20 ), ;
          cerrnum   C (   5 ), ;
          cerrtxt   C ( 200 ), ;
          cerrprg   C (  50 ), ;
          merrdets  M (   4 ))
      ELSE
        USE tlc_error_log AGAIN IN 0
      ENDIF
      
      *** Set buffering OFF for the error log
      CURSORSETPROP("Buffering", 1, 'tlc_error_log' )
      
      *** Return
      RETURN
    ENDWITH
  ENDFUNC && OpenLog

  ********************************************************************
  *** [P] SENDEMAIL(): Use Outlook to send an E-Mail
  ********************************************************************
  PROTECTED FUNCTION SendEMail( tcDetails )
  LOCAL lnResult
    WITH This
      *** Use ShellExecute to get to whatever E-Mail Client is registered
      lnResult = .UserEMail( tcDetails )
      *** And ShellExec returns 32+ for success! 
      lnResult = lnResult - 32
      *** Did we get the message out?
      IF lnResult <= 0
        MESSAGEBOX( "Unable to Send E-Mail message at this time", 48, 'E-Mail Access Denied', 3000 )
      ENDIF
    ENDWITH
  ENDFUNC

  ********************************************************************
  *** [P] SHOWMESSAGE(): Displays a user friendly error message with 
  *** appropriate action options
  ********************************************************************
  PROTECTED FUNCTION ShowMessage( toActions, tlSendEMail )
  LOCAL lcTit, lnOpt, lcMsg, lnRes, llRetVal
    WITH This
      IF VARTYPE( toActions ) # "O"
        *** We have no data! This should NEVER happen
        lcTit = MSG_MinorTitle
        lnOpt = 64
        lcMsg = .cStatus + CRLF + MSG_PressOK
      ELSE
        lcTit = MSG_StdTitle 
        lnOpt = IIF( tlSendEmail, 36, 64 ) 
        lcMsg = toActions.cMessage + CRLF + ALLTRIM( .cStatus ) + CRLF
        IF tlSendEMail
          lcMsg = lcMsg + MSG_PromEMail + .cAddr
        ELSE
          lcMsg = lcMsg + MSG_VFPError 
        ENDIF
        IF LOWER( toActions.cAction ) = "retry"
          *** Offer the user the option to cancel the Retry
          lcMsg = lcMsg + CRLF + MSG_AppGoOn + CRLF
          IF lnOpt = 36
            *** Add Cancel Option to Yes/No and note about suppressing E-Mail
            lnOpt = 35
            lcMsg = lcMsg + MSG_NoEMail + CRLF
          ELSE
            *** Add a Cancel option to the OK button
            lnOpt = 65
          ENDIF 
        ELSE
          *** Warn the user that the application will close
          lcMsg = lcMsg + CRLF + MSG_WillClose + CRLF
          IF LOWER( toActions.cAction ) = "quitrb"
            *** And that a roll-back will be performed
            lcMsg = lcMsg + MSG_RollBack + CRLF
          ENDIF        
        ENDIF                     
      ENDIF      
      lnRes = MESSAGEBOX( lcMsg, lnOpt, lcTit )
      *** Return .T. only if they chose either YES or OK buttons
      RETURN INLIST( lnRes, 1, 6 )
    ENDWITH
  ENDFUNC

  ********************************************************************
  *** [P] WRITELOG(): Dump error details to specified table
  ********************************************************************
  PROTECTED FUNCTION WriteLog( tcDetails )
  LOCAL lnDSID
    WITH This
      *** Get back to the handler's datasession
      lnDSID = .DataSessionId
      SET DATASESSION TO (.nDSid)
      *** The table is already opened, so just write the details
      APPEND BLANK IN tlc_error_log
      REPLACE tErrDT WITH DATETIME(), ;
              cErrUser WITH SYS(0), ;
              cErrNum  WITH TRANSFORM( .aErrs[1] ), ;
              cErrTxt WITH SUBSTR(.aErrs[2],1,200), ;
              cErrPrg WITH RIGHT(.cProgName,50), ;
              mErrDets WITH tcDetails IN tlc_error_log
      *** We added the record, so get it's ID value for reporting
      .cStatus = "This error has been recorded under ID# " ;
               + TRANSFORM( tlc_error_log.ierr_pk ) + CRLF 
      *** Restore the datasession
      SET DATASESSION TO (lnDSid)
    ENDWITH
  ENDFUNC

  ********************************************************************
  *** [P] USEREMAIL: Send E-Mail using whatever means we can
  ********************************************************************
  PROTECTED FUNCTION UserEMail( tcDetails )
  LOCAL ARRAY laJunk[1]
  LOCAL lcMail, lnRes
    WITH This
      *** And Embed CRLF characters in text string for the body
      tcDetails = STRTRAN( tcDetails, CHR(13)+CHR(10), "%0d%0a" )
      
      *** First the TO and Copy To addresses
      lcMail = "mailto:" + DEF_EMAILTO
      IF NOT EMPTY( DEF_CCMAILTO )
        lcMail = lcMail + IIF( AT( '?', lcMail ) = 0, "?", "&" ) + "CC=" + DEF_CCMAILTO
      ENDIF
      *** Now concatenate the values correctly      
      lcMail = lcMail + IIF( AT( '?', lcMail ) = 0, "?", "&" ) + "Subject=" + DEF_EMSUBJECT
      lcMail = lcMail + IIF( AT( '?', lcMail ) = 0, "?", "&" ) + "Body=" + tcDetails
      **********************************************************
      *** Check that the library has been set up and open it if not already done.
      **********************************************************
      lnRes = ADLLS( laJunk )
      IF lnRes = 0 OR  NOT ( ASCAN( laJunk, 'Send', 1, -1, 1, 15 ) > 0)
        *** We don't have the function available
        DECLARE INTEGER ShellExecute IN shell32 ;
                INTEGER lnhwnd, STRING lcOperation, ;
                STRING lcFile, STRING lcParameters, ; 
                STRING lcDirectory, INTEGER lnShowCmd
      ENDIF
      *** Send the message
      lnRes = ShellExecute( 0, "open", lcMail, "", "", 3 )
      *** Return Final Status    
      RETURN lnRes
    ENDWITH
  ENDFUNC && UserEMail

ENDDEFINE
