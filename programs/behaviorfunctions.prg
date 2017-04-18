Function Subscribe(toObj As Form)
Local i, lcBehavior, loBehavior, x, othis As Session, labehaviors(1), oCollection As Collection, loObj, llok
Local j, laRequired(1), lRegister, lcModifier, lcMess

If Vartype(toObj) <> 'O'
	Return
Endif

*- The calling object requires a "cBehaviors" property
*- If not this will create one and transfer the contents of the "Tag" property to it
*- the Tag property is used for base classes to utilize this technique.
Try
	toObj.cBehaviors = Chrtran(toObj.cBehaviors,',',' ')
	llok = .T.
Catch
	Try
		toObj.AddProperty('cBehaviors',toObj.Tag)
		llok = .T.
	Catch
	Finally
		If llok
			toObj.Tag = ''
			toObj.cBehaviors = Chrtran(toObj.cBehaviors,',',' ')
		Endif
	Endtry
Endtry


If llok And Not Empty(toObj.cBehaviors)
*- Remove excess spaces
	For i = 1 To Occurs('  ',toObj.cBehaviors)
		toObj.cBehaviors = Strtran(toObj.cBehaviors,'  ',' ',1,Occurs('  ',toObj.cBehaviors))
	Endfor

*- Attach a collection to the _screen object to serve as a global behavior server
	Try
*- already there
		oCollection = _Screen.Behaviors
	Catch
*- Not yet there so add it
		_Screen.AddObject('Behaviors','Collection')
		oCollection = _Screen.Behaviors
*- Add Andy's error handler for test
		If Empty(On('error')) Or Lower(On('error')) <> 'errorhandler'
*			On Error Do ErrorHandler
		Endif
	Finally
	Endtry

	For i = 1 To Alines(labehaviors,toObj.cBehaviors,.T.,' ')

		If Not Empty(labehaviors(i)) And Vartype(labehaviors(i)) = 'C'

			lcBehavior = Upper(Ltrim(labehaviors(i)))
			lcodifier = ''
			IF '-' $ lcBehavior
				lcModifier = SUBSTR(lcBehavior,AT('-',lcBehavior)+1,1)
				lcBehavior = UPPER(LEFT(lcBehavior,AT('-',lcBehavior)-1))
			ENDIF 
			IF EMPTY(lcModifier) OR NOT INLIST(lcModifier,'A','B')	&& After or Before
				lcModifier = ''
			ENDIF 

*- Add the behaviors which are not yet already registered
			Try
				loBehavior = oCollection.Item(lcBehavior)
			Catch
				Try

					loBehavior = Createobject('cBehaviorWrapper',lcBehavior)	&& Wrap the behavior in a session
					oCollection.Add(loBehavior,lcBehavior)						&& Place it in the be screen collection
					loBehavior.cRequiredProperties = Strtran(loBehavior.cRequiredProperties,'  ',' ',1,99)
					loBehavior.cRequiredMethods = Strtran(loBehavior.cRequiredMethods,'  ',' ',1,99)
					loBehavior.cEvents = Strtran(loBehavior.cEvents,'  ',' ',1,99)
				Catch
					If Version(2) = 2
					TEXT TO lcMess TEXTMERGE NOSHOW PRETEXT 3
						Failed to Register unknown Behavior <<lcBehavior>> 
						FOR <<toObj.Class>> - <<toObj.Name>>
						
						Behavior <<lcBehavior>> ignored
						endtext
						Messagebox(lcMess)
					Endif
				Endtry
			Finally
			Endtry

			If Vartype(loBehavior) = 'O' AND Vartype(loBehavior.oBehavior) = 'O'
				lok = .T.
				If Not Empty(loBehavior.cRequiredProperties)	&& does the registering object have the necessary prperties
					For j = 1 To Alines(laRequired,loBehavior.cRequiredProperties,.T.,' ')
						If !Pemstatus(toObj,laRequired(j),5)
							lok = .F.
							Exit
						Endif
					Endfor
				Endif
				If lok And Not Empty(loBehavior.cRequiredMethods)	&& does the registering object have the necessary methods
					For j = 1 To Alines(laRequired,loBehavior.cRequiredMethods,.T.,' ')
						If !Pemstatus(toObj,laRequired(j),5)
							lok = .F.
							Exit
						Endif
					Endfor
				Endif
				If lok	&& does the registering object have the necessary properties and methods
					lok = loBehavior.lRegister	
					For x = 1 To Alines(laEvents,loBehavior.cEvents,.T.,' ')
						If Not Empty(laEvents(x)) And Pemstatus(toObj,laEvents(x),5)
							IF lcModifier = 'A'		&& Call delegate after user code
								lcModifier = '1'
							ELSE
								lcModifier = '0'	&& Call delegate before user code
							ENDIF 
							Bindevent(toObj,Alltrim(laEvents(x)),loBehavior,'process',VAL(lcModifier)+4)	&& Bind the evetn with the behavior
							lok = .T.
						Endif
					Endfor
					If lok && Pre-Visible "One Shot" activities 
						loBehavior.Register(toObj)
					Endif
					lok = .T.
				Endif
			Endif

		Endif

	Endfor
Endif

*- Drill into all objects for registration
If Pemstatus(toObj,'objects',5)
	For Each loObj In toObj.Objects
		Subscribe(loObj)
	Endfor
Endif
Return
Endfunc

*- Data session wrapper for library based behaviors
Define Class cBehaviorWrapper As Session
	cBehavior			=	''	&& Class Name of wrapped behavior
	oBehavior			=	''	&& Object of wrapped behavior

	cRequiredProperties	=	''	&& Space delimited list of required properties 
	cRequiredMethods	=	''	&& Space delimited list of required Methods 
	cEvents				=	''	&& Space delimited list of events to register

	lRegister			=	.f.	&& Is there a special INIT registration required
	DataSession			=	2	&& Private data session

	Function Init
	Lparameters tcBehavior

	If Not Empty(tcBehavior)
		This.cBehavior = tcBehavior
	Endif

	Function cBehavior_assign
	Lparameters tcBehavior
	If Not Empty(tcBehavior) And Vartype(tcBehavior) = 'C'
		This.oBehavior = Createobject(tcBehavior)
		This.cRequiredProperties = UPPER(Strtran(This.oBehavior.cRequiredProperties,'  ',' '))
		This.cRequiredMethods = UPPER(Strtran(This.oBehavior.cRequiredMethods,'  ',' '))
		this.cEvents = UPPER(STRTRAN(this.oBehavior.cEvents,'  ',' '))
		this.lRegister = ' REGISTER ' $ ' '+ this.cEvents + ' '
		THIS.cEvents = STRTRAN(' '+THIS.cEvents+' ',' REGISTER ',' ')
	Else
		If Version(2)=2
			Messagebox('empty behavior')
			Set Step On
		Endif
		This.oBehavior = .Null.
	Endif

	FUNCTION Destroy
		this.oBehavior = .null.
	ENDFUNC
	
	Function Register
	Lparameters toObject
	This.oBehavior.oCaller = toObject
	This.oBehavior.cEvent = 'REGISTER'
	Try
		This.oBehavior.oform = This.Findform(toObject.Parent)
	Catch
		This.oBehavior.oform = toObject
	Endtry
	this.oBehavior.oParent = this
	This.onProcess()
	this.oBehavior.oParent = ''


	Function Process
	Lparameters p1,p2,p3,p4,p5
	Local laEvents(1)
	= Aevents(laEvents,0)	&& Accquire the caller and event
	This.oBehavior.oCaller = laEvents(1,1)	&& Store the object reference
	This.oBehavior.cEvent = laEvents(1,2)	&& Store the event which invoked this handler
	this.oBehavior.oParent = this
	Try
		This.oBehavior.oform = This.Findform(This.oBehavior.oCaller.Parent)
	Catch
		This.oBehavior.oform = This.oBehavior.oCaller
	Endtry
	This.onProcess(p1,p2,p3,p4,p5,laEvents(1,1))
	this.oBehavior.oParent = ''

	Function onProcess
	Lparameters p1,p2,p3,p4,p5,toCaller
	Local lccmd, i, lok, lnOldSession, launbind(1)
	If This.oBehavior.lSwapDatasession	&& If the behavior needs to be in the same datasession as the caller
		lnOldSession = This.DataSessionId
		This.DataSessionId = This.oBehavior.oform.DataSessionId
	Endif
	lccmd = 'this.oBehavior.Process('
	For i = 1 To 5
		lccmd = lccmd + 'p'+Transform(i)+','
	Endfor
	lccmd = Trim(lccmd,0,',') + ')'
	lok = Evaluate(lccmd)
	If Not Empty(lnOldSession)
		This.DataSessionId = lnOldSession
	Endif
	If Not Empty(This.oBehavior.cUnbind)	&& Unbind "One-Shot" behavior methods - for setup and such
		For i = 1 To Alines(launbind,This.oBehavior.cUnbind,5,' ')
			Unbindevents(This,'process',toCaller,launbind(i))
		Endfor
	Endif
	This.oBehavior.oform = .Null.
	This.oBehavior.oCaller = .Null.
	Return lok

	Function Findform	&& drill back in hierarchy to find form parent
	Lparameters toObj
	Do While toObj.BaseClass <> 'Form'
		toObj = toObj.Parent
	Enddo
	Return toObj


Enddefine

Function SingleInstance
********************************************************************
*** Name.....: FIRSTTIME.PRG
*** Author...: Marcia Akins & Andy Kramek
*** Date.....: 01/11/2004
*** Notice...: Copyright (c) 2004 Tightline Computers, Inc
*** Compiler.: Visual FoxPro 09.00.0000.2124 for Windows
*** Function.: Determine whether an instance of the application is already running
*** .........: This function uses the creation of a named MUTEX to determine whether
*** .........: or not the application is already running. This function should be called
*** .........: near the top of the application's main program to create a named object
*** .........: that can be checked every time the application is started.  If the named
*** .........: object exists, the function will try to activate the FoxPro running application.
*** Returns..: Logical
********************************************************************
Lparameters tcAppName
Local lcMsg, lcAppName, lnMutexHandle, lnhWnd, llRetVal

lcMsg = ''
Set Asserts On
If Empty( Nvl( tcAppName, '' ) )
	TEXT TO lcMsg NOSHOW
    This is another Brain Dead Programmer Error.
    You must pass the name of an application to FirstTime().
    Have a nice day now...
	ENDTEXT
	Assert .F. Message lcMsg
	Return
Endif

*** Format the passed in program name
lcAppName = Upper( Alltrim( tcAppName ) ) + Chr( 0 )

*** Declare API functions
Declare Integer CreateMutex In WIN32API Integer lnAttributes, Integer lnOwner, String @lcAppName
Declare Integer GetProp In WIN32API Integer lnhWnd, String @lcAppName
Declare Integer SetProp In WIN32API Integer lnhWnd, String @lcAppName, Integer lnValue
Declare Integer CloseHandle In WIN32API Integer lnMutexHandle
Declare Integer GetLastError In WIN32API
Declare Integer GetWindow In USER32 Integer lnhWnd, Integer lnRelationship
Declare Integer GetDesktopWindow In WIN32API
Declare BringWindowToTop In Win32APi Integer lnhWnd
Declare ShowWindow In WIN32API Integer lnhWnd, Integer lnStyle

*** Try and create a new MUTEX with the name of the passed application
lnMutexHandle = CreateMutex( 0, 1, @lcAppName )

*** If the named MUTEX creation fails because it exists already, try to display
*** the existing application
If GetLastError() = 183

*** Get the hWnd of the first top level window on the Windows Desktop.
	lnhWnd = GetWindow( GetDesktopWindow(), 5 )

*** Loop through the windows.
	Do While lnhWnd > 0

*** Is this the one that we are looking for?
*** Look for the property we added the first time
*** we launched the application
		If GetProp( lnhWnd, @lcAppName ) = 1
*** Activate the app and exit stage left
			BringWindowToTop( lnhWnd )
			ShowWindow( lnhWnd, 3 )
			Exit
		Endif
		lnhWnd = GetWindow( lnhWnd, 2  )
	Enddo

*** Close the 'failed to open' MUTEX handle
	CloseHandle( lnMutexHandle )

Else

*** Add a property to the FoxPro App so we can identify it again
	SetProp( _vfp.HWnd, @lcAppName, 1)
	llRetVal = .T.
Endif

Return llRetVal
