* Procedure File
* File API


PROCEDURE NewDS

	SET TALK OFF
	SET EXCLUSIVE OFF
	*Set Multilock on for the Buffering in VFP
	IF SET("multilock")='OFF'
		SET MULTILOCK ON
	ENDIF

	SET DELETE ON
	SET SAFETY OFF
	SET EXACT OFF
	SET MACKEY TO
	ON KEY LABEL ALT+f10 *

	SET HOURS TO 24
	SET CENTURY ON
	IF TYPE("kCentury")#"N"
		PUBLIC kCentury
		kCentury = 10
	ENDIF
	*    SET CENTURY TO (VAL(LEFT(STR(YEAR(DATE()),4),2))) ROLLOVER MAX(0,(YEAR(DATE())+kCentury)-2000)
	SET CENTURY TO 19 ROLLOVER MAX(0,(YEAR(DATE())+kCentury)-2000)
	*
	* more required settings.
	SET NOTIFY OFF

	SET MESSAGE TO " "
	SET BELL OFF
	SET ESCAPE OFF
	*

	SET ANSI OFF
	SET AUTOSAVE ON
	SET CARRY OFF
	SET CENTURY ON
	SET CONFIRM OFF
	SET LOCK OFF
	SET MEMOWIDTH TO 250
	SET NEAR OFF
	SET NULLDISPLAY TO ""
	SET REPROCESS TO AUTOMATIC
	SET UNIQUE OFF

	SET UDFPARMS TO VALUE

	*- set lock priority to Lock all before writing anything.
	= SYS(3052,1,.T.)
	= SYS(3052,2,.T.)

	RETURN
ENDPROC

PROCEDURE EndTrans
	LPARAMETERS lSave, cMsg

	PRIVATE adbfs
	z=AUSED(adbfs)
	IF lSave
		FOR x=1 TO z
			SELECT (adbfs[x,1])
			IF NOT TableUpdateM(.T., .T., ALIAS())
				DO ("pcDone")
				RETURN EndTrans()
			ENDIF
		NEXT
		IF TXNLEVEL()>0
			END TRANSACTION
		ENDIF
	ELSE
		*
		*- notify the rollback is going to occur.
		*
		*
		FOR x=1 TO z
			SELECT (adbfs[x,1])
			DO CASE
				CASE ISREADONLY()
				CASE CURSORGETPROP('Buffer')<=1
				CASE INLIST( CURSORGETPROP( 'Buffering'), 2, 3) AND EMPTY(CHRTRAN(GETFLDSTATE(-1) ,'1',' '))
				CASE INLIST( CURSORGETPROP( 'Buffering'), 4, 5) AND GETNEXTMODIFIED(-1)=0
				OTHERWISE
					=TABLEREVERT(.T.)
			ENDCASE
		NEXT
		IF TXNLEVEL()>0
			ROLLBACK
		ENDIF
	ENDIF
	RETURN lSave
ENDPROC


PROCEDURE TableUpdateM
	LPARAMETERS lAll, lForce, cTable

	cTable = IIF(EMPTY(cTable), ALIAS(), ALLTRIM(cTable))

	cMesg = []
	DO CASE
		CASE EMPTY(cTable)
			cMesg = [No Table to Update]
		CASE ISREADONLY(cTable)
			*- read only cursor
		CASE CURSORGETPROP('Buffer', cTable ) <= 1
			SELECT (cTable)
			IF SYS(2011) = 'Record Locked'
				UNLOCK
			ENDIF

		CASE INLIST( CURSORGETPROP( 'Buffering', cTable), 4, 5) AND GETNEXTMODIFIED(0, cTable) = 0
			*- table buffer, and no changes
		CASE INLIST( CURSORGETPROP( 'Buffering', cTable), 2, 3) AND ( EMPTY(STRTRAN(GETFLDSTATE(-1, cTable),'1',' ')) OR ISNULL(GETFLDSTATE(-1, cTable)) )
			*- record buffering, and no changes
		OTHERWISE
			LOCAL errcnt
			errcnt = 0
			DO WHILE NOT TABLEUPDATE( lAll, lForce, cTable ) AND errcnt<5
				errcnt = errcnt + 1
			ENDDO
			IF errcnt=5
				cMesg = [Unable to update table ]+cTable+[ after 5 attempts.]
			ENDIF
	ENDCASE
	*
	IF NOT EMPTY(cMesg)
		RETURN .F.
	ENDIF
	RETURN .T.
ENDPROC



******************

PROCEDURE DATASESS
	*- Creates a New data session when called

	LOCAL o

	PRIVATE llSession
	llSession = SET('DataSession')
	*- create private datasession
	o= CREATEOBJECT("DataSession")

	*- switch to new datasession
	SET DATASESSION TO (o.DATASESSIONID)

	*- establish enviroment setting for datasession

	*-- Perform all of the SET commands for the OCIS system. We don't need to save
	*-- the values previously set because we are not executed from within a FoxPro
	*-- environment.

	SET TALK OFF
	SET CENTURY ON
	SET DELETE ON
	SET EXCLUSIVE OFF
	SET SAFETY OFF
	SET EXACT OFF
	SET ANSI OFF
	SET OPTIMIZE ON
	SET UNIQUE OFF
	SET MULTILOCK ON
	SET REPROCESS TO 0

	*- REturn instance of class DataSession
	RETURN o
ENDPROC



DEFINE CLASS DATASESSION AS FORM

	DATASESSION 	= 2			&& Private
	LastDataSession = 1
	BUFFERMODE		= 2			&& optimistic
	VISIBLE 		= .F.
	NAME 			= 'NewDataSession'

	*- hide these properties
	PROTECTED VISIBLE
	PROTECTED SHOW

	PROCEDURE DESTROY
		SET DATASESSION TO (THIS.DATASESSIONID)
		LOCAL x,z
		z=AUSED(atables)
		FOR x=1 TO z
			SELECT (atables[x,2])
			DO CASE
				CASE EMPTY(ALIAS())
				CASE ISREADONLY(ALIAS())
				CASE CURSORGETPROP('Buffer',ALIAS())<=1
				CASE NOT TABLEUPDATE(.T., .T., ALIAS())
					*- error
			ENDCASE
			USE
		NEXT
		*
		SET DATASESSION TO (THIS.LastDataSession)
	ENDPROC

	PROCEDURE INIT
		SET TALK OFF
		THIS.LastDataSession = llSession
	ENDPROC

	PROCEDURE ERROR
		LPARAMETERS nError, cMethod, nLine
		IF nError=1540
			SET DATASESSION TO 1
			RETURN
		ENDIF
		errcmd = ON("ERROR")
		&errcmd
		RETURN
	ENDPROC

ENDDEFINE


*( ORCA GLR 09/10/1996 09:51:03 TABLEX.PRG :

*- Procedure file of TableX procedural Add-in hooks.
****************
*- support procedures
*		Checks to see if the current hook is to be performed for this table.
*
PROCEDURE CheckTableList
	LPARAMETERS xTable, cField

	xTable = PADR(UPPER(m.xTable),8)
	cField = ALLTRIM(UPPER(m.cField))
	SELECT * FROM TABLES WHERE TableName = m.xTable INTO CURSOR q0Tables
	IF _TALLY > 0
		IF TYPE( 'q0tables.'+m.cField ) ="L"
			RETURN EVALUATE( 'q0tables.'+m.cField )
		ENDIF
	ENDIF
	RETURN .F.		&& failed.

	**********************



	**********************************************
PROCEDURE AClose( _open )
	PRIVATE _Close, x,z
	z=AUSED(_Close)
	FOR x=1 TO z
		IF ASCAN( _open, _Close[x,1] )=0
			USE IN (_Close[x,2])
		ENDIF
	NEXT
	RETURN


PROCEDURE fClear
	PRIVATE z
	FOR z=1 TO 300
		=FCLOSE(z)
	NEXT
	RETURN

PROCEDURE fcnt0
	DO fClear
	PRIVATE x,F
	x=0
	DO WHILE .T.
		F=FOPEN('c:\TEMP\'+LTRIM(STR(x))+'.TTT')
		IF F<0
			F=FCREATE('c:\TEMP\'+LTRIM(STR(x))+'.TTT')
		ENDIF
		IF F>=0
			x=x+1
		ELSE
			EXIT
		ENDIF
	ENDDO
	DO fClear
	RETURN x

	*-- The following pair of functions are used to save and restore the state of the work
	*-- areas. An array is passed by reference to each function, and is used to save and
	*-- restore whether or not a table is open in each work area. The restore function
	*-- closes tables in any work area which were not open when the save function
	*-- was called.

PROCEDURE SaveWA
	PARAMETER saveArray
	DIMENSION saveArray[225]
	PRIVATE i

	FOR i = 1 TO 225
		SELECT (i)
		saveArray[i] = ALIAS()
	ENDFOR
	RETURN

PROCEDURE RestoreWA
	PARAMETER saveArray
	PRIVATE i

	FOR i = 1 TO 225
		SELECT (i)
		IF NOT EMPTY(ALIAS()) AND EMPTY(saveArray[i])
			USE
		ENDIF
	ENDFOR
	DIMENSION saveArray[1]
	RETURN


	*!*****************************************************************************
	*!
	*!      Procedure: FIELDNO
	*!
	*!*****************************************************************************
PROCEDURE FieldNo
	PARAMETER cName, cAlias
	PRIVATE z, FieldArray, fl
	fl = SELECT()
	cAlias = IIF(EMPTY(cAlias), ALIAS(), m.cAlias)
	cAlias = IIF(TYPE("m.cAlias")="N", ALIAS(m.cAlias), m.cAlias)
	SELECT (m.cAlias)
	z=AFIELD(FieldArray)
	SELECT (fl)
	IF z>0
		z= ASCAN(FieldArray, UPPER(ALLTRIM(m.cName)))
		IF z > 0
			RETURN ASUBSCRIPT(FieldArray, z, 1)
		ENDIF
	ENDIF
	RETURN 0

	*****************************
	*	Database routines
	*****************************

	*!*****************************************************************************
	*!
	*!      Procedure: ISEMPTY
	*!
	*!      Called by: PRINTER.PRG
	*!
	*!          Calls: NORMALAL           (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE isempty
	*- test to see if  a dbf is empty

	PARAMETER ALIAS

	PRIVATE al, rec, result
	al=SELECT()
	m.ALIAS = NormalAL(m.ALIAS)
	IF EMPTY(m.ALIAS)
		RETURN .F.
	ENDIF ( EMPTY(m.ALIAS) )
	SELECT (m.ALIAS)
	rec=RECNO()

	GOTO TOP
	result = EOF()
	IF BETWEEN(rec,1,RECCOUNT())
		GOTO (rec)
	ELSE
		GOTO BOTTOM
		IF NOT EOF()
			SKIP
		ENDIF ( NOT EOF() )
	ENDIF ( BETWEEN(rec,1,RECCOUNT()) )
	SELECT (al)
	RETURN(result)



	*!*****************************************************************************
	*!
	*!      Procedure: NEWREC
	*!
	*!          Calls: FINDDELETE         (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE NewRec
	PARAMETER xAlias
	xAlias = NormalAlias(m.xAlias)
	IF FindDelete(m.xAlias)
		=RLOCK()
		BLANK
		RECALL
	ELSE
		APPEND BLANK
		=RLOCK()
	ENDIF
	RETURN

	*!*****************************************************************************
	*!
	*!      Procedure: SETORDER
	*!
	*!      Called by: GETIT              (procedure in UTILITY.PRG)
	*!
	*!          Calls: E0ORDER            (procedure in UTILITY.PRG)
	*!               : NORMALAL           (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE setorder
	PARAMETER ORDER, ALIAS
	PRIVATE onerr, err
	err=.F.
	onerr = ON("ERROR")
	ON ERROR DO e0order
	m.ALIAS = NormalAL( m.ALIAS )
	SET ORDER TO (ORDER) IN (m.ALIAS)
	ON ERROR &onerr
	RETURN(NOT err)

PROCEDURE e0order
	err = .T.
	RETURN

	*!*****************************************************************************
	*!
	*!      Procedure: SETFILTER
	*!
	*!*****************************************************************************
PROCEDURE setfilter
	PARAMETER EXPR
	EXPR = IIF(EMPTY(EXPR),"",EXPR)
	SET FILTER TO &EXPR
	RETURN .T.

	*!*****************************************************************************
	*!
	*!      Procedure: FORMATFIELDS
	*!
	*!*****************************************************************************
PROCEDURE FormatFields
	PARAMETER fldnum

	PRIVATE f1, t1, fld1
	fldnum = IIF(EMPTY(fldnum), 1, MIN(MAX(fldnum, 0),FCOUNT()) )
	fld1 = FIELD(fldnum)

	t1 = TYPE(fld1)

	DO CASE
		CASE t1$'CM'
			f1 = ALLTRIM(EVALUATE(fld1))
		CASE t1='N'
			f1 = LTRIM(STR(EVALUATE(fld1)))
		CASE t1='D'
			f1 = DTOC(EVALUATE(fld1))
		CASE t1='L'
			f1 = IIF(EVALUATE(fld1),'Yes','No')
		OTHERWISE
			f1 = ''
	ENDCASE

	RETURN(f1)


	*!*****************************************************************************
	*!
	*!      Procedure: FINDDELETE
	*!
	*!      Called by: NEWREC             (procedure in UTILITY.PRG)
	*!
	*!          Calls: NORMALAL           (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE FindDelete
	PARAMETER xAlias
	xAlias = NormalAL( xAlias)
	PRIVATE fl, result, FILT
	fl = SELECT()
	SELECT (xAlias)
	FILT = FILTER()
	_del = SET("DELETED")
	SET DELETED OFF
	LOCATE FOR DELETED()
	result = FOUND()
	IF NOT EMPTY(FILT)
		SET FILTER TO &FILT
	ENDIF
	IF _del="ON"
		SET DELETED ON
	ENDIF
	SELECT (fl)
	RETURN(result)

	*!*****************************************************************************
	*!
	*!      Procedure: GETIT
	*!
	*!          Calls: NORMALALIAS        (procedure in UTILITY.PRG)
	*!               : SETORDER           (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE GetIt
	PARAMETER xret, xfindexpr, xAlias, xorder
	IF PARAMETERS()<2
		RETURN ''
	ENDIF ( PARAMETERS()<2 )
	xret   = ALLTRIM(xret)
	xAlias = ALLTRIM(NormalAlias(xAlias))
	DO setorder WITH xorder, xAlias
	=SEEK( xfindexpr, xAlias )
	RETURN EVALUATE(xAlias+'.'+xret)


	*!*****************************************************************************
	*!
	*!      Procedure: WORKAREA
	*!
	*!          Calls: NORMALAL           (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE WORKAREA
	PARAMETER ALIAS
	m.ALIAS = NormalAL( m.ALIAS )
	SELECT (m.ALIAS)
	RETURN(.T.)


	*!*****************************************************************************
	*!
	*!      Procedure: NORMALAL
	*!
	*!      Called by: ISEMPTY            (procedure in UTILITY.PRG)
	*!               : SETORDER           (procedure in UTILITY.PRG)
	*!               : FINDDELETE         (procedure in UTILITY.PRG)
	*!               : WORKAREA           (procedure in UTILITY.PRG)
	*!               : GOTO               (procedure in UTILITY.PRG)
	*!               : GOTOEOF            (procedure in UTILITY.PRG)
	*!
	*!          Calls: NORMALALIAS        (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE NormalAL
	PARAMETER inalias
	RETURN NormalAlias( inalias )

	*!*****************************************************************************
	*!
	*!      Procedure: NORMALALIAS
	*!
	*!      Called by: GETIT              (procedure in UTILITY.PRG)
	*!               : GETDEF             (procedure in UTILITY.PRG)
	*!               : COUNTREC           (procedure in UTILITY.PRG)
	*!               : NORMALAL           (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE NormalAlias
	PARAMETER inalias
	DO CASE
		CASE NOT TYPE("InAlias")$"CN"
			RETURN(ALIAS())
		CASE TYPE("InAlias")="N"
			RETURN(IIF(BETWEEN(inalias,1,SELECT(1)), ALIAS(inalias), ALIAS()))
		CASE EMPTY(inalias)
			RETURN(ALIAS())
		OTHERWISE
			RETURN(UPPER(ALLTRIM(inalias)))
	ENDCASE
	RETURN("")



	*!*****************************************************************************
	*!
	*!      Procedure: DELFILE
	*!
	*!          Calls: JUSTSTEM           (procedure in UTILITY.PRG)
	*!               : ADDBS()            (function in UTILITY.PRG)
	*!               : JUSTPATH()         (function in UTILITY.PRG)
	*!               : JUSTEXT            (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE delfile
	PARAMETER DBF
	PRIVATE x, dbf0, d
	m.dbf0 = JUSTSTEM(DBF)
	FOR x=1 TO SELECT(1)
		IF dirdeli+m.dbf0+'.' $ DBF(x)
			m.DBF = DBF(x)
			USE IN (x)
			EXIT
		ENDIF ( dirdeli+m.dbf0+'.' $ DBF(x) )
	NEXT ( x )
	d=(ADDBS(JUSTPATH(m.DBF))+JUSTSTEM(m.DBF)+'.'+IIF(EMPTY(JUSTEXT(m.DBF)), 'DBF', JUSTEXT(m.DBF)))
	IF x<=SELECT(1) OR FILE(d)
		ERASE (d)
		d=(ADDBS(JUSTPATH(m.DBF))+JUSTSTEM(m.DBF)+'.CDX')
		ERASE (d)
		d=(ADDBS(JUSTPATH(m.DBF))+JUSTSTEM(m.DBF)+'.FPT')
		ERASE (d)
	ENDIF ( x<=SELECT(1) )
	RETURN



PROCEDURE DELETETABLE( cName )
	IF FILE(ADDBS(dbfpath)+cName+".dbf")
		CLOSE DATABASE ALL
		OPEN DATABASE ocis EXCLUSIVE
		IF INDBC(cName,'TABLE')
			REMOVE TABLE (cName) DELETE recycle
		ELSE
			ERASE (ADDBS(dbfpath)+cName+".*") recycle
		ENDIF
	ENDIF
	RETURN NOT FILE(ADDBS(dbfpath)+cName+".dbf")

	*!*****************************************************************************
	*!
	*!      Procedure: GOTO
	*!
	*!      Called by: COUNTREC           (procedure in UTILITY.PRG)
	*!               : ENVIR.PRG
	*!
	*!          Calls: NORMALAL           (procedure in UTILITY.PRG)
	*!               : GOTOEOF            (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE GOTO
	PARAMETER rec, inalias
	inalias = NormalAL( inalias )
	IF EMPTY(inalias)
		RETURN 0
	ENDIF ( EMPTY(inalias) )
	DO CASE
		CASE TYPE("REC")="C"
			DO CASE
				CASE ALLTRIM(UPPER(rec))="TOP"
					GOTO TOP IN (inalias)
				CASE ALLTRIM(UPPER(rec))="BOTTOM"
					GOTO BOTTOM IN (inalias)
			ENDCASE
		OTHERWISE
			IF BETWEEN(rec,1,RECCOUNT(inalias))
				GOTO (rec) IN (inalias)
			ELSE
				DO gotoeof WITH inalias
				RETURN 0
			ENDIF
	ENDCASE
	RETURN(RECNO(inalias))

	*!*****************************************************************************
	*!
	*!      Procedure: GOTOEOF
	*!
	*!      Called by: GOTO               (procedure in UTILITY.PRG)
	*!
	*!          Calls: NORMALAL           (procedure in UTILITY.PRG)
	*!
	*!*****************************************************************************
PROCEDURE gotoeof
	PARAMETER ALIAS
	m.ALIAS = NormalAL( m.ALIAS )
	GOTO BOTTOM IN (m.ALIAS)
	IF NOT EOF(m.ALIAS)
		SKIP IN (m.ALIAS)
	ENDIF ( NOT EOF(m.ALIAS) )
	xrec = 0
	RETURN


	*!*****************************************************************************
	*!
	*!      Procedure: GETFIELD
	*!
	*!*****************************************************************************
PROCEDURE getfield
	*- GETFIELD - build a comma delimited string of fields.
	*			- Parameter "Except" is a comma delimited string of fields NOT to be included in resulting list


	PARAMETER EXCEPT

	PRIVATE x,Y,z, result

	EXCEPT = UPPER(IIF(EMPTY(EXCEPT),'',','+EXCEPT+','))
	EXCEPT = CHRTRAN(EXCEPT,' ','')
	result = ''

	FOR x=1 TO FCOUNT()
		IF NOT ','+FIELD(x)+',' $ EXCEPT
			result = result+IIF(EMPTY(result),'',',')+FIELD(x)
		ENDIF ( NOT ','+FIELD(x)+',' $ EXCEPT )
	ENDFOR ( x )

	RETURN( result )


PROCEDURE tCURSORSETPROP( cType, nLevel, cAlias)
	IF TXNLEVEL()=0
		RETURN CURSORSETPROP( cType, nLevel, cAlias)
	ENDIF
	RETURN .T.


PROCEDURE CopyInternalFile( cfile, cdest )
	LOCAL F,s,o
	F=FOPEN( cfile )
	IF F>=0
		s=FSEEK(F,0,2)
		=FSEEK(F,0,0)
		o=FCREATE(ADDBS(cdest)+cfile)
		=FWRITE(o, FREAD( F, s ))
	ENDIF
	=FCLOSE(F)
	=FCLOSE(o)
	RETURN FILE(ADDBS(cdest)+cfile)

PROCEDURE CopyFile2( cpath, cfile )
	IF NOT FILE(ADDBS(cpath)+cfile)
		DO CASE
			CASE FILE(ADDBS(dbfpath)+cfile)
				COPY FILE (ADDBS(dbfpath)+cfile) TO (ADDBS(cpath)+cfile)
			CASE FILE(ADDBS('bmp')+cfile)
				COPY FILE (ADDBS('bmp')+cfile) TO (ADDBS(cpath)+cfile)
		ENDCASE
	ENDIF
	RETURN FILE(ADDBS(cpath)+cfile)

