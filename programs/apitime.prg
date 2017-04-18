* Date / Time / Age Procedure file

PROCEDURE AgeUnit( lDOB, lAge, lAgeUnit, ltoday, lFormat )

	lDOB = CTOD(TRANSFORM(lDOB,""))
	ltoday = CTOD(TRANSFORM(ltoday,""))
	ltoday = IIF(EMPTY(ltoday),DATE(),ltoday)
	lFormat = ALLTRIM(UPPER(IIF(EMPTY(lFormat), "", lFormat)))

	DO CASE
		CASE EMPTY(lDOB) AND EMPTY(lAge)
			RETURN .F.
			*
			*- no future DOB
		CASE (DATE() < lDOB)
			?? CHR(7)
			*-- BEV 06/28/1996 14:40:30 TXTDOB.VALID : SEAL 1623
			*-- Use message box to inform user and get confirmation.
			*-- wait window nowait "Date of Birth is in the Future."
			=MESSAGEBOX("Date of Birth cannot be in the future",48,THISFORM.CAPTION )

			RETURN .F.

		CASE EMPTY(lDOB)
			lAge = VAL(TRANS(lAge,""))
			lAgeUnit = IIF(EMPTY(lAgeUnit),"",ALLTRIM(lAgeUnit))
		OTHERWISE
			DO CASE
				CASE lDOB > GOMONTH(DATE(),-2)	 				&& (DATE() - lDOB) < 90
					*- within the last 2 months
					lAge = DATE() - lDOB
					lAgeUnit = "Days"

				CASE lDOB > GOMONTH(DATE(),-24)					&& (DATE() - lDOB) <  (3*365)
					*- between 2 months and 2 years
					lAge = FLOOR( (DATE() - lDOB)/(365/12) )
					lAgeUnit = "Months"

				OTHERWISE
					ldYear = YEAR(DATE()) - YEAR(lDOB)
					IF RIGHT(DTOS(lDOB),4) > RIGHT(DTOS(DATE()),4)
						ldYear = ldYear - 1
					ENDIF
					lAge = ldYear
					lAgeUnit = "Years"
			ENDCASE
	ENDCASE

	DO CASE
		CASE EMPTY(lFormat)
		CASE lFormat="SHORT"
			RETURN ALLTRIM(STR(lAge))+" "+LEFT(lAgeUnit,1)
		CASE lFormat="LONG"
			RETURN ALLTRIM(STR(lAge))+" "+lAgeUnit
	ENDCASE
	RETURN .T.
ENDPROC

PROCEDURE age
	PARAMETER ldDOB, ldDateFrom
	* age calculation routine

	PRIVATE mage

	IF EMPTY(ldDateFrom)
		ldDateFrom = DATE()
	ENDIF


	mcentury = SET('century')
	SET CENTURY ON

	IF EMPTY(ldDOB)
		mage = 0
	ELSE
		mage = YEAR(ldDateFrom)-YEAR(ldDOB)
		* test if birthday has passed
		IF DTOS(ldDateFrom) < DTOS(CTOD(LEFT(DTOC(ldDOB),6)+STR(YEAR(ldDateFrom),4)))
			mage = mage-1
		ENDIF
	ENDIF

	IF mcentury = 'OFF'
		SET CENTURY OFF
	ENDIF

	RETURN mage
ENDPROC

PROCEDURE AgeMonth
	PARAMETER TODAY,DOB


	IF YEAR(DOB)<YEAR(TODAY)
		MTODAY = MONTH(TODAY)+12
	ELSE
		MTODAY = MONTH(TODAY)
	ENDIF

	RETURN MTODAY-MONTH(DOB)
ENDPROC

PROCEDURE gElapsed
	PARAMETERS ldatein, ltimein, ldateout, ltimeout

	*-- Compute the elapsed time.

	* get hours between date in and date out
	timemin = (ldateout - ldatein) * 24 * 60

	thenmin = FLOOR(ltimein / 100) * 60 + (ltimein % 100)
	nowmin  = VAL(LEFT(PADL(ALLT(STR(ltimeout)),4,"0"),2))*60 + VAL(RIGHT(PADL(ALLT(STR(ltimeout)),4,"0"),2))

	IF (nowmin > thenmin)
		timemin = timemin + nowmin - thenmin
	ELSE
		timemin = timemin - (thenmin - nowmin)
	ENDIF

	* lelapsed = PADL(ALLTRIM(STR(FLOOR(timemin/60))),2,' ') + ':' + PADL(ALLTRIM(STR(timemin % 60)),2,'0')

	RETURN ROUND(timemin,0)
ENDPROC


PROCEDURE CompTime
	*( ORCA GLR 03/14/1996 12:58:48 COMPTIME.PRG :

	PARAMETERS date1, time1, date2, time2

	*- computes the difference in date times, and returns the difference in miinutes.
	*
	*- time is expected to be in milatry.
	*
	*- RETURN Character

	*---------------
	* format examples
	*
	* as four parameters
	*- comptime( {01/11/1995}, "05:23", {01/12/1995}, "20:23")
	*- comptime( "01/11/95", "05:23", "01/12/95", "20:23")
	*- comptime( "01/11/95", "0523", "01/12/95", "2023")
	*- comptime( {01/12/95}, 523, {01/12/95}, 533)
	*- comptime( "01/12/95", 523, "01/12/95", 533)
	* as two parameters
	*- comptime( "01/11/95 05:23", "01/12/95 20:23")
	*- comptime( "01/11/95 0523", "01/12/95 2023")
	*
	*------------------
	*** RAS 03/31/1999 Correct date format problem.
	#DEFINE kBaseDate	{^1800-01-01}
	*!*		#DEFINE kBaseDate	{01/01/1800}

	PRIVATE lDate0, lTime0, lminutes

	IF PARAMETERS()=2
		date2 = time1
		time1 = ALLTRIM(SUBSTR(date1,AT(' ',date1)))
		date1 = ALLTRIM(SUBSTR(date1,1,AT(' ',date1)))
		time2 = ALLTRIM(SUBSTR(date2,AT(' ',date2)))
		date2 = ALLTRIM(SUBSTR(date2,1,AT(' ',date2)))
	ENDIF
	*
	IF EMPTY(date1) OR EMPTY(date2)
		RETURN [ ]
	ENDIF
	*
	time1 = IIF(TYPE("time1")="N", PADL(ALLTRIM(STR(time1)),4,'0'), time1)
	time2 = IIF(TYPE("time2")="N", PADL(ALLTRIM(STR(time2)),4,'0'), time2)
	*
	time1 = IIF( NOT ':'$time1, LEFT(time1,2)+':'+RIGHT(time1,2), time1)
	time2 = IIF( NOT ':'$time2, LEFT(time2,2)+':'+RIGHT(time2,2), time2)

	lDate0 = IIF(TYPE('date1')='D', date1, CTOD(date1))
	lDate0 = lDate0 - kBaseDate
	lDate0 = lDate0 * 24 * 60
	*
	lTime0 = (VAL(LEFT(time1,2))*60)+VAL(RIGHT(time1,2))
	*
	lminutes = lDate0+lTime0
	*
	lDate0 = IIF(TYPE('date2')='D', date2, CTOD(date2))
	lDate0 = lDate0 - kBaseDate
	lDate0 = lDate0 * 24 * 60
	*
	lTime0 = (VAL(LEFT(time2,2))*60)+VAL(RIGHT(time2,2))
	*
	lminutes = (lDate0+lTime0) - lminutes
	*
	RETURN ALLTRIM(STR(lminutes))
ENDPROC


PROCEDURE Time24
	*-- This program takes a time in a numeric format (up to 4 digits) and outputs
	*-- a string indicating which hour it falls into ("00:00-00:59", "01:00-01:59", ...)

	PARAMETER theTime

	leadingTwo = LEFT(PADL(ALLTRIM(STR(theTime)),4,"0"),2)
	RETURN leadingTwo + ":00-" + leadingTwo + ":59"
ENDPROC

PROCEDURE ValTime
	*( ORCA GLR 02/08/1996 13:19:04 VALTIME.PRG :
	*
	* Validates Time Strings
	*

	PARAMETER lTime, ldate

	*( ORCA GLR 03/27/1996 10:17:42 VALTIME.PRG : correct format

	*- standardize incoming time.
	IF TYPE('ltime')='N'
		lTime = LTRIM(STR(lTime))
		lTime = PADL(lTime,4,'0')
	ENDIF
	hasSeconds = LEN(lTime)>5
	ldate = IIF( EMPTY(ldate), DATE(), ldate)
	IF ':'$lTime
		lTime = STRTRAN(PADR(lTime,5+IIF(hasSeconds,3,0),'0'), ':','')
	ELSE
		lTime = PADL(lTime,4+IIF(hasSeconds,2,0),'0')
	ENDIF
	*
	IF EMPTY(LEFT(m.lTime,2)) AND EMPTY(RIGHT(m.lTime,2))
		RETURN .T.
	ENDIF

	*
	*- insert colon
	lTime = PADL( LTRIM(STR(VAL(LEFT(m.lTime,2)))), 2, '0') ;
		+':'+ PADL( LTRIM(STR(VAL(SUBSTR(lTime,3,2)))), 2, '0') ;
		+ IIF(hasSeconds, ':'+PADL( LTRIM(STR(VAL(SUBSTR(lTime,5,2)))), 2, '0'), '')

	*- varify if proper time
	DO CASE
		CASE hasSeconds
			DO CASE
				CASE (NOT BETWEEN(VAL(SUBSTR(m.lTime,1,2)), 0, 23)) ;
						OR (NOT BETWEEN(VAL(SUBSTR(m.lTime,4,2)), 0, 59)) ;
						OR (NOT BETWEEN(VAL(SUBSTR(m.lTime,7,2)), 0, 59))
					=MESSAGEBOX(  "Invalid Time! - 00:00:00 -> 23:59:59" )
					*-         _CUROBJ = _CUROBJ
					RETURN 0

					*- No Future time if today
				CASE VAL(SUBSTR(m.lTime,1,2)+SUBSTR(lTime,4,2)+SUBSTR(lTime,7,2)) > VAL(SUBSTR(TIME(),1,2)+SUBSTR(TIME(),4,2)+SUBSTR(TIME(),7,2)) AND ldate >= DATE()
					=MESSAGEBOX( "Time entered ("+ALLT(lTime)+") is greater than current time ("+TIME()+")" )
					RETURN 0
			ENDCASE

		CASE (NOT BETWEEN(VAL(LEFT(m.lTime,2)),  0, 23)) ;
				OR (NOT BETWEEN(VAL(RIGHT(m.lTime,2)), 0, 59))
			=MESSAGEBOX(  "Invalid Time! - 00:00 -> 23:59" )
			_CUROBJ = _CUROBJ
			RETURN 0

			*- No Future time if today
		CASE VAL(LEFT(m.lTime,2)+SUBSTR(lTime,4,2)) > VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2)) AND ldate >= DATE()
			=MESSAGEBOX( "Time entered ("+ALLT(lTime)+") is greater than current time ("+LEFT(TIME(),5)+")" )
			RETURN 0

	ENDCASE

	*!*	if not _IsVFP
	*!*		if not empty(varread())
	*!*			store ltime to (varread())
	*!*			show gets
	*!*		endif
	*!*	endif
	RETURN .T.
ENDPROC

*-- Three functions for dealing with timestamps:
*--   GetTS()      Returns the current 14 character timestamp.
*--   TS2Date(ts)  Returns the date for a given timestamp (MM/DD/YYYY).
*--   TS2Time(ts)  Returns the time for a given timestamp (HH:MM:SS)

PROCEDURE ts2dt
	LPARAMETERS m.seen
	*- convert TimeStamp as character to a datatime data type
	RETURN CTOT( SUBSTR(m.seen, 5,2)+'/'+SUBSTR(m.seen, 7,2)+'/'+SUBSTR(m.seen, 1,4)+' '+SUBSTR(m.seen, 9,2)+':'+SUBSTR(m.seen,11,2)+':'+SUBSTR(m.seen,13,2) )

	*!*	PROCEDURE dt2ts
	*!*	    LPARAMETERS m.seen
	*!*	    *- convert statetime to Timestamp format
	*!*	    RETURN DTOS(m.seen)+SUBSTR(TTOC(m.seen),12,2)+SUBSTR(TTOC(m.seen),15,2)
ENDPROC

PROCEDURE getts
	RETURN DTOS(DATE()) + SUBSTR(TIME(),1,2) + SUBSTR(TIME(),4,2) + SUBSTR(TIME(),7,2)
ENDPROC

PROCEDURE ts2date
	PARAMETER ts
	RETURN SUBSTR(ts,5,2) + '/' + SUBSTR(ts,7,2) + '/' + SUBSTR(ts,1,4)
ENDPROC

PROCEDURE ts2time
	PARAMETER ts
	*-- 05/22/1996 RHW - should never display seconds
	*	RETURN SUBSTR(ts,9,2) + ':' + SUBSTR(ts,11,2) + ':' + SUBSTR(ts,13,2)
	RETURN SUBSTR(ts,9,2) + ':' + SUBSTR(ts,11,2)
ENDPROC

* --------------- *

PROCEDURE formtime
	PARAMETER theTime

	PRIVATE lcampm
	* formats a time into a am/pm string
	lcampm = " am"

	IF theTime > 1199
		lcampm = " pm"
	ENDIF
	IF theTime > 1299
		theTime = theTime - 1200
	ENDIF

	RETURN STUFF(PADL(ALLTRIM(STR(theTime)),4,'0'),3,0,':')+lcampm
ENDPROC


PROCEDURE AmPm
	PARAMETER m.time
	*- converts miltary to standard
	DO CASE
		CASE EMPTY(m.time)
		CASE EMPTY(LEFT(m.time,2))
		CASE VAL(LEFT(m.time,2))>12
			m.time = PADL(LTRIM(STR(VAL(LEFT(m.time,2))-12)),2,'0')+SUBSTR(m.time,3)+'p'
		CASE VAL(LEFT(m.time,2))=0
			m.time = '12'+SUBSTR(m.time,3)+'a'
		CASE VAL(LEFT(m.time,2))=12
			m.time = '12'+SUBSTR(m.time,3)+'p'
		OTHERWISE
			m.time = m.time+'a'
	ENDCASE
	RETURN m.time
ENDPROC



*==========================================================================
*-- reformat a YYYYMMDDHHMMSS into HHMM DD/MM/YY
*-- used to be HH:MM DD/MM/YY
PROCEDURE fmtTS()
	LPARAMETER tsd
	RETURN	SUBSTR(tsd,9,2);
		+ SUBSTR(tsd,11,2)+" ";
		+ SUBSTR(tsd,5,2)+"/";
		+ SUBSTR(tsd,7,2)+"/";
		+ SUBSTR(tsd,3,2)
ENDPROC

*==========================================================================
PROCEDURE mttost
	RETURN IIF( VAL(LEFT(TIME(),2))=12 ,LEFT(TIME(),2) + ":" + SUBSTR(TIME(),4,2)+ " PM",IIF(VAL(LEFT(TIME(),2))>=12,PADL(ALLT(STR(INT(VAL(LEFT(TIME(),2))-12))),2,"0"),LEFT(TIME(),2)) + ":" + SUBSTR(TIME(),4,2)+ IIF(VAL(LEFT(TIME(),2))>=12," PM"," AM"))
ENDPROC


PROCEDURE ConvDate
	PARAMETER cDate
	RETURN CTOD(STUFF(RIGHT(ALLTRIM(m.cDate),4),3,0,'/')+'/'+LEFT(ALLTRIM(m.cDate),4))
ENDPROC


*!*****************************************************************************
*!
*!      Procedure: APPDATE
*!
*!      Called by: LOGO               (procedure in UTILITY.PRG)
*!
*!*****************************************************************************
PROCEDURE appdate
	PARAMETER APP
	APP = IIF(EMPTY(m.APP), SYS(16,1), m.APP)
	DIMENSION aapp[1,5]
	aapp[1,3]={}
	=ADIR(aapp,m.APP)
	RETURN aapp[1,3]
ENDPROC

*!*****************************************************************************
*!
*!      Procedure: APPTIME
*!
*!*****************************************************************************
PROCEDURE apptime
	APP = IIF(TYPE("m.APP")#"C", SYS(16,1), m.APP)
	DIMENSION aapp[1,5]
	aapp[1,2] = ""
	=ADIR(aapp,m.APP)
	RETURN aapp[1,4]
ENDPROC



*!*****************************************************************************
*!
*!      Procedure: VALIDTIME
*!
*!          Calls: MESG               (procedure in UTILITY.PRG)
*!
*!*****************************************************************************
PROCEDURE ValidTime
	PARAMETER timevar
	PRIVATE TIME, hr, mn
	TIME = EVALUATE(timevar)
	DO CASE
		CASE '.'$TIME
			DO CASE
				CASE VAL(TIME)>0
					hr = VAL(SUBSTR(TIME, 1, AT('.',TIME)))
					mn = VAL(SUBSTR(TIME, AT('.',TIME)))
					STORE STR(hr,3)+':'+PADL(LTRIM(STR(60*mn,2)),2,'0') TO (timevar)
					SHOW GET (timevar)
				OTHERWISE
					WAIT WINDOW NOWAIT "Invalid Time!"
					RETURN .F.
			ENDCASE
		CASE ":"$TIME
			DO CASE
				CASE BETWEEN(VAL(SUBSTR(TIME, AT(':',TIME)-1)), 0, 59)
				CASE BETWEEN(VAL(SUBSTR(TIME, 1, AT(':',TIME)+1)), 0, 9999)
				OTHERWISE
					WAIT WINDOW NOWAIT "Invalid Time!"
					RETURN .F.
			ENDCASE
			hr = VAL(SUBSTR(TIME, 1, AT(':',TIME)-1))
			mn = VAL(SUBSTR(TIME, AT(':',TIME)+1))
			STORE STR(hr,3)+':'+PADL(LTRIM(STR(mn,2)),2,'0') TO (timevar)
			SHOW GET (timevar)
	ENDCASE
	RETURN .T.
ENDPROC



*!*****************************************************************************
*!
*!      Procedure: SEC2TIME
*!
*!*****************************************************************************
PROCEDURE sec2time
	PARAMETER mn0
	PRIVATE hr,mn,sc
	hr = FLOOR(mn0/3600)
	mn0 = mn0 - hr*3600
	mn = FLOOR(mn0/60)
	mn0 = mn0 - mn*60
	sc = mn0
	RETURN STR(hr,3)+':'+PADL(LTRIM(STR(mn,2)),2,'0')+':'+PADL(LTRIM(STR(sc,2)),2,'0')
ENDPROC

*!*****************************************************************************
*!
*!      Procedure: MN2TIME
*!
*!*****************************************************************************
PROCEDURE mn2time
	PARAMETER mn0
	PRIVATE hr,mn
	hr = FLOOR(mn0/60)
	mn = mn0 - (hr*60)
	RETURN STR(hr,3)+':'+PADL(LTRIM(STR(mn,2)),2,'0')
ENDPROC

*!*****************************************************************************
*!
*!      Procedure: TIME2MN
*!
*!*****************************************************************************
PROCEDURE time2mn
	PARAMETER time0
	PRIVATE hr,mn
	hr = VAL(SUBSTR(time0, 1, AT(':',time0)-1))
	mn = VAL(SUBSTR(time0, AT(':',time0)+1))
	RETURN (hr*60)+mn
ENDPROC

PROCEDURE stod
	PARAMETER cDate
	*!*	    RETURN CTOD(STUFF(RIGHT(ALLTRIM(m.cDate),4),3,0,'/')+'/'+LEFT(ALLTRIM(m.cDate),4))
	RETURN ConvDate( cDate )
ENDPROC


PROCEDURE getsts

	LOCAL st
	st=SPACE(16)
	DECLARE GetSystemTime IN "kernel32" STRING @ lpSystemTime

	GetSystemTime( @st )

	RETURN SystemTime( st )
ENDPROC

PROCEDURE SystemTime( st )
	LOCAL st, Y,mn,dw,d,h,m,s,ms,dt,tm,dtm
	Y=(ASC( SUBSTR( st, 2, 1) )*256) + ASC( SUBSTR( st, 1, 1) )
	mn=ASC( SUBSTR( st, 3, 1) )
	dw=ASC( SUBSTR( st, 5, 1) )
	d=ASC( SUBSTR( st, 7, 1) )
	h=ASC( SUBSTR( st, 9, 1) )
	m=ASC( SUBSTR( st, 11, 1) )
	s=ASC( SUBSTR( st, 13, 1) )
	ms=(ASC( SUBSTR( st, 16, 1) )*256) + ASC( SUBSTR( st, 15, 1) )

	dt = CTOD( LTRIM(STR(mn))+'/'+LTRIM(STR(d))+'/'+LTRIM(STR(Y)) )
	tm = LTRIM(STR(h))+':'+LTRIM(STR(m))+':'+LTRIM(STR(s))			&& +'.'+ltrim(str(ms))

	dtm = CTOT( DTOC(dt)+' '+tm)
	RETURN dtm
ENDPROC





PROCEDURE dt2ts
	PARAMETER dt
	RETURN DTOS( dt )+PADL(LTRIM(STR(HOUR(dt))),2,'0')+PADL(LTRIM(STR(MINUTE(dt))),2,'0')+PADL(LTRIM(STR(SEC(dt))),2,'0')

	*-= converts DT (timestamp format) to seconds
PROCEDURE dt2sec
	PARAMETERS dt

	dt = PADL(dt,16,'0')
	*
	acc = 0
	*** RAS 03/31/1999 Correct date format problem.
	acc = acc + (( stod( SUBSTR( dt, 1, 8 ) ) - {^1900-01-01} ) * (24*60*60))
	*!*	    acc = acc + (( stod( SUBSTR( dt, 1, 8 ) ) - {01/01/1900} ) * (24*60*60))
	acc = acc + (VAL(SUBSTR(dt,9,2)) * (60*60))
	acc = acc + (VAL(SUBSTR(dt,11,2)) * 60)
	acc = acc + VAL(SUBSTR(dt,13,2))
	*
	RETURN acc

PROCEDURE cNow
	LOCAL T
	SET HOUR TO 24
	SET CENTURY ON
	T = TTOC(DATETIME())
	T = STRTRAN( LEFT(T, 10+1+5 ), ':','')
	RETURN T

PROCEDURE ts2mil
	LPARAMETER dt
	IF EMPTY(dt)
		RETURN ""
	ENDIF
	DO CASE
		CASE TYPE("dt")="D"
			RETURN DTOC(dt)
		CASE TYPE("dt")="T"
			RETURN STRTRAN(LEFT(TTOC(dt),RAT(':',TTOC(dt))-1),':','')
	ENDCASE
	RETURN ""
