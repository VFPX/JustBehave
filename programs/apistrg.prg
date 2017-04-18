*- String functions


*!*****************************************************************************
*!
*!      Procedure: REDUCE
*!
*!*****************************************************************************

PROCEDURE Reduce
	*- reduces multiple spaces to a single space
	LPARAMETER cStr

	DO WHILE '  '$cStr
		cStr = STRTRAN(cStr, '  ', ' ')
	ENDDO

	RETURN cStr
ENDPROC

*
* WORDNUM - Returns w_num-th word from string strg
*
*!*****************************************************************************
*!
*!       procedure: WORDNUM
*!
*!*****************************************************************************
PROCEDURE wordnum
	PARAMETERS m.strg,m.w_num, m.w_deli
	PRIVATE strg,s1,w_num,ret_str

	m.s1 = ALLTRIM(m.strg)

	* Replace tabs with spaces
	m.s1 = CHRTRAN(m.s1,CHR(9)," ")

	* Reduce multiple spaces to a single space
	*DO WHILE AT('  ',m.s1) > 0
	*   m.s1 = STRTRAN(m.s1,'  ',' ')
	*ENDDO

	m.w_deli = IIF(TYPE("w_deli")#"C", ' ', m.w_deli)
	ret_str = ""
	DO CASE
		CASE m.w_num > 1
			DO CASE
				CASE AT(m.w_deli,m.s1,m.w_num-1) = 0       && No word w_num.  Past end of string.
					m.ret_str = ""
				CASE AT(m.w_deli,m.s1,m.w_num) = 0         && Word w_num is last word in string.
					m.ret_str = SUBSTR(m.s1,AT(m.w_deli,m.s1,m.w_num-1)+1)
				OTHERWISE                                  && Word w_num is in the middle.
					m.strt_pos = AT(m.w_deli,m.s1,m.w_num-1)
					m.ret_str  = SUBSTR(m.s1,strt_pos,AT(m.w_deli,m.s1,m.w_num)+1 - strt_pos)
			ENDCASE
		CASE m.w_num = 1
			IF AT(m.w_deli,m.s1) > 0                   && Get first word.
				m.ret_str = SUBSTR(m.s1,1,AT(m.w_deli,m.s1)-1)
			ELSE                                       && There is only one word.  Get it.
				m.ret_str = m.s1
			ENDIF ( AT(m.w_deli,m.s1) > 0 )
	ENDCASE
	m.ret_str = STRTRAN(m.ret_str, m.w_deli, '')
	RETURN ALLTRIM(m.ret_str)
ENDPROC


*!*****************************************************************************
*!
*!      Procedure: WORDS
*!
*!*****************************************************************************
PROCEDURE words
	PARAMETER STR, m.w_deli
	m.w_deli = IIF(EMPTY(m.w_deli), ' ', m.w_deli)
	RETURN OCCUR(m.w_deli,ALLTRIM(Reduce(STR)))+1
ENDPROC




*!*****************************************************************************
*!
*!      Procedure: StrFilter
*!
*!*****************************************************************************

PROCEDURE StrFilter
	LPARAMETER cStr, cChars
	LOCAL lstr
	lstr = CHRTRAN(cStr, cChars, '' )
	lstr = CHRTRAN(cStr, lstr, '')
	RETURN lstr
ENDPROC



****************************************************************************

*!*****************************************************************************
*!
*!      Procedure: FromList
*		Returns the value from a list of values
*!
*!*****************************************************************************
PROCEDURE fromList( nvalue, p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12 )
	IF BETWEEN( nvalue, 1,PARAMETERS()-1)
		RETURN EVALUATE('m.p'+ALLTRIM(STR(nvalue)))
	ELSE
		RETURN ""
	ENDIF
ENDPROC

*!*****************************************************************************
*!
*!      Procedure: WordIndex
*		return the position of a word in a delimited string.
*!
*!*****************************************************************************
PROCEDURE WordIndex( cWordList, cWord, cDelim )

	LOCAL X, T, Y
	Y=0
	T = STRTRAN( cWordList, cDelim, CHR(13)+CHR(10))
	_MLINE = 0
	FOR X=1 TO MEMLINES( T )
		IF UPPER(ALLTRIM(MLINE(T,1,_MLINE)))=UPPER(ALLTRIM(cWord))
			Y=X
			EXIT
		ENDIF
	NEXT
	RETURN Y
ENDPROC

*!*****************************************************************************
*!
*!      Procedure: REVERSE
*!
*!*****************************************************************************
PROCEDURE reverse
	PARAMETER TEXT

	PRIVATE X, T
	T=''
	FOR X=1 TO LEN(TEXT)
		T=T+CHR( 255-ASC(SUBSTR(TEXT,X,1)) )
	NEXT ( X )
	RETURN T
ENDPROC




PROCEDURE STRTOFILE( cTxt, cFile, lAppend )

	LOCAL F, lsize
	IF FILE( cFile )
		F = FOPEN( cFile, 1)
	ELSE
		F = FCREATE( cFile, 0 )
	ENDIF
	IF F>=0
		IF lAppend
			=FSEEK(F,0,2)
			lsize = FWRITE(F, cTxt)
			IF lsize >=0
				=FCLOSE(F)
				RETURN .T.
			ENDIF
		ELSE
			IF FCHSIZE(F, 0)=0
				lsize = FWRITE(F, cTxt)
				IF lsize >=0
					=FCLOSE(F)
					RETURN .T.
				ENDIF
			ENDIF
		ENDIF
	ENDIF
	=FCLOSE(F)
	RETURN 0
ENDPROC


PROCEDURE FILETOSTR( cFile )

	LOCAL F, lsize, lstr
	lstr = ""
	IF FILE( cFile )
		F = FOPEN( cFile, 0)
		IF F>=0
			lsize = FSEEK(F,0,2)
			=FSEEK(F,0,0)
			lstr  = FREAD(F,lsize)
			=FCLOSE(F)
		ENDIF
	ENDIF
	RETURN lstr
ENDPROC

*****************************************************************************

*!*****************************************************************************
*!
*!      Procedure: M0LINE
*!
*!      Called by: MHS_RECV.PRG
*!               : LOGO               (procedure in UTILITY.PRG)
*!
*!*****************************************************************************
PROCEDURE m0line
	PARAMETER txt, pos, BEGIN

	DO CASE
		CASE PARAMETERS()<2
			*- simulate error for error trap
			=MLINE("X",.F.)
			RETURN ""
		CASE pos=0
			RETURN ""
	ENDCASE
	pos	    = MAX(IIF(EMPTY(pos),0,pos),0)
	BEGIN   = MAX(IIF(EMPTY(BEGIN),1,BEGIN),1)
	DO CASE
		CASE EMPTY(txt)
			RETURN ""
		CASE pos>MEMLINE(txt)
			RETURN ""
		CASE BEGIN>LEN(txt)
			RETURN ""
	ENDCASE

	PRIVATE X, Y,  T
	T = txt+REPLICATE(crlf,100)
	X = IIF(pos<=1, BEGIN, AT(crlf, SUBSTR(T ,BEGIN), pos-1)+2)
	Y = AT(crlf, SUBSTR(T ,X), 1)-1
	_MLINE = X+Y+2
	BEGIN = X+Y+2

	RETURN SUBSTR( T , X, Y )
ENDPROC

*!*****************************************************************************
*!
*!      Procedure: AT0LINE
*!
*!      Called by: MHS_RECV.PRG
*!
*!*****************************************************************************
PROCEDURE at0line
	PARAMETER s, txt, noccurance
	noccurance = MAX(1, VAL(TRANS(noccurance,"")) )
	PRIVATE N
	N = OCCURS(crlf, SUBSTR(txt, 1, ATC(s,txt,noccurance)))
	RETURN IIF(ATC(s,txt)=0,0,N+1)
ENDPROC

*!*****************************************************************************
*!
*!      Procedure: PADF
*!
*!*****************************************************************************
PROCEDURE Padf
	PARAMETER VAR, fld
	RETURN PADR(m.VAR, LEN(m.fld))
ENDPROC



* returns the starting position for a centered string
*!*****************************************************************************
*!
*!       procedure: CENTER
*!
*!*****************************************************************************
PROCEDURE CENTER
	PARAMETERS cStr,nRepLen
	RETURN ROUND((nRepLen - LEN(cStr))/2,0)
ENDPROC




PROCEDURE RemoveTrailCrlf( cText )
	cText = ALLTRIM(cText)
	DO WHILE ASC(RIGHT(cText,1))<32 AND LEN(cText)>0
		cText = ALLTRIM(LEFT(cText, LEN(cText)-1))
	ENDDO
	RETURN cText
ENDPROC


PROCEDURE strtran_c
	PARAMETERS expc1,expc2,expc3,expn1,expn2
	PRIVATE EXPR,at_pos,at_pos2,i,j

	IF EMPTY(m.expc1).OR.EMPTY(m.expc2)
		RETURN m.expc1
	ENDIF
	m.EXPR=m.expc1
	IF TYPE('m.expn1')#'N'
		m.expn1=1
	ENDIF
	IF TYPE('m.expn2')#'N'
		m.expn2=LEN(m.expc1)
	ENDIF
	IF m.expn1<1.OR.m.expn2<1
		RETURN m.expc1
	ENDIF
	m.i=0
	m.j=0
	m.at_pos2=1
	DO WHILE .T.
		m.at_pos=ATC(m.expc2,SUBSTR(m.EXPR,m.at_pos2))
		IF m.at_pos=0
			EXIT
		ENDIF
		m.i=m.i+1
		IF m.i<m.expn1
			m.at_pos2=m.at_pos+m.at_pos2+LEN(m.expc2)-1
			LOOP
		ENDIF
		m.EXPR=LEFT(m.EXPR,m.at_pos+m.at_pos2-2)+m.expc3+;
			SUBSTR(m.EXPR,m.at_pos+m.at_pos2+LEN(m.expc2)-1)
		m.j=m.j+1
		IF m.j>=m.expn2
			EXIT
		ENDIF
		m.at_pos2=m.at_pos+m.at_pos2+LEN(m.expc3)-1
		IF m.at_pos2>LEN(m.EXPR)
			EXIT
		ENDIF
	ENDDO
	RETURN m.EXPR
ENDPROC


PROCEDURE GetINIstr( cINI, cSection, cEntry, noccurance )
	noccurance = MAX(VAL(TRANSFORM(noccurance,"")),1)
	LOCAL ini, X, mw, cresult
	cresult = ""
	mw=SET("MEMOWIDTH")
	SET MEMOWIDTH TO 10000
	*- read file
	ini = FILETOSTR( cINI )
	* isolate the section
	X = ATC( '['+ALLTRIM(cSection)+']', m.ini )
	IF X>0
		ini = SUBSTR( m.ini, X+LEN('['+ALLTRIM(cSection)+']')+1 )
		X = ATC('[', m.ini)
		IF X>0
			m.ini = LEFT(m.ini, X-1)
		ENDIF
		*- look for entry
		X = at0line( ALLTRIM(cEntry), m.ini, noccurance )
		IF X>0
			m.ini = m0line( m.ini, X )
			IF UPPER(ALLTRIM(wordnum( m.ini, 1, '=')))= UPPER(cEntry)
				m.cresult = wordnum( m.ini, 2, '=')
			ENDIF
		ENDIF
	ENDIF
	SET MEMOWIDTH TO (mw)
	RETURN m.cresult

ENDPROC

PROCEDURE RemoveWord( tcString, tcWord )
	LOCAL lcNewString, i
	lcNewString = []

	FOR i=1 TO ALINES(aWords, tcString, .T., ' ')
		IF NOT LOWER(aWords[i])==LOWER(tcWord)
			lcNewString = LTRIM(lcNewString + [ ] + aWords[i])
		ENDIF
	NEXT

	RETURN lcNewString
ENDPROC
