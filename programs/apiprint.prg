

***********************************************
PROCEDURE prswtch

	PARAMETERS newport

	* program to test printing to more than one printer

	* This program illustrates how to register and use the
	* windows API functions that allow manipulation of Window's
	* WIN.INI file.

	* See the Windows SDK reference manual for complete details
	* on how these functions work.

	PRIVATE onerr
	onerr = ON("ERROR")
	ON ERROR DO e0printer

	* First, find the foxtools library
	IF ATC('FoxTools',SET('Library'))=0
		SET LIBRARY TO SYS(2004)+"foxtools.fll" ADDITIVE
	ENDIF

	#DEFINE HWND_BROADCAST 65535
	#DEFINE WM_WININICHANGE 26

	LOCAL Printers, NUMBER, Defa_print, portfound, buflen, retval, strg, bytes, wrfn, postfn, gtfn, retbuf, dcount, thisdevice
	Printers = ""
	NUMBER = 0
	Defa_print = ""
	portfound = .F.
	buflen = 2048
	retval = 0
	m.strg = ""
	m.bytes = 0
	m.wrfn = ""
	m.postfn = ""
	m.gtfn = ""
	m.retbuf = ""
	m.dcount = 0
	m.thisdevice = ""

	*( GLR 11/3/95 - remove system menu and toolbar for main desktop
	IF TYPE('_isVFP')="U"
		PRIVATE _IsVFP, _FPW26
		_IsVFP = ("VISUAL FOXPRO"$UPPER(VERSION()) AND _WINDOWS)
		_FPW26 = ("FOXPRO 2.6"$UPPER(VERSION()) AND _WINDOWS)
	ENDIF
	IF TYPE("_WinNT")="U"
		PRIVATE _WINNT, _WIN95, _WIN31
		_WINNT	= FILE(GETENV("WINDIR") +"\SYSTEM32\USER32.DLL")
		_WIN95	= FILE(GETENV("WINDIR") +"\SYSTEM\USER32.DLL")
		_WIN31	= FILE(GETENV("WINDIR") +"\SYSTEM\WIN32S\W32COMB.DLL")
	ENDIF

	LOCAL i,j

	DO CASE
		CASE _FPW26
			wrfn = regfn("WRITEPROFILESTRING", "CCC", "I")
			postfn = regfn("SENDMESSAGE", "IIIC", "L")
			gtfn = regfn("GETPROFILESTRING", "CCC@CI", "I")

			** create and load an array "Printers" with current printer info
			** from WIN.INI.  The Printers array is dimensioned as an (n,3)
			** type and collects data from the
			** [Printers] and [Devices] sections of WIN.INI.
			** The array contents look like this after being filled:
			**
			** Col 1                       Col 2                Col 3
			** -------------------         -----------------   --------------------
			** EPSON FX-80                 EPSON9,LPT2:         EPSON9,LPT2:,15,45
			** HP LaserJet 4 Plus/4M Plus  HPPCL5E,LPT1:        HPPCL5E,LPT1:,45,60
			**

			m.retbuf = REPLICATE(CHR(0), buflen)
			m.bytes = callfn(gtfn, "devices",0, CHR(0), @m.retbuf, buflen)

			m.dcount = 0
			m.retbuf = LEFT(m.retbuf, m.bytes)

			DO WHILE CHR(0) $ m.retbuf
				m.thisdevice = LEFT(m.retbuf, AT(CHR(0), m.retbuf)-1)

				IF LEFT(m.thisdevice,1)<>CHR(0)
					m.dcount = m.dcount + 1
					DIMENSION Printers[m.dcount,3]
					Printers[m.dcount,1] = m.thisdevice
				ENDIF
				IF AT(CHR(0), m.retbuf) < LEN(m.retbuf)
					m.retbuf = SUBSTR(m.retbuf,AT(CHR(0),m.retbuf) + 1)
				ELSE
					EXIT
				ENDIF
			ENDDO

			FOR m.j = 1 TO m.dcount
				retbuf = REPLICATE(CHR(0),256)

				m.bytes = callfn(gtfn, "devices", Printers[m.j,1], CHR(0), @m.retbuf, 256)

				m.retbuf = LEFT(m.retbuf, m.bytes)
				Printers[m.j, 2] = m.retbuf

				m.retbuf = REPLICATE(CHR(0), 256)

				m.bytes = callfn(gtfn,"PrinterPorts", Printers[m.j,1], CHR(0), @m.retbuf, 256)

				m.retbuf = LEFT(m.retbuf, m.bytes)
				Printers[m.j, 3] = m.retbuf
			NEXT

			m.number = m.dcount

			retbuf = REPLICATE(CHR(0),256)
			m.bytes = callfn(gtfn, "Windows", "device", CHR(0), @m.retbuf, 256)

			m.retbuf = LEFT(m.retbuf, m.bytes)

			** attempt to set new default printer **
			* search for port name in array *
			FOR m.i = 1 TO m.number
				IF AT(m.newport, Printers[m.i, 2]) <> 0
					m.strg = Printers[m.i, 1] + "," + Printers[m.i,2]
					portfound = .T.
					EXIT
				ENDIF
			NEXT

			IF portfound

				* write new device name in [Windows] section of WIN.INI
				retval = callfn(wrfn, "windows","device",m.strg)

				* call WriteProfileString with nulls to flush WIN.INI from
				* Window's cache

				retval = callfn(wrfn, "", "", "")

				* call SendMessage to inform other Windows apps that the WIN.INI changed
				retval = callfn(postfn, HWND_BROADCAST, WM_WININICHANGE, 0, "Windows")

			ENDIF

		CASE _IsVFP

			*-	postfn = regfn("SENDMESSAGE", "IIIC", "L")

			** create and load an array "Printers" with current printer info
			** from WIN.INI.  The Printers array is dimensioned as an (n,3)
			** type and collects data from the
			** [Printers] and [Devices] sections of WIN.INI.
			** The array contents look like this after being filled:
			**
			** Col 1                       Col 2                Col 3
			** -------------------         -----------------   --------------------
			** EPSON FX-80                 EPSON9,LPT2:         EPSON9,LPT2:,15,45
			** HP LaserJet 4 Plus/4M Plus  HPPCL5E,LPT1:        HPPCL5E,LPT1:,45,60
			**

			m.retbuf = REPLICATE(CHR(0), buflen)
			m.bytes = GetProStrg( "devices",0, CHR(0), @m.retbuf, buflen)

			m.dcount = 0
			m.retbuf = LEFT(m.retbuf, m.bytes)

			DO WHILE CHR(0) $ m.retbuf
				m.thisdevice = LEFT(m.retbuf, AT(CHR(0), m.retbuf)-1)

				IF LEFT(m.thisdevice,1)<>CHR(0)
					m.dcount = m.dcount + 1
					DIMENSION Printers[m.dcount,3]
					Printers[m.dcount,1] = m.thisdevice
				ENDIF
				IF AT(CHR(0), m.retbuf) < LEN(m.retbuf)
					m.retbuf = SUBSTR(m.retbuf,AT(CHR(0),m.retbuf) + 1)
				ELSE
					EXIT
				ENDIF
			ENDDO

			FOR m.j = 1 TO m.dcount
				retbuf = REPLICATE(CHR(0),256)

				m.bytes = GetProStrg( "devices", Printers[m.j,1], CHR(0), @m.retbuf, 256)

				m.retbuf = LEFT(m.retbuf, m.bytes)
				Printers[m.j, 2] = m.retbuf

				m.retbuf = REPLICATE(CHR(0), 256)

				m.bytes = GetProStrg("PrinterPorts", Printers[m.j,1], CHR(0), @m.retbuf, 256)

				m.retbuf = LEFT(m.retbuf, m.bytes)
				Printers[m.j, 3] = m.retbuf
			NEXT

			m.number = m.dcount

			retbuf = REPLICATE(CHR(0),256)
			m.bytes = GetProStrg( "Windows", "device", CHR(0), @m.retbuf, 256)

			m.retbuf = LEFT(m.retbuf, m.bytes)

			** attempt to set new default printer **
			* search for port name in array *
			FOR m.i = 1 TO m.number
				*( ORCA GLR 02/16/1996 16:37:59 PRSWTCH.PRG : multiple fashions
				DO CASE
					CASE INLIST( ALLTRIM(UPPER(m.newport)), 'LPT1:', 'LPT2:', 'LPT3:') AND ATC(m.newport, Printers[m.i, 2]) <> 0
						m.strg = Printers[m.i, 1] + "," + Printers[m.i,2]
						portfound = .T.
						EXIT
					CASE ATC(m.newport, Printers[m.i, 1]) <> 0			&& by Name
						m.strg = Printers[m.i, 1] + "," + Printers[m.i,2]
						portfound = .T.
						EXIT
					CASE ATC(m.newport, Printers[m.i, 2]) <> 0			&& by NE##:
						m.strg = Printers[m.i, 1] + "," + Printers[m.i,2]
						portfound = .T.
						EXIT
				ENDCASE
			NEXT

			IF portfound

				* write new device name in [Windows] section of WIN.INI
				retval = PutProStrg( "windows","device",m.strg)

				* call WriteProfileString with nulls to flush WIN.INI from
				* Window's cache

				retval = PutProStrg( "", "", "")

				* call SendMessage to inform other Windows apps that the WIN.INI changed
				*-		retval = callfn(postfn, HWND_BROADCAST, WM_WININICHANGE, 0, "Windows")

			ENDIF
	ENDCASE
	ON ERROR &onerr

	RETURN(portfound)
ENDPROC


PROCEDURE e0printer
	RETURN
ENDPROC


