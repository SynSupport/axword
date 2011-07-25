AXWord:		        ActiveX control and support code for creating
			an example Dunning letter.

Author:			David Evans
Company:		Synergex
E-Mail:			david.evans@synergex.com
Submision Name:		AXWord
Date Posted:		12/07/98
Minimum Synergy:	6
Platforms:		Windows

Description:            This is an example of using Synergy to access
			Microsoft Word (OLE) through and ActiveX control
			written in Visual Basic.
                        
Filenames:
	ReadMe.txt	AXword info: This file
	AXWord.bat	Batch file to compile and link
	AXWord.dbl	Synergy Language source file (mainline)
	AXWord.def	Synergy Language header file
	AXWord.ocx	ActiveX control
	AXWord.vbp	Visual Basic project file
	AXWord.vbw	Visual Basic vbw file
	AXWord.wsc	Synergy UI Toolkit script file
	DumbDB.dbl	Synergy Language source file (database stubs)
	AXWordCTL.ctl	Visual Basic control source code
	Dunning.dot	Dunning letter with tags (Word for Windows (r) template)
	MSVBV50.dll	Visual Basic Runtime

Registered CTRL name:	AXWordExample.AXWordCTL

AXWord CTRL Methods:
	Boolean Generate()

AXWord CTRL Properties:
	String m_account
	String m_invs
	String m_pmts
	String m_pdue
	String m_cdue
	String m_tdue
	String m_tdate
	String m_status
	String m_ad
	String m_rep
	String m_DOTFile

Discussion:

  To compile the Synergy Language code:
  Use the AXWord.bat file located in this directory.
  
  To run AXWord:
  set AXWORD environment variable to point to directory where AXWord
  is installed, so the control can find the ".dot" file.
  	Example "set AXWORD=C:\AXWORD"
  The AXWord.ocx must be registered using axutl or regsvr32.
  Run "dbr AXWord".
  
  The example will allow the user to select 1 of 2 accounts and move the
  data for these accounts into the ActiveX control.  WinWord will then replace
  the tags in the DOT file and display the modified document for user changes.
  The user then may FAX, Print, or Save the modified document.
  
  The control was created with Visual Basic 5.  To rebuild control, open the
  project file AXWord.vbp with Visual Basic 5 and select File:Make AXWord.ocx.
  
  Because WinWord uses OLE automation, Synergy cannot directly access 
  WinWord's methods and properties.  A Visual Basic wrapper has been put 
  around the Word for Windows OLE control to access the methods needed to 
  get the job done.
