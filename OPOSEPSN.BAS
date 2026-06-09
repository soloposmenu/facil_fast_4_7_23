Attribute VB_Name = "OPOSEPSN"

' /////////////////////////////////////////////////////////////////////
' //
' // EPSON.H
' //
' //   General header file for EPSON Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' // 96-01-05 OPOS Release 1.0                                     Kata
' // 96-04-25 OPOS Release 1.1; OPOS-J 1.0
' // 96-07-08 DirectIO command for device specific details setting
' //
' /////////////////////////////////////////////////////////////////////

' // * default numeric property on error *

Public Const PROP_ERRNUMBER As Long = 0&  

' // * default string property on error *

Public Const PROP_ERRSTRING As String = "[Error]" 

' /////////////////////////////////////////////////////////////////////
' // OPOS "ResultCodeExtended" Property Base Constants
' /////////////////////////////////////////////////////////////////////

Public Const EPSNERREXT As Long = 10000&    ' // EPSON specific error base
Public Const CASHERREXT As Long = 20000&    ' // Cash Drawer specific error base
Public Const COINERREXT As Long = 30000&    ' // Coin Dispenser specific error base
Public Const TOTERREXT As Long = 40000&     ' // Hard Totals specific error base
Public Const LOCKERREXT As Long = 50000&    ' // Keylock specific error base
Public Const DISPERREXT As Long = 60000&    ' // Line Display specific error base
Public Const MICRERREXT As Long = 70000&    ' // MICR specific error base
Public Const MSRERREXT As Long = 80000&     ' // MSR specific error base
Public Const KBDERREXT As Long = 90000&     ' // POS Keyboard specific error base
Public Const PTRERREXT As Long = 100000&    ' // POS Printer specific error base
Public Const SCALERREXT As Long = 110000&   ' // Scale specific error base
Public Const SCANERREXT As Long = 120000&   ' // Scanner specific error base
Public Const SIGERREXT As Long = 130000&    ' // Signature Capture specific error base
Public Const PORTERREXT As Long = 200000&   ' // CPortDrv error base
Public Const PARAMERREXT As Long = 300000&      ' // Parameter error base

' /////////////////////////////////////////////////////////////////////
' // "ResultCodeExtended" Property Constants (EPSON specific)
' /////////////////////////////////////////////////////////////////////

Public Const OPOS_EX_BADCO As Long = 1& + EPSNERREXT        ' // invalid CO I/F
Public Const OPOS_EX_BADPORT As Long = 2& + EPSNERREXT      ' // invalid Port
Public Const OPOS_EX_BADDEVICE As Long = 3& + EPSNERREXT    ' // invalid DeviceName
Public Const OPOS_EX_BADPROPIDX As Long = 4& + EPSNERREXT   ' // invalid property index
Public Const OPOS_EX_BADPROPVAL As Long = 5& + EPSNERREXT   ' // invalid property value
Public Const OPOS_EX_NOTSUPPORTED As Long = 6& + EPSNERREXT     ' // function not supported
Public Const OPOS_EX_NOASB As Long = 7& + EPSNERREXT            ' // no ASB data returned
Public Const OPOS_EX_BUSY As Long = 8& + EPSNERREXT             ' // Async Output busy
Public Const OPOS_EX_INCAPABLE As Long = 9& + EPSNERREXT        ' // incapable of the function
Public Const OPOS_EX_INVALIDMODE As Long = 10& + EPSNERREXT     ' // invalid device mode
Public Const OPOS_EX_REOPEN As Long = 11& + EPSNERREXT          ' // re-open
Public Const OPOS_EX_BADPEEKRANGE As Long = 12& + EPSNERREXT    ' // PeekRange invalid
Public Const OPOS_EX_BADDISPRANGE As Long = 13& + EPSNERREXT    ' // DispatchRange invalid
Public Const OPOS_EX_NOTCLAIMED As Long = 14& + EPSNERREXT      ' // Not claimed (Release method used)
Public Const OPOS_EX_TIMEOUT As Long = 15& + EPSNERREXT         ' // sync output timeout
Public Const OPOS_EX_PORTUSED As Long = 16& + EPSNERREXT        ' // Port used by another
Public Const OPOS_EX_PORTBUSY As Long = 17& + EPSNERREXT        ' // HOST Port busy
Public Const OPOS_EX_MICRMODE As Long = 18& + EPSNERREXT        ' // MICR mode
Public Const OPOS_EX_DEVBUSY As Long = 19& + EPSNERREXT         ' // device busy
Public Const OPOS_EX_BADINF As Long = 20& + EPSNERREXT          ' // invalid INF file

Public Const OPOS_EX_SOVERSION As Long = 100& + EPSNERREXT      ' // invalid SO version

Public Const OPOS_EX_BADPARAM As Long = PARAMERREXT             ' // invalid parameter (general)

' /////////////////////////////////////////////////////////////////////
' // DirectIO commands
' /////////////////////////////////////////////////////////////////////

' /////////////////////////////////////////////////////////////////////
' //
' // EPSNLOCK.H
' //
' //   LOCK header file for EPSON Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' //
' /////////////////////////////////////////////////////////////////////

' /////////////////////////////////////////////////////////////////////
' // "ResultCodeExtended" Property Constants for Scanner
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_ELOCK_NOTSUPPORTDEVICE As Long = 1& + LOCKERREXT      ' // Device is not supported
Public Const OPOS_ELOCK_KBHOOKOPEN As Long = 2& + LOCKERREXT            ' // KeyBoardHookDriver Open Error
Public Const OPOS_ELOCK_KBHOOKSTART As Long = 3& + LOCKERREXT           ' // KeyBoardHookDriver Start Error
Public Const OPOS_ELOCK_KBHOOKSTOP As Long = 4& + LOCKERREXT            ' // KeyBoardHookDriver Stop Error
Public Const OPOS_ELOCK_KBHOOKCLOSE As Long = 5& + LOCKERREXT           ' // KeyBoardHookDriver Close Error
Public Const OPOS_ELOCK_FAILURE As Long = 6& + LOCKERREXT               ' // 

' /////////////////////////////////////////////////////////////////////
' //
' // EPSNDISP.H
' //
' //   Line Display header file for EPSON Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' // 95-11-29 OPOS Release 1.0                                     Kata
' //
' /////////////////////////////////////////////////////////////////////

' /////////////////////////////////////////////////////////////////////
' // "ResultCodeExtended" Property Constants for Line Display
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_EDISP_TOOMANYWINDOWS As Long = 1& + DISPERREXT    ' // 

' /////////////////////////////////////////////////////////////////////
' // DirectIO Method Constants
' /////////////////////////////////////////////////////////////////////

Public Const DISP_DI_SYNC As Long = 0&  

Public Const DISP_DI_DUMMY As Long = 0&  

' /////////////////////////////////////////////////////////////////////
' //
' // EPSNMICR.H
' //
' //    header file for MICR Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' // 96-05-14 OPOS Release 1.0                                     RYU
' //
' /////////////////////////////////////////////////////////////////////

' //************************************************** new 96/05/15 **************
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_EMICR_ERRORDEVICESTATUS As Long = MICRERREXT + 1&  
Public Const OPOS_EMICR_DATAERROR As Long = MICRERREXT + 2&  
Public Const OPOS_EMICR_COMPORT As Long = MICRERREXT + 3&  
Public Const OPOS_EMICR_DATAEND As Long = MICRERREXT + 4&  
' //************************************************** old end 96/05/15 **********

' /////////////////////////////////////////////////////////////////////
' //
' // EPSNMSR.H
' //
' //   MSR header file for EPSON Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' //
' /////////////////////////////////////////////////////////////////////

' /////////////////////////////////////////////////////////////////////
' // "ResultCodeExtended" Property Constants for Scanner
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_EMSR_NOTSUPPORTDEVICE As Long = 1& + MSRERREXT    ' // Device is not supported
Public Const OPOS_EMSR_KBHOOKOPEN As Long = 2& + MSRERREXT          ' // KeyBoardHookDriver Open Error
Public Const OPOS_EMSR_KBHOOKSTART As Long = 3& + MSRERREXT         ' // KeyBoardHookDriver Start Error
Public Const OPOS_EMSR_KBHOOKSTOP As Long = 4& + MSRERREXT          ' // KeyBoardHookDriver Stop Error
Public Const OPOS_EMSR_KBHOOKCLOSE As Long = 5& + MSRERREXT         ' // KeyBoardHookDriver Close Error
Public Const OPOS_EMSR_FAILURE As Long = 6& + MSRERREXT             ' // 

' /////////////////////////////////////////////////////////////////////
' // "ErrorEvent" Event: "ResultCodeExtended" Parameter Constants
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_EMSR_SEPARATOR As Long = 10& + MSRERREXT              ' // 
Public Const OPOS_EMSR_ACCOUNTNUMBER As Long = 11& + MSRERREXT          ' // 
Public Const OPOS_EMSR_TITLE As Long = 12& + MSRERREXT                  ' // 
Public Const OPOS_EMSR_SECONDFIELD As Long = 13& + MSRERREXT            ' // 
Public Const OPOS_EMSR_EXPIRATIONDATA As Long = 14& + MSRERREXT         ' // 
Public Const OPOS_EMSR_SERVICECODE As Long = 15& + MSRERREXT            ' // 

' /////////////////////////////////////////////////////////////////////
' //
' // EPSNPTR.H
' //
' //   POS Printer header file for EPSON Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' //
' /////////////////////////////////////////////////////////////////////

' /////////////////////////////////////////////////////////////////////
' // DirectIO Method Constants
' /////////////////////////////////////////////////////////////////////

Public Const PTR_DI_SYNC As Long = 0&  
Public Const PTR_DI_IMMEDIATE As Long = 1&  
Public Const PTR_DI_REALTIME As Long = 2&  
Public Const PTR_DI_DOWNLOAD_BMP As Long = 3&  
Public Const PTR_DI_NOTDOWNLOAD_BMP As Long = 4&  
Public Const PTR_DI_ASCII_STRING As Long = 5&  
' //******************** add Ver1.40 Label **************************************
Public Const PTR_DI_LABEL_REMOVE As Long = 6&  
Public Const PTR_DI_LABEL_PRINTMODE_RIGHT_SPACE As Long = 7&  
Public Const PTR_DI_LABEL_PRINTMODE_RIGHT_ZERO As Long = 8&  
Public Const PTR_DI_LABEL_PRINTMODE_LEFT_SPACE As Long = 9&  
Public Const PTR_DI_LABEL_COUNTMODE As Long = 10&  
Public Const PTR_DI_LABEL_PRINT_COUNT As Long = 11&  
Public Const PTR_DI_LABEL_COUNTVALUE As Long = 12&  
' //******************** add end Ver1.40 Label **********************************

Public Const PTR_DI_DUMMY As Long = 0&  

' /////////////////////////////////////////////////////////////////////
' // "ResultCodeExtended" Property Constants for Printer
' /////////////////////////////////////////////////////////////////////
' //************************************************** add 96/04/09 **************
Public Const OPOS_EPTR_BADSTATION As Long = 1& + PTRERREXT      ' // invalid station
Public Const OPOS_EPTR_NOSTATION As Long = 2& + PTRERREXT       ' // station not present
                                                                ' //************************************************** add end 96/04/09 **********

Public Const OPOS_EPTR_UNRECOVERABLE As Long = 3& + PTRERREXT  
Public Const OPOS_EPTR_CUTTER As Long = 4& + PTRERREXT  
Public Const OPOS_EPTR_MECHANICAL As Long = 5& + PTRERREXT  
Public Const OPOS_EPTR_OVERHEAT As Long = 6& + PTRERREXT  
Public Const OPOS_EPTR_ROTATE90 As Long = 7& + PTRERREXT  

' /////////////////////////////////////////////////////////////////////
' // "StatusUpdateEvent" Event: "Data" Parameter Constants
' /////////////////////////////////////////////////////////////////////

Public Const PTR_SUE_UNRECOVERABLE_ERROR As Long = 1& + 2000&   ' // (Several)
Public Const PTR_SUE_UNRECOVERABLE_OK As Long = 2& + 2000&      ' // (Several)
Public Const PTR_SUE_CUTTER_ERROR As Long = 3& + 2000&          ' // (Several)
Public Const PTR_SUE_CUTTER_OK As Long = 4& + 2000&             ' // (Several)
Public Const PTR_SUE_MECHANICAL_ERROR As Long = 5& + 2000&      ' // (Several)
Public Const PTR_SUE_MECHANICAL_OK As Long = 6& + 2000&         ' // (Several)

' /////////////////////////////////////////////////////////////////////
' //
' // EPSNSCAL.H
' //
' //   SCALE header file for EPSON Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' //
' /////////////////////////////////////////////////////////////////////

' /////////////////////////////////////////////////////////////////////
' // "ResultCodeExtended" Property Constants for Scale
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_ESCAL_NOTRESPONSE As Long = 1& + SCALERREXT   ' // 
Public Const OPOS_ESCAL_BADDATA As Long = 2& + SCALERREXT       ' // 
Public Const OPOS_ESCAL_WEIGHTOVER As Long = 3& + SCALERREXT    ' // 
Public Const OPOS_ESCAL_LOCKED As Long = 4& + SCALERREXT        ' // 

' /////////////////////////////////////////////////////////////////////
' // "DirectIO" Method Constants for Scale
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_SCAL_RESET As Long = 0&           ' // 
Public Const OPOS_SCAL_WEIGHTREQUEST As Long = 1&   ' // 
Public Const OPOS_SCAL_CANCEL As Long = 2&          ' // 
Public Const OPOS_SCAL_STATUS As Long = 3&          ' // 
Public Const OPOS_SCAL_MONITOR As Long = 4&         ' //
                                                    ' //const LONG OPOS_SCAL_DISP_DATA		= 5;	//
Public Const OPOS_SCAL_DISP_STATUS As Long = 5&     ' //

' /////////////////////////////////////////////////////////////////////
' //
' // EPSNSCAN.H
' //
' //   SCANNER header file for EPSON Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' //
' /////////////////////////////////////////////////////////////////////

' /////////////////////////////////////////////////////////////////////
' // "ResultCodeExtended" Property Constants for Scanner
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_ESCAN_PORTCOMM As Long = 1& + SCANERREXT      ' // 
Public Const OPOS_ESCAN_DATAFORMAT As Long = 2& + SCANERREXT    ' // 
Public Const OPOS_ESCAN_DATAEND As Long = 3& + SCANERREXT       ' // 
Public Const OPOS_ESCAN_FAILURE As Long = 4& + SCANERREXT       ' // 
Public Const OPOS_ESCAN_NOTRESPONSE As Long = 5& + SCANERREXT   ' // 
Public Const OPOS_ESCAN_LOCKED As Long = 6& + SCANERREXT        ' // 

' /////////////////////////////////////////////////////////////////////
' // "DirectIO" Method Constants for Scanner
' /////////////////////////////////////////////////////////////////////
Public Const OPOS_SCAN_OUTPUT As Long = 0&              ' // 

Public Const OPOS_SCAN_MULTI As Long = 1&           ' // 

Public Const OPOS_SCAN_RESET As Long = 10&              ' // 
Public Const OPOS_SCAN_ENABLE As Long = 11&             ' // 
Public Const OPOS_SCAN_NOINDICATION As Long = 12&       ' // 
Public Const OPOS_SCAN_STATUS As Long = 13&             ' // 
Public Const OPOS_SCAN_SWITCHREAD As Long = 14&         ' //
Public Const OPOS_SCAN_NOTONFILE As Long = 15&          ' //
Public Const OPOS_SCAN_REDLIGHTFLASH As Long = 16&      ' //
                                                        ' //const LONG OPOS_SCAN_DISP_DATA		= 17;	//
Public Const OPOS_SCAN_DISP_STATUS As Long = 17&        ' //

' /////////////////////////////////////////////////////////////////////
' //
' // EPSNSIG.H
' //
' //   The header file for SigCap Applications.
' //
' // Modification history
' // ------------------------------------------------------------------
' // 96-05-23 OPOS Release 1.1                                     RYU
' //
' /////////////////////////////////////////////////////////////////////

Public Const OPOS_ESIG_MEMRAW As Long = SIGERREXT + 1&  
Public Const OPOS_ESIG_MEMASC As Long = SIGERREXT + 2&  
Public Const OPOS_ESIG_SIGBEGIN As Long = SIGERREXT + 3&  
Public Const OPOS_ESIG_SIGNOTBEGIN As Long = SIGERREXT + 4&  
Public Const OPOS_ESIG_DATAERROR As Long = SIGERREXT + 5&  
Public Const OPOS_ESIG_BUFFERFULL As Long = SIGERREXT + 6&  
Public Const OPOS_ESIG_DATAEND As Long = SIGERREXT + 7&  
Public Const OPOS_ESIG_PORTERROR As Long = SIGERREXT + 8&  

Public Const SIG_DI_TYPE1 As Long = 0&          ' // for WHOLE 
Public Const SIG_DI_TYPE2 As Long = 1&          ' // for PACKET

' // DirectIO Method Constants(VectorArray Property Type Convert)
Public Const SIG_DI_CONV As Long = 2&       ' // Convert to Decimal Number and Convert ASCII Code 

