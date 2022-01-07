#include-once

#Region APPLICATION
    Global Const $__ROCKETBZ_ROOT64_FOLDER = "C:\Program Files\BlueZone\"
    Global Const $__ROCKETBZ_ROOT32_FOLDER = "C:\Program Files (x86)\BlueZone\"
    Global Const $__ROCKETBZ_VERSION = "7.1"

    Global Const $__ROCKETBZ_ROOT_FOLDER = @AutoItX64 ? $__ROCKETBZ_ROOT64_FOLDER : $__ROCKETBZ_ROOT32_FOLDER

    Global Const $__ROCKETBZ_MAINFRAME_DISPLAY = "bzmd.exe"
    Global Const $__ROCKETBZ_VT_DISPLAY = "bzvt.exe"
    Global Const $__ROCKETBZ_ISERIES_DISPLAY = "bzad.exe"

    Global $__ROCKETBZ_CURRENT_HOSTTYPE = ""
    Global $__ROCKETBZ_WAITHOSTTIME = 75
    Global $__ROCKETBZ_DEBUG = True
#EndRegion APPLICATION

#Region ERROR_CODES
    ;-- BlueZone error codes correspond with pre-defined Windows HLLAPI return codes.
    ;-- https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_error-codes.htm

    Global Const $__ROCKETBZ_ERR_OK = 0 ;-- Successful.
    Global Const $__ROCKETBZ_ERR_NOTCONNECTED = 1 ;-- Not Connected To Presentation Space.
    Global Const $__ROCKETBZ_ERR_BLOCKNOTAVAIL = 1 ;-- Requested size is not available.
    Global Const $__ROCKETBZ_ERR_PARAMETERERROR = 2 ;-- Parameter Error/Invalid Function.
    Global Const $__ROCKETBZ_ERR_BLOCKIDINVALID = 2 ;-- Invalid Block ID was specified.
    Global Const $__ROCKETBZ_ERR_FTXCOMPLETE = 3 ;-- File Transfer Complete.
    Global Const $__ROCKETBZ_ERR_FTXSEGMENTED = 4 ;-- File Transfer Complete / segmented.
    Global Const $__ROCKETBZ_ERR_PSBUSY = 4 ;-- Presentation Space is Busy.
    Global Const $__ROCKETBZ_ERR_INHIBITED = 5 ;-- Inhibited/Keyboard Locked.
    Global Const $__ROCKETBZ_ERR_TRUNCATED = 6 ;-- Data Truncated.
    Global Const $__ROCKETBZ_ERR_POSITIONERROR = 7 ;-- Invalid Presentation Space Position.
    Global Const $__ROCKETBZ_ERR_NOTAVAILABLE = 8 ;-- Unavailable Operation.
    Global Const $__ROCKETBZ_ERR_SYSERROR = 9 ;-- System Error.
    Global Const $__ROCKETBZ_ERR_WOULDBLOCK = 10 ;-- Blocking error.
    Global Const $__ROCKETBZ_ERR_UNAVAILABLE = 11 ;-- Resource is unavailable.
    Global Const $__ROCKETBZ_ERR_PSENDED = 12 ;-- The session stopped.
    Global Const $__ROCKETBZ_ERR_UNDEFINEDKEY = 20 ;-- Undefined Key Combination.
    Global Const $__ROCKETBZ_ERR_OIAUPDATE = 21 ;-- OIA Updated.
    Global Const $__ROCKETBZ_ERR_PSUPDATE = 22 ;-- PS Updated.
    Global Const $__ROCKETBZ_ERR_BOTHUPDATE = 23 ;-- Both PS And OIA Updated.
    Global Const $__ROCKETBZ_ERR_NOFIELD = 24 ;-- No Such Field Found.
    Global Const $__ROCKETBZ_ERR_NOKEYSTROKES = 25 ;-- No Keystrokes are available.
    Global Const $__ROCKETBZ_ERR_PSCHANGED = 26 ;-- PS or OIA changed.
    Global Const $__ROCKETBZ_ERR_FTXABORTED = 27 ;-- File transfer aborted.
    Global Const $__ROCKETBZ_ERR_ZEROLENFIELD = 28 ;-- Field length is zero.
    Global Const $__ROCKETBZ_ERR_INVALIDTYPE = 30 ;-- Invalid Cursor Type.
    Global Const $__ROCKETBZ_ERR_KEYOVERFLOW = 31 ;-- Keystroke overflow.
    Global Const $__ROCKETBZ_ERR_SFACONN = 32 ;-- Other application already connected.
    Global Const $__ROCKETBZ_ERR_TRANCANCLI = 34 ;-- Message sent inbound to host canceled*/
    Global Const $__ROCKETBZ_ERR_TRANCANCL = 35 ;-- Outbound trans from host canceled.
    Global Const $__ROCKETBZ_ERR_HOSTCLOST = 36 ;-- Contact with host was lost.
    Global Const $__ROCKETBZ_ERR_OKDISABLED = 37 ;-- The function was successful.
    Global Const $__ROCKETBZ_ERR_NOTCOMPLETE = 38 ;-- The requested fn was not completed.
    Global Const $__ROCKETBZ_ERR_SFDDM = 39 ;-- One DDM session already connected.
    Global Const $__ROCKETBZ_ERR_SFDPEND = 40 ;-- Disconnected w async requests pending*/
    Global Const $__ROCKETBZ_ERR_BUFFINUSE = 41 ;-- Specified buffer currently in use.
    Global Const $__ROCKETBZ_ERR_NOMATCH = 42 ;-- No matching request found.
    Global Const $__ROCKETBZ_ERR_LOCKERROR = 43 ;-- API already locked or unlocked.
    
    Global Const $__ROCKETBZ_ERR_INVALIDFUNCTIONNUM = 301 ;-- Invalid function number.
    Global Const $__ROCKETBZ_ERR_FILENOTFOUND = 302 ;-- File Not Found.
    Global Const $__ROCKETBZ_ERR_ACCESSDENIED = 305 ;-- Access Denied.
    Global Const $__ROCKETBZ_ERR_MEMORY = 308 ;-- Insufficient Memory.
    Global Const $__ROCKETBZ_ERR_INVALIDENVIRONMENT = 310 ;-- Invalid environment.
    Global Const $__ROCKETBZ_ERR_INVALIDFORMAT = 311 ;-- Invalid format.
    
    Global Const $__ROCKETBZ_ERR_INVALIDPSID = 9998 ;-- Invalid Presentation Space ID.
    Global Const $__ROCKETBZ_ERR_INVALIDRC = 9999 ;-- Invalid Row or Column Code.
    
    Global Const $__ROCKETBZ_ERR_ALREADY = 61440 ;-- An async call is already outstanding
    Global Const $__ROCKETBZ_ERR_INVALID = 61441 ;-- Async Task Id is invalid
    Global Const $__ROCKETBZ_ERR_CANCEL = 61442 ;-- Blocking call was canceled
    Global Const $__ROCKETBZ_ERR_SYSNOTREADY = 61443 ;-- Underlying subsystem not started
    Global Const $__ROCKETBZ_ERR_VERNOTSUPPORTED = 61444 ;-- Application version not supported
#EndRegion ERROR_CODES