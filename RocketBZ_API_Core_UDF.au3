#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: RocketBZ_Core_API_UDF
    Description...........: UDF for Rocket BlueZone terminal emulator API
    Dependencies..........: Rocket Bluezone 7.1 and it's architecture (x32/x64);
                            <RocketBZ_Globals.au3>

    Documentation.........: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_acon_bz-host-automation-object.htm
                            https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_bz-object-methods.htm

    Author................: exorcistas@github.com
    Modified..............: 2020-11-16
    Version...............: v1.8.2b
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <RocketBZ_Globals.au3>

#AutoIt3Wrapper_UseX64=N    ;-- set 'Y' if application is 64-bit

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
%% HOST %%
    __RocketBZ_CreateObject($_sHostType = $__RocketBZ_MAINFRAME_DISPLAY)
    __RocketBZ_Connect($_oRocketBZHost, $_sSessionShortName = "", $_iConnectRetryTimeout = 10)
    __RocketBZ_ConnectToHost($_oRocketBZHost, $_iSessionTypeVal = 0, $_iSessionIdentifierVal = 0)
    __RocketBZ_Connected($_oRocketBZHost)
    __RocketBZ_GetSessionId($_oRocketBZHost)
    __RocketBZ_GetSessionName($_oRocketBZHost)
    __RocketBZ_FullName($_oRocketBZHost)
    __RocketBZ_CloseSession($_oRocketBZHost, $_iSessionTypeVal = 0, $_iSessionIdentifierVal = 0)
    __RocketBZ_Disconnect($_oRocketBZHost)

%% CORE %%
    __RocketBZ_WriteScreen($_oRocketBZHost, $_sWriteStr, $_iRowVal, $_iColumnVal)
    __RocketBZ_ReadScreen($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
    __RocketBZ_PSGetText($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
    __RocketBZ_SendKey($_oRocketBZHost, $_sKeyStr)
    __RocketBZ_WaitReady($_oRocketBZHost, $_iTimeoutVal = 0, $_iExtraWaitVal = $__RocketBZ_WAITHOSTTIME)
    __RocketBZ_WaitForReady($_oRocketBZHost)
    __RocketBZ_SetCursor($_oRocketBZHost, $_iRowVal, $_iColumnVal)
    __RocketBZ_WaitForText($_oRocketBZHost, $_sFindStr, $_iRowVal = 1, $_iColumnVal = 1, $_iTimeoutValue = 0)
    __RocketBZ_FindText($_oRocketBZHost, $_sFindStr, $_iRowVal = 1, $_iColumnVal = 1, $_bIgnoreCase = True)
    __RocketBZ_FoundTextRow($_oRocketBZHost)
    __RocketBZ_FoundTextColumn($_oRocketBZHost)
    __RocketBZ_PSSearch($_oRocketBZHost, $_sSearchStr, ByRef $_iRow, ByRef $_iCol, $_iStartPosVal = 1)
    __RocketBZ_LockKeyboard($_oRocketBZHost, $_bLock = False)
    __RocketBZ_Status($_oRocketBZHost)
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region HOST

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_CreateObject()
        Description ...: Creates object for Rocket BlueZone terminal emulator

        Return values..: Success - returns object value
                         Failure - returns @error

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-03
    #ce ===============================================================================================================================
    Func __RocketBZ_CreateObject($_sHostType = $__RocketBZ_MAINFRAME_DISPLAY)
        Local $_ProcID = ProcessExists($_sHostType)
            If $_ProcID = 0 Then Return SetError($__ROCKETBZ_ERR_UNAVAILABLE, 1, -1)

        Local $_sSystemName = "BZWhll.WhllObj"
        Local $_oRocketBZHost = ObjCreate($_sSystemName)
        Local $_iErrCode = @error

            If $__ROCKETBZ_DEBUG Then
                Local $_sArch = (@AutoItX64) ? "x64" : "x32"
                ConsoleWrite("[__RocketBZ_CreateObject]:   " & $_iErrCode & @CRLF & _
                            @TAB & "HostType:    " & $_sHostType & @CRLF & _
                            @TAB & "Process ID: " & $_ProcID & @CRLF & _
                            @TAB & "Architecture:    " & $_sArch & @CRLF & @CRLF)
            EndIf
            $__ROCKETBZ_CURRENT_HOSTTYPE = $_sHostType

        Return SetError($_iErrCode, 2, $_oRocketBZHost)
    EndFunc   ;==>__RocketBZ_CreateObject

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_Connect($_oRocketBZHost, $_sSessionShortName = "", $_iConnectRetryTimeout = 10)
        Description ...: Establishes a link between the BlueZone Host Automation object and the BlueZone Display session.
                         Connect must be called before any other BlueZone Host Automation object methods that access data in the host screen.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-Connect.htm
        Parameters.....:
                         {$_sSessionShortName}
                         Uniquely identifies the BlueZone Display session.
                         The session name corresponds to the HLLAPI Short Name Session Identifier configured in the Options → API settings in the BlueZone Display emulator.
                         As an option, when the BlueZone session is embedded, using an exclamation point (!) can be used in place of an actual SessionShortName.
                         When ! is used, the BZHAO auto-determines the session name.
                         If the SessionShortName parameter is empty string, the BZHAO connects to the session that is running in the foreground.
                         If no session is running in the foreground, then the first session found is used.

                         {$_iConnectRetryTimeout}
                         Optional: Used to set the time in seconds that Connect spends attempting to connect to the BlueZone session.
                         Also, as an option, a zero (0) can be used to cause the Connect to abort if the first attempt fails.

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: The object can only be connected to one display session at a time.
                         When connecting to multiple sessions, any previously connected session is automatically disconnected.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_Connect($_oRocketBZHost, $_sSessionShortName = "", $_iConnectRetryTimeout = 10)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.Connect($_sSessionShortName, $_iConnectRetryTimeout)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_Connect]:   " & $_iReturnCode & @CRLF & _
                            @TAB & "SessionShortName:    " & $_sSessionShortName & @CRLF & _
                            @TAB & "ConnectRetryTimeout:  " & $_iConnectRetryTimeout & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_Connect

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_ConnectToHost($_oRocketBZHost, $_iSessionTypeVal = 0, $_iSessionIdentifierVal = 0)
        Description ...: Forces a specific BlueZone session to connect to a host.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-ConnectToHost.htm
        Parameters.....:
                         {$_iSessionTypeVal}
                         0 - Mainframe; 1 - iSeries; 2 - VT; 3 - UTS; 4 - T27; 6 - 6530

                         {$_iSessionIdentifierVal}
                         1 for S1; 2 for S2; 3 for S3; etc.
                         0 is a special case where the object will try to find an embedded BlueZone session in the browser or search for any BlueZone window to link.

        Returns........: 0 for success, or a non-zero for failure

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_ConnectToHost($_oRocketBZHost, $_iSessionTypeVal = 0, $_iSessionIdentifierVal = 0)
        If Not IsObj($_oRocketBZHost) Then SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.ConnectToHost($_iSessionTypeVal, $_iSessionIdentifierVal)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_ConnectToHost]:   " & $_iReturnCode & _
                            @TAB & "SessionTypeVal:    " & $_iSessionTypeVal & _
                            @TAB & "SessionIdentifierVal:  " & $_iSessionIdentifierVal & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_ConnectToHost

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_Connected($_oRocketBZHost)
        Description ...: Retrieves the session's host connection status.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-Connected.htm
        Parameters.....: -

        Returns........: True if the session is connected to the host system, or False otherwise.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_Connected($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_bIsConnected = $_oRocketBZHost.Connected()
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_Connected]:   " & $_bIsConnected & @CRLF & @CRLF)
            EndIf

        Return $_bIsConnected
    EndFunc   ;==>__RocketBZ_Connected

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_GetSessionId($_oRocketBZHost)
        Description ...: Returns the session identifier (1,2,3,etc.) of the currently connected session.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-GetSessionId.htm
        Parameters.....: -

        Returns........: Returns the session identifier
        Remarks........: This method can only be used after a successful Connect() and can be used in subsequent calls that require a session name parameter.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_GetSessionId($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_sSessionId = $_oRocketBZHost.GetSessionId()
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_GetSessionId]:   " & $_sSessionId & @CRLF & @CRLF)
            EndIf

        Return $_sSessionId
    EndFunc   ;==>__RocketBZ_GetSessionId

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_GetSessionName($_oRocketBZHost)
        Description ...: Returns the HLLAPI Short Name (''A'',''B'',''C'',etc.) of the currently connected BlueZone session.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-GetSessionName.htm
        Parameters.....: -

        Returns........: Returns the session name
        Remarks........: This method can only be used after a successful Connect() and can be used in subsequent calls that require a session identifier parameter.
                         While not necessary, this method can improve the performance of the script,
                         because the BZHAO does not need to enumerate child windows of the browser when attempting to auto-locate the BlueZone sessions.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_GetSessionName($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

       Local  $_sSessionName = $_oRocketBZHost.GetSessionName()
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_GetSessionName]:   " & $_sSessionName & @CRLF & @CRLF)
            EndIf

        Return $_sSessionName
    EndFunc   ;==>__RocketBZ_GetSessionName

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_FullName($_oRocketBZHost)
        Description ...: Returns the full filename of the BlueZone object.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-FullName.htm
        Parameters.....: -

        Returns........: Returns the full filename (dll path used)

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_FullName($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

       Local  $_sFullName = $_oRocketBZHost.FullName()
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_FullName]:   " & $_sFullName & @CRLF & @CRLF)
            EndIf

        Return $_sFullName
    EndFunc   ;==>__RocketBZ_FullName

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_CloseSession($_oRocketBZHost, $_iSessionTypeVal = 0, $_iSessionIdentifierVal = 1)
        Description ...: Disconnects from the host system and closes the BlueZone Display session window.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-CloseSession.htm
        Parameters.....:
                         {$_iSessionTypeVal}
                         0 - Mainframe; 1 - iSeries; 2 - VT; 3 - UTS; 4 - T27; 6 - 6530

                         {$_iSessionIdentifierVal}
                         1 for S1; 2 for S2; 3 for S3; etc.

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: CloseSession can be used to end a session started with the OpenSession method.
                         Starting with BlueZone version 3.3, the CloseSession method works with BlueZone Web-to-Host when using the Embedded Client Mode.
                         Note: The embedded BlueZone session must first be connected to the program or script by using the Connect method.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_CloseSession($_oRocketBZHost, $_iSessionTypeVal = 0, $_iSessionIdentifierVal = 1)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode =  $_oRocketBZHost.CloseSession($_iSessionTypeVal, $_iSessionIdentifierVal)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_CloseSession]:   " & $_iReturnCode & _
                            @TAB & "SessionTypeVal:    " & $_iSessionTypeVal & _
                            @TAB & "SessionIdentifierVal:  " & $_iSessionIdentifierVal & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_CloseSession

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_Disconnect($_oRocketBZHost)
        Description ...: The Disconnect method is used in conjunction with the Connect method,
                         to halt communication between the Host Automation Object and the currently connected BlueZone session.
                         To reestablish communication with the same BlueZone Session or a new BlueZone session, use the Connect method. See the Connect method for more information.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-Disconnect.htm
        Parameters.....: None.

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_Disconnect($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.Disconnect()
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_Disconnect]:   " & $_iReturnCode & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_Disconnect

#EndRegion HOST

#Region CORE

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_WriteScreen($_oRocketBZHost, $_sWriteStr, $_iRowVal, $_iColumnVal)
        Description ...: Pastes specified text in the host screen.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-WriteScreen.htm
        Parameters.....:
                         {$_sWriteStr} Text to place in host screen.
                         {$_iRowVal} Row position.
                         {$_iColumnVal} Column position.

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: The WriteScreen method can only paste text in unprotected fields in the host screen.
                         In a BlueZone VT session, the WriteStr parameter is only echoed to the VT client screen.
                         The WriteStr parameter is never sent to the host.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_WriteScreen($_oRocketBZHost, $_sWriteStr, $_iRowVal, $_iColumnVal)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.WriteScreen($_sWriteStr, $_iRowVal, $_iColumnVal)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_WriteScreen]:   " & $_iReturnCode & @CRLF & _
                            @TAB & "WriteStr:   " & $_sWriteStr & @CRLF & _
                            @TAB & "RowVal, ColumnVal: " & $_iRowVal & "/" & $_iColumnVal & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_WriteScreen

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_ReadScreen($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
        Description ...: Retrieves data from the host screen.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-ReadScreen.htm
        Parameters.....:
                         {$_iLengthVal} Number of characters to read.
                         {$_iRowVal} Row position.
                         {$_iColumnVal} Column position.

        Returns........: String from screen, return code 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: When the ReadScreen function returns, the BufferStr variable contains the host screen data.

        Author ........: exorcistas@github.com
        Modified.......: 2020-11-05
    #ce ===============================================================================================================================
    Func __RocketBZ_ReadScreen($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_sBufferStr = ""  ;-- Variable to contain host screen data.
        Local $_iReturnCode = $_oRocketBZHost.ReadScreen($_sBufferStr, $_iLengthVal, $_iRowVal, $_iColumnVal)

            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_ReadScreen]:   " & $_iReturnCode & @CRLF & _
                            @TAB & "BufferStr:   " & $_sBufferStr & @CRLF & _
                            @TAB & "RowVal, ColumnVal, LengthVal: " & $_iRowVal & "/" & $_iColumnVal & "/" & $_iLengthVal & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2, $_sBufferStr)
    EndFunc   ;==>__RocketBZ_ReadScreen

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_PSGetText($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
        Description ...: Reads text from the host screen into a variable.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-PSGetText.htm
        Parameters.....:
                         {$_iLengthVal} Number of characters to read.
                         {$_iRowVal} Row position.
                         {$_iColumnVal} Column position.

        Returns........: String containing the text for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: The screen position starts at 1 in the upper-left corner of the window (row 1, column 1),
                         and ends at the bottom-right of the window (max row times max column).
                         For example, for a Model 2 - 24 x 80 screen, the last position is 1920.

        Author ........: exorcistas@github.com
        Modified.......: 2020-11-11
    #ce ===============================================================================================================================
    Func __RocketBZ_PSGetText($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)
        If NOT (IsNumber($_iRowVal) OR IsNumber($_iColumnVal) OR IsNumber($_iLengthVal)) Then Return SetError($__ROCKETBZ_ERR_INVALIDRC, 2, -1)

        If $_iRowVal <> 1 Then
            $_iRowVal = $_iRowVal - 1
            $_iRowVal = $_iRowVal * 80
        Else
            $_iColumnVal = $_iColumnVal - 1
        EndIf

        Local $_iPositionVal = $_iRowVal + $_iColumnVal
        Local $_sResponse = $_oRocketBZHost.PSGetText($_iLengthVal,$_iPositionVal)

        If $__ROCKETBZ_DEBUG Then
            ConsoleWrite("[__RocketBZ_PSGetText]:   " & @CRLF & _
                        @TAB & "RowVal, ColumnVal, LengthVal: " & $_iRowVal & "/" & $_iColumnVal & "/" & $_iLengthVal & @CRLF & _
                        @TAB & "PositionVal:   " & $_iPositionVal & @CRLF & _
                        @TAB & "Response:   " & $_sResponse & @CRLF & @CRLF)
        EndIf

        Return $_sResponse
    EndFunc

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_SendKey($_oRocketBZHost, $_sKeyStr)
        Description ...: Sends a sequence of keys to the display session.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-SendKey.htm

        Parameters.....: {$_sKeyStr} String of key codes. See the Remarks section for a complete listing of valid key codes and descriptions.

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: The SendKey function affects the host screen as if the user were typing on the keyboard.
                         Refer to the following tables for a list of host function keys and their corresponding codes.
                         If a character is used in the code, then the case of the character is important.
                         IBM 3270/5250 send key table:  https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_ibm-send-key-table.htm
                         BlueZone VT send key table:    https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_vt-send-key-table.htm
                         UTS send key table:            https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_uts-send-keys.htm

                         Note: If you want to use the “at” sign (@) in the data string you must use the two-byte code “@@”.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_SendKey($_oRocketBZHost, $_sKeyStr)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.SendKey($_sKeyStr)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_SendKey]:   " & $_iReturnCode & @CRLF & _
                            @TAB & "KeyStr:   " & $_sKeyStr & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_SendKey

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_WaitReady($_oRocketBZHost, $_iTimeoutVal = 5, $_iExtraWaitVal = $__RocketBZ_WAITHOSTTIME)
        Description ...: Suspends script execution until the host screen is ready for keyboard input. (Not recommended for VT sessions.)
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-WaitReady.htm
        Parameters.....:
                         {$_iTimeoutVal} The number of seconds to wait before returning with a session is busy error code.
                         {$_iExtraWaitVal}
                         The number of milliseconds to validate for a keyboard unlocked status. For a detailed explanation, see the Remarks section below.
                         Note: The above features behave differently when scripting non-IBM hosts. See Non-IBM Remarks below.

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: The WaitReady function must be called each time after sending a attention identifier key (such as a PF key) to the display session.
                         If ExtraWaitVal is a value in the range of 1, 2, 3, on up to 50,
                         then script execution is suspended until the host machine sends the specified number of keyboard restores.
                         Refer to the BlueZone status bar to determine the keyboard restore count for a given screen.
                         In the following screen shot Ready (2) on the status bar means two keyboard restores were detected when this particular screen was written by the host.
                         If ExtraWaitVal is set to 51 or higher, the operation of the parameter changes to specify the number of milliseconds to wait
                         after the keyboard lock has been detected prior to executing the next script command.

                         Note: WaitReady only works after an AID key is sent to the host.
                         If WaitReady is used after data is put in a field, but not sent to the host,
                         then the wait count is never reached and the command times out (first parameter).
                         When converting macros and replacing WaitHostQuiet with WaitReady, ensure the preceding command is Enter, PFKey, Attn, SysReq, or some other AID key
                         that causes the host to write to the screen.

                         Non-IBM remark: TimeoutVal and ExtraWaitVal behave differently when scripting non-IBM hosts.
                         That is because keyboard locked status is not supported on non-IBM hosts.
                         When scripting on non-IBM hosts, set TimeoutVal to 0 and treat ExtraWaitVal as a pause before the scripts moves on to the next command.

        Author ........: exorcistas@github.com
        Modified.......: 2020-11-09
    #ce ===============================================================================================================================
    Func __RocketBZ_WaitReady($_oRocketBZHost, $_iTimeoutVal = 0, $_iExtraWaitVal = $__RocketBZ_WAITHOSTTIME)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.WaitReady($_iTimeoutVal, $_iExtraWaitVal)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_WaitReady]:   " & $_iReturnCode & @CRLF & _
                            @TAB & "TimeoutVal:   " & $_iTimeoutVal & @CRLF & _
                            @TAB & "ExtraWaitVal:  " & $_iExtraWaitVal & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_WaitReady

        #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_WaitForReady($_oRocketBZHost)
        Description ...: Used after sending an AID key to wait for the host keyboard to unlock. (Not suggested for VT hosts as the hosts do not lock keyboards.)
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-WaitForReady.htm

        Parameters.....: -

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_WaitForReady($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.WaitForReady()
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_WaitForReady]:   " & $_iReturnCode & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_WaitForReady

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_SetCursor($_oRocketBZHost, $_iRowVal, $_iColumnVal)
        Description ...: Sets the host screen cursor position.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-SetCursor.htm
        Parameters.....:
                         {$_iRowVal} Row position.
                         {$_iColumnVal} Column position.

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: BlueZone VT attempts to move the cursor on the screen by sending cursor movement commands to the host.
                         Not all VT applications/screens support cursor movement commands.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_SetCursor($_oRocketBZHost, $_iRowVal, $_iColumnVal)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.SetCursor($_iRowVal, $_iColumnVal)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_SetCursor]:   " & $_iReturnCode & @CRLF & _
                            @TAB & "RowVal, ColumnVal: " & $_iRowVal & "/" & $_iColumnVal & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_SetCursor

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_WaitForText($_oRocketBZHost, $_sFindStr, $_iRowVal = 1, $_iColumnVal = 1, $_iTimeoutValue = 0)
        Description ...: Suspends script execution until the desired text is found in the host screen.
                         The WaitForText function can be used to verify that a specific host screen is being displayed before continuing with script execution.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-WaitForText.htm
        Parameters.....:
                         {$_sTextStr} The text string that you want to search for in the host screen
                         {$_iRowVal} Row start position where the search is to begin
                         {$_iColumnVal} Column start position where the search is to begin
                         {$_iTimeoutValue} The number of seconds to wait before returning with a session is busy error code.

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: The WaitForText function can be used to verify that a specific host screen is being displayed before continuing with script execution.
                         This method is Base-0 for VT/6530 sessions, and Base-1 for all other session types.

        Author ........: exorcistas@github.com
        Modified.......: 2020-11-12
    #ce ===============================================================================================================================
    Func __RocketBZ_WaitForText($_oRocketBZHost, $_sFindStr, $_iRowVal = 1, $_iColumnVal = 1, $_iTimeoutValue = 0)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.WaitForText($_sFindStr, $_iRowVal, $_iColumnVal, $_iTimeoutValue)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_WaitForText]:   " & $_iReturnCode & @CRLF & _
                            @TAB & "FindStr:   " & $_sFindStr & @CRLF & _
                            @TAB & "TimeoutValue:   " & $_iTimeoutValue & @CRLF & _
                            @TAB & "RowVal, ColumnVal: " & $_iRowVal & "/" & $_iColumnVal & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_WaitForText

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_FindText($_oRocketBZHost, $_sFindStr, $_iRowVal = 1, $_iColumnVal = 1, $_bIgnoreCase = True)
        Description ...: Looks for a specific string on the screen.
        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-FindText.htm
        Parameters.....:
                         {$_sFindStr} Variable search string
                         {$_iRowVal} Row start position
                         {$_iColumnVal} Column start position
                         {$_bIgnoreCase} An optional parameter to disable case sensitive find.

        Returns........: False for failure; True for success
        Remarks........: Use FoundTextRow and FoundTextColumn to retrieve the starting row and column of where the text was found.
                         In VT/6530 sessions you can use a negative row start position to search back into the scrollback history.
                         This method is Base-0 for VT/6530 sessions, and Base-1 for all other session types.

        Author ........: exorcistas@github.com
        Modified.......: 2020-11-04
    #ce ===============================================================================================================================
    Func __RocketBZ_FindText($_oRocketBZHost, $_sFindStr, $_iRowVal = 1, $_iColumnVal = 1, $_bIgnoreCase = True)
		Local $_iRow
		Local $_iColumn

        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        If $__ROCKETBZ_CURRENT_HOSTTYPE = $__ROCKETBZ_MAINFRAME_DISPLAY Then
            $_iRow = $_iRowVal-1
            $_iColumn = $_iColumnVal-1
        EndIf

        Local $_bTextExists = $_oRocketBZHost.FindText($_sFindStr, $_iRow, $_iColumn, $_bIgnoreCase)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_FindText]:   " & $_bTextExists & @CRLF & _
                            @TAB & "FindStr:   " & $_sFindStr & @CRLF & _
                            @TAB & "IgnoreCase:   " & $_bIgnoreCase & @CRLF & _
                            @TAB & "RowVal, ColumnVal: " & $_iRowVal & "/" & $_iColumnVal & @CRLF & @CRLF)
            EndIf

        Return $_bTextExists
    EndFunc   ;==>__RocketBZ_FindText

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_FoundTextRow($_oRocketBZHost)
        Description ...: Returns the row position of the beginning of the string found by the FindText method.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-FoundTextRow.htm
        Parameters.....: -

        Returns........: Returns the row position of the beginning of the string found by the FindText method.
                         This method is Base-0 for VT/6530 sessions, and Base-1 for all other session types.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_FoundTextRow($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $iRetVal = $_oRocketBZHost.FoundTextRow()
            If $__ROCKETBZ_CURRENT_HOSTTYPE = $__ROCKETBZ_MAINFRAME_DISPLAY Then $iRetVal = $iRetVal+1

            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_FoundTextRow]:   " & $iRetVal & @CRLF & @CRLF)
            EndIf

        Return $iRetVal
    EndFunc   ;==>__RocketBZ_FoundTextRow

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_FoundTextColumn($_oRocketBZHost)
        Description ...: Returns the column position of the beginning of the string found by the FindText method.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-FoundTextColumn.htm
        Parameters.....: -

        Returns........: Returns the column position of the beginning of the string found by the FindText method.
                         This method is Base-0 for VT/6530 sessions, and Base-1 for all other session types.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce ===============================================================================================================================
    Func __RocketBZ_FoundTextColumn($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $iRetVal = $_oRocketBZHost.FoundTextColumn()
            If $__ROCKETBZ_CURRENT_HOSTTYPE = $__ROCKETBZ_MAINFRAME_DISPLAY Then $iRetVal = $iRetVal+1

            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_FoundTextColumn]:   " & $iRetVal & @CRLF & @CRLF)
            EndIf

        Return $iRetVal
    EndFunc   ;==>__RocketBZ_FoundTextColumn

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_PSSearch($_oRocketBZHost, $_sSearchStr, ByRef $_iRow, ByRef $_iCol, $_iStartPosVal = 1)
        Description ...: Performs a case-sensitive search for an occurrence of text in the host screen.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-PSSearch.htm
        Parameters.....:

        Returns........: The position in the host screen where the text was found, or 0 if not found.
        Remarks........: The screen position starts at 1 in the upper-left corner of the window (row 1, column 1),
                         and ends at the bottom-right of the window (max row times max column).
                         For example, for a Model 2 - 24 x 80 screen, the last position is 1920.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-06
    #ce ===============================================================================================================================
    Func __RocketBZ_PSSearch($_oRocketBZHost, $_sSearchStr, ByRef $_iRow, ByRef $_iCol, $_iStartPosVal = 1)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.PSSearch($_sSearchStr, $_iStartPosVal)
            If $_iReturnCode > 0 Then
                $_iRow = Floor($_iReturnCode / 80)
                $_iCol = $_iReturnCode - ($_iRow * 80)
                $_iRow += 1
            EndIf

            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_Search]:   " & $_iReturnCode & @CRLF & _
                            "SearchStr:  " & $_sSearchStr & @CRLF & _
                            "StartPosVal:   " & $_iStartPosVal & @CRLF & _
                            "Row/Col:  " & $_iRow & "/" & $_iCol & @CRLF & @CRLF)
            EndIf

        Return $_iReturnCode
    EndFunc   ;==>__RocketBZ_PSSearch

    #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_LockKeyboard($_oRocketBZHost, $_bLock = False)
        Description ...: Used to lock or unlock the BlueZone session's keyboard.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-LockKeyboard.htm
        Parameters.....: {$_bLock} True to lock keybord, False to unlock

        Returns........: 0 for success; or a non-zero error code. Refer to Error codes for a complete listing of error code descriptions.
        Remarks........: Does not display 'keyboard locked' icon on application.

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-06
    #ce ===============================================================================================================================
    Func __RocketBZ_LockKeyboard($_oRocketBZHost, $_bLock = False)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.LockKeyboard($_bLock)
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_LockKeyboard]:   " & $_iReturnCode & @CRLF & _
                            @TAB & "Lock:" & $_bLock & @CRLF & @CRLF)
            EndIf

        Return SetError($_iReturnCode, 2)
    EndFunc   ;==>__RocketBZ_LockKeyboard

        #cs #FUNCTION# ====================================================================================================================
        Name...........: __RocketBZ_Status($_oRocketBZHost)
        Description ...: Returns the status of the host session.

        Documentation..: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_obj-Status.htm
        Parameters.....: -

        Returns........: 0 - Ready
                         4 - Presentation Space is busy
                         5 - Keyboard is locked

        Remarks........: While the Status method can return the current session status,
                         the WaitReady method is the preferred way of waiting for a Ready session status after sending an AID-key to the host.
                         Not recommended for VT or 6530 sessions.

        Author ........: exorcistas@github.com
        Modified.......: 2020-10-12
    #ce ===============================================================================================================================
    Func __RocketBZ_Status($_oRocketBZHost)
        If Not IsObj($_oRocketBZHost) Then Return SetError($__ROCKETBZ_ERR_NOTCONNECTED, 1, -1)

        Local $_iReturnCode = $_oRocketBZHost.Status()
            If $__ROCKETBZ_DEBUG Then
                ConsoleWrite("[__RocketBZ_Status]:   " & $_iReturnCode & @CRLF & @CRLF)
            EndIf

        Return $_iReturnCode
    EndFunc   ;==>__RocketBZ_Status

#EndRegion CORE
