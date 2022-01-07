#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: RocketBZ_Extended_API_UDF
    Description...........: Extended UDF for Rocket BlueZone terminal emulator API
    Dependencies..........: Rocket Bluezone 7.1 x32/x64;
                            <RocketBZ_Globals.au3>;
                            <RocketBZ_API_Core_UDF.au3>

    Documentation.........: https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_acon_bz-host-automation-object.htm
                            https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_bz-object-methods.htm

    Author................: exorcistas@github.com
    Modified..............: 2020-12-04
    Version...............: v1.8.2b
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <RocketBZ_Globals.au3>
#include <RocketBZ_API_Core_UDF.au3>

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
%% HOST_EXTENDED %%
    _RocketBZ_Ext_AttachActiveMainframeSession()
    _RocketBZ_Ext_AttachActiveVTSession()

%% EXTENDED_CORE %%
    _RocketBZ_Ext_GetText($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
    _RocketBZ_Ext_SendKeysAndWait($_oRocketBZHost, $_sKeys, $_iRepeatTimes = 1)
    _RocketBZ_Ext_SendEnter($_oRocketBZHost, $_iRepeatTimes = 1)
    _RocketBZ_Ext_EraseField($_oRocketBZHost, $_iRow, $_iCol)
    _RocketBZ_Ext_EraseFieldAndPutString($_oRocketBZHost, $_sString, $_iRow, $_iCol, $_bEraseField = True)
    _RocketBZ_Ext_WaitScreenRefresh($_oRocketBZHost, $_sExpectedTextOnScreen = "", $_iTimeoutValue = 0)
    _RocketBZ_Ext_UnlockMainframeKeyboard($_oRocketBZHost)
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region HOST_EXTENDED

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _RocketBZ_Ext_AttachActiveMainframeSession()
        Description ...: Creates RocketBZ COM object and attaches to active Mainframe session
        Syntax.........: -
        Parameters.....: -
        Return values..: {$_oRocketBZHost} connected host object; sets @error+@extended on exception

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_AttachActiveMainframeSession()
        Local $_oRocketBZHost = __RocketBZ_CreateObject($__RocketBZ_MAINFRAME_DISPLAY)
        If @error Then Return SetError(@error, @extended)

        Local $_iReturnCode = __RocketBZ_Connect($_oRocketBZHost)
        If @error Then Return SetError(@error, @extended)

        _RocketBZ_Ext_WaitScreenRefresh($_oRocketBZHost)
        If @error Then Return SetError(@error, @extended)

        Return $_oRocketBZHost
    EndFunc   ;==>_RocketBZ_Ext_AttachActiveMainframeSession

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _RocketBZ_Ext_AttachActiveVTSession()
        Description ...: Creates RocketBZ COM object and attaches to active VT session
        Syntax.........: -
        Parameters.....: -
        Return values..: {$_oRocketBZHost} connected host object; sets @error+@extended on exception

        Author ........: exorcistas@github.com
        Modified.......: 2020-12-04
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_AttachActiveVTSession()
        Local $_oRocketBZHost = __RocketBZ_CreateObject($__ROCKETBZ_VT_DISPLAY)
        If @error Then Return SetError(@error, @extended)

        Local $_iReturnCode = __RocketBZ_Connect($_oRocketBZHost)
        If @error Then Return SetError(@error, @extended)

        _RocketBZ_Ext_WaitScreenRefresh($_oRocketBZHost)
        If @error Then Return SetError(@error, @extended)

        Return $_oRocketBZHost
    EndFunc   ;==>_RocketBZ_Ext_AttachActiveVTSession

#EndRegion HOST_EXTENDED

#Region EXTENDED_CORE

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _RocketBZ_Ext_GetText($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
        Description ...: Reads contents from screen
        Syntax.........: -
        Parameters.....: {$_oRocketBZHost} connected host object
                        {$_iRowVal, $_iColumnVal} row and column as start position to read (min value = 1)
                        {$_iLengthVal} lenght of string to read (min value = 1)

        Return values..: {$_sText} string from screen; sets @error+@extended on exception

        Author ........: exorcistas@github.com
        Modified.......: 2020-11-11
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_GetText($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
        Local $_sText = __RocketBZ_PSGetText($_oRocketBZHost, $_iRowVal, $_iColumnVal, $_iLengthVal)
        If @error Then Return SetError(@error, @extended)

        Return $_sText
    EndFunc   ;==>_RocketBZ_Ext_GetText

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _RocketBZ_Ext_SendKeysAndWait($_oRocketBZHost, $_sKeys, $_iRepeatTimes = 1)
        Description ...: Sends keyboard signals to emulator
        Syntax.........: see function: __RocketBZ_SendKey

        Parameters.....: {$_oRocketBZHost} connected host object
                        {$_sKeys} keyboard value to send
                        {$_iRepeatTimes} times to repeat key signal

        Return values..: sets @error+@extended on exception

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_SendKeysAndWait($_oRocketBZHost, $_sKeys, $_iRepeatTimes = 1)
        For $i = 1 To $_iRepeatTimes
            _RocketBZ_Ext_WaitScreenRefresh($_oRocketBZHost)
            If @error Then Return SetError(@error, @extended)

            __RocketBZ_SendKey($_oRocketBZHost, $_sKeys)
            If @error Then Return SetError(@error, @extended)
        Next

        _RocketBZ_Ext_WaitScreenRefresh($_oRocketBZHost)
        If @error Then Return SetError(@error, @extended)
    EndFunc   ;==>_RocketBZ_Ext_SendKeysAndWait

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _RocketBZ_Ext_SendEnter($_oRocketBZHost, $_iRepeatTimes = 1)
        Description ...: Sends Enter key signal to emulator

        Syntax.........: see function: _RocketBZ_Ext_SendKeysAndWait
        Parameters.....: {$_oRocketBZHost} connected host object
            {$_iRepeatTimes} times to repeat key signal
        Return values..: sets @error+@extended on exception

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_SendEnter($_oRocketBZHost, $_iRepeatTimes = 1)
        _RocketBZ_Ext_SendKeysAndWait($_oRocketBZHost, "<Enter>", $_iRepeatTimes)
        Return SetError(@error, @extended)
    EndFunc   ;==>_RocketBZ_Ext_SendEnter

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _RocketBZ_Ext_EraseField($_oRocketBZHost, $_iRow, $_iCol)
        Description ...: Erases field contents from given position
        Syntax.........: see function: _RocketBZ_Ext_SendKeysAndWait

        Parameters.....: {$_oRocketBZHost} connected host object
                        {$_iRow, $_iCol} row, column start position of field to erase

        Return values..: sets @error+@extended on exception

        Author ........: exorcistas@github.com
        Modified.......: 2020-07-01
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_EraseField($_oRocketBZHost, $_iRow, $_iCol)
        _RocketBZ_Ext_WaitScreenRefresh($_oRocketBZHost)
        If @error Then Return SetError(@error, @extended)

        __RocketBZ_SetCursor($_oRocketBZHost, $_iRow, $_iCol)
        If @error Then Return SetError(@error, @extended)

        _RocketBZ_Ext_SendKeysAndWait($_oRocketBZHost, "<EraseEOF>")
        If @error Then Return SetError(@error, @extended)
    EndFunc   ;==>_RocketBZ_Ext_EraseField

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _RocketBZ_Ext_EraseFieldAndPutString($_oRocketBZHost, $_sString, $_iRow, $_iCol, $_bEraseField = True)
        Description ...: Clears field from given position and sets text in same position
        Syntax.........: see functions: _RocketBZ_Ext_EraseField; _RocketBZ_Ext_SendKeysAndWait

        Parameters.....: {$_oRocketBZHost} connected host object
                        {$_sString} text to set
                        {$_iRow, $_iCol} row, column start position of field to erase and set text
                        {$_bEraseField} An optional parameter to disable field erasing

        Return values..: sets @error+@extended on exception

        Author ........: exorcistas@github.com
        Modified.......: 2020-11-16
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_EraseFieldAndPutString($_oRocketBZHost, $_sString, $_iRow, $_iCol, $_bEraseField = True)
        If $_bEraseField = True Then
            _RocketBZ_Ext_EraseField($_oRocketBZHost, $_iRow, $_iCol)
            If @error Then Return SetError(@error, @extended)
        EndIf

        __RocketBZ_WriteScreen($_oRocketBZHost, $_sString, $_iRow, $_iCol)
        If @error Then Return SetError(@error, @extended)
    EndFunc   ;==>_RocketBZ_Ext_EraseFieldAndPutString

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _RocketBZ_Ext_WaitScreenRefresh($_oRocketBZHost, $_sExpectedTextOnScreen = "", $_iTimeoutValue = 0)
        Description ...: Waits for screen/application readiness, may look for specific text appear
        Syntax.........: see functions: __RocketBZ_WaitReady; _RocketBZ_Ext_WaitScreenRefresh; __RocketBZ_WaitForText
        Parameters.....:
                        {$_oRocketBZHost} connected host object
                        {$_sExpectedTextOnScreen} text to wait appear on screen
                        {$_iTimeoutValue} time to wait for specific text

        Return values..: sets @error+@extended on exception

        Author ........: exorcistas@github.com
        Modified.......: 2020-11-10
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_WaitScreenRefresh($_oRocketBZHost, $_sExpectedTextOnScreen = "", $_iTimeoutValue = 0)
        If $_sExpectedTextOnScreen <> "" Then
            __RocketBZ_WaitForText($_oRocketBZHost, $_sExpectedTextOnScreen, 1, 1, $_iTimeoutValue)
        Else
            __RocketBZ_WaitReady($_oRocketBZHost)
        EndIf

        Return SetError(@error, @extended)
    EndFunc   ;==>_RocketBZ_Ext_WaitScreenRefresh

    #cs #FUNCTION# ====================================================================================================================
    Name...........:  _RocketBZ_Ext_UnlockMainframeKeyboard($_oRocketBZHost)
    Description ...: Sends <Reset> (Left Control) key to unlock Mainframe
    Syntax.........: see: __RocketBZ_Status; _RocketBZ_Ext_SendKey
    Parameters.....: {$_oRocketBZHost} connected host object

    Return values..: sets @error+@extended on exception

    Author ........: exorcistas@github.com
    Modified.......: 2020-10-12
    #ce #FUNCTION# ====================================================================================================================
    Func _RocketBZ_Ext_UnlockMainframeKeyboard($_oRocketBZHost)
        Local $_iStatus = __RocketBZ_Status($_oRocketBZHost)
        If $_iStatus = 5 Then __RocketBZ_SendKey($_oRocketBZHost, "<Reset>")

        Return SetError(@error, @extended)
    EndFunc   ;==>_RocketBZ_Ext_UnlockMainframeKeyboard

#EndRegion EXTENDED_CORE
