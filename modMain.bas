Attribute VB_Name = "modMain"
'-----------------------------------------------------------------------
' Copyright ICEnetware Ltd 2002-2005
' Module    : modMain
' Created   : 05/09/2005
' Author    : Gordon Wilmot
' Purpose   :
'-----------------------------------------------------------------------
' Dependancies :
' Assumptions  :
' Last Updated :
'-----------------------------------------------------------------------
Option Explicit

' Some background articles

' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dninvb00/html/callback.asp
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/dataexchange/datacopy/datacopyreference/datacopymessages/wm_copydata.asp
' http://share.skype.com/developer_zone/documentation/api_v1.3_documentation
' http://share.skype.com/developer_zone/developer_blog/new_api_features_for_skype_fr_windows_%281.4%29/
' http://share.skype.com/directory/learning_skypes_plug_in_architecture/view/

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

' Messaging Constants
Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_COPYDATA As Long = &H4A

' API Calls
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Skype Message Names
Public Const SkypeControlAPIDiscover As String = "SkypeControlAPIDiscover"
Public Const SkypeControlAPIAttach As String = "SkypeControlAPIAttach"
' Skype Message Handles
Public WM_SKYPECONTROLAPIDISCOVER As Long
Public WM_SKYPECONTROLAPIATTACH As Long
' Reply Status
Public Const WM_FAILURE As Integer = 0
Public Const WM_SUCCESS As Integer = 1
' Skype Attach Messages
Public Const SKYPECONTROLAPI_ATTACH_SUCCESS As Long = 0               ' Client is successfully attached and API window handle can be found in wParam parameter;
Public Const SKYPECONTROLAPI_ATTACH_PENDING_AUTHORIZATION As Long = 1 ' Skype has acknowledged connection request and is waiting for confirmation from the user.
                                                                      ' The client is not yet attached and should wait for SKYPECONTROLAPI_ATTACH_SUCCESS message;
Public Const SKYPECONTROLAPI_ATTACH_REFUSED As Long = 2               ' User has explicitly denied access to client;
Public Const SKYPECONTROLAPI_ATTACH_NOT_AVAILABLE As Long = 3         ' API is not available at the moment. For example, this happens when no user is currently logged in.

' Application Parameters
Public glSkypeHandler As Long    ' Holds Handle to send messages to Skype
Public glfrmMain As Long         ' Holds handle for frmMain
Private nlPrevWndProc As Long    ' Holds Previous Handle for windows procedure

Public Sub ConfigureSkypeMsgs()
'-----------------------------------------------------------------------
' Procedure    : modMain.ConfigureSkypeMsgs
' Author       : Gordon Wilmot
' Date Created : 05/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
On Error GoTo Catch

' Register the Skype Messages
WM_SKYPECONTROLAPIDISCOVER = RegisterWindowMessage(SkypeControlAPIDiscover)
WM_SKYPECONTROLAPIATTACH = RegisterWindowMessage(SkypeControlAPIAttach)

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: ConfigureSkypeMsgs" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub

Public Function HookCallbackProc() As Boolean
'-----------------------------------------------------------------------
' Procedure    : modMain.HookCallbackProc
' Author       : Gordon Wilmot
' Date Created : 05/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
' If the function succeeds, the return value is the previous value of the specified 32-bit integer
' If the function fails, the return value is zero.
' Unfortunately the "AddressOf" function can only be used for module level procedures
On Error GoTo Catch

nlPrevWndProc = SetWindowLong(glfrmMain, GWL_WNDPROC, AddressOf WindowProc)
If nlPrevWndProc = WM_FAILURE Then HookCallbackProc = False Else HookCallbackProc = True

Finally:
    Exit Function

Catch:
    MsgBox "Internal Error: HookCallbackProc" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Function

Public Function UnHookCallbackProc() As Boolean
'-----------------------------------------------------------------------
' Procedure    : modMain.UnHookCallbackProc
' Author       : Gordon Wilmot
' Date Created : 05/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
Dim lhHandle As Long    ' This is the handle returned by the API call

On Error GoTo Catch

' sets it back
lhHandle = SetWindowLong(glfrmMain, GWL_WNDPROC, nlPrevWndProc)
If lhHandle = WM_FAILURE Then UnHookCallbackProc = False Else UnHookCallbackProc = True

Finally:
    Exit Function

Catch:
    MsgBox "Internal Error: UnHookCallbackProc" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Function

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'-----------------------------------------------------------------------
' Procedure    : modMain.WindowProc
' Author       : Gordon Wilmot
' Date Created : 05/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
On Error GoTo Catch

Select Case uMsg
    Case WM_COPYDATA
        ProcessSkypeMsg lParam
        WindowProc = WM_SUCCESS
    Case WM_SKYPECONTROLAPIDISCOVER
        ' Don't really care 'bout this message as it's the broadcast message
        WindowProc = WM_SUCCESS
    Case WM_SKYPECONTROLAPIATTACH
        ' This is where SKYPE replies with its handle
        Select Case lParam
            Case SKYPECONTROLAPI_ATTACH_SUCCESS
                If wParam <> 0 Then glSkypeHandler = wParam
                ' Add to the display showing that it's received data
                frmMain.AddText "<- " & "SKYPE Control API Attach Success"
            Case SKYPECONTROLAPI_ATTACH_PENDING_AUTHORIZATION
                ' Add to the display showing that it's received data
                frmMain.AddText "<- " & "SKYPE Control API Attach Pending Authorization"
            Case SKYPECONTROLAPI_ATTACH_REFUSED
                ' Add to the display showing that it's received data
                frmMain.AddText "<- " & "SKYPE Control API Attach Refused"
            Case SKYPECONTROLAPI_ATTACH_NOT_AVAILABLE
                ' Add to the display showing that it's received data
                frmMain.AddText "<- " & "SKYPE Control API Attach Not Available"
        End Select
        WindowProc = WM_SUCCESS
    Case Else
        ' Just pass on the message
        WindowProc = CallWindowProc(nlPrevWndProc, hw, uMsg, wParam, lParam)
End Select

Finally:
    Exit Function

Catch:
    MsgBox "Internal Error: WindowProc" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Function

Sub ProcessSkypeMsg(ByRef lParam As Long)
'-----------------------------------------------------------------------
' Procedure    : modMain.ProcessSkypeMsg
' Author       : Gordon Wilmot
' Date Created : 05/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
Dim ltCDS As COPYDATASTRUCT   ' Holds the Pointer structure
Dim laMessage() As Byte       ' Holds the Message in a byte array
Dim lsMessage As String       ' Holds the Message as a string

On Error GoTo Catch

' Get the datastructure
CopyMemory ltCDS, ByVal lParam, Len(ltCDS)
' Get the buffer ready
ReDim laMessage(1 To ltCDS.cbData) As Byte
' Get the message into the byte buffer
CopyMemory laMessage(1), ByVal ltCDS.lpData, ltCDS.cbData
' Move it into a string
lsMessage = StrConv(laMessage, vbUnicode)
' Chop it to remove null terminators
lsMessage = Left$(lsMessage, InStr(1, lsMessage, Chr$(0)) - 1)
' Add to the display showing that it's received data
frmMain.AddText "<- " & CStr(lsMessage)

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: ProcessSkypeMsg" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub

