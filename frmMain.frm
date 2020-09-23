VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SKYPE Communication Example"
   ClientHeight    =   4080
   ClientLeft      =   5805
   ClientTop       =   4320
   ClientWidth     =   4410
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4410
   Begin VB.Frame fraCommunicate 
      Caption         =   "Communicate"
      Height          =   1050
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   4335
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   285
         Left            =   2895
         TabIndex        =   7
         Top             =   270
         Width           =   1300
      End
      Begin VB.TextBox txtSend 
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Text            =   "PING"
         Top             =   630
         Width           =   2580
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   285
         Left            =   2895
         TabIndex        =   5
         Top             =   630
         Width           =   1300
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect"
         Height          =   285
         Left            =   1485
         TabIndex        =   4
         Top             =   270
         Width           =   1300
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   1300
      End
   End
   Begin VB.Frame fraData 
      Caption         =   "Conversation with Skype"
      Height          =   2925
      Left            =   45
      TabIndex        =   0
      Top             =   1125
      Width           =   4335
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   225
         Width           =   4110
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------
' Copyright ICEnetware Ltd 2002-2005
' Module    : frmMain
' Created   : 05/09/2005
' Author    : Gordon Wilmot
' Purpose   :
'-----------------------------------------------------------------------
' Dependancies :
' Assumptions  :
' Last Updated :
'-----------------------------------------------------------------------

' This example of communication with the SKYPE API using VB6 was developed by
' ICEnetware (www.icenetware.com)
' Email: Gordon.Wilmot@icenetware.com if you have any comments

Const MAX_DISPLAY_CHARS As Integer = 10000

' Resize Variables
Private nlMinHeight As Long    ' Holds the minimum form Height
Private nlMinWidth As Long     ' Holds the minimum form Width
Private nlRightMargin As Long  ' Holds the right Margin offset
Private nlBottomMargin As Long ' Holds the bottom Margin offset

Option Explicit

Private Sub cmdConnect_Click()
'-----------------------------------------------------------------------
' Procedure    : frmMain.cmdConnect_Click
' Author       : Gordon Wilmot
' Date Created : 06/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long    ' Reply from API call
Dim llParam As Long    ' Dummy parameter for API call

On Error GoTo Catch

' Hook up the callback
If HookCallbackProc() Then

    ' Reset the Handle
    glSkypeHandler = 0
    
    ' Log the attempt
    AddText "// " & "Connecting at " & Now
    
    ' OK Search for SKYPE
    llReply = SendMessage(HWND_BROADCAST, WM_SKYPECONTROLAPIDISCOVER, Me.hwnd, llParam)
    If llReply = WM_FAILURE Then
        MsgBox "Can't Send Broadcast Message.", vbExclamation
    Else
        
        ' Reset UI
        txtSend.Enabled = True
        cmdSend.Enabled = True
        cmdDisconnect.Enabled = True
        cmdConnect.Enabled = False
        
    End If
Else
    MsgBox "Failed to Hook CallBack!", vbExclamation
End If

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: cmdConnect" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub

Private Sub cmdDisconnect_Click()
'-----------------------------------------------------------------------
' Procedure    : frmMain.cmdDisconnect_Click
' Author       : Gordon Wilmot
' Date Created : 06/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
On Error GoTo Catch

' UnHook the routine
If UnHookCallbackProc() Then
    
    ' Reset the Handle
    glSkypeHandler = 0
    
    ' Reset UI
    txtSend.Enabled = False
    cmdSend.Enabled = False
    cmdDisconnect.Enabled = False
    cmdConnect.Enabled = True
    
    ' Log the disconnect
    AddText "// " & "Disconnecting at " & Now
Else
    MsgBox "Failed to Hook CallBack!", vbExclamation
End If

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: cmdDisconnect" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub

Private Sub cmdExit_Click()
'-----------------------------------------------------------------------
' Procedure    : frmMain.cmdExit_Click
' Author       : Administrator
' Date Created : 07/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
On Error GoTo Catch

' See if we need to unhook
If glSkypeHandler <> 0 Then UnHookCallbackProc
Unload Me

Finally:
    Exit Sub

Catch:
    Resume Finally

End Sub

Private Sub cmdSend_Click()
'-----------------------------------------------------------------------
' Procedure    : frmMain.cmdSend_Click
' Author       : Gordon Wilmot
' Date Created : 06/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
Dim ltCDS As COPYDATASTRUCT   ' Holds the Pointer structure
Dim laMessage() As Byte       ' Holds the message in a byte array
Dim lsMessage As String       ' Holds the message as a string
Dim llReply As Long           ' API reply

On Error GoTo Catch

' Check if text is entered
If txtSend.Text <> vbNullString Then

    ' Get Text to send
    lsMessage = txtSend.Text & Chr$(0)
    ReDim laMessage(1 To Len(lsMessage))
    
    ' Copy the string into a byte array, converting it to ASCII
    CopyMemory laMessage(1), ByVal lsMessage, Len(lsMessage)
    
    ' Set up the structure
    ltCDS.dwData = 0
    ltCDS.cbData = UBound(laMessage)
    ltCDS.lpData = VarPtr(laMessage(1))
    
    'Send the string to SKYPE
    llReply = SendMessage(glSkypeHandler, WM_COPYDATA, Me.hwnd, ltCDS)
    If llReply = WM_FAILURE Then
        MsgBox "Can't Send Message, try re-connecting.", vbExclamation
        ' Auto disconnect
        cmdDisconnect.Value = True
    Else
        ' Just adding to display box
        AddText "-> " & CStr(txtSend.Text)
    End If
Else

End If

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: cmdSend" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub
     
Private Sub Form_Load()
'-----------------------------------------------------------------------
' Procedure    : frmMain.Form_Load
' Author       : Gordon Wilmot
' Date Created : 06/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
On Error GoTo Catch

' Configure up default values
glfrmMain = Me.hwnd
ConfigureSkypeMsgs
txtSend.Enabled = False
cmdSend.Enabled = False

' Centre the form
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

' Set Resize parameters
nlMinHeight = Me.Height
nlMinWidth = Me.Width
' Get the offsets
nlBottomMargin = Me.Height - (fraData.Top + fraData.Height)
nlRightMargin = Me.Width - (fraData.Left + fraData.Width)

' Get the caption right
Me.Caption = Me.Caption & " v" & CStr(App.Major) & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "00")

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: Form_Load" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub

Private Sub Form_Resize()
'-----------------------------------------------------------------------
' Procedure    : frmMain.Form_Resize
' Author       : Gordon Wilmot
' Date Created : 08/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
' Just re-sizes form
Dim lMeHeight As Long     ' Holds the Form height to base the resize on
Dim lMeWidth As Long      ' Holds the Form width to base the resize on

On Error GoTo Catch

Select Case Me.WindowState
    Case vbMinimized
    Case Else
        
        ' See if we're at minimums
        If Me.Height < nlMinHeight Then lMeHeight = nlMinHeight Else lMeHeight = Me.Height
        If Me.Width < nlMinWidth Then lMeWidth = nlMinWidth Else lMeWidth = Me.Width
    
        ' Form specific Code
        fraData.Width = lMeWidth - (fraData.Left + nlRightMargin)
        fraData.Height = lMeHeight - (fraData.Top + nlBottomMargin)
        fraCommunicate.Width = lMeWidth - (fraData.Left + nlRightMargin)
        
        ' Hardcoded offsets
        txtData.Height = fraData.Height - 300
        txtData.Width = fraData.Width - 200
        cmdSend.Left = fraData.Width - cmdSend.Width - 100
        cmdExit.Left = fraData.Width - cmdExit.Width - 100
        txtSend.Width = fraData.Width - cmdExit.Width - 300
        
End Select

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: Form_Resize" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub

Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------
' Procedure    : frmMain.Form_Unload
' Author       : Gordon Wilmot
' Date Created : 06/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
On Error GoTo Catch

If glSkypeHandler <> 0 Then UnHookCallbackProc

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: Form_Unload" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub

Public Sub AddText(ByVal Data As String)
'-----------------------------------------------------------------------
' Procedure    : frmMain.AddText
' Author       : Gordon Wilmot
' Date Created : 05/09/2005
'-----------------------------------------------------------------------
' Purpose      :
' Notes        :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsMessage As String   ' Holds the message to be displayed

On Error GoTo Catch

lsMessage = txtData.Text
If lsMessage = vbNullString Then
    ' If nothing just add it
    lsMessage = Data
Else
    ' Add a CRLF & then add to the end
    lsMessage = lsMessage & vbCrLf & Data
End If
' Just keep last x characters
lsMessage = Right$(lsMessage, MAX_DISPLAY_CHARS)
' Update the text box
txtData.Text = lsMessage
' Make sure the last part displayed
txtData.SelStart = Len(lsMessage)

Finally:
    Exit Sub

Catch:
    MsgBox "Internal Error: AddText" & vbCrLf & Err.Description, vbExclamation
    Resume Finally

End Sub

