Attribute VB_Name = "privilege"
Option Explicit
Option Base 0
'本模块为提升进程权限用

'调用过程 PromotePrivileges ，提升本进程权限到最高


Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" _
   (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, _
   ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
   PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, _
   ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)



Private Const ANYSIZE_ARRAY = 1
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_DEBUG_NAME = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED = &H2

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type


Public Sub PromotePrivileges()

    Dim hToken As Long
    Dim mLUID As LUID
    Dim mPriv As TOKEN_PRIVILEGES
    Dim mNewPriv As TOKEN_PRIVILEGES
    Dim hProc As Long

    hProc = GetCurrentProcess()
    OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken
    LookupPrivilegeValue "", SE_DEBUG_NAME, mLUID
    mPriv.PrivilegeCount = 1
    mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    mPriv.Privileges(0).pLuid = mLUID
    AdjustTokenPrivileges hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)
    CloseHandle hToken

End Sub




