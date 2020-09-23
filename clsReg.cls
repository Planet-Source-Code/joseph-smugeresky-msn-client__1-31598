VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

Public Enum CONST_KEYS
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum
   
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long


Public Function GetRegSetting(CKeys As CONST_KEYS, strFolder As String, strKeyDesc As String) As String
       
        Dim lngReturn As Long, lngKeyHwnd As Long
        Dim varValue As Variant
        
        lngReturn = RegOpenKeyEx(CKeys, "Software\" & strFolder, 0, KEY_ALL_ACCESS, lngKeyHwnd)
        lngReturn = GetRegSettingEx(lngKeyHwnd, strKeyDesc, varValue)
        GetRegSetting = CStr(varValue)
        RegCloseKey (lngKeyHwnd)
   
End Function

Public Function GetRegSettingEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long

        Dim cch As Long
        Dim lrc As Long
        Dim lType As Long
        Dim lValue As Long
        Dim sValue As String

        On Error GoTo QueryValueExError

        lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       
        If lrc <> ERROR_NONE Then Error 5

        Select Case lType

            Case REG_SZ:
                sValue = String(cch, 0)
                lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = Left$(sValue, cch - 1)
                Else
                    vValue = Empty
                End If
           
            Case REG_DWORD:
            
                lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
                If lrc = ERROR_NONE Then vValue = lValue
                
            Case Else
            
               lrc = -1
        
        End Select

QueryValueExExit:
       GetRegSettingEx = lrc
       Exit Function

QueryValueExError:
       Resume QueryValueExExit

End Function


Public Sub CreateRegFolder(CKeys As CONST_KEYS, strNewFolder As String)
        
        Dim lngNewKey As Long, lngReturn As Long
        
        lngReturn = RegCreateKeyEx(CKeys, "Software\" & strNewFolder, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngNewKey, lngReturn)
       
        RegCloseKey (lngNewKey)
        
End Sub

Public Sub SetRegValue(CKeys As CONST_KEYS, strFolder As String, strKeyDesc As String, vValueSetting As Variant)
        
        Dim lngReturn As Long, lngKeyHwnd As Long

        lngReturn = RegOpenKeyEx(CKeys, "Software\" & strFolder, 0, KEY_ALL_ACCESS, lngKeyHwnd)
        lngReturn = SetRegValueEx(lngKeyHwnd, strKeyDesc, REG_SZ, vValueSetting)
       
        RegCloseKey (lngKeyHwnd)
   
End Sub

Public Function SetRegValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
       
        Dim lValue As Long
        Dim sValue As String
       
        Select Case lType
           
            Case REG_SZ
               
                sValue = vValue & Chr$(0)
                SetRegValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
            
            Case REG_DWORD
               
                lValue = vValue
                SetRegValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
           
           End Select
   
End Function


