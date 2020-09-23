Attribute VB_Name = "Module1"
Public TxtSave As String
Public E As String
Public E1 As String
Public E2 As String
Public E3 As String
Public E4 As String
Public E5 As String
Public In0 As Integer
Public In1 As Integer
Public In2 As Integer
Public In3 As Integer
Public In4 As Integer
Public In5 As Integer

Public Bolopen As Boolean
Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Const ERROR_SUCCESS = 0&

Public Const REG_OPTION_VOLATILE = 1
Public Const REG_OPTION_NON_VOLATILE = 0
 
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const KEY_WRITE = &H20006
Public Const KEY_READ = &H20019
Public Const KEY_ALL_ACCESS = &HF003F

Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2

