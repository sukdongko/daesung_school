Attribute VB_Name = "basTNS"
'Option Explicit
'Global DB As New ADODB.Connection
'Global SQL  As String
'Global L As Long
'Global M As Long
'Global N As Long
'
''############################
''# tnsnames.ora ¼¼ÆÃ
''############################
'
'Private Const TNS_Path1 = "C:\oracle\instantclient_11_2_0_3\network\ADMIN\tnsnames.ora"
'Private Const TNS_Path2 = "C:\ORACLE\instantclient_11_2\network\ADMIN\tnsnames.ora"
'Private Const TNS_Path3 = "C:\oracle\product\11.2.0\client_1\Network\Admin\tnsnames.ora"
'Private Const TNS_Path4 = "C:\oracle\ora81\network\ADMIN\tnsnames.ora"
'Private Const TNS_Path5 = "C:\oracle\ora92\network\ADMIN\tnsnames.ora"
'
''Private Const TNS = "DAC=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=ms.mimacstudy.com)(PORT=30074))(LOAD_BALANCE=no))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=DS)(INSTANCE_NAME=DS2)(FAILOVER_MODE=(TYPE=SELECT)(METHOD=BASIC)(RETRIES=180)(DELAY=5))))"
''Private Const DB_Name = "DAC"
'
'Public Const S = vbNewLine ' Chr(13) & Chr(10)
'
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
'            (ByVal lpApplicationName As String, _
'             ByVal lpKeyName As Any, _
'             ByVal lpDefault As String, _
'             ByVal lpReturnedString As String, _
'             ByVal nSize As Long, _
'             ByVal lpFileName As String) _
'             As Long
'
'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
'            (ByVal lpApplicationName As String, _
'             ByVal lpKeyName As Any, _
'             ByVal lpString As Any, _
'             ByVal lpFileName As String) As Long
'
'Public Function GetINIValue(asINIFileName As String, asApplicationName As String, asKeyName As String) As String
'    Dim sReturn         As String * 100
'
'    Call GetPrivateProfileString(asApplicationName, asKeyName, "", sReturn, 100, asINIFileName)
'    GetINIValue = Left(sReturn, InStr(sReturn, Chr(0)) - 1)
'End Function
'
'Public Function SetINIValue(asINIFileName As String, asApplicationName As String, asKeyName As String, asKeyValue As String) As Boolean
'    Call WritePrivateProfileString(asApplicationName, asKeyName, asKeyValue, asINIFileName)
'End Function
'
'
'
'
