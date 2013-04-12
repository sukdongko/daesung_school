Attribute VB_Name = "basDataBase"
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
' 5  �����ý��۸� :
'   ��   ��   �� : BASDATABASE
'   �� ��  �� �� : �����ͺ��̽� ����
'
'   ��   ��   �� : 2007/08/20
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit

'>> ODBC API ����
    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
                   (ByVal hwndParent As Long, _
                    ByVal fRequest As Long, _
                    ByVal lpszDriver As String, _
                    ByVal lpszAttributes As String) As Long
                    
    Private Const ODBC_ADD_DSN_USER = 1     ' Add data source
    Private Const ODBC_ADD_DSN_SYS = 4      ' Add data source
    Private Const ODBC_CONFIG_DSN = 2       ' Configure (edit) data source
    Private Const ODBC_REMOVE_DSN = 3       ' Remove data source
    Private Const vbAPINull As Long = 0&    ' NULL Pointer

    Private Const ODBCDrv = "Oracle ODBC Driver"
    Private Const ODBCDrv92 = "Oracle in OraHome92"             ' �뷮��
    Private Const ODBCDrv11g = "{Microsoft ODBC for Oracle}"    ' 11g
    
    Private OracleVer       As String                           ' ����Ŭ ����
    
    
    Public DBConn As ADODB.Connection
    
    
'>> REGISTRY �� �����ϴ� DECLARE STATEMENT
    Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    
    Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
    Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    
    Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

    ' ROOT Ű�� ó��
    Public Const ERROR_SUCCESS = 0&
    
    Public Const REG_OPTION_BACKUP_RESTORE = 4                      '�����̳� ������ ���� �ʿ��� �������� ������.
    Public Const REG_OPTION_VOLATILE = 1                            '�ֹ߼� �����Ͱ� �ƴϹǷ� �ý��� ���õ��� �սǵ��� �ʴ´�.
    Public Const REG_OPTION_NON_VOLATILE = &O0                      '�����ʹ� �޸𸮿��� ��������, ����ũ �󿡴� ���ϵ��� �ʴ´�.
    
    Public Const REG_BINARY = 3                                     '����������
    Public Const REG_DWORD = 4                                      '32 BIT NUMBER
    Public Const REG_DWORD_LITTLE_ENDIAN = 4                        '32 BIT NUMBER �ֻ��� ����Ʈ�� ��������Ʈ
    Public Const REG_DWORD_BIG_ENDIAN = 5                           '32 BIT NUMBER �ֻ��� ����Ʈ�� ��������Ʈ
    Public Const REG_EXPAND_SZ = 2                                  'ȯ�溯���� ���� Ȯ�� ���� ���� ���������� ���ڿ� (EX : %PATH%)
    
    Public Const REG_LINK = 6                                       '�����ڵ� �ɺ��� ����
    Public Const REG_MULTI_SZ = 7                                   '�η� ������ ���ڿ��� ����Ʈ
    Public Const REG_NONE = 0                                       '���ǵ��� ���� ����
    Public Const REG_RESOURCE_LIST = 8                              '�����̽� �����̹� �ڿ� ����
    Public Const REG_SZ = 1                                         'NULL�� ������ ���ڿ�
    
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    Public Const HKEY_CURRENT_CONFIG = &H80000005
    Public Const HKEY_DYN_DATA = &H80000006
    
    Public Const STANDARD_RIGHT_ALL = &H1F0000
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_NOTIFY = &H10
    
    Public Const KEY_CREATE_LINK = &H20
    Public Const KEY_CREATE_SUB_KEY = &H4
    
    Public Const SYNCHRONIZE = &H100000
    Public Const KEY_ALL_ACCESS = ( _
                    STANDARD_RIGHT_ALL Or _
                    KEY_QUERY_VALUE Or _
                    KEY_SET_VALUE Or _
                    KEY_CREATE_SUB_KEY Or _
                    KEY_ENUMERATE_SUB_KEYS Or _
                    KEY_NOTIFY Or _
                    KEY_CREATE_LINK _
                    ) And _
                    (Not SYNCHRONIZE)
    
    Public Type SECURITY_ATTRIBUTES
        nLength                 As Long
        lpSecurityDescriptor    As Long
        bInHeritHandle          As Long
    End Type
    
    Public Type ACL
        AclRevision             As Byte
        Sbz1                    As Byte
        AclSize                 As Integer
        AceCount                As Integer
        Sbz2                    As Integer
    End Type
    
    Public Type SECURITY_DESCRIPTOR
        Revision                As Byte
        Sbz1                    As Byte
        Control                 As Long
        Owner                   As Long
        Group                   As Long
        Sacl                    As ACL
        Dacl                    As ACL
    End Type
    
'>> SERVER ����
    Public Const PORT = "15800"
    Public Const Dev_LoginHost = "dmdev.mimacstudy.com"
    Public Const Mimac_LoginHost = "ms.mimacstudy.com"
    'Public Const PassWord = "sybaQ#12"

    Public Const hKey = HKEY_LOCAL_MACHINE
'    Public Const SubKey = "SOFTWARE\ORACLE\KEY_OraClient11g_home1\"
    
    
    Private Oracle_Pass         As String                   '<< oracle pass
    
    Public Const TNS_Path1 = "C:\oracle\instantclient_11_2_0_3\network\ADMIN\tnsnames.ora"
'    Public Const TNS_Path2 = "C:\ORACLE\instantclient_11_2\network\ADMIN\tnsnames.ora"
'    Public Const TNS_Path3 = "C:\oracle\product\11.2.0\client_1\Network\Admin\tnsnames.ora"
'    Public Const TNS_Path4 = "C:\oracle\ora81\network\ADMIN\tnsnames.ora"
'    Public Const TNS_Path5 = "C:\oracle\ora92\network\ADMIN\tnsnames.ora"
    
    

Public Function DataBase_Connection() As Boolean

    Dim bSuccess As Boolean
    Dim sTns As String
    Dim DB_Name As String
    Dim tns_Path As String
    
    On Error GoTo Error1                     'error ó��
    
    '>>>>>>>>>> tnsnames.ora���� ���� ��������
    tns_Path = Get_TNSNames_Path()
    If "" = tns_Path Then
        DataBase_Connection = False
        Exit Function
    End If
    
    
    '>>>>>>>>>>> tnsnames.ora���Ͽ� DB�������� ����.
    'tnsname.ora���Ͽ� �߰��� DB��������
    Select Case UCase(Trim(basModule.connDB))
        Case "MIMAC"
            sTns = "MI2_CLASS= (DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = ms.mimacstudy.com)(PORT = 30022)))(CONNECT_DATA = (SERVICE_NAME = DS)(INSTANCE_NAME = DS2)))"
            DB_Name = "MI2_CLASS"
        Case Else
            sTns = "DMDB =(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.0.5)(PORT = 1521))    (CONNECT_DATA =      (SERVER = DEDICATED)      (SERVICE_NAME = dm)    )  )"
            DB_Name = "DMDB"
    End Select
    
    bSuccess = writeOraTnsName(tns_Path, sTns, DB_Name)
    
    
    
    '>>>>>>>>>> DB ����
    If bSuccess Then
        DataBase_Connection = connectionOledb()           '<< ���� DB����
    Else
        MsgBox "tnsnames.ora���Ͽ� DB�������� �����߿� �����߻�", vbCritical + vbOKOnly, "������ ����"
    End If
    
    DataBase_Connection = True
    Exit Function
    
    
Error1:
    DataBase_Connection = False
    MsgBox "DatabaseConnection Error", "�����ͺ��̽� ����"

End Function



Public Function connectionOledb() As Boolean
    
    On Error GoTo ErrorADODB                     'error ó��
    
    Dim strDB As String
    
    strDB = Chr(13) & Chr(10)
    
    Set DBConn = New ADODB.Connection
    'READGAME  READGAME/eoqkrskfk7
    'MSDAORA.1.1
    'OraOLEDB.Oracle
    
        
    ' Data Source = ������ �����ͺ��̽�
    DBConn.ConnectionString = strDB       '�����ͺ��̽��� ������ �õ��մϴ�.
    DBConn.ConnectionTimeout = 5          '���� �ð����� ������ ���� ������ �ڵ����� �����ϴ�.
    'DB.Properties("Prompt") = adPromptNever   '�̰��� ADO���� �⺻ ������Ʈ �����Դϴ�.
    'DB.CursorLocation = adUseClient           'Ŀ����ġ�� Client �ʿ� �ֽ��ϴ�.
        
    DBConn.Open                                   '�����ͺ��̽��� ���ϴ�.
    'MsgBox "���� ����"
    DoEvents
'    Do While DB.State And adStateConnecting
'        DoEvents
'    Loop

    connectionOledb = True

    Exit Function
    
ErrorADODB:
    MsgBox "connectionOledb�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "�л���ȸ"

    'End
End Function





'>> �켱���� 1. DAESUNG.INI���Ͽ� PATH_ORACLE_TNS
'>> �켱���� 2. ȯ�溯�� Path   (������Ʈ������ ����)
'>> TNS���ο� tnsnames.ora������ �������� ����   (\network\admin������ �������� ����)
Public Function Get_TNSNames_Path() As String

    Dim tns_Path As String
    
    '>>>>>>>>> ����Ŭ ���� ��������
    ' ȯ�溯�� Path�� ������Ʈ������ �о��ͼ� ����Ŭ���θ� ������.
    tns_Path = Get_TnsFile_Path_Registry()
    
    
    '>>>>>>>>>> DAESUNG.INI���Ͽ��� ����Ŭ ���� ��������
    ' DAESUNG.INI���Ͽ��� tns����(PATH_ORACLE_TNS)�� ������ �װɷ� �ϰ� ������. Path�� �����ִ°ɷ� �Ѵ�.
    Dim sData               As String * 255
    Dim sTmp                As String
    
    sData = ""
    Call basModule.GetPrivateProfileString("SCHOOL", "PATH_ORACLE_TNS", "", sData, 255, App.Path & "\DAESUNG.INI")             '>> ����Ŭ �ν��Ͻ� ����
    sTmp = Trim(Replace(sData, Chr(0), "", 1, -1, vbTextCompare))
    If "" <> sTmp Then
        tns_Path = sTmp
    End If
    
    If "" = tns_Path Then
        MsgBox "����Ŭ ���μ����� ������ �ֽ��ϴ�. " & Chr(13) & "���� INI���Ͽ� PATH_ORACLE_TNS�� �߰��Ұ��� �����մϴ�.", vbCritical + vbOKOnly, "TNS���� ��������"
        Get_TNSNames_Path = ""
        Exit Function
    End If
    
    
    '>>>>>>>>>> tnsnames.ora���� ����
    ' tns_Path���ο� tnsnames.ora������ �������� ����
    ' \network\admin ������ �������� ���ϻ��� ����
    If "" = Dir(tns_Path) Then
    
        If False = Create_Tnsnames(tns_Path) Then
            Get_TNSNames_Path = ""
            MsgBox "tnsnames.ora���� ��������", "Create_Tnsnames"
            Exit Function
        End If
    End If
    
    Get_TNSNames_Path = tns_Path
End Function



'RETURN VALLUE : ã������ ���� , ��ã���� ""
'tnsnames.ora���� ���� �������� : ȯ�溯�� PATH�� ORACLE ����(������Ʈ������ �о���)�� ����
Private Function Get_TnsFile_Path_Registry() As String

    Dim RetStr      As String
    Dim hSubKey     As Long
    Dim Rtn         As Long
    Dim strLength   As Long
    Dim dType       As Long
    Dim sReturn     As String
    Dim bFindPath   As Boolean
    
    
    '>> Registry Open
    Rtn = RegOpenKeyEx(hKey, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\", 0, KEY_ALL_ACCESS, hSubKey)
    
    '���� Ű���� ���´�.
    strLength = 256
    RetStr = String(strLength, 0)

    '>> oracle path
    Rtn = RegQueryValueEx(hSubKey, "Path", 0, dType, ByVal RetStr, strLength)
    
    
    '>> PathŰ�� �����Ұ���
    bFindPath = False
    If Rtn = ERROR_SUCCESS Then
    
        '>> path���� ���ڿ����� ";"�� ���еǾ�����
        Dim strPaths
        Dim Path
        
        sReturn = Left(RetStr, strLength - 1)   '�ڿ� �������� ���ڿ��� �����Ѵ�
        strPaths = Split(sReturn, ";")
        
        'PATH���� ���ڵ��߿� oracle���ΰ� �ִ��� Ȯ��. '(network\admin ������ ������ OK)
        For Each Path In strPaths
            
            If "" <> Dir(Path & "\" & "network\admin", vbDirectory) Then
                '���ΰ� ����.
                sReturn = Path & "\network\admin\tnsnames.ora"
                bFindPath = True
                Exit For
            End If
        Next
        
        If bFindPath = False Then
            'MsgBox "ȯ�溯�� PATH���� PATH\network\admin ������ ���� ���� �ʽ��ϴ�"
         End If
    Else
        'MsgBox "������Ʈ���� �о����� ���� " & vbCrLf & "rtn : " & Rtn & "  dType :" & dType & Left(RetStr, strLength - 1)
    End If
    
    Rtn = RegCloseKey(hSubKey)
    
    If False = bFindPath Then
        sReturn = ""
    End If

    Get_TnsFile_Path_Registry = sReturn
    
End Function


'>>>>>>>>>>>> tnsnames.ora������ �������� ���� ����
Private Function Create_Tnsnames(tns_Path As String) As Boolean

    Dim strFolder
    Dim strFile
    
    strFile = Right(tns_Path, 12)
    If strFile <> "tnsnames.ora" Then
        MsgBox tns_Path & " �� tnsnames.ora������ �ƴմϴ�", vbCritical + vbOKOnly, "tnsnames ����"
        Create_Tnsnames = False
        Exit Function
    End If
    
    ' ������ �����ϴ���
    strFolder = Split(UCase(tns_Path), "NETWORK\ADMIN\")
    strFolder(0) = strFolder(0) & "NETWORK\ADMIN\"
    
    If "" = Dir(strFolder(0), vbDirectory) Then
    
        ' ���� ���� FALSE(����) ��ȯ
        MsgBox tns_Path & "�� �������� �ʽ��ϴ�."
        Create_Tnsnames = False
        Exit Function
        
    Else
        ' ������ �ִµ� ������ ����. >> ���ϻ���
        Dim FileNumber As Integer
        FileNumber = FreeFile
        
        Open tns_Path For Output As FileNumber
            ' �ؿ� FileStream.ReadAll�Ҷ� ���Ͼȿ� �ƹ� TEXT�� ������ "���� ���� �Ѿ�� �Է��Դϴ�" ���� �����߻�
            Print #FileNumber, " "
        Close FileNumber
    End If
    
    Create_Tnsnames = True
End Function



'########################################
'# ORACLE Tnsnames �ڵ� �߰�
'########################################
Public Function writeOraTnsName(tns_Path As String, sTns As String, DB_Name As String) As Boolean

    Dim FS, FileStream, OutStream
    Dim strTxt As String, arrTxt() As String
    Dim i As Integer
    Dim tns_nm_flag As Boolean
    
    
    On Error GoTo err_rtn
    
    
    '>>>>>>>>>> ���Ͽ� DBName�� �������� �߰��ؼ� ������ �������. <
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    If tns_Path <> "" Then
        
         Set FileStream = FS.OpenTextFile(tns_Path)
        
         strTxt = FileStream.ReadAll
         arrTxt = Split(strTxt, vbCrLf)
         FileStream.Close
        
         tns_nm_flag = False
         For i = 0 To UBound(arrTxt)
             If UCase(Left(arrTxt(i), Len(DB_Name))) = UCase(DB_Name) Then
                 arrTxt(i) = sTns
                 tns_nm_flag = True
             End If
         Next

         If tns_nm_flag = False Then
            ReDim Preserve arrTxt(UBound(arrTxt) + 1)
            arrTxt(UBound(arrTxt)) = sTns
         End If
         
         strTxt = Join(arrTxt, vbCrLf)
        
         Set OutStream = FS.OpenTextFile(tns_Path, 2, True)
         OutStream.Write strTxt
         OutStream.Close
         
         writeOraTnsName = True       '������ ó��
    Else
    
         Set OutStream = FS.OpenTextFile(tns_Path, 2, True)
         OutStream.Write sTns
         OutStream.Close
    
    End If

    Set FS = Nothing
   
   writeOraTnsName = True                '������ ó��
   Exit Function
   
err_rtn:
   Set FS = Nothing
   writeOraTnsName = False               '�������� ó�� (���ΰ� �������� �ʴ� ����.)
   
End Function




'Public Function Find_DB_Tnsnames_Location() As String
'    '>> Registry OPEN
'    Dim hSubKey     As Long
'    Dim dType       As Long
'    Dim strLength   As Long
'    Dim RetStr      As String
'    Dim Rtn         As Long
'
'    '>> Registry Save
'    Dim iSecurity   As SECURITY_ATTRIBUTES
'    Dim strSubkey   As String
'    Dim KeyRet      As Long
'    Dim dPosition   As Long
'    Dim iRet        As Long
'
'    Dim sTmp        As String
'    Dim sReturn     As String
'
'
'    sReturn = ""
'
'    '>> registry open
'
'    Rtn = RegOpenKeyEx(hKey, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\", 0, KEY_ALL_ACCESS, hSubKey)
'    If Rtn <> ERROR_SUCCESS Then
'        '������Ʈ���� ���������־��� �Ѵ�.
'        iRet = RegCreateKeyEx(hKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, iSecurity, KeyRet, dPosition)
'        If iRet <> ERROR_SUCCESS Then
'            MsgBox "����Ŭ�� ��ġ�Ͽ� �ֽʽÿ�.", vbExclamation + vbOKOnly, "DB Connection"
'            Rtn = RegCloseKey(hSubKey)
'
'            Find_DB_Tnsnames_Location = sReturn
'            Exit Function
'        End If
'    End If
'
'    '���� Ű���� ���´�.
'    strLength = 256
'    RetStr = String(strLength, 0)
'
'
'    '>> oracle path
'        Rtn = RegQueryValueEx(hSubKey, "Path", 0, dType, ByVal RetStr, strLength)
'
'        If Rtn = ERROR_SUCCESS And dType = REG_SZ Then
'            '�ڿ� �������� ���ڿ��� �����Ѵ�
'            sReturn = Left(RetStr, strLength - 1)
'            sReturn = sReturn & "\network\admin\tnsnames.ora"
'
'            Rtn = RegQueryValueEx(hSubKey, "ORACLE_HOME_NAME", 0, dType, ByVal RetStr, strLength)
'
'            OracleVer = "ORAHOME8"
'            If strLength > 0 Then OracleVer = Trim(Left(RetStr, strLength - 1))
'
'            If InStr(1, UCase(OracleVer), "ORAHOME8", vbTextCompare) > 0 Then OracleVer = "ORA8"
'            If InStr(1, UCase(OracleVer), "ORAHOME9", vbTextCompare) > 0 Then OracleVer = "ORA9"
'
'        Else
'            MsgBox "����Ŭ ���μ����� ������ �ֽ��ϴ�.", vbExclamation + vbOKOnly, "DB Connection"
'
'        End If
'        Rtn = RegCloseKey(hSubKey)
'
'        Find_DB_Tnsnames_Location = sReturn
'
'
'End Function
