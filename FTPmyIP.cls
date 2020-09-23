VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' CodeGuru FTP Class
'
' Chris Eastwood July 1999
'
' This class wraps the functionality of the Win32 WinInet.DLL
'
' It could easily be expanded to provide HTTP/Gopher and other internet
' standard file protocols.
'

Private Const MAX_PATH = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Private Const ERROR_NO_MORE_FILES = 18
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
    
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
    (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
    (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
' Initializes an application's use of the Win32 Internet functions
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
    (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
    ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' Use registry access settings.
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_INVALID_PORT_NUMBER = 0

Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H1
Private Const INTERNET_FLAG_PASSIVE = &H8000000

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
    (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
    ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long


Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

' Type of service to access.
Private Const INTERNET_SERVICE_FTP = 1
'private Const INTERNET_SERVICE_GOPHER = 2
'private Const INTERNET_SERVICE_HTTP = 3

Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_MULTIPART = &H200000

Private Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFileName As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
    
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias _
    "FtpRenameFileA" (ByVal hFtpSession As Long, _
    ByVal sExistingName As String, _
    ByVal sNewName As String) As Boolean
   
' Closes a single Internet handle or a subtree of Internet handles.
Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer


'
' Our Defined Errors
'
Public Enum errFtpErrors
    errCannotConnect = vbObjectError + 2001
    errNoDirChange = vbObjectError + 2002
    errCannotRename = vbObjectError + 2003
    errCannotDelete = vbObjectError + 2004
    errNotConnectedToSite = vbObjectError + 2005
    errGetFileError = vbObjectError + 2006
    errInvalidProperty = vbObjectError + 2007
    errFatal = vbObjectError + 2008
End Enum

'
' File Transfer types
'
Public Enum FileTransferType
    ftAscii = FTP_TRANSFER_TYPE_ASCII
    ftBinary = FTP_TRANSFER_TYPE_BINARY
End Enum

'
' Error messages
'
Private Const ERRCHANGEDIRSTR As String = "Cannot Change Directory to %s. It either doesn't exist, or is protected"
Private Const ERRCONNECTERROR As String = "Cannot Connect to %s using User and Password Parameters"
Private Const ERRNOCONNECTION As String = "Not Connected to FTP Site"
Private Const ERRNODOWNLOAD As String = "Couldn't Get File %s from Server"
Private Const ERRNORENAME As String = "Couldn't Rename File %s"
Private Const ERRNODELETE As String = "Couldn't Delete File %s from Server"
Private Const ERRALREADYCONNECTED As String = "You cannot change this property while connected to an FTP server"
Private Const ERRFATALERROR As String = "Cannot get Connection to WinInet.dll !"

'
' Session Identifier to Windows
'
Private Const SESSION As String = "CGFtp Instance"
'
' Our INET handle
'
Private mlINetHandle As Long
'
' Our FTP Connection Handle
'
Private mlConnection As Long
'
' Standard FTP properties for this class
'
Private msHostAddress As String
Private msUser As String
Private msPassword As String
Private msDirectory As String

Private Sub Class_Initialize()
'
' Create Internet session handle
'
    mlINetHandle = InternetOpen(SESSION, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    
    If mlINetHandle = 0 Then
        mlConnection = 0
        Err.Raise errFatal, "CGFTP::Class_Initialise", ERRFATALERROR
    End If
    
    mlConnection = 0
    
End Sub


Private Sub Class_Terminate()
'
' Kill off any connection
'
    If mlConnection <> 0 Then
        InternetCloseHandle mlConnection
    End If
'
' Kill off API Handle
'
    If mlINetHandle <> 0 Then
        InternetCloseHandle mlINetHandle
    End If
    mlConnection = 0
    mlINetHandle = 0
    
End Sub

Public Property Let Host(ByVal sHostName As String)
'
' Set the Host Name - only if not connected
'
    If mlConnection <> 0 Then
        Err.Raise errInvalidProperty, "ACNFTP:Host_Let", ERRALREADYCONNECTED
    End If
    msHostAddress = sHostName
End Property

Public Property Get Host() As String
'
' Get Host Name
'
    Host = msHostAddress
End Property

Public Property Let User(ByVal sUserName As String)
'
' Set the user - only if not connected
'
    If mlConnection <> 0 Then
        Err.Raise errInvalidProperty, "CGFTP::User_Let", ERRALREADYCONNECTED
    End If
    msUser = sUserName
End Property

Public Property Get User() As String
'
' Get the user information
'
    User = msUser
End Property

Public Property Let Password(ByVal sPassword As String)
'
' Set the password - only if not connected
'
    If mlConnection <> 0 Then
        Err.Raise errInvalidProperty, "CGFTP::Password_Let", ERRALREADYCONNECTED
    End If
    msPassword = sPassword
End Property

Public Property Get Password() As String
'
' Get the password
'
    Password = msPassword
End Property

Public Property Get Directory() As String
'
' Get the directory
'
    Directory = msDirectory
End Property

Public Property Let Directory(ByVal sDirectory As String)
'
' Set the directory - only if connected
'
On Error GoTo vbErrorHandler

    Dim sError As String
    
    If Not (mlConnection = 0) Then
        RemoteChDir sDirectory
        msDirectory = sDirectory
    Else
        On Error GoTo 0
        Err.Raise errNotConnectedToSite, "CGFTP::Directory_Let", ERRNOCONNECTION
    End If

    Exit Property

vbErrorHandler:
    
    Err.Raise errNoDirChange, "CGFTP::Directory[Let]", Err.Description

End Property

Public Property Get Connected() As Boolean
'
' Are we connected to an FTP Server ? T/F
'
    Connected = (mlConnection <> 0)
End Property

Public Function Connect(Optional Host As String, _
    Optional User As String, _
    Optional Password As String) As Boolean
'
' Connect to the FTP server
'
On Error GoTo vbErrorHandler

    Dim sError As String
'
' If we already have a connection then raise an error
'
    If mlConnection <> 0 Then
        On Error GoTo 0
        Err.Raise errInvalidProperty, "CGFTP::Connect", "You are already connected to FTP Server " & msHostAddress
        Exit Function
    End If
'
' Overwrite any existing properties if they were supplied in the
' arguments to this method
'
    If Len(Host) > 0 Then
        msHostAddress = Host
    End If
    
    If Len(User) > 0 Then
        msUser = User
    End If
    
    If Len(Password) > 0 Then
        msPassword = Password
    End If

'
' Connect !
'

    If Len(msHostAddress) = 0 Then
        Err.Raise errInvalidProperty, "CGFTP::Connect", "No Host Address Specified!"
    End If
    
    mlConnection = InternetConnect(mlINetHandle, msHostAddress, INTERNET_INVALID_PORT_NUMBER, _
        msUser, msPassword, INTERNET_SERVICE_FTP, 0, 0)
'
' Check for connection errors
'
    If mlConnection = 0 Then
        sError = Replace(ERRCONNECTERROR, "%s", msHostAddress)
        On Error GoTo 0
        sError = sError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
        Err.Raise errCannotConnect, "CGFTP::Connect", sError
    End If
    
    Connect = True

    Exit Function

vbErrorHandler:

    Err.Raise Err.Number, "cFTP::Connect", Err.Description

End Function

Public Function Disconnect() As Boolean
'
' Disconnect, only if connected !
'
    If mlConnection <> 0 Then
        InternetCloseHandle mlConnection
        mlConnection = 0
    Else
        Err.Raise errNotConnectedToSite, "CGFTP::Disconnect", ERRNOCONNECTION
    End If
    msHostAddress = ""
    msUser = ""
    msPassword = ""
    msDirectory = ""
            
End Function


Public Function GetFile(ByVal ServerFileAndPath As String, ByVal DestinationFileAndPath As String, Optional TransferType As FileTransferType = ftAscii) As Boolean
'
' Get the specified file to the desired location using the specified
' file transfer type
'
    Dim bRet As Boolean
    Dim sFileRemote As String
    Dim sDirRemote As String
    Dim sFileLocal As String
    Dim sTemp As String
    Dim lPos As Long
    Dim sError As String

On Error GoTo vbErrorHandler
'
' If not connected, raise an error
'
    If mlConnection = 0 Then
        On Error GoTo 0
        Err.Raise errNotConnectedToSite, "CGFTP::GetFile", ERRNOCONNECTION
    End If
    
'
' Get the file
'
    bRet = FtpGetFile(mlConnection, ServerFileAndPath, DestinationFileAndPath, False, INTERNET_FLAG_RELOAD, TransferType, 0)
    
    If bRet = False Then
        sError = ERRNODOWNLOAD
        sError = Replace(sError, "%s", ServerFileAndPath)
        On Error GoTo 0
        GetFile = False
        Err.Raise errGetFileError, "CGFTP::GetFile", sError
    End If
    
    GetFile = True

    Exit Function

vbErrorHandler:
    GetFile = False
    Err.Raise errGetFileError, "cFTP::GetFile", Err.Description

End Function

Public Function PutFile(ByVal LocalFileAndPath As String, ByVal ServerFileAndPath As String, Optional TransferType As FileTransferType) As Boolean
    Dim bRet As Boolean
    Dim sFileRemote As String
    Dim sDirRemote As String
    Dim sFileLocal As String
    Dim sTemp As String
    Dim lPos As Long
    Dim sError As String

On Error GoTo vbErrorHandler
'
' If not connected, raise an error!
'
    If mlConnection = 0 Then
        On Error GoTo 0
        Err.Raise errNotConnectedToSite, "CGFTP::PutFile", ERRNOCONNECTION
    End If

    bRet = FtpPutFile(mlConnection, LocalFileAndPath, ServerFileAndPath, _
        TransferType, 0)
        
    If bRet = False Then
        sError = ERRNODOWNLOAD
        sError = Replace(sError, "%s", ServerFileAndPath)
        On Error GoTo 0
        PutFile = False
        sError = sError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
        Err.Raise errCannotRename, "CGFTP::PutFile", sError
    End If
    
    PutFile = True

    Exit Function

vbErrorHandler:
    Err.Raise Err.Number, "cFTP::PutFile", Err.Description

End Function

Public Function RenameFile(ByVal ExistingName As String, ByVal NewName As String) As Boolean
    Dim bRet As Boolean
    Dim sError As String

On Error GoTo vbErrorHandler
'
' If not connected, raise an error
'
    If mlConnection = 0 Then
        On Error GoTo 0
        Err.Raise errNotConnectedToSite, "CGFTP::RenameFile", ERRNOCONNECTION
    End If
   
    bRet = FtpRenameFile(mlConnection, ExistingName, NewName)
'
' Raise an error if we couldn't rename the file (most likely that
' a file with the new name already exists
'
    If bRet = False Then
        sError = ERRNORENAME
        sError = Replace(sError, "%s", ExistingName)
        On Error GoTo 0
        RenameFile = False
        sError = sError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
        Err.Raise errCannotRename, "CGFTP::RenameFile", sError
    End If
    
    RenameFile = True
    

    Exit Function

vbErrorHandler:
    Err.Raise Err.Number, "cFTP::RenameFile", Err.Description

End Function

Public Function DeleteFile(ByVal ExistingName As String) As Boolean
    Dim bRet As Boolean
    Dim sError As String

On Error GoTo vbErrorHandler
'
' Check for a connection
'
    If mlConnection = 0 Then
        On Error GoTo 0
        Err.Raise errNotConnectedToSite, "CGFTP::DeleteFile", ERRNOCONNECTION
    End If
    
    bRet = FtpDeleteFile(mlConnection, ExistingName)
'
' Raise an error if the file couldn't be deleted
'
    If bRet = False Then
        sError = ERRNODELETE
        sError = Replace(sError, "%s", ExistingName)
        On Error GoTo 0
        Err.Raise errCannotDelete, "CGFTP::DeleteFile", sError
    End If
    
    DeleteFile = True

    Exit Function

vbErrorHandler:
    Err.Raise Err.Number, "cFTP::DeleteFile", Err.Description

End Function

Private Sub RemoteChDir(ByVal sDir As String)
On Error GoTo vbErrorHandler
'
' Remote Change Directory Command through WININET
'
    Dim sPathFromRoot As String
    Dim bRet As Boolean
    Dim sError As String
'
' Needs standard Unix Convention
'
    sDir = Replace(sDir, "\", "/")
'
' Check for a connection
'
    If mlConnection = 0 Then
        On Error GoTo 0
        Err.Raise errNotConnectedToSite, "CGFTP::RemoteChDir", ERRNOCONNECTION
        Exit Sub
    End If
    
    If Len(sDir) = 0 Then
        Exit Sub
    Else
        sPathFromRoot = sDir
        If Len(sPathFromRoot) = 0 Then
            sPathFromRoot = "/"
        End If
        bRet = FtpSetCurrentDirectory(mlConnection, sPathFromRoot)
'
' If we couldn't change directory - raise an error
'
        If bRet = False Then
            sError = ERRCHANGEDIRSTR
            sError = Replace(sError, "%s", sDir)
            On Error GoTo 0
            Err.Raise errNoDirChange, "CGFTP::ChangeDirectory", sError
        End If
    End If

    Exit Sub

vbErrorHandler:
    Err.Raise Err.Number, "cFTP::RemoteChDir", Err.Description

End Sub

Private Function GetINETErrorMsg(ByVal ErrNum As Long) As String
    Dim lError As Long
    Dim lLen As Long
    Dim sBuffer As String
'
' Get Extra Info from the WinInet.DLL
'
    If ErrNum = ERROR_INTERNET_EXTENDED_ERROR Then
'
' Get Message Size and Number
'
        InternetGetLastResponseInfo lError, vbNullString, lLen
        sBuffer = String$(lLen + 1, vbNullChar)
'
' Get Message
'
        InternetGetLastResponseInfo lError, sBuffer, lLen
        GetINETErrorMsg = vbCrLf & sBuffer
    End If
End Function


