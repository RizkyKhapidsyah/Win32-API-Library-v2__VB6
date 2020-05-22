VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Constants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************
'
' Global module containing all the constants and enums needed
' by the DLL objects, all are public with the exception of
' ErrorConstants, which is for internal use only
'
' OWNER: Neil Stevens
' CREATED: 30/01/2003
' HISTORY:
'   1.0.0 - Created
'
'*************************************************************
Option Explicit

'*** FOR INTERNAL USE ONLY
Public Enum ErrorConstants
    eNoError = 0
    eErrorInvalidFunction = 1&
    eErrorFileNotFound = 2&
    eErrorPathNotFound = 3&
    eErrorTooManyOpenFiles = 4&
    eErrorAccessDenied = 5&
    eErrorInvalidHandle = 6&
    eErrorArenaTrashed = 7&
    eErrorNotEnoughMemory = 8&
    eErrorInvalidBlock = 9&
    eErrorBadEnvironment = 10&
    eErrorBadFormat = 11&
    eErrorInvalidAccess = 12&
    eErrorInvalidData = 13&
    eErrorInvalidDrive = 15&
    eErrorCurrentDirectory = 16&
    eErrorNotSameDevice = 17&
    eErrorNoMoreFiles = 18&
    eErrorWriteProtect = 19&
    eErrorBadUnit = 20&
    eErrorNotReady = 21&
    eErrorBadCommand = 22&
    eErrorCRC = 23&
    eErrorBadLength = 24&
    eErrorSeek = 25&
    eErrorNotDosDisk = 26&
    eErrorSectorNotFound = 27&
    eErrorOutOfPaper = 28&
    eErrorWriteFault = 29&
    eErrorReadFault = 30&
    eErrorGenFailure = 31&
    eErrorSharingViolation = 32&
    eErrorLockViolation = 33&
    eErrorWrongDisk = 34&
    eErrorFCBUnavailable = 35&
    eErrorSharingBufferExceeded = 36&
    eErrorNotSupported = 50&
    eErrorUnexpNetErr = 59&
    eErrorBadDevType = 66&
    eErrorBadNetName = 67&
    eErrorFileExists = 80&
    eErrorDUPFCB = 81&
    eErrorCannotMake = 82&
    eErrorFailI24 = 83&
    eErrorOutOfStructures = 84&
    eErrorAlreadyAssigned = 85&
    eErrorInvalidPassword = 86&
    eErrorInvalidParameter = 87&
    eErrorNetWriteFault = 88&
    eErrorNoProcSlots = 89&
    eErrorNotFrozen = 90&
    eErrorTSTOVFL = 91&
    eErrorTSTDUP = 92&
    eErrorNoItems = 93&
    eErrorInterrupt = 95&
    eErrorTooManySemaphores = 100&
    eErrorExclSemAlreadyOwned = 101&
    eErrorSemIsSet = 102&
    eErrorTooManySemRequests = 103&
    eErrorInvalidAtInterruptTime = 104&
    eErrorSEMOwnerDied = 105&
    eErrorSEMUsageLimit = 106&
    eErrorDiskChange = 107&
    eErrorDriveLocked = 108&
    eErrorBrokenPipe = 109&
    eErrorOpenFailed = 110&
    eErrorBufferOverflow = 111&
    eErrorDiskFull = 112&
    eErrorNoMoreSearchHandles = 113&
    eErrorInvalidTargetHandle = 114&
    eErrorProtectionViolation = 115&
    eErrorVIOKBDRequest = 116&
    eErrorInvalidCategory = 117&
    eErrorInvalidVerifySwitch = 118&
    eErrorBadDriverLevel = 119&
    eErrorCallNotImplemented = 120&
    eErrorSEMTimeout = 121&
    eErrorInsufficientBuffer = 122&
    eErrorInvalidName = 123&
    eErrorInvalidLevel = 124&
    eErrorNoVolumeLabel = 125&
    eErrorModNotFound = 126&
    eErrorProcNotFound = 127&
    eErrorWaitNoChildren = 128&
    eErrorChildNotComplete = 129&
    eErrorDirectAccessHandle = 130&
    eErrorNegativeSeek = 131&
    eErrorSeekOnDevice = 132&
    eErrorIsJoinTarget = 133&
    eErrorIsJoined = 134&
    eErrorIsSubsted = 135&
    eErrorNotJoined = 136&
    eErrorNotSubsted = 137&
    eErrorJoinToJoin = 138&
    eErrorSubstToSubst = 139&
    eErrorJoinToSubst = 140&
    eErrorSubstToJoin = 141&
    eErrorBusyDrive = 142
    eErrorSameDrive = 143
    eErrorDirNotRoot = 144
    eErrorDirNotEmpty = 145
    eErrorIsSubstPath = 146
    eErrorIsJoinPath = 147
    eErrorPathBusy = 148
    eErrorIsSubstTarget = 149
    eErrorSystemTrace = 150
    eErrorInvalidEventCount = 151
    eErrorTooManyMuxwaiters = 152
    eErrorInvalidListFormat = 153
    eErrorLabelTooLong = 154
    eErrorTooManyTCBS = 155
    eErrorSignalRefused = 156
    eErrorDiscarded = 157
    eErrorNotLocked = 158
    eErrorBadThreadIDAddr = 159
    eErrorBadArguments = 160
    eErrorBadPathname = 161
    eErrorSignalPending = 162
    eErrorUncertainMedia = 163
    eErrorMaxThrdsReached = 164
    eErrorMonitorsNotSupported = 165
    eErrorBusy = 170&
    eErrorInvalidSegmentNumber = 180
    eErrorInvalidCallgate = 181
    eErrorInvalidOrdinal = 182
    eErrorAlreadyExists = 183
    eErrorNoChildProcess = 184
    eErrorChildAliveNowait = 185
    eErrorInvalidFlagNumber = 186
    eErrorSemNotFound = 187
    eErrorInvalidStartingCodeseg = 188
    eErrorInvalidStackseg = 189
    eErrorInvalidModuleType = 190
    eErrorInvalidExeSignature = 191
    eErrorExeMarkedInvalid = 192
    eErrorBadExeFormat = 193
    eErrorIteratedDataExceeds64k = 194
    eErrorInvalidMinAllocSize = 195
    eErrorDynLinkFromInvalidRing = 196
    eErrorIOPLNotEnabled = 197
    eErrorInvalidSegDPL = 198
    eErrorAutoDataSegExceeds64k = 199
    eErrorRing2SegMustBeMovable = 200
    eErrorRelocChainXeedsSeglim = 201
    eErrorInfLoopInRelocChain = 202
    eErrorEnvVarNotFound = 203
    eErrorNotCurrentCtry = 204
    eErrorNoSignalSent = 205
    eErrorFilenameExcedRange = 206
    eErrorRing2StackInUse = 207
    eErrorMetaExpansionTooLong = 208
    eErrorInvalidSignalNumber = 209
    eErrorThread1Inactive = 210
    eErrorInfoNotAvail = 211
    eErrorLocked = 212
    eErrorBadDynalink = 213
    eErrorTooManyModules = 214
    eErrorNestingNotAllowed = 215
    eErrorMoreData = 234
    eErrorNoMoreItems = 259&
    eErrorInvalidAddress = 487&
    eErrorBadDevice = 1200&
    eErrorConnectionUnavail = 1201&
    eErrorDeveiceAlreadyRemembered = 1202&
    eErrorNoNetOrBadPath = 1203&
    eErrorBadProvider = 1204&
    eErrorCannotOpenProfile = 1205&
    eErrorBadProfile = 1206&
    eErrorNotContainer = 1207&
    eErrorExtendedOrder = 1208&
    eErrorNoNetwork = 1222&
    eErrorBadUsername = 2202&
    eErrorNotConnected = 2250&
    eErrorOpenFiles = 2401&
    eErrorDeviceInUse = 2404&
    eErrorKeyNotFound = &H2
End Enum


'*** Public constants and enums
' General public constants
Public Enum GeneralConstants
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2

    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOACTIVATE = &H10
    SWP_SHOWWINDOW = &H40
    TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

    MAX_COMPUTERNAME_LENGTH = &HF
    MAX_PATH = (&HFF + 5)
    MAX_USERNAME_LENGTH = (&HFF + 1)

    WM_USER = &H400

    NONESSENTIAL = &H0
    NOHANDLE = &H0
    Default = &H80000000
End Enum

' Computer and User constants
Public Enum CSIDL
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_LOCAL_APPDATA = &H1B
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
End Enum

Public Enum CSIDL_FOLDERPATH
    SHGFP_TYPE_CURRENT = &H0
    SHGFP_TYPE_DEFAULT = &H1
End Enum

' System tray constants
Const NOTIFYICON_OLDVERSION = 0
Const NOTIFYICON_VERSION = 3

Public Enum TrayStyleConstants
    tsSimple = &H0
    tsBalloon = &H1
End Enum

Public Enum TrayMessageIconConstants
    tmiNone = &H0
    tmiInfo = &H1
    tmiWarning = &H2
    tmiError = &H3
End Enum

Public Enum MouseMessageConstants
    mmMouseOver = &H200
    mmLeftButtonDown = &H201
    mmLeftButtonUp = &H202
    mmLeftButtonDoubleClick = &H203
    mmRightButtonUp = &H204
    mmRightButtonDown = &H205
    mmRightButtonDoubleClick = &H206
    mmMiddleButtonUp = &H207
    mmMiddleButtonDown = &H208
    mmMiddleButtonDoubleClick = &H209
End Enum

Public Enum TrayMessageConstants
    tmAdd = &H0
    tmModify = &H1
    tmDelete = &H2
    tmSetFocus = &H4
    tmSetVersion = &H8
End Enum

Public Enum TrayIconFlags
    tiMessage = &H1
    tiIcon = &H2
    tiTip = &H4
    tiState = &H8
    tiInfo = &H10
End Enum

' Registry constants
Const SYNCHRONIZE = &H100000

Public Enum RegistryOptionsConstants
    roRegOptionNonVolatile = 0
    roRegOptionVolatile = 1
End Enum

Public Enum RegistryRightsConstants
    rrReadControl = &H20000
    rrStandardRightsRead = (rrReadControl)
    rrStandardRightsWrite = (rrReadControl)
    rrStandardRightsAll = &H1F0000
End Enum

Public Enum RegistryDataTypeConstants
    dtRegSz = &H1
    dtRegExpandSz = &H2
    dtRegBinary = &H3
    dtRegDword = &H4
    dtRegMultiSz = &H7
End Enum

Public Enum RegistryKeyConstants
    rkKeyQueryValue = &H1
    rkKeySetValue = &H2
    rkKeyCreateSubKey = &H4
    rkKeyEnumerateSubKeys = &H8
    rkKeyNotify = &H10
    rkKeyCreateLink = &H20
    rkKeyRead = ((rrStandardRightsRead Or rkKeyQueryValue Or rkKeyEnumerateSubKeys Or rkKeyNotify) And (Not SYNCHRONIZE))
    rkKeyWrite = ((rrStandardRightsWrite Or rkKeySetValue Or rkKeyCreateSubKey) And (Not SYNCHRONIZE))
    rkKeyExecute = (rkKeyRead)
    rkKeyAllAccess = ((rrStandardRightsAll Or rkKeyQueryValue Or rkKeySetValue Or rkKeyCreateSubKey Or rkKeyEnumerateSubKeys Or rkKeyNotify Or rkKeyCreateLink) And (Not SYNCHRONIZE))
End Enum

Public Enum RegistryHKeyConstants
    hKeyClassesRoot = &H80000000
    hKeyCurrentUser = &H80000001
    hKeyLocalMachine = &H80000002
    hKeyUsers = &H80000003
    hKeyPerformanceData = &H80000004
    hKeyCurrentConfig = &H80000005
    hKeyDynData = &H80000006
End Enum

' Network constants
Public Enum ResourceTypeConstants
    rtResourceTypeAny = &H0&
    rtResourceTypeDisk = &H1&
    rtResourceTypePrint = &H2&
    rtResourceTypeUnknown = &HFFFF&
End Enum

Public Enum ResourceDisplayTypeConstants
    rdResourceDisplayTypeGeneric = &H0&
    rdResourceDisplayTypeDomain = &H1&
    rdResourceDisplayTypeServer = &H2&
    rdResourceDisplayTypeShare = &H3&
    rdResourceDisplayTypeFile = &H4&
    rdResourceDisplayTypeGroup = &H5&
End Enum

Public Enum ResourceUsageConstants
    ruResourceUsageAll = &H0&
    ruResourceUsageConnectable = &H1&
    ruResourceUsageContainer = &H2&
    ruResourceUsageReserved = &H80000000
End Enum

Public Enum NetworkResourceConstants
    nrResourceConnected = &H1&
    nrResourceGlobalNet = &H2&
    nrResourceRemembered = &H3&
End Enum

Public Enum NetworkErrorConstants
    wnAccessDenied = eErrorAccessDenied
    wnAlreadyConnected = eErrorAlreadyAssigned
    wnBadHandle = eErrorInvalidHandle
    wnBadLocalName = eErrorBadNetName
    wnBadPassword = eErrorInvalidPassword
    wnBadPointer = eErrorInvalidAddress
    wnBadProfile = eErrorBadProfile
    wnBadProvider = eErrorBadProvider
    wnBadUser = eErrorBadUsername
    wnBadValue = eErrorInvalidParameter
    wnBadDevType = eErrorBadDevType
    wnCannotOpenProfile = eErrorCannotOpenProfile
    wnConnectionClosed = eErrorConnectionUnavail
    wnDeviceError = eErrorGenFailure
    wnDeviceInUse = eErrorDeviceInUse
    wnExtendedError = eErrorExtendedOrder
    wnFunctionBusy = eErrorBusy
    wnMoreData = eErrorMoreData
    wnNetError = eErrorUnexpNetErr
    wnNoError = eNoError
    wnNoMoreEntries = eErrorNoMoreItems
    wnNoNetOrBadPath = eErrorNoNetOrBadPath
    wnNoNetwork = eErrorNoNetwork
    wnNotConnected = eErrorNotConnected
    wnNotContainer = eErrorNotContainer
    wnNotSupported = eErrorNotSupported
    wnOpenFiles = eErrorOpenFiles
    wnOutOfMemory = eErrorNotEnoughMemory
    wnSuccess = eNoError
    wnWindowsError = eErrorUnexpNetErr
End Enum

Public Enum ConnectOptionConstants
    coConnectUpdateProfile = &H1
    coConnectUpdateRecent = &H2
    coConnectTemporary = &H4
    coConnectInteractive = &H8
    coConnectPrompt = &H10
    coConnectNeedDrive = &H20
    coConnectRefCount = &H40
    coConnectRedirect = &H80
    coConnectLocalDrive = &H100
    coConnectCurrentMedia = &H200
    coConnectDeferred = &H400
End Enum

' Internet connection constants
Public Enum InternetConnectionConstants
    icInternetConnectionModem = &H1
    icInternetConnectionLan = &H2
    icInternetConnectionProxy = &H4
    icInternetConnectionModemBusy = &H8
    icInternetRasInstalled = &H10
    icInternetConnectionOffline = &H20
    icInternetConnectionConfigured = &H40
End Enum

' Event Log constants
Public Enum EventLogTypeConstants
    etEventLogSuccess = &H0
    etEventLogErrorType = &H1
    etEventLogWarningType = &H2
    etEventLogInfomationType = &H4
    etEventLogAuditType = &H8
    etEventLogAuditFailure = &H10
End Enum

Public Enum EventLogCategoryConstants
    catDevices = 1
    catDisk = 2
    catPrinters = 3
    catServices = 4
    catShell = 5
    catSystemEvent = 6
    catNetwork = 7
End Enum

Public Enum EventLogConstants
    elApplication
    elSecurity
    elSystem
End Enum

' OS constants
Public Enum OperatingSystemConstants
    osWindows95 = &H0
    osWindows98 = &H1
    osWindowsME = &H2
    osWindowsNT351 = &H3
    osWindowsNT400 = &H4
    osWindows2000 = &H5
    osWindowsXP = &H6
    osWindowsNetServer = &H7
    osUnknown = &H8
End Enum

Public Enum MajorVersionNumberConstants
    mnMajorVersionWin2000 = &H5
    mnMajorVersionWin95 = &H4
    mnMajorVersionWin98 = &H4
    mnMajorVersionWinME = &H4
    mnMajorVersionWinNetServer = &H5
    mnMajorVersionWinNT351 = &H3
    mnMajorVersionWinNT400 = &H4
    mnMajorVersionWinXP = &H5
End Enum

Public Enum MinorVersionNumberConstants
    mnMinorVersionWin2000 = &H0
    mnMinorVersionWin95 = &H0
    mnMinorVersionWin98 = &HA
    mnMinorVersionWinME = &H5A
    mnMinorVersionWinNetServer = &H1
    mnMinorVersionWinNT351 = &H33
    mnMinorVersionWinNT400 = &H0
    mnMinorVersionWinXP = &H1
End Enum

' Browse For Folder constants
Public Enum BrowseForFolderConstants
      bffReturnOnlyFSDirs = &H1
      bffDontGoBelowDomain = &H2
      bffStatusText = &H4
      bffReturnFSAncesters = &H8
      bffEditBox = &H10
      bffBrowseForComputer = &H1000
      bffBrowseForPrinter = &H2000
      bffBrowseIncludeFiles = &H4000
End Enum

' Windows socket constants
Public Enum WSAErrorConstants
    WSABASEERR = 10000
    WSAEINTR = (WSABASEERR + 4)
    WSAEBADF = (WSABASEERR + 9)
    WSAEACCESS = (WSABASEERR + 13)
    WSAEFAULT = (WSABASEERR + 14)
    WSAEINVAL = (WSABASEERR + 22)
    WSAEMFILE = (WSABASEERR + 24)
    WSAEWOULDBLOCK = (WSABASEERR + 35)
    WSAEINPROGRESS = (WSABASEERR + 36)
    WSAEALREADY = (WSABASEERR + 37)
    WSAENOTSOCK = (WSABASEERR + 38)
    WSAEDESTADDRREQ = (WSABASEERR + 39)
    WSAEMSGSIZE = (WSABASEERR + 40)
    WSAEPROTOTYPE = (WSABASEERR + 41)
    WSAENOPROTOOPT = (WSABASEERR + 42)
    WSAEPROTONOSUPPORT = (WSABASEERR + 43)
    WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
    WSAEOPNOTSUPP = (WSABASEERR + 45)
    WSAEPFNOSUPPORT = (WSABASEERR + 46)
    WSAEAFNOSUPPORT = (WSABASEERR + 47)
    WSAEADDRINUSE = (WSABASEERR + 48)
    WSAEADDRNOTAVAIL = (WSABASEERR + 49)
    WSAENETDOWN = (WSABASEERR + 50)
    WSAENETUNREACH = (WSABASEERR + 51)
    WSAENETRESET = (WSABASEERR + 52)
    WSAECONNABORTED = (WSABASEERR + 53)
    WSAECONNRESET = (WSABASEERR + 54)
    WSAENOBUFS = (WSABASEERR + 55)
    WSAEISCONN = (WSABASEERR + 56)
    WSAENOTCONN = (WSABASEERR + 57)
    WSAESHUTDOWN = (WSABASEERR + 58)
    WSAETOOMANYREFS = (WSABASEERR + 59)
    WSAETIMEDOUT = (WSABASEERR + 60)
    WSAECONNREFUSED = (WSABASEERR + 61)
    WSAELOOP = (WSABASEERR + 62)
    WSAENAMETOOLONG = (WSABASEERR + 63)
    WSAEHOSTDOWN = (WSABASEERR + 64)
    WSAEHOSTUNREACH = (WSABASEERR + 65)
    WSAENOTEMPTY = (WSABASEERR + 66)
    WSAEPROCLIM = (WSABASEERR + 67)
    WSAEUSERS = (WSABASEERR + 68)
    WSAEDQUOT = (WSABASEERR + 69)
    WSAESTALE = (WSABASEERR + 70)
    WSAEREMOTE = (WSABASEERR + 71)
    WSASYSNOTREADY = (WSABASEERR + 91)
    WSAVERNOTSUPPORTED = (WSABASEERR + 92)
    WSANOTINITIALIZED = (WSABASEERR + 93)
    WSAEDISCON = (WSABASEERR + 101)
    WSAENOMORE = (WSABASEERR + 102)
    WSAECANCELLED = (WSABASEERR + 103)
    WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
    WSAEINVALIDPROVIDER = (WSABASEERR + 105)
    WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
    WSASYSCALLFAILED = (WSABASEERR + 107)
    WSASERVICENOTFOUND = (WSABASEERR + 108)
    WSATYPENOTFOUND = (WSABASEERR + 109)
    WSA_E_NO_MORE = (WSABASEERR + 110)
    WSA_E_CANCELLED = (WSABASEERR + 111)
    WSAEREFUSED = (WSABASEERR + 112)
End Enum

Public Enum WinsockConstants
    INADDR_NONE = &HFFFF
    SOCKET_ERROR = -1
    INVALID_SOCKET = -1
    INADDR_ANY = &H0
    
    FD_SETSIZE = 64
    SOCKADDR_SIZE = 16
    
    MAX_GET_HOST_STRUCT = 1024
    
    SD_RECEIVE = &H0
    SD_SEND = &H1
    SD_BOTH = &H2
    
    MSG_OOB = &H1
    MSG_PEEK = &H2
    MSG_DONTROUTE = &H4
    MSG_PARTIAL = &H8000
    
    FD_READ = &H1&
    FD_WRITE = &H2&
    FD_OOB = &H4&
    FD_ACCEPT = &H8&
    FD_CONNECT = &H10&
    FD_CLOSE = &H20&
    
    SOL_SOCKET = 65535
    
    SO_DEBUG = &H1&
    SO_ACCEPT_CONN = &H2&
    SO_REUSE_ADDR = &H4&
    SO_KEEP_ALIVE = &H8&
    SO_DONT_ROUTE = &H10&
    SO_BROADCAST = &H20&
    SO_USE_LOOP_BACK = &H40&
    SO_LINGER = &H80&
    SO_OOB_INLINE = &H100&
    SO_DONT_LINGER = Not SO_LINGER
    SO_EXCLUSIVE_ADDR_REUSE = Not SO_REUSE_ADDR
    SO_SNDBUF = &H1001&
    SO_RCVBUF = &H1002&
    SO_ERROR = &H1007&
    SO_TYPE = &H1008&
    
    WSA_DESCRIPTION_LEN = 257
    WSA_SYS_STATUS_LEN = 129
    
    WSAHOSTNOTFOUND = (WSABASEERR + 1001)
    WSATRYAGAIN = (WSABASEERR + 1002)
    WSANORECOVERY = (WSABASEERR + 1003)
    WSANODATA = (WSABASEERR + 1004)
    
    GWL_WNDPROC = -4
    
    GMEM_FIXED = &H0
    GMEM_MOVEABLE = &H2
End Enum

Public Enum SocketTypeConstants
    stSockStream = 1
    stSockDGram = 2
    stSockRaw = 3
    stSockRDM = 4
    stSockSeqPacket = 5
End Enum

Public Enum AddressFamilyConstants
    afUNSPECIFIED = 0
    afUNIX = 1
    afINET = 2
    afIMPLINK = 3
    afPUP = 4
    afCHAOS = 5
    afNS = 6
    afIPX = afNS
    afISO = 7
    afOSI = afISO
    afECMA = 8
    afDATAKIT = 9
    afCCITT = 10
    afSNA = 11
    afDECNET = 12
    afDLI = 13
    afLAT = 14
    afHYLINK = 15
    afAPPLETALK = 16
    afNETBIOS = 17
    afVOICEVIEW = 18
    afFIREFOX = 19
    afUNKNOWN1 = 20
    afBAN = 21
    afATM = 22
    afINET6 = 23
    afCLUSTER = 24
    af12844 = 25
    afMAX = 26
End Enum

Public Enum ProtocolConstants
    IPPROTO_IP = 0
    IPPROTO_ICMP = 1
    IPPROTO_IGMP = 2
    IPPROTO_GGP = 3
    IPPROTO_TCP = 6
    IPPROTO_PUP = 12
    IPPROTO_UDP = 17
    IPPROTO_IDP = 22
    IPPROTO_ND = 77
    IPPROTO_RAW = 255
    IPPROTO_MAX = 256
End Enum

Public Enum WinsockProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
End Enum

Public Enum WinsockErrorConstants
    sckAddressInUse = 10048
    sckAddressNotAvailable = 10049
    sckAlreadyComplete = 10037
    sckAlreadyConnected = 10056
    sckBadState = 40006
    sckConnectAborted = 10053
    sckConnectionRefused = 10061
    sckConnectionReset = 10054
    sckGetNotSupported = 394
    sckHostNotFound = 11001
    sckHostNotFoundTryAgain = 11002
    sckInProgress = 10036
    sckInvalidArg = 40014
    sckInvalidArgument = 10014
    sckInvalidOp = 40020
    sckInvalidPropertyValue = 380
    sckMsgTooBig = 10040
    sckNetReset = 10052
    sckNetworkSubsystemFailed = 10050
    sckNetworkUnreachable = 10051
    sckNoBufferSpace = 10055
    sckNoData = 11004
    sckNonRecoverableError = 11003
    sckNotConnected = 10057
    sckNotInitialized = 10093
    sckNotSocket = 10038
    sckOpCanceled = 10004
    sckOutOfMemory = 7
    sckOutOfRange = 40021
    sckPortNotSupported = 10043
    sckSetNotSupported = 383
    sckSocketShutdown = 10058
    sckSuccess = 40017
    sckTimedout = 10060
    sckUnsupported = 40018
    sckWouldBlock = 10035
    sckWrongProtocol = 40026
End Enum

Public Enum WinsockStateConstants
    sckClosed = 0
    sckOpen
    sckListening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckConnecting
    sckConnected
    sckClosing
    sckError
End Enum

Public Enum WinsockInternalStateConstants
    istConnecting
    istSendingDatagram
End Enum

Public Enum MailFormatConstants
    mfTextPlain = 0
    mfTextHtml = 1
End Enum

Public Enum CharsetConstants
    cUsAscii = 0
End Enum

Public Enum ImportanceConstants
    iLow = 0
    iNormal = 1
    iHigh = 2
End Enum
