VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mvarHKey As Long
Dim mvarKeyRoot As String
Dim mvarSubKey As String
Dim Security As SECURITY_ATTRIBUTES

Dim mMem As Memory
Dim mStrings As Strings

Public Property Let hKey(ByVal vData As RegistryHKeyConstants)
    mvarHKey = vData
End Property

Public Property Get hKey() As RegistryHKeyConstants
    hKey = mvarHKey
End Property

Public Property Let KeyRoot(ByVal vData As String)
    mvarKeyRoot = vData
End Property

Public Property Get KeyRoot() As String
    KeyRoot = mvarKeyRoot
End Property

Public Property Let SubKey(ByVal vData As String)
    mvarSubKey = vData
End Property

Public Property Get SubKey() As String
    SubKey = mvarSubKey
End Property

'********************************************************************************
'
' FUNCTION: KeyExists()
'
' RETURN: Boolean
'
' PARAMS: None
'
' USE: Check to see if the registry key exists using the class properties
'          and returns True if the key exists
'
'********************************************************************************
Public Function KeyExists() As Boolean
    Dim handle As Long
    Dim ret As Long
    
    If RegOpenKeyEx(mvarHKey, mvarKeyRoot & "\" & mvarSubKey, 0, rkKeyRead, handle) Then
        KeyExists = False
        Exit Function
    End If
    KeyExists = True
End Function

'********************************************************************************
'
' FUNCTION: CreateKey()
'
' RETURN: String
'
' PARAMS: Optional KeyName As Variant
'
' USE: Create a key in the registry, if KeyName is empty then uses
'          class property, returns 0 if successful
'
'********************************************************************************
Public Function CreateKey(Optional KeyName As Variant) As String
    Dim handle As Long
    Dim disp As Long
    Dim RetVal As Long
    
    KeyName = IIf(IsMissing(KeyName), mvarSubKey, CStr(KeyName))
    RetVal = RegCreateKeyEx(mvarHKey, mvarKeyRoot & "\" & KeyName, 0, "", roRegOptionNonVolatile, rkKeyAllAccess, Security, handle, disp)
    If RetVal Then Exit Function
    RegCloseKey (handle)
    CreateKey = RetVal
End Function

'********************************************************************************
'
' FUNCTION: DeleteKey()
'
' RETURN: Long
'
' PARAMS: KeyName As String
'
' USE: Delete a key from the registry, returns 0 if successful
'
'********************************************************************************
Public Function DeleteKey(KeyName As String) As Long
    Dim RetVal As Long
    Dim handle As Long
    
    RetVal = RegDeleteKey(mvarHKey, mvarKeyRoot & "\" & KeyName)
    If RetVal Then Exit Function
    RegCloseKey (handle)
    DeleteKey = RetVal
End Function

'********************************************************************************
'
' FUNCTION: DeleteValue()
'
' RETURN: Long
'
' PARAMS: ValueName As String
'
' USE: Deletes a value in the registry, opens the subkey specified in
'          the class properties, returns 0 if successful
'
'********************************************************************************
Public Function DeleteValue(ValueName As String) As Long
    Dim handle As Long
    Dim RetVal As Long
    
    RetVal = RegOpenKeyEx(mvarHKey, mvarKeyRoot & "\" & mvarSubKey, roRegOptionNonVolatile, rkKeyAllAccess, handle)
    
    If RetVal <> 0 Then
        ' Failed to open the key
        DeleteValue = RetVal
        Exit Function
    End If
    
    DeleteValue = RegDeleteValue(handle, ValueName)
    RegCloseKey (handle)
End Function

'********************************************************************************
'
' FUNCTION: EnumRegistryKeys()
'
' RETURN: Collection
'
' PARAMS: hKey As hKey, KeyName As String
'
' USE: Enumerates the value names under a given key, returns a
'          collection of the values.
'
'********************************************************************************
Public Function EnumRegistryKeys(ByVal hKey As RegistryHKeyConstants, ByVal KeyName As String) As Collection
    Dim handle As Long
    Dim Length As Long
    Dim Index As Long
    Dim subkeyName As String
    Dim fFileTime As FILETIME
    
    ' Initialize the results collection
    Set EnumRegistryKeys = New Collection
    
    ' Open the key, exit if not found
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, rkKeyRead, handle) Then Exit Function
        ' in all cases the subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' Set the max length for a key name
        Length = 260
        subkeyName = Space$(Length)
        ' Get the nth key, exit the loop if not found
        If RegEnumKeyEx(hKey, Index, subkeyName, Length, 0, "", vbNull, fFileTime) = eErrorNoMoreItems Then Exit Do
        ' Add to the results collection
        subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
        EnumRegistryKeys.Add subkeyName, subkeyName
        ' Prepare to query for next key
        Index = Index + 1
    Loop
    
    ' Close the key, if it was opened
    If handle Then RegCloseKey (handle)
End Function

'********************************************************************************
'
' FUNCTION: EnumRegistryValues()
'
' RETURN: Collection
'
' PARAMS: hKey As hKey, KeyName As String
'
' USE: Enumerates the values under a given key, returns a collection
'          of 2-element variant array, element(0) is the value name,
'          element(1) is the value's value
'
'********************************************************************************
Public Function EnumRegistryValues(ByVal hKey As RegistryHKeyConstants, ByVal KeyName As String) As Collection
    Dim handle As Long
    Dim Index As Long
    Dim valueType As Long
    Dim Name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim Length As Long
    Dim valueInfo(0 To 1) As Variant
    Dim RetVal As Long
    Dim i As Integer
    Dim vTemp As Variant
    
    ' Initialize the results collection
    Set EnumRegistryValues = New Collection
    
    ' Open the key, exit if not found
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, rkKeyRead, handle) Then Exit Function
        ' In all cases, subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' Set the max length for the key name
        nameLen = 260
        Name = Space$(nameLen)
        ' Prepare the recieving buffer for the value
        Length = 4096
        ReDim resBinary(0 To Length - 1) As Byte
        
        ' Read the values name and data
        ' exit the loop if nothing found
        RetVal = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, _
            resBinary(0), Length)
        
        ' Enlarge the buffer if more space is needed
        If RetVal = eErrorMoreData Then
            ReDim resBinary(0 To Length - 1) As Byte
            RetVal = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, _
                valueType, resBinary(0), Length)
        End If
        ' Exit the loop if any other error occurs, typically no more values
        If RetVal Then Exit Do
        
        ' Retreive the value's name
        valueInfo(0) = Left$(Name, nameLen)
        
        ' Return a value corresponding to the value type
        Select Case valueType
            Case dtRegDword
                RtlMoveMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case dtRegSz
                ' Copy everything except the trailing null char
                If Length <> 0 Then
                    resString = Space(Length - 1)
                    RtlMoveMemory resString, resBinary(0), Length - 1
                    valueInfo(1) = resString
                Else
                    valueInfo(1) = ""
                End If
            Case dtRegExpandSz
                ' Copy everything except the trailing null char
                ' expand the environment variables to it's value
                ' Ignore a blank string
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    RtlMoveMemory resString, resBinary(0), Length - 1
                    Length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    valueInfo(1) = mStrings.TrimNull(resString)
                Else
                    valueInfo(1) = ""
                End If
            Case dtRegBinary
                ' Shrink the buffer if necessary
                If Length < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To Length - 1) As Byte
                End If
                ' Convert to display as string, i.e: 00 01 01 00 01
                For i = 0 To UBound(resBinary)
                    resString = resString & " " & Format(Trim(Hex(resBinary(i))), "0#")
                Next i
                valueInfo(1) = LTrim(resString) ' Remove any leading spaces
            Case dtRegMultiSz
                ' Copy everything except the 2 trailing null chars
                resString = Space$(Length - 2)
                RtlMoveMemory resString, resBinary(0), Length - 2
                
                ' Convert from null delimited (vbNullChar) stream of strings
                ' to a comma delimited stream of strings
                ' The list view control likes it better this way
                resString = Replace(resString, vbNullChar, ",", , , vbBinaryCompare)
                valueInfo(1) = resString
            Case Else
                ' Unsupported value type - Do Nothing
        End Select
        
        ' Add the array to the results collection
        ' The collection key is the values name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        
        Index = Index + 1
    Loop
    
    ' Close the key, if it was opened
    If handle Then RegCloseKey (handle)
End Function

'********************************************************************************
'
' FUNCTION: GetRegistryValue()
'
' RETURN: Variant
'                REG_DWORD: Long
'                REG_SZ: String
'                REG_EXPAND_SZ: String with expanded environment
'                                              variable
'                REG_BINARY: Byte Array
'                REG_MULTI_SZ: null-delimited stream of strings
'
' PARAM: ValueName As String, Optional DefaultValue As Variant
'
' USE: Read a registry value from the registry, if no value exists then
'          it returns the value of DefaultValue, to get the default value of
'          a key use KeyName = "".
'
'********************************************************************************
Public Function GetRegistryValue(ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim TestString As String
    Dim resBinary() As Byte
    Dim Length As Long
    Dim RetVal As Long
    Dim valueType As Long
    
    ' Prepare the default result
    GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(mvarHKey, mvarKeyRoot & "\" & mvarSubKey, roRegOptionNonVolatile, rkKeyRead, handle) Then
        ' Don't overwrite the default value!!
        Exit Function
    End If
    
    ' Prepare a 1k receiving resBinary
    Length = 1024
    ReDim resBinary(0 To Length - 1) As Byte
    
    ' Read the registry key
    RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
    ' If resBinary was too small, try again
    If RetVal = eErrorMoreData Then
        ' Enlarge resBinary and read the value again
        ReDim resBinary(0 To Length - 1) As Byte
        RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
    End If
    
    ' If the value is not found then exit
    If RetVal = eErrorKeyNotFound Then
        RegCloseKey (handle)
        Exit Function
    End If
    
    ' Return the value corresponding to the value type
    Select Case valueType
        Case dtRegDword
            RtlMoveMemory resLong, resBinary(0), 4
            GetRegistryValue = resLong
        Case dtRegSz
            ' Copy everything except the trailing null char
            ' Ignore blank strings
            If Length <> 0 Then
                resString = Space$(Length - 1)
                RtlMoveMemory resString, resBinary(0), Length - 1
                GetRegistryValue = resString
            End If
        Case dtRegExpandSz
            ' Copy everything except the trailing null char
            ' Expand the environment variable to its value
            ' Ignore blank strings
            If Length <> 0 Then
                resString = Space$(Length - 1)
                RtlMoveMemory resString, resBinary(0), Length - 1
                Length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                GetRegistryValue = Left$(resString, Length)
            End If
        Case dtRegBinary
            ' Resize the results resBinary
            If Length <> UBound(resBinary) + 1 Then
                ReDim Preserve resBinary(0 To Length - 1) As Byte
            End If
            GetRegistryValue = resBinary()
        Case dtRegMultiSz
            ' Copy everything except the 2 trailing null chars
            resString = Space$(Length - 2)
            RtlMoveMemory resString, resBinary(0), Length - 2
            
            ' A non-existant value for REG_MULTI_SZ will return a string of nulls
            ' with a length = 1022
            ' This is because at the beginning of the routine we define length = 1024
            ' resString = Space(length - 2) = 1022
            ' So if we trimm all nulls and are left with an empty string then
            ' the value doesn't exist so the default value is returned
            ' Set resString to a temporary variable because TrimNull will truncate it
            TestString = resString
            If Len(mStrings.TrimNull(TestString)) > 0 Then GetRegistryValue = resString
        Case Else
            ' Unsupported value type - Do Nothing
    End Select
    
    ' Close the key, if it was opened
    If handle Then RegCloseKey (handle)
End Function

'********************************************************************************
'
' FUNCTION: SetRegistryValue()
'
' RETURN: Boolean
'
' PARAMS: ValueName As String, Value As Variant, DType As DataType
'
' USE: Writes a value to the registry, if the value does not exist
'          it creates the value first. Returns True if successful
'          Use KeyName = "" for the default value
'          Supports all data types
'
'********************************************************************************
Public Function SetRegistryValue(ByVal ValueName As String, Value As Variant, DType As RegistryDataTypeConstants) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim Length As Long
    Dim RetVal As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(mvarHKey, mvarKeyRoot & "\" & mvarSubKey, roRegOptionNonVolatile, rkKeyWrite, handle) Then
        SetRegistryValue = False
        Exit Function
    End If
    
    ' Three cases, according to the data type passed
    Select Case DType
        Case dtRegDword
            lngValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegDword, lngValue, 4)
        Case dtRegSz
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegSz, strValue, Len(strValue))
        Case dtRegBinary
            binValue = Value
            Length = UBound(binValue) - LBound(binValue) + 1
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegBinary, binValue(LBound(binValue)), Length)
        Case dtRegExpandSz
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegExpandSz, ByVal strValue, Len(strValue))
        Case dtRegMultiSz
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegMultiSz, ByVal strValue, Len(strValue))
        Case Else
            ' Unsupported value type - Do Nothing
    End Select
    
    ' Close the key and signal success
    RegCloseKey (handle)
    ' Signal success if the value was written correctly
    SetRegistryValue = (RetVal = 0)
End Function

'********************************************************************************
'
' FUNCTION: ReadRemoteRegistryValue()
'
' RETURN: Variant
'
' PARAMS: sRemoteComputer As String, hKey As hKey,
'                 ValueName As String, Optional KeyPath As String
'
' USE: Read a value from a remote registry, use KeyName = "" for
'          the default value
'
'********************************************************************************
Public Function ReadRemoteRegistryValue(ByVal sComputerName As String, ByVal hKey As RegistryHKeyConstants, _
                                                            ByVal ValueName As String, Optional KeyPath As String) As Variant
    Dim handle As Long
    Dim lReturnCode, lHive, lhRemoteRegistry As Long
    Dim valueType As Long
    Dim resLong As Long
    Dim resString As String
    Dim TestString As String
    Dim resBinary() As Byte
    Dim Length As Long
    Dim RetVal As Long
    Dim RegPath As String
    
    RegPath = IIf(IsMissing(KeyPath), mvarKeyRoot & "\" & mvarSubKey, KeyPath)
    
    If RegConnectRegistry(sComputerName, hKey, lhRemoteRegistry) Then
        ReadRemoteRegistryValue = CVar("Error!")
        Exit Function
    End If
    lReturnCode = RegOpenKeyEx(lhRemoteRegistry, RegPath, 0, rkKeyAllAccess, handle)
    
    ' Prepare a 1k receiving resBinary
    Length = 1024
    ReDim resBinary(0 To Length - 1) As Byte
    
    ' Read the registry key
    RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
    ' If resBinary was too small try again
    If RetVal = eErrorMoreData Then
        ' Enlarge resBinary and read the registry key again
        ReDim resBinary(0 To Length - 1) As Byte
        RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
    End If
    
    ' If key not found then exit
    If RetVal = eErrorKeyNotFound Then
        RegCloseKey (handle)
        Exit Function
    End If
    
    ' Return a value corresponding to the value type
    Select Case valueType
        Case dtRegDword
            RtlMoveMemory resLong, resBinary(0), 4
            ReadRemoteRegistryValue = resLong
        Case dtRegSz
            ' Copy everything except the trailing null char
            ' Ignore blank strings
            If Length <> 0 Then
                resString = Space$(Length)
                RtlMoveMemory resString, resBinary(0), Length - 1
                ReadRemoteRegistryValue = resString
            End If
        Case dtRegExpandSz
            ' Copy everything except the trailing null char
            ' Expand the environment variable to its value
            ' Ignore blank strings
            If Length <> 0 Then
                resString = Space$(Length - 1)
                RtlMoveMemory resString, resBinary(0), Length - 1
                
                Length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                ReadRemoteRegistryValue = resString
            End If
        Case dtRegBinary
            If Length <> UBound(resBinary) + 1 Then
                ReDim Preserve resBinary(0 To Length - 1) As Byte
            End If
            ReadRemoteRegistryValue = resBinary()
        Case dtRegMultiSz
            ' Copy everything except the trailing 2 null chars
            resString = Space$(Length - 2)
            RtlMoveMemory resString, resBinary(0), Length - 2
            
            ' A non-existant value for REG_MULTI_SZ will return a string of nulls
            ' with a length = 1022
            ' This is because at the beginning of the routine we define length = 1024
            ' resString = Space(length - 2) = 1022
            ' So if we trimm all nulls and are left with an empty string then
            ' the value doesn't exist so the default value is returned
            ' Set resString to a temporary variable because TrimNull will truncate it
            TestString = resString
            If Len(mStrings.TrimNull(TestString)) > 0 Then ReadRemoteRegistryValue = resString
        Case Else
            ' Unsupported value type - Do Nothing
    End Select
    
    ' Close the registry key
    RegCloseKey (handle)
End Function

'********************************************************************************
'
' FUNCTION: WriteRemoteRegistryValue()
'
' RETURN: Boolean
'
' PARAMS: sRemoteComputer As String, hKey As hKey,
'                 ValueName As String, Value As Variant,
'                 DType As DataType, Optional KeyPath As String
'
' USE: Write to a remote registry, use KeyName = "" to write to
'          the default value
'
'********************************************************************************
Public Function WriteRemoteRegistryValue(ByVal sComputerName As String, ByVal hKey As RegistryHKeyConstants, ByVal ValueName As String, _
                                                            Value As Variant, DType As RegistryDataTypeConstants, Optional KeyPath As String) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim Length As Long
    Dim RetVal As Long
    Dim RegPath As String
    Dim lhRemoteRegistry As Long
    Dim lReturnCode As Long
    
    RegPath = IIf(IsMissing(KeyPath), mvarKeyRoot & "\" & mvarSubKey, KeyPath)
    
    If RegConnectRegistry(sComputerName, hKey, lhRemoteRegistry) Then
        WriteRemoteRegistryValue = False
        Exit Function
    End If
    lReturnCode = RegOpenKeyEx(lhRemoteRegistry, RegPath, 0, rkKeyAllAccess, handle)
    
    ' Three cases, according to the data type passed
    Select Case DType
        Case dtRegDword
            lngValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegDword, lngValue, 4)
        Case dtRegSz
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegSz, strValue, Len(strValue))
        Case dtRegBinary
            binValue = Value
            Length = UBound(binValue) - LBound(binValue) + 1
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegBinary, binValue(LBound(binValue)), Length)
        Case dtRegExpandSz
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegExpandSz, ByVal strValue, Len(strValue))
        Case dtRegMultiSz
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, dtRegMultiSz, ByVal strValue, Len(strValue))
        Case Else
            ' Unsupported value type - Do Nothing
    End Select
    
    ' Close the registry key and signal success
    RegCloseKey (handle)
    ' Signal success if the value was written correctly
    WriteRemoteRegistryValue = (RetVal = 0)
End Function

'********************************************************************************
'
' FUNCTION: GetIniSetting()
'
' RETURN: Long
'
' PARAMS: Section As String, Key As String, KeyType As Long,
'                 Optional default As Variant, Optional IniLocation As String
'
' USE: Retreives a value from an INI file, uses App.Path & "\" App.EXEName & ".ini"
'          if IniLocation = ""
'
'********************************************************************************
Public Function GetIniSetting(ByVal Section As String, ByVal Key As String, ByVal KeyType As RegistryDataTypeConstants, _
                                        Optional Default As Variant, Optional IniLocation As String) As Variant
    Dim lngSetting As Long
    Dim strSetting As String * 128
    Dim l, sLet
    Dim iLen As Long
    Dim Length As Long
    
    ' Initialize return value to default
    GetIniSetting = Default
    
    ' If IniLocation is missing then set to application name
    IniLocation = IIf(IsMissing(IniLocation), App.Path & "\" & App.EXEName & ".ini", IniLocation)
    Default = IIf(IsMissing(Default), "", Default)
    
    ' Case dependant on the key type
    Select Case KeyType
        Case dtRegSz
            Length = GetPrivateProfileString(Section, Key, Default, strSetting, 128, IniLocation)
            l = InStr(1, Trim(strSetting), Chr(0))
            iLen = Len(Trim(strSetting))
            GetIniSetting = CStr(Left(Trim(strSetting), (l - 1)))
        Case dtRegDword
            lngSetting = GetPrivateProfileInt(Section, Key, Default, IniLocation)
    End Select
End Function

'********************************************************************************
'
' FUNCTION: SetIniSetting()
'
' RETURN: Long
'
' PARAMS: Section As String, Key As String, KeyType As Long,
'                 Setting As Variant, Optional IniLocation As String
'
' USE: saves a value to an INI file, uses App.Path & "\" App.EXEName & ".ini"
'          if IniLocation = ""
'
'********************************************************************************
Public Function SetIniSetting(ByVal Section As String, ByVal Key As String, ByVal KeyType As RegistryDataTypeConstants, _
                                        ByVal Setting As String, Optional IniLocation As String) As Long
    ' If IniLocation is empty then set to app name
    IniLocation = IIf(IsMissing(IniLocation), App.Path & "\" & App.EXEName & ".ini", IniLocation)
    
    Select Case KeyType
        Case dtRegSz
            SetIniSetting = WritePrivateProfileString(Section, Key, Setting, IniLocation)
        Case dtRegDword
            SetIniSetting = WritePrivateProfileString(Section, Key, Setting, IniLocation)
    End Select
End Function

Private Sub Class_Initialize()
    Set mMem = New Memory
    Set mStrings = New Strings
End Sub
