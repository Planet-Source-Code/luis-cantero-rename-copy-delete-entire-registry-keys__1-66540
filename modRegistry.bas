Attribute VB_Name = "modRegistry"
'Author: Luis Cantero
'© 2002-2006 L.C. Enterprises
'http://LCen.com

Option Explicit

Public Enum ROOT_KEYS
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum REG_ERROR
    ERROR_NONE = 0
    ERROR_BADDB = 1
    ERROR_BADKEY = 2
    ERROR_CANTOPEN = 3
    ERROR_CANTREAD = 4
    ERROR_CANTWRITE = 5
    ERROR_OUTOFMEMORY = 6
    ERROR_INVALID_PARAMETER = 7
    ERROR_ACCESS_DENIED = 8
    ERROR_INVALID_PARAMETERS = 87
    ERROR_NO_MORE_ITEMS = 259
    ERROR_MORE_DATA = 234
End Enum

Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const BUFFER_SIZE As Long = 255

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Enum REG_LTYPES
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_MULTI_SZ = 7
End Enum

Private Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegQueryValueExString Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Function CopyKey(MainKeySource As ROOT_KEYS, SubKeySource As String, MainKeyDest As Long, SubKeyDest As String, blnCopySubKeys As Boolean) As Boolean

  Dim hSourceKey As Long
  Dim hDestKey As Long
  Dim intRetValue As Integer '0 = Not set, 1 = True, 2 = False

  Dim strKeyName As String, lngKeyIndex As Long
  Dim hTempKey As Long, lngKeyRet As Long
  Dim filFT As FILETIME
  Dim strValueName As String, lngValueIndex As Long
  Dim lngValueRet As Long
  Dim lngValueType As Long, strValueData As String, lngValueDataRet As Long
  Dim arrTempValue() As Byte
  Dim lngTempDword As Long
  Dim i As Integer

    On Error GoTo Problems

    'Fails on recursion
    Call RegCreateKeyEx(MainKeyDest, SubKeyDest, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hDestKey, lngKeyRet)
    If Not KeyExists(MainKeyDest, SubKeyDest) Then intRetValue = 2 'Error

    If Not intRetValue = 2 Then 'Destination key exists
        'Open source
        If RegOpenKeyEx(MainKeySource, SubKeySource, 0, KEY_ALL_ACCESS, hSourceKey) = ERROR_NONE Then 'Open OK
            'Create a buffer
            strValueName = Space$(BUFFER_SIZE)
            lngValueRet = BUFFER_SIZE
            strValueData = Space$(BUFFER_SIZE)
            lngValueDataRet = BUFFER_SIZE

            'Enumerate the values
            While RegEnumValue(hSourceKey, lngValueIndex, strValueName, lngValueRet, 0, lngValueType, ByVal strValueData, lngValueDataRet) <> ERROR_NO_MORE_ITEMS
                strValueName = Trim$(left$(strValueName, lngValueRet))

                If strValueName <> "" Then 'Value found
                    Select Case lngValueType
                      Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                        strValueData = Trim$(left$(strValueData, lngValueDataRet - 1)) 'Trim terminating null

                        'Copy and check if successful
                        If SetKeyValue(MainKeyDest, SubKeyDest, strValueName, strValueData, lngValueType) <> ERROR_NONE Then intRetValue = 2 'Error

                      Case REG_BINARY
                        strValueData = left$(strValueData, lngValueDataRet)

                        'Convert to byte array
                        ReDim arrTempValue(Len(strValueData) - 1)
                        For i = 1 To Len(strValueData)
                            arrTempValue(i - 1) = CByte(Asc(Mid$(strValueData, i, 1)))
                        Next i

                        'Copy and check if successful
                        If SetKeyValue(MainKeyDest, SubKeyDest, strValueName, arrTempValue, lngValueType) <> ERROR_NONE Then intRetValue = 2 'Error

                      Case REG_DWORD
                        strValueData = left$(strValueData, lngValueDataRet)
                        'Copy 4 Bytes to the long variable
                        Call CopyMemory(lngTempDword, ByVal strValueData, Len(strValueData) + 1)

                        'Copy and check if successful
                        If SetKeyValue(MainKeyDest, SubKeyDest, strValueName, lngTempDword, lngValueType) <> ERROR_NONE Then intRetValue = 2 'Error

                    End Select
                End If

                'Prepare for the next value
                lngValueIndex = lngValueIndex + 1
                strValueName = Space$(BUFFER_SIZE)
                lngValueRet = BUFFER_SIZE
                strValueData = Space$(BUFFER_SIZE)
                lngValueDataRet = BUFFER_SIZE
            Wend

            'Create a buffer
            strKeyName = Space$(BUFFER_SIZE)
            lngKeyRet = BUFFER_SIZE

            'Enumerate the keys
            While RegEnumKeyEx(hSourceKey, lngKeyIndex, strKeyName, lngKeyRet, ByVal 0&, vbNullString, 0&, filFT) <> ERROR_NO_MORE_ITEMS
                strKeyName = Trim$(left$(strKeyName, lngKeyRet))

                If strKeyName <> "" Then 'Key found
                    Call RegCreateKeyEx(MainKeyDest, SubKeyDest & "\" & strKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hTempKey, lngKeyRet)
                    'Close temp dest key
                    Call RegCloseKey(hTempKey)

                    If blnCopySubKeys Then 'Recursion
                        If Not CopyKey(MainKeySource, SubKeySource & "\" & strKeyName, MainKeyDest, SubKeyDest & "\" & strKeyName, blnCopySubKeys) Then intRetValue = 2 'Error
                    End If
                End If

                'Prepare for the next key
                lngKeyIndex = lngKeyIndex + 1
                strKeyName = Space$(BUFFER_SIZE)
                lngKeyRet = BUFFER_SIZE
            Wend

            'Close source
            Call RegCloseKey(hSourceKey)

            'Set to true only if no errors occurred yet
            If intRetValue = 0 Then intRetValue = 1 'OK
          Else 'Error ocurred
            intRetValue = 2 'Error
        End If
    End If

    'Close destination
    Call RegCloseKey(hDestKey)

    'Return
    Select Case intRetValue
      Case 0, 2
        CopyKey = False

      Case 1
        CopyKey = True

    End Select

Exit Function

Problems:
    MsgBox Err.Description & " (CopyKey)", vbExclamation, "Error " & Err.number

End Function

Public Function CreateNewKey(MainKey As ROOT_KEYS, SubKey As String) As REG_ERROR

  Dim hNewKey As Long
  Dim lngRetVal As Long

    On Error GoTo Problems

    Call RegCreateKeyEx(MainKey, SubKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lngRetVal)
    Call RegCloseKey(hNewKey)

    'Return
    CreateNewKey = lngRetVal

Exit Function

Problems:
    MsgBox Err.Description & " (CreateNewKey)", vbExclamation, "Error " & Err.number

End Function

'Deletes single or entire subkeys.
'IMPORTANT: Information about the security measure controlled by "intMinimumPathDepth" can be found below.
Public Function DeleteKey(MainKey As ROOT_KEYS, ByVal SubKey As String, Optional blnIncludeSubkeys As Boolean = False, Optional intMinimumPathDepth As Integer = 3) As REG_ERROR

    On Error GoTo Problems

  Dim hKey As Long
  Dim strKeyName As String
  Dim lngKeyRet As Long
  Dim filFT As FILETIME

    'Try to delete key
    DeleteKey = RegDeleteKey(MainKey, SubKey)

    'IMPORTANT: Delete subkeys if wanted and if the path has a minimum length,
    'this is a security measure to avoid accidentally deleting important subkeys such as "Software".
    'A MinimumPathDepth of "1" when deleting "Software" will delete "Software" and all its subkeys.
    'A MinimumPathDepth of "2" when deleting "Software\Microsoft" will delete "Microsoft" and all its subkeys,
    'nothing shorter.
    'A MinimumPathDepth of "3" (default setting) when deleting "Software\Microsoft\Windows"
    'will delete "Windows" and all its subkeys, nothing shorter.
    'This feature only affects keys with subkeys, empty keys will be deleted regardless of their depth.
    If DeleteKey <> ERROR_NONE And blnIncludeSubkeys And GetPathDepth(SubKey) >= intMinimumPathDepth Then
        'Open source
        If RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey) = ERROR_NONE Then  'Open OK
            'Create a buffer
            strKeyName = Space$(BUFFER_SIZE)
            lngKeyRet = BUFFER_SIZE

            'Enumerate the keys (Index = 0 to always delete first item)
            While RegEnumKeyEx(hKey, 0, strKeyName, lngKeyRet, ByVal 0&, vbNullString, 0&, filFT) <> ERROR_NO_MORE_ITEMS
                strKeyName = Trim$(left$(strKeyName, lngKeyRet))

                If strKeyName <> "" Then 'Key found, try to delete children
                    'Recursion
                    Call DeleteKey(MainKey, SubKey & "\" & strKeyName, blnIncludeSubkeys)
                End If

                'Prepare for the next key
                strKeyName = Space$(BUFFER_SIZE)
                lngKeyRet = BUFFER_SIZE
            Wend

            'Close source
            Call RegCloseKey(hKey)

            'Try again now that all children have been deleted
            DeleteKey = RegDeleteKey(MainKey, SubKey)
        End If
    End If

Exit Function

Problems:
    MsgBox Err.Description & " (DeleteKey)", vbExclamation, "Error " & Err.number

End Function

Public Function DeleteValue(MainKey As ROOT_KEYS, SubKey As String, strValueName As String) As REG_ERROR

  Dim hKey As Long
  Dim lngRetVal As Long

    On Error GoTo Problems

    lngRetVal = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey)

    If lngRetVal = ERROR_NONE Then
        lngRetVal = RegDeleteValue(hKey, strValueName)
        Call RegCloseKey(hKey)
    End If

    'Return
    DeleteValue = lngRetVal

Exit Function

Problems:
    MsgBox Err.Description & " (DeleteValue)", vbExclamation, "Error " & Err.number

End Function

'Examples: "Software" = 1, "Software\Microsoft" = 2
Private Function GetPathDepth(strSubKey As String) As Integer

  Dim arrTemp() As String

    arrTemp = Split(strSubKey, "\")

    'Return
    GetPathDepth = UBound(arrTemp) + 1

End Function

Public Function KeyCount(MainKey As ROOT_KEYS, SubKey As String) As Long

  Dim filFT As FILETIME
  Dim hKey As Long
  Dim lngRetVal As Long
  Dim lngCounter As Long
  Dim strKeyName As String, strClassName As String
  Dim lngKeyLen As Long, lngClassLen As Long

    On Error GoTo Problems

    lngRetVal = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey)

    If lngRetVal = ERROR_NONE Then
        Do
            strKeyName = Space$(BUFFER_SIZE)
            lngKeyLen = BUFFER_SIZE
            strClassName = Space$(BUFFER_SIZE)
            lngClassLen = BUFFER_SIZE

            lngRetVal = RegEnumKeyEx(hKey, lngCounter, strKeyName, lngKeyLen, ByVal 0&, strClassName, lngClassLen, filFT)
            lngCounter = lngCounter + 1
        Loop While lngRetVal = ERROR_NONE

        Call RegCloseKey(hKey)
    End If

    'Return
    KeyCount = lngCounter - 1 'Adjust count

Exit Function

Problems:
    MsgBox Err.Description & " (KeyCount)", vbExclamation, "Error " & Err.number

End Function

Public Function KeyExists(MainKey As ROOT_KEYS, SubKey As String) As Boolean

  Dim hKey As Long
  Dim lngRetVal As Long

    On Error GoTo Problems

    lngRetVal = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey)

    If lngRetVal = ERROR_NONE Then
        Call RegCloseKey(hKey)

        KeyExists = True
      Else
        KeyExists = False
    End If

Exit Function

Problems:
    MsgBox Err.Description & " (KeyExists)", vbExclamation, "Error " & Err.number

End Function

Public Function QueryValue(MainKey As ROOT_KEYS, SubKey As String, strValueName As String, lngType As REG_LTYPES) As Variant

  Dim hKey As Long
  Dim lngRetVal As Long
  Dim varRetrievedValue As Variant

    lngRetVal = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey)

    If lngRetVal = ERROR_NONE Then
        lngRetVal = QueryValueEx(hKey, strValueName, varRetrievedValue, lngType)
        Call RegCloseKey(hKey)
    End If

    'Return
    QueryValue = varRetrievedValue

End Function

Private Function QueryValueEx(ByVal lngKeyHandle As Long, ByVal strValueName As String, varRetrievedValue As Variant, lngType As REG_LTYPES) As REG_ERROR

  Dim lngDataRet As Long
  Dim lngRetVal As Long
  Dim lngValue As Long
  Dim strValue As String

    ReDim bData(0) As Byte

    On Error GoTo Problems

    Select Case lngType
      Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
        lngRetVal = RegQueryValueExNULL(lngKeyHandle, strValueName, 0&, lngType, 0&, lngDataRet)
        strValue = String$(lngDataRet, 0)
        lngRetVal = RegQueryValueExString(lngKeyHandle, strValueName, 0&, lngType, strValue, lngDataRet)

        If lngRetVal = ERROR_NONE Then
            varRetrievedValue = left$(strValue, lngDataRet)
          Else
            varRetrievedValue = Empty
        End If

      Case REG_BINARY
        lngRetVal = RegQueryValueEx(lngKeyHandle, strValueName, 0&, lngType, bData(0), lngDataRet)
        If lngRetVal = ERROR_NONE Or lngRetVal = ERROR_MORE_DATA Then
            ReDim bData(0 To lngDataRet - 1)
            lngRetVal = RegQueryValueEx(lngKeyHandle, strValueName, CLng(0), lngType, bData(0), lngDataRet)
        End If

        varRetrievedValue = bData

      Case REG_DWORD
        lngRetVal = RegQueryValueExNULL(lngKeyHandle, strValueName, 0&, lngType, 0&, lngDataRet)
        lngRetVal = RegQueryValueExLong(lngKeyHandle, strValueName, 0&, lngType, lngValue, lngDataRet)

        If lngRetVal = ERROR_NONE Then varRetrievedValue = lngValue

      Case Else
        varRetrievedValue = -1

    End Select

ExitThisFunction:
    If right$(varRetrievedValue, 1) = Chr$(0) Then
        varRetrievedValue = left$(varRetrievedValue, Len(varRetrievedValue) - 1)
    End If

    'varRetrievedValue contains the retrieved value
    'lngRetVal contains the return value
    QueryValueEx = lngRetVal

Exit Function

Problems:
    MsgBox Err.Description & " (QueryValue)", vbExclamation, "Error " & Err.number
    Resume ExitThisFunction

End Function

Public Function RenameKey(MainKeySource As ROOT_KEYS, SubKeySource As String, MainKeyDest As Long, SubKeyDest As String, blnCopySubKeys As Boolean, Optional intMinimumPathDepth As Integer = 3) As Boolean

  'Copy old key to new one

    If CopyKey(MainKeySource, SubKeySource, MainKeyDest, SubKeyDest, True) Then
        'If successful copying, delete old key recursively
        If DeleteKey(MainKeySource, SubKeySource, True, intMinimumPathDepth) = ERROR_NONE Then RenameKey = True
    End If

End Function

Public Function SetKeyValue(MainKey As ROOT_KEYS, SubKey As String, strValueName As String, ValueSetting As Variant, lngType As REG_LTYPES) As REG_ERROR

  Dim lngValue As Long
  Dim strValue As String
  Dim hKey As Long
  Dim lngRetVal As Long
  Dim lngLength As Long
  Dim i As Integer
  Dim bData() As Byte
  Dim bData2() As Byte

    On Error GoTo Problems

    lngRetVal = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey)

    If lngRetVal = ERROR_NONE Then
        Select Case lngType
          Case REG_SZ, REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
            strValue = ValueSetting & Chr$(0)

            SetKeyValue = RegSetValueExString(hKey, strValueName, 0&, lngType, strValue, Len(strValue))

          Case REG_BINARY 'Free form binary
            lngLength = (UBound(ValueSetting) - LBound(ValueSetting)) + 1
            ReDim bData(LBound(ValueSetting) To UBound(ValueSetting))
            ReDim bData2(LBound(ValueSetting) To UBound(ValueSetting))

            If TypeName(ValueSetting) = "Byte()" Then 'Already a byte array
                bData = ValueSetting
              Else
                For i = LBound(ValueSetting) To UBound(ValueSetting)
                    bData(i) = CByte(ValueSetting(i))
                Next i
            End If

            SetKeyValue = RegSetValueEx(hKey, strValueName, 0&, lngType, bData(LBound(ValueSetting)), lngLength)

          Case REG_DWORD
            lngValue = ValueSetting

            SetKeyValue = RegSetValueExLong(hKey, strValueName, 0&, lngType, lngValue, 4)

          Case Else
            SetKeyValue = -1

        End Select

        Call RegCloseKey(hKey)
      Else 'Error opening key
        SetKeyValue = lngRetVal
    End If

Exit Function

Problems:
    MsgBox Err.Description & " (SetKeyValue)", vbExclamation, "Error " & Err.number

End Function

Public Function ValueCount(MainKey As ROOT_KEYS, SubKey As String) As Long

  Dim hKey As Long
  Dim lngRetVal As Long
  Dim lngCounter As Long
  Dim lngType As Long
  Dim strValueName As String, Valuelen As Long
  Dim strData As String, lngDatalen As Long

    On Error GoTo Problems

    lngRetVal = RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey)

    If lngRetVal = ERROR_NONE Then
        Do
            strValueName = Space$(BUFFER_SIZE)
            Valuelen = BUFFER_SIZE
            strData = Space$(BUFFER_SIZE)
            lngDatalen = BUFFER_SIZE

            lngRetVal = RegEnumValue(hKey, lngCounter, strValueName, Valuelen, 0, lngType, strData, lngDatalen)
            lngCounter = lngCounter + 1
        Loop While lngRetVal = ERROR_NONE

        Call RegCloseKey(hKey)
    End If

    'Return
    ValueCount = lngCounter - 1 'Adjust count

Exit Function

Problems:
    MsgBox Err.Description & " (ValueCount)", vbExclamation, "Error " & Err.number

End Function

':) Ulli's VB Code Formatter V2.13.6 (2006-09-08 21:30:17) 68 + 490 = 558 Lines
