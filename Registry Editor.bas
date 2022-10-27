Attribute VB_Name = "RegistryEditorModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API structures used by this program:
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

'The Microsoft Windows API constants used by this program:
Public Const REG_BINARY As Long = 3&
Public Const REG_NONE As Long = 0&
Public Const REG_OPENED_EXISTING_KEY As Long = &H2&
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_CALL_NOT_IMPLEMENTED As Long = 120&
Private Const ERROR_FILE_NOT_FOUND As Long = 2&
Private Const ERROR_INVALID_HANDLE As Long = 6&
Private Const ERROR_INVALID_PARAMETER As Long = 87&
Private Const ERROR_MORE_DATA As Long = 234&
Private Const ERROR_NO_MORE_ITEMS As Long = 259&
Private Const ERROR_SUCCESS  As Long = 0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const KEY_ALL_ACCESS As Long = &HF003F
Private Const KEY_WOW64_64KEY As Long = &H100&
Private Const REG_DWORD As Long = 4&
Private Const REG_DWORD_BIG_ENDIAN As Long = 5&
Private Const REG_EXPAND_SZ As Long = 2&
Private Const REG_LINK As Long = 6&
Private Const REG_MULTI_SZ As Long = 7&
Private Const REG_OPTION_NON_VOLATILE As Long = 0&
Private Const REG_QWORD As Long = 11&
Private Const REG_SZ As Long = 1&

'The Microsoft Windows API functions used by this program:
Private Declare Function FileTimeToSystemTime Lib "Kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKeyA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValueA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValueA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegOpenKeyExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryInfoKeyA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegSetValueExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long

'The constants and structures used by this program:

'This structure defines the hive key information.
Private Type HiveKeyStr
   KeyName As String    'The hive key's name.
   PredefinedH As Long  'The hive key's predefined handle.
End Type

'This structure defines the registry key information.
Public Type KeyStr
   KeyAccessible As Boolean   'Indicates whether the key is accessible.
   KeyClass As String         'The key's class.
   KeyDateTime As FILETIME    'The key's modification/creation date and time.
   KeyName As String          'The key's name.
End Type

'This structure defines the registry value information.
Public Type ValueStr
   ValueData As String  'The value's data.
   ValueName As String  'The value's name.
   ValueType As Long    'The value's data type.
End Type

Public Const DEFAULT_VALUE_COLOR As Long = vbCyan       'Defines the background color for displaying default values.
Public Const NO_INDEX As Long = -1&                     'Defines an empty selection.
Public Const NO_KEY As Long = 0&                        'Defines a null registry key.
Public Const PIXELS_PER_CHARACTER_X As Long = 8&        'Defines the width of a character in pixels used by the character scale mode.
Public Const ROOT_KEY As Long = -1&                     'Defines the registry's root key.
Private Const ESCAPE_CHARACTER As String = "/"           'Defines the escape character used when escaping registry values.
Private Const HKEY_CLASSES_ROOT As Long = &H80000000     'Defines the HKEY_CLASSES registry hive key handle.
Private Const HKEY_CURRENT_CONFIG As Long = &H80000005   'Defines the HKEY_CURRENT_CONFIG registry hive key handle.
Private Const HKEY_CURRENT_USER As Long = &H80000001     'Defines the HKEY_CURRENT_USER registry hive key handle.
Private Const HKEY_DYN_DATA As Long = &H80000006         'Defines the HKEY_DYN_DATA registry hive key handle.
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002    'Defines the HKEY_LOCAL_MACHINE registry hive key handle.
Private Const HKEY_PERFORMANCE_DATA As Long = &H80000004 'Defines the HKEY_PERFORMANCE registry hive key handle.
Private Const HKEY_USERS As Long = &H80000003            'Defines the HKEY_USERS registry hive key handle.
Private Const MAX_LONG_STRING As Long = &HFFFF&          'Defines the maximum length in bytes allowed for a long string.
Private Const MAX_REG_VALUE_DATA As Long = &HFFFFF       'Defines the maximum length in bytes allowed for a registry value's data.
Private Const MAX_REG_VALUE_NAME As Long = &H3FFF&       'Defines the maximum length in bytes allowed for a registry value's name.
Private Const MAX_SHORT_STRING As Long = &HFF&           'Defines the maximum length in bytes allowed for a short string.

'This procedure checks whether an error occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long, Optional CheckReturnValue As Boolean = False, Optional Ignored As Long = ERROR_SUCCESS) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear

On Error GoTo ErrorTrap

   If CheckReturnValue Then ErrorCode = ReturnValue
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored) Then
      Description = String$(MAX_LONG_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      If Not CheckReturnValue Then Message = Message & "Return value: " & CStr(ReturnValue)
      MsgBox Message, vbExclamation
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function
'This procedure closes the specified registry key.
Public Function CloseKey(KeyH As Long) As Long
On Error GoTo ErrorTrap
Dim ReturnValue As Long
   
   ReturnValue = ERROR_SUCCESS
   
   If Not GetHiveKey(, KeyH).PredefinedH = NO_KEY Then
      ReturnValue = CheckForError(RegCloseKey(KeyH), CheckReturnValue:=True)
   End If
   
EndRoutine:
   CloseKey = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure asks the user a yes/no question and returns whether the answer was a "yes".
Public Function Confirmed(Question As String) As Boolean
On Error GoTo ErrorTrap
   Confirmed = (MsgBox(Question, vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes)
EndRoutine:
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure creates the specified registry key.
Public Function CreateKey(Key As KeyStr, ParentKeyH As Long, Optional ByRef Disposition As Long) As Long
On Error GoTo ErrorTrap
Dim KeyH As Long
Dim ReturnValue As Long

   ReturnValue = CheckForError(RegCreateKeyExA(ParentKeyH, Key.KeyName, CLng(0), Key.KeyClass, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SecurityAttributes(), KeyH, Disposition), CheckReturnValue:=True)
   If Not KeyH = NO_KEY Then CloseKey KeyH
   
EndRoutine:
   CreateKey = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure deletes the specified registry key.
Public Function DeleteKey(KeyName As String, ParentKeyH As Long) As Long
On Error GoTo ErrorTrap
Dim ReturnValue As Long

   ReturnValue = CheckForError(RegDeleteKeyA(ParentKeyH, KeyName), CheckReturnValue:=True)
   
EndRoutine:
   DeleteKey = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure deletes the specified registry value.
Public Function DeleteValue(ValueName As String, ParentKeyH As Long) As Long
On Error GoTo ErrorTrap
Dim ReturnValue As Long

   ReturnValue = CheckForError(RegDeleteValueA(ParentKeyH, ValueName), CheckReturnValue:=True)
   
EndRoutine:
   DeleteValue = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure converts non-displayable characters in the specified text to escape sequences.
Public Function Escape(Text As String, Optional EscapeAll As Boolean = False) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Escaped As String
Dim Index As Long
Dim NextCharacter As String

   Escaped = vbNullString
   Index = 1
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If EscapeAll Then
         Escaped = Escaped & ESCAPE_CHARACTER & String$(2 - Len(Hex$(Asc(Character))), "0") & Hex$(Asc(Character))
      Else
         If Character = ESCAPE_CHARACTER Then
            Escaped = Escaped & String$(2, ESCAPE_CHARACTER)
         ElseIf Character = vbTab Or Character >= " " Then
            Escaped = Escaped & Character
         Else
            Escaped = Escaped & ESCAPE_CHARACTER & String$(2 - Len(Hex$(Asc(Character))), "0") & Hex$(Asc(Character))
         End If
      End If
      
      Index = Index + 1
   Loop
   
EndRoutine:
   Escape = Escaped
   Exit Function
   
ErrorTrap:
   HandleError
   Escaped = vbNullString
   Resume EndRoutine
End Function

'This procedure checks whether the return value for escape sequence procedures indicates an error.
Public Function EscapeSequenceError(ErrorAt As Long) As Boolean
On Error GoTo ErrorTrap
Dim EscapeError As Boolean

   EscapeError = (ErrorAt > 0)
   If EscapeError Then MsgBox "Bad escape sequence at character #" & CStr(ErrorAt) & ".", vbExclamation
   
EndRoutine:
   EscapeSequenceError = EscapeError
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function




'This procedure converts the specified date and time to text.
Public Function FileTimeToText(FileDateTime As FILETIME) As String
On Error GoTo ErrorTrap
Dim SystemDateTime As SYSTEMTIME
Dim Text As String

   Text = vbNullString
   If CBool(CheckForError(FileTimeToSystemTime(FileDateTime, SystemDateTime))) Then
      With SystemDateTime
         Text = Text & Format$(.wYear, "0000") & "-"
         Text = Text & Format$(.wMonth, "00") & "-"
         Text = Text & Format$(.wDay, "00") & " "
         Text = Text & Format$(.wHour, "00") & ":"
         Text = Text & Format$(.wMinute, "00") & ":"
         Text = Text & Format$(.wSecond, "00")
      End With
   End If
   
EndRoutine:
   FileTimeToText = Text
   Exit Function
   
ErrorTrap:
   HandleError
   Text = vbNullString
   Resume EndRoutine
End Function
'This procedure returns the specified registry hive key's information.
Private Function GetHiveKey(Optional Index As Long = NO_INDEX, Optional HiveKeyH As Long = NO_KEY, Optional HiveKeyName As String = vbNullString) As HiveKeyStr
On Error GoTo ErrorTrap
Dim HiveKey As HiveKeyStr
   
   HiveKey.KeyName = vbNullString
   HiveKey.PredefinedH = NO_KEY
   
   Do
      Select Case Index
         Case 0
            HiveKey.KeyName = "HKEY_CLASSES_ROOT"
            HiveKey.PredefinedH = HKEY_CLASSES_ROOT
         Case 1
            HiveKey.KeyName = "HKEY_CURRENT_CONFIG"
            HiveKey.PredefinedH = HKEY_CURRENT_CONFIG
         Case 2
            HiveKey.KeyName = "HKEY_CURRENT_USER"
            HiveKey.PredefinedH = HKEY_CURRENT_USER
         Case 3
            HiveKey.KeyName = "HKEY_DYN_DATA"
            HiveKey.PredefinedH = HKEY_DYN_DATA
         Case 4
            HiveKey.KeyName = "HKEY_LOCAL_MACHINE"
            HiveKey.PredefinedH = HKEY_LOCAL_MACHINE
         Case 4
            HiveKey.KeyName = "HKEY_PERFORMANCE_DATA"
            HiveKey.PredefinedH = HKEY_PERFORMANCE_DATA
         Case 5
            HiveKey.KeyName = "HKEY_USERS"
            HiveKey.PredefinedH = HKEY_USERS
         Case Is > 5
            HiveKey.KeyName = vbNullString
            HiveKey.PredefinedH = NO_KEY
            Exit Do
      End Select
      If HiveKeyH = NO_KEY And HiveKeyName = vbNullString Then
         Exit Do
      Else
         If Not HiveKeyH = NO_KEY Then
            If HiveKey.PredefinedH = HiveKeyH Then Exit Do
         ElseIf Not HiveKeyName = vbNullString Then
            If HiveKey.KeyName = HiveKeyName Then Exit Do
         End If
         Index = Index + 1
      End If
   Loop
   
EndRoutine:
   GetHiveKey = HiveKey
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the specified registry key.
Public Function GetKey(KeyName As String, ParentKeyH As Long) As KeyStr
On Error GoTo ErrorTrap
Dim ClassLength As Long
Dim Key As KeyStr
Dim KeyH As Long
Dim KeyClass As String
Dim KeyDateTime As FILETIME
Dim ReturnValue As Long

   KeyH = OpenKey(KeyName, ParentKeyH)
   If Not KeyH = NO_KEY Then
      KeyClass = String$(MAX_SHORT_STRING, vbNullChar)
      ClassLength = Len(KeyClass)
   
      ReturnValue = CheckForError(RegQueryInfoKeyA(KeyH, KeyClass, ClassLength, CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), KeyDateTime), CheckReturnValue:=True, Ignored:=ERROR_INVALID_PARAMETER)
   
      If ReturnValue = ERROR_INVALID_PARAMETER Then
         ReturnValue = CheckForError(RegQueryInfoKeyA(KeyH, vbNullString, CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), CLng(0), KeyDateTime), CheckReturnValue:=True, Ignored:=ERROR_INVALID_PARAMETER)
         ClassLength = 0
         KeyClass = vbNullString
      End If
      CloseKey KeyH
   
      With Key
         .KeyAccessible = Not (ReturnValue = ERROR_ACCESS_DENIED)
         .KeyClass = Left$(KeyClass, ClassLength)
         .KeyDateTime.dwHighDateTime = KeyDateTime.dwHighDateTime
         .KeyDateTime.dwLowDateTime = KeyDateTime.dwLowDateTime
         .KeyName = KeyName
      End With
   End If
   
EndRoutine:
   GetKey = Key
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the information for the specified registry key.
Public Function GetKeyInformation(KeyName As String, KeyIndex As Long) As KeyStr
On Error GoTo ErrorTrap
Dim Key As KeyStr
Dim ParentKeyH As Long
   
   ParentKeyH = WalkKeyStack()
   If Not ((Not ParentKeyH = NO_KEY) And (KeyIndex = 0)) Then
      Key = GetKey(KeyName, ParentKeyH)
   End If
   
EndRoutine:
   CloseKey ParentKeyH
   GetKeyInformation = Key
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the registry keys contained by the specified key.
Public Function GetKeys(Optional ParentKeyH As Long = ROOT_KEY) As KeyStr()
On Error GoTo ErrorTrap
Dim ClassLength As Long
Dim HiveKeyIndex As Long
Dim HiveKeyH As Long
Dim KeyClass As String
Dim KeyDateTime As FILETIME
Dim KeyH As Long
Dim KeyIndex As Long
Dim KeyName As String
Dim Keys() As KeyStr
Dim NameLength As Long
Dim ReturnValue As Long

   Erase Keys()
   
   If ParentKeyH = ROOT_KEY Then
      HiveKeyIndex = 0
      Do
         HiveKeyH = GetHiveKey(HiveKeyIndex).PredefinedH
         If HiveKeyH = NO_KEY Then Exit Do
         ReturnValue = RegOpenKeyExA(HiveKeyH, vbNullString, CLng(0), KEY_ALL_ACCESS Or KEY_WOW64_64KEY, KeyH)
         Select Case ReturnValue
            Case ERROR_SUCCESS
               CloseKey KeyH
   
               If CheckForError(SafeArrayGetDim(Keys())) = 0 Then
                  ReDim Keys(0 To 0) As KeyStr
               Else
                  ReDim Preserve Keys(LBound(Keys()) To UBound(Keys()) + 1) As KeyStr
               End If
               Keys(UBound(Keys())).KeyAccessible = True
               Keys(UBound(Keys())).KeyClass = vbNullString
               Keys(UBound(Keys())).KeyDateTime.dwHighDateTime = 0
               Keys(UBound(Keys())).KeyDateTime.dwLowDateTime = 0
               Keys(UBound(Keys())).KeyName = GetHiveKey(HiveKeyIndex).KeyName
            Case Not ERROR_CALL_NOT_IMPLEMENTED, ERROR_INVALID_HANDLE
               CheckForError ReturnValue, CheckReturnValue:=True
         End Select
         HiveKeyIndex = HiveKeyIndex + 1
      Loop
   Else
      KeyIndex = 0
      Do
         KeyClass = String$(MAX_SHORT_STRING, vbNullChar)
         KeyName = String$(MAX_SHORT_STRING, vbNullChar)
         ClassLength = Len(KeyClass)
         NameLength = Len(KeyName)
         ReturnValue = CheckForError(RegEnumKeyExA(ParentKeyH, KeyIndex, KeyName, NameLength, CLng(0), KeyClass, ClassLength, KeyDateTime), CheckReturnValue:=True, Ignored:=ERROR_NO_MORE_ITEMS)
         If ReturnValue = ERROR_NO_MORE_ITEMS Or Not ReturnValue = ERROR_SUCCESS Then Exit Do
         If CheckForError(SafeArrayGetDim(Keys())) = 0 Then
            ReDim Keys(0 To 0) As KeyStr
         Else
            ReDim Preserve Keys(LBound(Keys()) To UBound(Keys()) + 1) As KeyStr
         End If
         Keys(UBound(Keys())).KeyAccessible = Not (ReturnValue = ERROR_ACCESS_DENIED)
         Keys(UBound(Keys())).KeyClass = Left$(KeyClass, ClassLength)
         Keys(UBound(Keys())).KeyDateTime.dwHighDateTime = KeyDateTime.dwHighDateTime
         Keys(UBound(Keys())).KeyDateTime.dwLowDateTime = KeyDateTime.dwLowDateTime
         Keys(UBound(Keys())).KeyName = Left$(KeyName, NameLength)
         KeyIndex = KeyIndex + 1
      Loop
   End If
   
EndRoutine:
   GetKeys = Keys()
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the specified registry value.
Public Function GetValue(ValueName As String, ParentKeyH As Long) As ValueStr
On Error GoTo ErrorTrap
Dim DataLength As Long
Dim ReturnValue As Long
Dim Value As ValueStr
Dim ValueData As String
Dim ValueType As Long

   ValueData = String$(MAX_REG_VALUE_DATA, vbNullChar)
   DataLength = Len(ValueData)
   
   ReturnValue = CheckForError(RegQueryValueExA(ParentKeyH, ValueName, CLng(0), ValueType, ValueData, DataLength), CheckReturnValue:=True)
   If ReturnValue = ERROR_SUCCESS Then
      If Not IsNumber(ValueType) And DataLength > 0 Then DataLength = DataLength - 1
      Value.ValueData = Left$(ValueData, DataLength)
      Value.ValueName = ValueName
      Value.ValueType = ValueType
   Else
      Value.ValueData = vbNullString
      Value.ValueName = vbNullString
      Value.ValueType = REG_NONE
   End If
   
EndRoutine:
   GetValue = Value
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function
'This procedure returns the registry values contained by the specified key.
Public Function GetValues(ParentKeyH As Long) As ValueStr()
On Error GoTo ErrorTrap
Dim DataLength As Long
Dim HiveKeyH As Long
Dim Index As Long
Dim NameLength As Long
Dim ReturnValue As Long
Dim ValueData As String
Dim ValueName As String
Dim Values() As ValueStr
Dim ValueType As Long

   Erase Values()
   
   If Not ParentKeyH = NO_KEY Then
      Index = 0
      Do
         ValueData = String$(MAX_REG_VALUE_DATA, vbNullChar)
         DataLength = Len(ValueData)
         ValueName = String$(MAX_REG_VALUE_NAME, vbNullChar)
         NameLength = Len(ValueName)
         ReturnValue = CheckForError(RegEnumValueA(ParentKeyH, Index, ValueName, NameLength, CLng(0), ValueType, ValueData, DataLength), CheckReturnValue:=True, Ignored:=ERROR_NO_MORE_ITEMS)
         If Not IsNumber(ValueType) And DataLength > 0 Then DataLength = DataLength - 1
         If ReturnValue = ERROR_NO_MORE_ITEMS Or Not ReturnValue = ERROR_SUCCESS Then Exit Do
         If CheckForError(SafeArrayGetDim(Values())) = 0 Then
            ReDim Values(0 To 0) As ValueStr
         Else
            ReDim Preserve Values(LBound(Values()) To UBound(Values()) + 1) As ValueStr
         End If
   
         Values(UBound(Values())).ValueData = Left$(ValueData, DataLength)
         Values(UBound(Values())).ValueName = Left$(ValueName, NameLength)
         Values(UBound(Values())).ValueType = ValueType
         Index = Index + 1
      Loop
   End If
   
EndRoutine:
   GetValues = Values()
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   Err.Clear
   
   On Error Resume Next
   Description = Description & vbCr & "Error code: " & CStr(ErrorCode)
   MsgBox Description, vbExclamation
End Sub




'This procedure indicates whether the specified value type is numeric.
Public Function IsNumber(ValueType As Long) As Boolean
On Error GoTo ErrorTrap
Dim Numeric As Boolean
Dim TypeV As Variant

   Numeric = False
   For Each TypeV In Array(REG_BINARY, REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_QWORD)
      If CLng(TypeV) = ValueType Then
         Numeric = True
         Exit For
      End If
   Next TypeV
   
EndRoutine:
   IsNumber = Numeric
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure adds or returns a parent key handle.
Public Function KeyStack(Optional PushKey As String = vbNullString, Optional PopKey As Boolean = False, Optional ByRef Index As Long = NO_INDEX, Optional Refresh As Boolean = False) As String
On Error GoTo ErrorTrap
Dim KeyName As String
Static KeyNames() As String

   KeyName = vbNullString
   
   If Not PushKey = vbNullString Then
      If CheckForError(SafeArrayGetDim(KeyNames())) = 0 Then
         ReDim KeyNames(0 To 0) As String
      Else
         ReDim Preserve KeyNames(LBound(KeyNames()) To UBound(KeyNames()) + 1) As String
      End If
      KeyNames(UBound(KeyNames())) = PushKey
   ElseIf PopKey Then
      If Not CheckForError(SafeArrayGetDim(KeyNames())) = 0 Then
         If UBound(KeyNames()) = LBound(KeyNames()) Then
            Erase KeyNames()
         Else
            ReDim Preserve KeyNames(LBound(KeyNames()) To UBound(KeyNames()) - 1) As String
         End If
      End If
   ElseIf Refresh Then
      Erase KeyNames()
   End If
   
   If Not CheckForError(SafeArrayGetDim(KeyNames())) = 0 Then
      If Index = NO_INDEX Then Index = UBound(KeyNames())
      If Index <= UBound(KeyNames()) Then KeyName = KeyNames(Index)
   End If
   
EndRoutine:
   KeyStack = KeyName
   Exit Function
   
ErrorTrap:
   HandleError
   KeyName = vbNullString
   Resume EndRoutine
End Function

'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   KeyStack , , , Refresh:=True
   
   RegistryEditorWindow.Show
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the path to the selected registry key.
Public Function KeyStackToText() As String
On Error GoTo ErrorTrap
Dim Index As Long
Dim KeyName As String
Dim Text As String

   Index = 0
   Text = "\"
   Do
      KeyName = KeyStack(, , Index)
      If KeyName = vbNullString Then Exit Do
      Text = Text & KeyName & "\"
      Index = Index + 1
   Loop
   
EndRoutine:
   KeyStackToText = Text
   Exit Function
   
ErrorTrap:
   HandleError
   Text = vbNullString
   Resume EndRoutine
End Function




'This procedure opens the registry key under the specified key with the specified name.
Private Function OpenKey(ChildKeyName As String, Optional ParentKeyH As Long = ROOT_KEY) As Long
On Error GoTo ErrorTrap
Dim KeyH As Long

   KeyH = NO_KEY
   If ParentKeyH = ROOT_KEY Then
      KeyH = GetHiveKey(, , ChildKeyName).PredefinedH
   Else
      CheckForError RegOpenKeyExA(ParentKeyH, ChildKeyName, CLng(0), KEY_ALL_ACCESS Or KEY_WOW64_64KEY, KeyH), CheckReturnValue:=True
   End If
   
EndRoutine:
   OpenKey = KeyH
   Exit Function
   
ErrorTrap:
   HandleError
   KeyH = NO_KEY
   Resume EndRoutine
End Function

'This procedure returns the selected key's handle.
Public Function OpenSelectedKey(SelectedKeyName As String, SelectedKeyIndex As Long) As Long
On Error GoTo ErrorTrap
Dim ParentKeyH As Long

   ParentKeyH = ROOT_KEY
   
   If Not SelectedKeyIndex = NO_INDEX Then
      If (Not KeyStack() = vbNullString) And (SelectedKeyIndex = 0) Then
         KeyStack , PopKey:=True
      Else
         KeyStack PushKey:=SelectedKeyName
      End If
      
      ParentKeyH = WalkKeyStack()
   End If
   
EndRoutine:
   OpenSelectedKey = ParentKeyH
   Exit Function
   
ErrorTrap:
   HandleError
   ParentKeyH = NO_KEY
   Resume EndRoutine
End Function

'This procedure pads the specified registry value's data if required.
Private Function PadData(Value As ValueStr) As String
On Error GoTo ErrorTrap
Dim PaddedData As String

   PaddedData = Value.ValueData
   Select Case Value.ValueType
      Case REG_DWORD
         PaddedData = String$(4 - Len(PaddedData), vbNullChar) & PaddedData
      Case REG_DWORD_BIG_ENDIAN
         PaddedData = PaddedData & String$(4 - Len(PaddedData), vbNullChar)
      Case REG_QWORD
         PaddedData = PaddedData & String$(8 - Len(PaddedData), vbNullChar)
   End Select
   
EndRoutine:
   PadData = PaddedData
   Exit Function
   
ErrorTrap:
   HandleError
   Value.ValueType = REG_NONE
   Resume EndRoutine
End Function


'This procedure manages the security attributes used to access the registry.
Private Function SecurityAttributes() As SECURITY_ATTRIBUTES
On Error GoTo ErrorTrap
Static CurrentAttributes As SECURITY_ATTRIBUTES

   With CurrentAttributes
      If .nLength = 0 Then
         .bInheritHandle = CLng(True)
         .lpSecurityDescriptor = CLng(0)
         .nLength = Len(CurrentAttributes)
      End If
   End With
   
EndRoutine:
   SecurityAttributes = CurrentAttributes
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure creates/modifies the specified registry value.
Public Function SetValue(Value As ValueStr, ParentKeyH As Long) As Long
On Error GoTo ErrorTrap
Dim DataLength As Long
Dim ReturnValue As Long

   With Value
      If IsNumber(.ValueType) Then
         .ValueData = PadData(Value)
         DataLength = Len(.ValueData)
      Else
         DataLength = Len(.ValueData) + 1
      End If
   
      If Not Value.ValueType = REG_NONE Then
         ReturnValue = CheckForError(RegSetValueExA(ParentKeyH, .ValueName, CLng(0), .ValueType, .ValueData, DataLength), CheckReturnValue:=True)
      End If
   End With
   
EndRoutine:
   SetValue = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure updates the specified key list.
Public Sub UpdateKeyList(KeyList As ListBox, Keys() As KeyStr, ParentKeyH As Long)
On Error GoTo ErrorTrap
Dim Index As Long

   If Not KeyList Is Nothing Then
      With KeyList
         .Enabled = False
         .Clear
         If Not ParentKeyH = ROOT_KEY Then .AddItem ".."
         
         If Not SafeArrayGetDim(Keys()) = 0 Then
            For Index = LBound(Keys()) To UBound(Keys())
               .AddItem Keys(Index).KeyName
            Next Index
         End If
      End With
   End If
   
EndRoutine:
   If Not KeyList Is Nothing Then
      With KeyList
         .Enabled = True
         If .Visible Then .SetFocus
      End With
   End If
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure converts any escape sequences in the specified text to characters.
Public Function Unescape(Text As String, Optional UnescapeAll As Boolean = False, Optional ErrorAt As Long = 0) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Hexadecimals As String
Dim Index As Long
Dim NextCharacter As String
Dim Unescaped As String

   ErrorAt = 0
   Index = 1
   Unescaped = vbNullString
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If Character = ESCAPE_CHARACTER Then
         If NextCharacter = ESCAPE_CHARACTER Then
            Unescaped = Unescaped & Character
            Index = Index + 1
         Else
            Hexadecimals = UCase$(Mid$(Text, Index + 1, 2))
            If Len(Hexadecimals) = 2 Then
               If Left$(Hexadecimals, 1) = "0" Then Hexadecimals = Right$(Hexadecimals, 1)
      
               If UCase$(Hex$(CLng(Val("&H" & Hexadecimals & "&")))) = Hexadecimals Then
                  Unescaped = Unescaped & Chr$(CLng(Val("&H" & Hexadecimals & "&")))
                  Index = Index + 2
               Else
                  ErrorAt = Index
                  Exit Do
               End If
            Else
               ErrorAt = Index
               Exit Do
            End If
         End If
      Else
         If UnescapeAll Then
            ErrorAt = Index
            Exit Do
         Else
            Unescaped = Unescaped & Character
         End If
      End If
      Index = Index + 1
   Loop
   
EndRoutine:
   Unescape = Unescaped
   Exit Function
   
ErrorTrap:
   HandleError
   Unescaped = vbNullString
   Resume EndRoutine
End Function


'This procedure updates the specified value table.
Public Sub UpdateValueTable(ValueTable As MSFlexGrid, Values() As ValueStr)
On Error GoTo ErrorTrap
Dim Index As Long

   If Not ValueTable Is Nothing Then
      With ValueTable
         .Enabled = False
         .Clear
         .Rows = .FixedRows + 1
         .Row = 0
         .Col = 0: .Text = "Type:"
         .Col = 1: .Text = "Name:"
         .Col = 2: .Text = "Data:"
         
         If Not SafeArrayGetDim(Values()) = 0 Then
            .Rows = Abs(UBound(Values()) - LBound(Values())) + .FixedRows + 1
            For Index = LBound(Values()) To UBound(Values())
               .Row = Index + .FixedRows
               If Values(Index).ValueName = vbNullString Then
                  .CellBackColor = DEFAULT_VALUE_COLOR: .Col = 0: .Text = ValueTypeName(Values(Index).ValueType)
                  .CellBackColor = DEFAULT_VALUE_COLOR: .Col = 1: .Text = "(Default)"
                  .CellBackColor = DEFAULT_VALUE_COLOR: .Col = 2: .Text = Escape(Values(Index).ValueData, EscapeAll:=IsNumber(Values(Index).ValueType))
               Else
                  .Col = 0: .Text = ValueTypeName(Values(Index).ValueType)
                  .Col = 1: .Text = Values(Index).ValueName
                  .Col = 2: .Text = Escape(Values(Index).ValueData, EscapeAll:=IsNumber(Values(Index).ValueType))
               End If
            Next Index
         
            .Col = 1: .Sort = flexSortGenericAscending
         End If
      End With
   End If
   
EndRoutine:
   If Not ValueTable Is Nothing Then ValueTable.Enabled = True
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns value indicating whether the specified registry value exists.
Public Function ValueExists(ValueName As String, ParentKeyH As Long) As Boolean
On Error GoTo ErrorTrap
Dim ReturnValue As Long
   
   ReturnValue = RegQueryValueExA(ParentKeyH, ValueName, CLng(0), CLng(0), vbNullString, CLng(0))
   If Not (ReturnValue = ERROR_FILE_NOT_FOUND Or ReturnValue = ERROR_MORE_DATA) Then CheckForError ReturnValue, CheckReturnValue:=True
   
EndRoutine:
   ValueExists = Not (ReturnValue = ERROR_FILE_NOT_FOUND)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the name of the specified registry value type.
Private Function ValueTypeName(ValueType As Long) As String
On Error GoTo ErrorTrap
Dim TypeName As String
   
   TypeName = vbNullString
   Select Case ValueType
      Case REG_BINARY
         TypeName = "REG_BINARY"
      Case REG_DWORD
         TypeName = "REG_DWORD"
      Case REG_DWORD_BIG_ENDIAN
         TypeName = "REG_DWORD_BIG_ENDIAN"
      Case REG_EXPAND_SZ
         TypeName = "REG_EXPAND_SZ"
      Case REG_LINK
         TypeName = "REG_LINK"
      Case REG_MULTI_SZ
         TypeName = "REG_MULTI_SZ"
      Case REG_QWORD
         TypeName = "REG_QWORD"
      Case REG_SZ
         TypeName = "REG_SZ"
      Case Else
         TypeName = CStr(ValueType)
   End Select
   
EndRoutine:
   ValueTypeName = TypeName
   Exit Function
   
ErrorTrap:
   HandleError
   TypeName = vbNullString
   Resume EndRoutine
End Function

'This procedure returns a list of all registry data type names.
Public Function ValueTypeNames(Optional Delimiter As String = vbCr)
On Error GoTo ErrorTrap
Dim DataType As Variant
Dim DataTypes As String

   DataTypes = vbNullString
   For Each DataType In Array(REG_SZ, REG_EXPAND_SZ, REG_BINARY, REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_LINK, REG_MULTI_SZ, REG_QWORD)
      DataTypes = DataTypes & CStr(DataType) & " - " & ValueTypeName(CLng(DataType)) & Delimiter
   Next DataType
   
EndRoutine:
   ValueTypeNames = DataTypes
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure returns the handle of the last key on the stack.
Public Function WalkKeyStack() As Long
On Error GoTo ErrorTrap
Dim ChildKeyH As Long
Dim ChildKeyName As String
Dim Index As Long
Dim ParentKeyH As Long

   Index = 0
   ParentKeyH = ROOT_KEY
   Do
      ChildKeyName = KeyStack(, , Index)
      If ChildKeyName = vbNullString Then Exit Do
      ChildKeyH = OpenKey(ChildKeyName, ParentKeyH)
      If ChildKeyH = NO_KEY Then
         ParentKeyH = NO_KEY
         Exit Do
      End If
      CloseKey ParentKeyH
      ParentKeyH = ChildKeyH
      Index = Index + 1
   Loop
   
EndRoutine:
   WalkKeyStack = ParentKeyH
   Exit Function
   
ErrorTrap:
   HandleError
   ParentKeyH = NO_KEY
   Resume EndRoutine
End Function


