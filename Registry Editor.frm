VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RegistryEditorWindow 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1416
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   5280
   Icon            =   "Registry Editor.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   5.9
   ScaleMode       =   4  'Character
   ScaleWidth      =   44
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox KeyInformationBox 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Displays the selected key's class and modification date and time."
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox KeyListBox 
      Height          =   300
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Lists the current key's child keys."
      Top             =   120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid ValueTableBox 
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Lists the current key's values."
      Top             =   120
      Width           =   3135
      _ExtentX        =   5525
      _ExtentY        =   1291
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.TextBox KeyPathBox 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Displays the current key's path."
      Top             =   960
      Width           =   5055
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu ProgramMenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu KeyMainMenu 
      Caption         =   "&Key"
      Begin VB.Menu DeleteKeyMenu 
         Caption         =   "&Delete Key"
         Shortcut        =   ^D
      End
      Begin VB.Menu NewKeyMenu 
         Caption         =   "&New Key"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu ValueMainMenu 
      Caption         =   "&Value"
      Begin VB.Menu DeleteValueMenu 
         Caption         =   "&Delete Value"
         Shortcut        =   ^R
      End
      Begin VB.Menu ModifyValueMenu 
         Caption         =   "&Modify Value"
         Shortcut        =   ^M
      End
      Begin VB.Menu NewValueMenu 
         Caption         =   "&New Value"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "RegistryEditorWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit

'This enumeration contains a list of actions that can be performed on keys and/or values.
Private Enum ActionsE
   ActionCreate         'Defines the create key/value action.
   ActionDelete         'Defines the delete key/value action.
   ActionModify         'Defines the modify value action.
End Enum

'This procedure displays the selected registry key's subkeys.
Private Sub DisplayKeySubkeys()
On Error GoTo ErrorTrap
Dim Keys() As KeyStr
Dim ParentKeyH As Long

   Screen.MousePointer = vbHourglass
   
   ParentKeyH = OpenSelectedKey(KeyListBox.List(KeyListBox.ListIndex), KeyListBox.ListIndex)
   Keys() = GetKeys(ParentKeyH)
   CloseKey ParentKeyH
   
   UpdateKeyList KeyListBox, Keys(), ParentKeyH
   KeyPathBox.Text = "Path: """ & KeyStackToText() & """"
   
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the selected registry key's information and values.
Private Sub DisplayKeyValuesAndInformation()
On Error GoTo ErrorTrap
Dim Key As KeyStr
Dim ParentKeyH As Long
Dim Values() As ValueStr

   Screen.MousePointer = vbHourglass
   
   If (KeyStack() = vbNullString And KeyListBox.ListIndex >= 0) Or (KeyListBox.ListIndex > 0) Then
      Key = GetKeyInformation(KeyListBox.List(KeyListBox.ListIndex), KeyListBox.ListIndex)
      With Key
         If .KeyAccessible Then
            KeyInformationBox.Text = "Class: """ & .KeyClass & """ - Date: " & FileTimeToText(.KeyDateTime)
            KeyStack PushKey:=KeyListBox.List(KeyListBox.ListIndex)
            
            ParentKeyH = WalkKeyStack()
            KeyStack , PopKey:=True
            Values() = GetValues(ParentKeyH)
            CloseKey ParentKeyH
         Else
            KeyInformationBox.Text = "Cannot access this key."
         End If
      End With
   Else
      KeyInformationBox.Text = "No key selected."
   End If
   
   UpdateValueTable ValueTableBox, Values()
   
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure creates/deletes a registry key after requesting the user for input.
Private Sub EditKey(Action As ActionsE)
On Error GoTo ErrorTrap
Dim Disposition As Long
Dim ErrorAt As Long
Dim Key As KeyStr
Dim ParentKeyH As Long

   ParentKeyH = WalkKeyStack()
   
   If ParentKeyH = NO_KEY Or ParentKeyH = ROOT_KEY Then
      MsgBox "Cannot create/delete a key here.", vbInformation
   Else
      Select Case Action
         Case ActionCreate
            Key.KeyName = Unescape(InputBox$("Key name:", Escape(Key.KeyName)), , ErrorAt)
            If Not (EscapeSequenceError(ErrorAt) Or Key.KeyName = vbNullString) Then
               Key.KeyClass = InputBox$("Key class:", , Escape(Key.KeyClass))
               If Not StrPtr(Key.KeyClass) = 0 Then
                  Key.KeyClass = Unescape(Key.KeyClass, , ErrorAt)
                  If Not EscapeSequenceError(ErrorAt) Then
                     CreateKey Key, ParentKeyH, Disposition
                     If Disposition = REG_OPENED_EXISTING_KEY Then MsgBox """" & Key.KeyName & """ already exists.", vbInformation
                  End If
               End If
            End If
         Case ActionDelete
            If (KeyStack() = vbNullString And KeyListBox.ListIndex >= 0) Or (KeyListBox.ListIndex > 0) Then
               Key = GetKey(KeyListBox.List(KeyListBox.ListIndex), ParentKeyH)
               If Confirmed("Delete """ & Key.KeyName & """?") Then DeleteKey Key.KeyName, ParentKeyH
            End If
      End Select
   
      UpdateKeyList KeyListBox, GetKeys(ParentKeyH), ParentKeyH
   End If
   
EndRoutine:
   CloseKey ParentKeyH
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure performs the specified action on a registry value after requesting the user to specify input.
Private Sub EditValue(Action As ActionsE)
On Error GoTo ErrorTrap
Dim CurrentName As String
Dim ParentKeyH As Long
Dim Value As ValueStr

   If (KeyStack() = vbNullString And KeyListBox.ListIndex >= 0) Or (KeyListBox.ListIndex > 0) Then
      KeyStack PushKey:=KeyListBox.List(KeyListBox.ListIndex)
      ParentKeyH = WalkKeyStack()
      KeyStack , PopKey:=True
     
      If ParentKeyH = NO_KEY Or ParentKeyH = ROOT_KEY Then
         MsgBox "Cannot create/delete/modify a value here.", vbInformation
      Else
         Select Case Action
            Case ActionCreate
               Value = RequestValueParameters(Value)
   
               If Not Value.ValueType = REG_NONE Then
                  If ValueExists(Value.ValueName, ParentKeyH) Then
                     If Confirmed("Overwrite """ & Value.ValueName & """?") Then
                        SetValue Value, ParentKeyH
                     End If
                  Else
                     SetValue Value, ParentKeyH
                  End If
               End If
            Case ActionDelete
               If ValueTableBox.Row > 0 Then
                  Value.ValueName = ValueTableBox.TextMatrix(ValueTableBox.Row, 1)
      
                  If Confirmed("Delete """ & Value.ValueName & """?") Then
                     If ValueTableBox.CellBackColor = DEFAULT_VALUE_COLOR Then
                        DeleteValue vbNullString, ParentKeyH
                     Else
                        DeleteValue Value.ValueName, ParentKeyH
                     End If
                  End If
               Else
                   MsgBox "No value has been selected.", vbInformation
               End If
            Case ActionModify
               If ValueTableBox.Row > 0 Then
                  If ValueTableBox.CellBackColor = DEFAULT_VALUE_COLOR Then
                     Value.ValueName = vbNullString
                  Else
                     Value.ValueName = ValueTableBox.TextMatrix(ValueTableBox.Row, 1)
                  End If
   
                  CurrentName = Value.ValueName
                  Value = RequestValueParameters(GetValue(Value.ValueName, ParentKeyH))
      
                  If Not Value.ValueType = REG_NONE Then
                     SetValue Value, ParentKeyH
                     If Not Value.ValueName = CurrentName Then DeleteValue CurrentName, ParentKeyH
                  End If
               Else
                   MsgBox "No value has been selected.", vbInformation
               End If
         End Select
      End If
   Else
      MsgBox "No subkey has been selected.", vbInformation
   End If
   
EndRoutine:
   CloseKey ParentKeyH
   DisplayKeyValuesAndInformation
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure requests the user to specify the parameters for a value and returns the result.
Private Function RequestValueParameters(Value As ValueStr) As ValueStr
On Error GoTo ErrorTrap
Dim ErrorAt As Long
Dim NewValue As ValueStr

   With NewValue
      .ValueName = InputBox$("Value name:", , Escape(Value.ValueName))
      If StrPtr(.ValueName) = 0 Then
         .ValueType = REG_NONE
      Else
         .ValueName = Unescape(.ValueName, , ErrorAt)
         If EscapeSequenceError(ErrorAt) Then
            .ValueType = REG_NONE
         Else
            .ValueType = CLng(Val(InputBox$("Value type:" & vbCr & ValueTypeNames(), , CStr(Value.ValueType))))
            If Not .ValueType = REG_NONE Then
               .ValueData = InputBox$("Value data:", , Escape(Value.ValueData, EscapeAll:=Numeric(.ValueType)))
               If StrPtr(.ValueData) = 0 Then
                  .ValueType = REG_NONE
               Else
                  .ValueData = Unescape(.ValueData, UnescapeAll:=Numeric(.ValueType), ErrorAt:=ErrorAt)
                  If EscapeSequenceError(ErrorAt) Then .ValueType = REG_NONE
               End If
            End If
         End If
      End If
   End With
      
EndRoutine:
   RequestValueParameters = NewValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure gives the command to delete a registry key.
Private Sub DeleteKeyMenu_Click()
On Error GoTo ErrorTrap
   EditKey ActionDelete
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to delete a registry value.
Private Sub DeleteValueMenu_Click()
On Error GoTo ErrorTrap
   EditValue ActionDelete
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap

   Me.Caption = ProgramInformation()
   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   
   DisplayKeySubkeys
   DisplayKeyValuesAndInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub
'This procedure adjusts this window to its new size.
Private Sub Form_Resize()
On Error Resume Next
Dim Column As Long

   KeyListBox.Width = (Me.ScaleWidth / 2) - 2
   KeyListBox.Height = Me.ScaleHeight - 5
   KeyListBox.Top = 0.5
   KeyListBox.Left = 1
   
   KeyInformationBox.Width = KeyListBox.Width
   KeyInformationBox.Left = KeyListBox.Left
   KeyInformationBox.Top = KeyListBox.Top + (KeyListBox.Height + 0.5)
   
   KeyPathBox.Width = Me.ScaleWidth - 2
   KeyPathBox.Left = 1
   KeyPathBox.Top = Me.ScaleHeight - 2
   
   ValueTableBox.Width = (Me.ScaleWidth / 2) - 1
   ValueTableBox.Height = Me.ScaleHeight - 3
   ValueTableBox.Top = 0.5
   ValueTableBox.Left = (Me.ScaleWidth - 1) - ValueTableBox.Width
   
   For Column = 0 To ValueTableBox.Cols - 1
      ValueTableBox.ColAlignment(Column) = flexAlignLeftCenter
      ValueTableBox.ColWidth(Column) = (ValueTableBox.Width * (Screen.TwipsPerPixelX * PIXELS_PER_CHARACTER_X)) * (0.95 / ValueTableBox.Cols)
   Next Column
End Sub

'This procedure closes this program when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   Cancel = Not Confirmed("Close this program?")
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   MsgBox App.Comments, vbInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to display the selected registry key's information and values.
Private Sub KeyListBox_Click()
On Error GoTo ErrorTrap
   DisplayKeyValuesAndInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to display the selected registry key.
Private Sub KeyListBox_DblClick()
On Error GoTo ErrorTrap
   DisplayKeySubkeys
   DisplayKeyValuesAndInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure handles the user's keystrokes for the registry key list.
Private Sub KeyListBox_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap

   Select Case KeyAscii
      Case vbKeyBack
         If Not KeyStack(, , Index:=0) = vbNullString Then
            KeyListBox.ListIndex = 0
            DisplayKeySubkeys
            DisplayKeyValuesAndInformation
         End If
      Case vbKeyReturn
         DisplayKeySubkeys
         DisplayKeyValuesAndInformation
   End Select
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to modify a registry value.
Private Sub ModifyValueMenu_Click()
On Error GoTo ErrorTrap
   EditValue ActionModify
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to create a new registry key.
Private Sub NewKeyMenu_Click()
On Error GoTo ErrorTrap
   EditKey ActionCreate
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to create a new registry value.
Private Sub NewValueMenu_Click()
On Error GoTo ErrorTrap
   EditValue ActionCreate
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure closes this program.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

