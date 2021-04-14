Attribute VB_Name = "Module1"
'PUT ALL THIS CODE (DECLARATIONS AND FUNCTIONS)
'INTO A .BAS MODULE.

Option Explicit
Option Compare Text
Private Declare Function EnumWindows Lib "user32" _
    (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias _
"GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, _
ByVal cch As Long) As Long

Private pbExact As Boolean
Private psAppString As String
Private piAppCount As Integer
Public Function AppInstances(AppNamePart As String, _
   Optional ExactMatchOnly As Boolean) As Integer

'from http://www.freevbcode.com/ShowCode.ASP?ID=488
'PURPOSE: Counts the Number of Instances of a Given Application
'PARAMETERS:
   'AppNamePart = Any Part of the WindowTitle for the App
   'ExactMatchOnly (Optional): If you want to
   'count only exact matches for AppNamePart,
   'set this parameter to true

'RETURNS: Number of Running Instances
'EXAMPLE:
  'dim iIEInstances as integer
   'iIEInstances = AppInstances("Microsoft Internet Explorer")

Dim lRet As Long

psAppString = AppNamePart
pbExact = ExactMatchOnly

lRet = EnumWindows(AddressOf CheckForInstance, 0)
AppInstances = piAppCount
End Function

Private Function CheckForInstance(ByVal lhWnd As Long, ByVal _
lParam As Long) As Long

Dim sTitle As String
Dim lRet As Long
Dim iNew As Integer


sTitle = Space(255)
lRet = GetWindowText(lhWnd, sTitle, 255)

sTitle = StripNull(sTitle)

If sTitle <> "" Then
    If pbExact Then
        If sTitle = psAppString Then piAppCount = piAppCount + 1
    Else
        If InStr(sTitle, psAppString) > 0 Then _
           piAppCount = piAppCount + 1
    End If
End If

CheckForInstance = True

End Function

Private Function StripNull(ByVal InString As String) As String

'Input: String containing null terminator (Chr(0))
'Returns: all character before the null terminator

Dim iNull As Integer
If Len(InString) > 0 Then
    iNull = InStr(InString, vbNullChar)
    Select Case iNull
    Case 0
        StripNull = InString
    Case 1
        StripNull = ""
    Case Else
       StripNull = Left$(InString, iNull - 1)
   End Select
End If

End Function


