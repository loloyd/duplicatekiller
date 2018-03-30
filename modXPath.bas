Attribute VB_Name = "modXPath"
Option Explicit
Public Function XDir(ByVal pStr As String, Optional ByVal pAttr As VbFileAttribute = -1, Optional ByVal pDoEvents As Boolean = True) As String
  XDir = ""
  If InStr(1, pStr, "\") = 0 Then
    XDir = "Error: Path separator " & Chr(34) & "\" & Chr(34) & " required in pStr argument!"
    Exit Function
  End If
  Dim vPath As String, vParent As String
  vParent = XParentPath(pStr)
  Select Case pDoEvents
    Case True
      If pAttr = -1 Then
        vPath = Dir(pStr)
        Do While vPath <> ""
          If vPath <> "." And vPath <> ".." Then XDir = XDir & vParent & vPath & "|"
          vPath = Dir
          DoEvents
        Loop
      Else
        vPath = Dir(pStr, pAttr)
        Do While vPath <> ""
          If vPath <> "." And vPath <> ".." Then
            On Error Resume Next
            If (GetAttr(vParent & vPath) And pAttr) <> 0 Then
              If Err.Number = 0 Then XDir = XDir & vParent & vPath & "|"
            End If
            On Error GoTo 0
          End If
          vPath = Dir
          DoEvents
        Loop
      End If
    Case False
      If pAttr = -1 Then
        vPath = Dir(pStr)
        Do While vPath <> ""
          XDir = XDir & vParent & vPath & "|"
          vPath = Dir
        Loop
      Else
        vPath = Dir(pStr, pAttr)
        Do While vPath <> ""
          On Error Resume Next
          If (GetAttr(vParent & vPath) And pAttr) <> 0 Then
            If Err.Number = 0 Then XDir = XDir & vParent & vPath & "|"
          End If
          On Error GoTo 0
          vPath = Dir
        Loop
      End If
    End Select
End Function
Public Function XParentPath(ByVal pPath As String) As String
  Dim vI As Long
  XParentPath = ""
  If IsNull(pPath) Or Trim(pPath) = "" Then XParentPath = "Error: Non-empty string required in pPath argument!"
  If Right(pPath, 1) = "\" Then pPath = Left(pPath, Len(pPath) - 1)
  XParentPath = Left(pPath, (InStrRev(pPath, "\")))
End Function
Public Function XProperPath(ByVal pPath As String) As String
  XProperPath = pPath & IIf(Right(pPath, 1) = "\", "", "\")
End Function
Public Function XFileLen(ByVal pFile As String) As Currency
  If FileLen(pFile) >= 0 Then
    XFileLen = FileLen(pFile)
  Else
    XFileLen = CCur(FileLen(pFile) And &H7FFFFFFF) + 2147483648#
  End If
End Function

