VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Duplicate Killer"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRemoveEmptyDirectories 
      Caption         =   "&Remove empty directories"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox chkKillDuplicates 
      Caption         =   "&Kill detected duplicates"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdDetectDuplicates 
      Caption         =   "&Detect Duplicates Only"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.ListBox lstSource 
      Height          =   2010
      Left            =   120
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   3720
      Width           =   10695
   End
   Begin VB.ListBox lstTarget 
      Height          =   2010
      Left            =   120
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1320
      Width           =   10695
   End
   Begin VB.Label Label1 
      Caption         =   "Runtime Options:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblSource 
      Caption         =   "Source"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   10695
   End
   Begin VB.Label lblTarget 
      Caption         =   "Target"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   10695
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function CompareFilesIfMatch(ByVal pTarget As String, ByVal pSource As String) As Boolean
  Dim vTargetByte() As Byte, vSourceByte() As Byte
  Dim vFileLen As Currency
  Dim vDiskIOError As Boolean
  Dim vJ As Long, vK As Long
  CompareFilesIfMatch = True
  vDiskIOError = False
  On Error GoTo Err_Local
  vFileLen = FileLen(pTarget)
  Open pTarget For Binary Access Read Shared As #1
  Open pSource For Binary Access Read Shared As #2
'AppendToLog "Comparing T:" & vItem.Text & " and S:" & vItem.SubItems(4) & "..."
  If Not vDiskIOError Then
    If vFileLen < 16384 ^ 2 / 8 Then 'if file is relatively small
      ReDim vTargetByte(LOF(1))
      ReDim vSourceByte(LOF(2))
        Do While Not EOF(1)
          Get #1, , vTargetByte()
          Get #2, , vSourceByte()
        Loop
      Close #2
      Close #1
      For vJ = LBound(vTargetByte) To UBound(vTargetByte)
        If vTargetByte(vJ) <> vSourceByte(vJ) Then
          CompareFilesIfMatch = False
          Exit For
        End If
        If vJ Mod 1024 = 0 And UBound(vTargetByte) <> 0 Then
'pbComp.Value = Int(vJ / UBound(vTargetByte) * 100)
'lblPBComp.Caption = Format(vJ, "#,##0") & " (" & pbComp.Value & "%)"
          lblProgress.Caption = "Comparing... " & Format(vJ, "#,##0") & " (" & Int(vJ / UBound(vTargetByte) * 100) & "%)"
          DoEvents
        End If
      Next vJ
      vJ = vJ - 1
      lblProgress.Caption = "Comparing... " & Format(vJ, "#,##0") & " (" & Int(vJ / UBound(vTargetByte) * 100) & "%)"
      DoEvents
    Else 'if file is relatively big
      vJ = 1
      Do While Not EOF(1)
        If vFileLen - vJ + 1 > 2097152 Then
          ReDim vTargetByte(1 To 2097152)
          ReDim vSourceByte(1 To 2097152)
        Else
          ReDim vTargetByte(1 To (vFileLen - vJ + 1))
          ReDim vSourceByte(1 To (vFileLen - vJ + 1))
        End If
        Get #1, , vTargetByte()
        Get #2, , vSourceByte()
        For vK = LBound(vTargetByte) To UBound(vTargetByte)
          If vTargetByte(vK) <> vSourceByte(vK) Then
            CompareFilesIfMatch = False
'AppendToLog "Mismatch found at byte offset " & vK & ". Skipping this target file..."
            Exit For
          End If
          If vK Mod 1024 = 0 And UBound(vTargetByte) <> 0 Then
'pbComp.Value = Int((vK + vJ - 1) / vFileLen * 100)
'lblPBComp.Caption = Format(vK + vJ - 1, "#,##0") & " (" & pbComp.Value & "%)"
            lblProgress.Caption = "Comparing... " & Format(vK + vJ - 1, "#,##0") & " (" & Int((vK + vJ - 1) / vFileLen * 100) & "%)"
            DoEvents
          End If
        Next vK
        vK = vK - 1
        lblProgress.Caption = "Comparing... " & Format(vK + vJ - 1, "#,##0") & " (" & Int((vK + vJ - 1) / vFileLen * 100) & "%)"
        DoEvents
        If CompareFilesIfMatch And vJ + 2097152 < vFileLen Then
          vJ = vJ + 2097152
        Else
          Exit Do
        End If
      Loop
      Close #2
      Close #1
    End If 'whether if file is relatively small or relatively big
  End If 'whether if no disk error has been encountered
  If vDiskIOError Then CompareFilesIfMatch = False
  Exit Function
Err_Local:
  If Err.Number <> 0 Then
    'AppendToLog "Error encountered - " & Err.Number & " " & Err.Description
    If Err.Number = 75 Then DoEvents 'AppendToLog "Check file permissions!"
    vDiskIOError = True 'possible disk IO error encountered
    CompareFilesIfMatch = False
  End If
  Resume Next
End Function
Private Sub DragDropOnListBox(ByRef pListBox As ListBox, ByRef pData As DataObject)
  Dim vI As Long
  Dim vFileName As String
  Dim vFileTitle As String
  If pData.Files(1) <> "" Then
    For vI = 1 To pData.Files.Count
      vFileName = pData.Files(vI)
      vFileTitle = Mid(vFileName, Len(XParentPath(vFileName)) + 1)
      pListBox.AddItem vFileTitle & " | " & vFileName
    Next vI
  End If
End Sub
Private Sub DetectDuplicatesAndAct()
  Dim vIlstTarget As Long
  Dim vIlstSource As Long
  Dim vFileNameTarget As String
  Dim vFileNameSource As String
  Dim vFilePathTarget As String
  Dim vFileTitle As String
  Dim vItemRemovedFlag As Boolean
  Dim vFilesInsideFound As Boolean
  vItemRemovedFlag = False
  For vIlstTarget = 0 To lstTarget.ListCount - 1
    vFileNameTarget = Mid(lstTarget.List(vIlstTarget), InStrRev(lstTarget.List(vIlstTarget), "|") + 2)
    lstTarget.Selected(vIlstTarget) = True
    lblTarget.Caption = "Target: " & Format(FileLen(vFileNameTarget), "#,##0") & " bytes, " & Format(FileDateTime(vFileNameTarget), "yyyy/mm/dd hh:mm:ss")
    For vIlstSource = 0 To lstSource.ListCount - 1
      vFileNameSource = Mid(lstSource.List(vIlstSource), InStrRev(lstSource.List(vIlstSource), "|") + 2)
      lblSource.Caption = "Source: " & Format(FileLen(vFileNameSource), "#,##0") & " bytes, " & Format(FileDateTime(vFileNameSource), "yyyy/mm/dd hh:mm:ss")
      lstSource.Selected(vIlstSource) = True
      If FileLen(vFileNameTarget) = FileLen(vFileNameSource) And LCase(vFileNameTarget) <> LCase(vFileNameSource) Then
        If CompareFilesIfMatch(vFileNameTarget, vFileNameSource) Then 'if files match
          If (GetAttr(vFileNameTarget) And vbSystem) <> 0 Then SetAttr vFileNameTarget, vbNormal
          If (GetAttr(vFileNameTarget) And vbHidden) <> 0 Then SetAttr vFileNameTarget, vbNormal
          If (GetAttr(vFileNameTarget) And vbReadOnly) <> 0 Then SetAttr vFileNameTarget, vbNormal
'AppendToLog "Match found.  Killing T:" & vItem.Text & "..."
          If chkKillDuplicates.Value = vbChecked Then
            Kill vFileNameTarget
          Else
            If Left(lstTarget.List(vIlstTarget), Len("Duplicate Match!")) <> "Duplicate Match!" Then lstTarget.List(vIlstTarget) = "Duplicate Match! " & lstTarget.List(vIlstTarget)
          End If
          lstTarget.Selected(vIlstTarget) = False
          If Dir(vFileNameTarget) = "" Then
            lstTarget.RemoveItem vIlstTarget
            vItemRemovedFlag = True
            DoEvents
            If chkRemoveEmptyDirectories.Value = vbChecked Then
              vFilesInsideFound = False
              vFilePathTarget = vFileNameTarget
              Do While Not vFilesInsideFound
                vFilePathTarget = XParentPath(vFilePathTarget)
                vFileTitle = Dir(vFilePathTarget & "*")
                Do While vFileTitle <> ""
                  If vFileTitle <> "." And vFileTitle <> ".." Then
                    vFilesInsideFound = True
                    Exit Do
                  End If
                  vFileTitle = Dir
                Loop
                If Not vFilesInsideFound Then
                  On Error Resume Next
                  SetAttr XParentPath(vFileNameTarget), vbNormal
                  RmDir XParentPath(vFileNameTarget)
                  On Error GoTo 0
                End If
              Loop 'if files/folders are still found inside any of the parent paths
            End If 'whether the option to remove empty directories is enabled or not
            If vIlstTarget <= lstTarget.ListCount - 1 Then vIlstTarget = vIlstTarget - 1
          End If
          On Error GoTo 0
        Else 'if files mismatch
          lstTarget.Selected(vIlstTarget) = False
        End If 'whether files match or mismatch
      End If 'filetarget and filesource have the same sizes but are from different named location references
      lstSource.Selected(vIlstSource) = False
      If vItemRemovedFlag Then Exit For
    Next vIlstSource
    vItemRemovedFlag = False
    lblTarget.Caption = "Target"
    lblSource.Caption = "Source"
    If vIlstTarget >= lstTarget.ListCount - 1 Then Exit For
  Next vIlstTarget
  lblProgress.Caption = "Progress"
End Sub
Private Sub chkKillDuplicates_Click()
  cmdDetectDuplicates.Caption = IIf(chkKillDuplicates.Value = vbChecked, "&Detect Duplicates and Kill", "&Detect Duplicates Only")
End Sub
Private Sub cmdDetectDuplicates_Click()
  cmdDetectDuplicates.Enabled = False
  DetectDuplicatesAndAct
  cmdDetectDuplicates.Enabled = True
End Sub
Private Sub Form_Resize()
  lstTarget.Width = Me.Width - 520
  lstSource.Width = lstTarget.Width
  lblTarget.Width = lstTarget.Width
  lblSource.Width = lstTarget.Width
  Me.Height = 6435
End Sub
Private Sub lstSource_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim vI As Long
  If KeyCode = vbKeyDelete Then
    For vI = lstSource.ListCount - 1 To 0 Step -1
      If lstSource.Selected(vI) Then lstSource.RemoveItem vI
    Next vI
  End If
End Sub
Private Sub lstSource_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  DragDropOnListBox lstSource, Data
End Sub
Private Sub lstTarget_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim vI As Long
  If KeyCode = vbKeyDelete Then
    For vI = lstTarget.ListCount - 1 To 0 Step -1
      If lstTarget.Selected(vI) Then lstTarget.RemoveItem vI
    Next vI
  End If
End Sub
Private Sub lstTarget_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  DragDropOnListBox lstTarget, Data
End Sub
