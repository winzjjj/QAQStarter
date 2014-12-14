Attribute VB_Name = "mdlFiles"
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Public_sFiles() As String  '����sfiles
'���·�����Ⱥ��ļ����Գ����Ķ���
 Public Const MAX_PATH = 260
 Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
 Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
 Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
 Public Const FILE_ATTRIBUTE_HIDDEN = &H2
 Public Const FILE_ATTRIBUTE_NORMAL = &H80
 Public Const FILE_ATTRIBUTE_READONLY = &H1
 Public Const FILE_ATTRIBUTE_SYSTEM = &H4
 Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

'�Զ�����������FILETIME��WIN32_FIND_DATA�Ķ���
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
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

Public Function fDelInvaildChr(str As String) As String
    On Error Resume Next
    For i = Len(str) To 1 Step -1
        If Asc(Mid(str, i, 1)) <> 0 And Asc(Mid(str, i, 1)) <> 32 Then
            fDelInvaildChr = Left(str, i)
            Exit For
        End If
    Next
End Function
Public Function TrimPath(sPath As String) As String
  Dim i As Integer
    i = InStrRev(sPath, ".") + 1
    TrimPath = Mid(sPath, i)
End Function

Public Sub sDirTraversal(ByVal strPathName As String, ByVal List1 As ListBox, ByVal F_name As String)
    Dim sSubDir() As String '��ŵ�ǰĿ¼�µ���Ŀ¼,�±�ɸ�����Ҫ����
    Dim iIndex       As Integer '��Ŀ¼�����±�
    Dim i            As Integer '����ѭ����Ŀ¼�Ĳ���
     
    Dim lHandle      As Long 'FindFirstFileA �ľ��
    Dim tFindData    As WIN32_FIND_DATA '
    Dim strFileName  As String '�ļ���
     
    On Error Resume Next
    '��ʼ������
    i = 1
    iIndex = 0
    tFindData.cFileName = "" '��ʼ�������ַ���
     
    lHandle = FindFirstFile(strPathName & "\*.*", tFindData)

    If lHandle = 0 Then '��ѯ������������
        Exit Sub
    End If

    strFileName = fDelInvaildChr(tFindData.cFileName)

    If tFindData.dwFileAttributes = &H10 Then 'Ŀ¼
        If strFileName <> "." And strFileName <> ".." Then
            iIndex = iIndex + 1
            sSubDir(iIndex) = strPathName & "\" & strFileName '��ӵ�Ŀ¼����
            ReDim Preserve sSubDir(iIndex)
        End If

    Else

      If TrimPath(strPathName & "\" & strFileName) = F_name Then
            List1.AddItem strFileName
        End If
    End If

    Do While True
        tFindData.cFileName = ""

        If FindNextFile(lHandle, tFindData) = 0 Then '��ѯ������������
            FindClose (lHandle)
            Exit Do
        Else
            strFileName = fDelInvaildChr(tFindData.cFileName)

            If tFindData.dwFileAttributes = &H10 Then
                If strFileName <> "." And strFileName <> ".." Then
                    iIndex = iIndex + 1
                    sSubDir(iIndex) = strPathName & "\" & strFileName '��ӵ�Ŀ¼����
                End If

            Else

               If TrimPath(strPathName & "\" & strFileName) = F_name Then
                    List1.AddItem strPathName & "\" & strFileName
                End If
            End If
        End If

    Loop

    '�����Ŀ¼����Ŀ¼�������Ŀ¼����ݹ����
    'If iIndex > 0 Then

    '    For i = 1 To iIndex
    '        sDirTraversal sSubDir(i), newclist, F_name
    '    Next
'
'    End If

End Sub

Public Function TreeSearch(ByVal sPath As String, ByVal sFileSpec As String, sFiles() As String) As Long
    Static lngFiles As Long '�ļ���Ŀ
    Dim sDir As String
    Dim sSubDirs() As String '�����Ŀ¼����
    Dim lngIndex As Long
    Dim lngTemp&
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    sDir = Dir(sPath & sFileSpec)
   '��õ�ǰĿ¼���ļ�������Ŀ
    Do While Len(sDir)
      lngFiles = lngFiles + 1
      ReDim Preserve sFiles(1 To lngFiles)
      sFiles(lngFiles) = sPath & sDir
      sDir = Dir
    Loop
   '��õ�ǰĿ¼�µ���Ŀ¼����
    lngIndex = 0
    sDir = Dir(sPath & "*.*", vbDirectory)
    Do While Len(sDir)
      If Left(sDir, 1) <> "." And Left(sDir, 1) <> ".." Then '' ������ǰ��Ŀ¼���ϲ�Ŀ¼
     '�ҳ���Ŀ¼��
        If GetAttr(sPath & sDir) And vbDirectory Then
          lngIndex = lngIndex + 1
         '������Ŀ¼��
          ReDim Preserve sSubDirs(1 To lngIndex)
          sSubDirs(lngIndex) = sPath & sDir & "\"
        End If
      End If
      sDir = Dir
    Loop
    For lngTemp = 1 To lngIndex
      '���õݹ鷽������ÿһ����Ŀ¼���ļ�
      Call TreeSearch(sSubDirs(lngTemp), sFileSpec, sFiles())
    Next lngTemp
    TreeSearch = lngFiles
    Public_sFiles = sFiles
End Function

