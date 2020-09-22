Attribute VB_Name = "Search"
Option Explicit

'Find File declarations and types
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
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
        cFileName As String * 260
        cAlternate As String * 14
End Type
'
'Convert Time Declare and type
'Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'Public Type SYSTEMTIME
'        wYear As Integer
'        wMonth As Integer
'        wDayOfWeek As Integer
'        wDay As Integer
'        wHour As Integer
'        wMinute As Integer
'        wSecond As Integer
'        wMilliseconds As Integer
'End Type


Public Function FindAllFiles(Directory As String, Optional SearchFor As String)
    Dim Exists As Long
    Dim hFindFile As Long
    Dim FileData As WIN32_FIND_DATA
    Dim TotalFileSize As Double
    
  With Form1
    
    'Sets Exists to equal 1
    'You need this so the loop doesn't automatically exit
    
    Exists = 1
    
    'Makes sure theres a "\" at the end of the directory
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    
    'Sets the default search item to *.*
    If SearchFor = vbNullString Then SearchFor = "*.*"
    
    'If the search for text doesn't contain any * or ?
    'Add *'s before and after
    If InStr(1, SearchFor, "?") = 0 And InStr(1, SearchFor, "*") = 0 Then
        SearchFor = "*" & SearchFor & "*"
    End If
    
    'Finds the first file
    hFindFile = FindFirstFile(Directory & SearchFor, FileData)
    
    Do While hFindFile <> -1 And Exists <> 0
        'A loop until all the files have been added
        
        DoEvents
        
        If (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) _
        = vbDirectory Then
            'If the file IS a directory than add it
            'to the temp listbox with the prefix DIR
            'don't list "." and ".." dirs
            If (ClearNull(FileData.cFileName) <> ".") And (ClearNull(FileData.cFileName) <> "..") Then
               .FindFilesTmpResults.AddItem "[dir]  " & Directory & ClearNull(FileData.cFileName)
            End If
        ElseIf (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) _
        <> vbDirectory Then
            'If the file ISN'T a directory than add it
            'to the temp listbox with the prefix FILE
            .FindFilesTmpResults.AddItem "[file] " & Directory & ClearNull(FileData.cFileName)
            'get file size
            TotalFileSize = TotalFileSize + (FileLen(Directory & ClearNull(FileData.cFileName)) / 1024)
            Form1.Text7 = Format(TotalFileSize, "########,##")
            Form1.Text8 = Val(Form1.Text8) + (FileLen(Directory & ClearNull(FileData.cFileName)))
        End If
        
        'Finds next file
        Exists = FindNextFile(hFindFile, FileData)
    Loop
    
    Do While .FindFilesTmpResults.ListCount
        'Removes everything from the temp listbox (Which is
        'alphabetically sorted, and puts it into the Viewed
        'Listbox
        'This is done so all the files are sorted alphabetically
        .FindFilesResults.AddItem .FindFilesTmpResults.List(0)
        Form1.Frame1.Caption = "Building File List: " & .FindFilesResults.ListCount & " Files"
        .FindFilesTmpResults.RemoveItem 0
    Loop
    
    'Sets Exists to equal 1
    'You need this so the loop doesn't automatically exit
    Exists = 1
    
    'Find first file, this time includes directories in
    'the search
    hFindFile = FindFirstFile(Directory & "*", FileData)
    
    Do While hFindFile <> -1 And Exists <> 0
        'A loop until all the files have been added
      On Error GoTo skiptonextfile
        If (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) _
        = vbDirectory And (ClearNull(FileData.cFileName) <> "." And ClearNull(FileData.cFileName) <> "..") Then
           'If the file IS a directory and isn't "." or ".."
             'than adds it to the temp dir listbox
            
            .FindFilesTmpDirs.AddItem Directory & ClearNull(FileData.cFileName)
            DoEvents
        End If
nextfile:
      On Error GoTo 0
        
        'Finds next file
        Exists = FindNextFile(hFindFile, FileData)
    Loop

  End With
  
  Exit Function
  
skiptonextfile:
 Err.Clear
 Resume nextfile


End Function

Public Function ClearNull(StringToClear As String) As String
    Dim StartOfNulls As Long
    
    'This function clears all the nulls in the string and
    'Returns it, by using Instr to find the first null
    
    StartOfNulls = InStr(1, StringToClear, Chr(0))
    ClearNull = Left(StringToClear, StartOfNulls - 1)
End Function

Public Sub SearchFilesInDir(ByVal Directory As String, Optional SearchFor As String)

 Dim NextDir As String
    
  With Form1
    'Clears the result listbox
    .FindFilesResults.Clear
    .FindFilesTmpResults.Clear
    'Calls the FindAllFiles function
    FindAllFiles Directory, SearchFor

    Do While .FindFilesTmpDirs.ListCount
        'Searches through all the new directories and removes
        'Them from the temp dir listbox
        
        DoEvents
        NextDir = .FindFilesTmpDirs.List(0)
        .FindFilesTmpDirs.RemoveItem 0
        FindAllFiles NextDir, SearchFor
    Loop
    
  End With
End Sub

