Attribute VB_Name = "Module1"
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const MAX_PATH = 260
Private Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32" _
   Alias "lstrcatA" (ByVal lpString1 As String, _
   ByVal lpString2 As String) As Long
   
Private Declare Function SHBrowseForFolder Lib "shell32" _
   (lpBI As BrowseInfo) As Long
   
Private Declare Function SHGetPathFromIDList Lib "shell32" _
   (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'OPEN FILE DECLARATIONS===========================================
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias _
         "GetSaveFileNameA" (pSavefilename As OPENFILENAME) As Long
         
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    strFilter As String
    strCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Public Function OpenFileAPI(Optional StartingFolder As String, Optional sFilter As String, Optional dialogTitle As String, Optional savedialog As Boolean, Optional existingcontent As String)


         Dim stfol As String
         Dim findex() As String
         Dim OpenFile As OPENFILENAME
         Dim lReturn As Long
         
         findex = Split(sFilter, ",")
         sFilter = Replace(sFilter, ",", vbNullChar)
         
         OpenFile.lStructSize = Len(OpenFile)
         OpenFile.hwndOwner = Form1.hWnd
         OpenFile.hInstance = App.hInstance
         OpenFile.strFilter = sFilter
         OpenFile.nFilterIndex = 0
         OpenFile.strFile = String(257, 0)
         OpenFile.nMaxFile = Len(OpenFile.strFile) - 1
         OpenFile.strFileTitle = OpenFile.strFile
         OpenFile.nMaxFileTitle = OpenFile.nMaxFile
         stfol = "C:\Users\kavan\Desktop\Tify Sets\Fig5"
         If stfol <> "" Then
            OpenFile.strInitialDir = stfol
         Else
            OpenFile.strInitialDir = Environ("USERPROFILE")
         End If
         If dialogTitle = "" Then dialogTitle = "Open"
         OpenFile.strTitle = dialogTitle
         OpenFile.Flags = 0
         If savedialog = False Then
             lReturn = GetOpenFileName(OpenFile)
         Else
            lReturn = GetSaveFileName(OpenFile)
         End If
        
        
         If lReturn = 0 Then
            OpenFileAPI = existingcontent
            Exit Function
         Else
            
            OpenFileAPI = Replace(OpenFile.strFile, vbNullChar, "")
          
            If sFilter <> "" Then
                'work out which filter they chose
                selfil = findex((OpenFile.nFilterIndex * 2) - 1)
                rightexts = Right(selfil, 3)
                If Right(OpenFileAPI, 3) <> rightexts Then OpenFileAPI = OpenFileAPI & "." & rightexts
            End If
            
            
            
            
         
         End If

         End Function


