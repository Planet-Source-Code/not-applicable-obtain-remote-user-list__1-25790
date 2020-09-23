VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const ERROR_SUCCESS As Long = 0&

Private Declare Function NetUserEnum Lib "Netapi32" (servername As Byte, ByVal level As Long, ByVal filter As Long, buff As Long, ByVal buffsize As Long, entriesread As Long, totalentries As Long, resumehandle As Long) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32" (ByVal Buffer As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nBytes As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Public Function Get_Remote_Users_List(strServerName As String) As Collection
    Dim colReturn As Collection: Set colReturn = New Collection
    
    Dim arrServerName() As Byte

    Dim Users() As Long
    Dim buff As Long
    Dim buffsize As Long
    Dim entriesread As Long
    Dim totalentries As Long
    Dim cnt As Integer
   
    buffsize = 255
    
    If Left(strServerName, 2) <> "\\" Then strServerName = "\\" & strServerName
    
    strServerName = strServerName & Chr(0)
    
    arrServerName = strServerName
    
    If NetUserEnum(arrServerName(0), 0, FILTER_NORMAL_ACCOUNT, buff, buffsize, entriesread, totalentries, 0&) = ERROR_SUCCESS Then
        ReDim Users(0 To entriesread - 1) As Long
        
        CopyMemory Users(0), ByVal buff, entriesread * 4
    
        For cnt = 0 To entriesread - 1
            colReturn.Add GetPointerToByteStringW(Users(cnt))
        Next cnt
    
        NetApiBufferFree buff
    End If
    
    Set Get_Remote_Users_List = colReturn
End Function

Private Function GetPointerToByteStringW(lpString As Long) As String
    Dim buff() As Byte
    Dim nSize As Long
    
    If lpString Then
        nSize = lstrlenW(lpString) * 2
    
        If nSize Then
            ReDim buff(0 To (nSize - 1)) As Byte
            
            CopyMemory buff(0), ByVal lpString, nSize
            
            GetPointerToByteStringW = buff
        End If
    End If
End Function


Private Sub Form_Load()
    Dim x As Collection
    
    Dim i As Integer
    
    Set x = Get_Remote_Users_List("EXTERNAL")
    
    For i = 1 To x.Count
        Debug.Print x.Item(i)
    Next
End Sub

