VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Net Utils Tester"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   345
      Left            =   3090
      TabIndex        =   2
      Top             =   60
      Width           =   735
   End
   Begin VB.TextBox txtServerName 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   60
      Width           =   2895
   End
   Begin VB.ListBox lstUsers 
      Height          =   3570
      ItemData        =   "frmMain.frx":0000
      Left            =   90
      List            =   "frmMain.frx":0002
      TabIndex        =   0
      Top             =   450
      Width           =   3705
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_COMPUTERNAME As Long = 15

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub cmdGo_Click()
    'For those of you serious about speed (even though this app dosen't really need it), you shoudl try
    'and do the following ALL the time. Never simply declare an object as a new instance of a class, always
    'declare the variable as the type you wish, then set it to be a new instance as I have done here. The
    'Visual Basic Runtime keeps track of how a variable was declared, and if it was done in one step (not
    'what I have done here.. the other way), if check to make sure it is NOT Nothing EVERY TIME YOU USE IT!
    
    '-----------------------------------------------------------------------------------------------
    Dim objUsers As NetUtils.Users: Set objUsers = New NetUtils.Users
    
    Dim intCount As Integer
    
    Dim colUsers As Collection
    '-----------------------------------------------------------------------------------------------
    
    If Len(txtServerName.Text) < 0 Then Exit Sub
    
    lstUsers.Clear
    
    Set colUsers = objUsers.Get_Remote_Users_List(txtServerName.Text)
    
    For intCount = 1 To colUsers.Count
        lstUsers.AddItem colUsers.Item(intCount)
    Next
End Sub

Public Function Local_System_Name() As String
    Dim strTemp As String
    
    strTemp = Space$(MAX_COMPUTERNAME + 1)
    
    If GetComputerName(strTemp, Len(strTemp)) <> 0 Then
        Local_System_Name = TrimNull(strTemp)
    End If
End Function

Private Function TrimNull(strItem As String)
    Dim intPosition As Integer
   
    intPosition = InStr(strItem, Chr$(0))
    
    If intPosition Then
        TrimNull = Left$(strItem, intPosition - 1)
    Else
        TrimNull = strItem
    End If
End Function

Private Sub Form_Load()
    txtServerName.Text = Local_System_Name()
End Sub
