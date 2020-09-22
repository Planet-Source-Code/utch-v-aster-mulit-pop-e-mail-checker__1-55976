Attribute VB_Name = "mMyconsole"
Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetInputState Lib "user32" () As Long

Public m_State()        As Integer
Public m_Ready()        As Boolean
Public m_Server()       As String
Public m_User()         As String
Public m_Password()     As String
Public m_Title()        As String

Public CurrAccount      As Integer
Public Adding           As Boolean
Public FirstRun         As Boolean
Public inTray           As Boolean
Public LastSum          As Integer
Public DB               As Database

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2

Public Enum POP3States
  POP3_Connect
  POP3_USER
  POP3_PASS
  POP3_STAT
  POP3_TOP
  POP3_RETR
  POP3_DELE
  POP3_QUIT
End Enum

Public Sub Pause(Seconds As Single)
  Dim T   As Single
  Dim T2  As Single
  Dim Num As Single
  
  Num = Seconds * 1000
  T = GetTickCount()
  T2 = GetTickCount()
  
  Do Until T2 - T >= Num
    
    If GetInputState <> 0 Then DoEvents
    T2 = GetTickCount()
  Loop
End Sub

Public Sub PlaySound(Filename As String)
  Dim wFlags%, X, SoundName As String
  SoundName$ = Filename
  wFlags% = SND_ASYNC Or SND_NODEFAULT
  X = sndPlaySound(SoundName$, wFlags%)
End Sub

Public Function Encrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim I     As Integer
    Dim Char  As String
    
    Encrypt = ""
    If StringToEncrypt = "" Then Exit Function

    For I = 1 To Len(StringToEncrypt)
      Char = Asc(Mid(StringToEncrypt, I, 1))
      Encrypt = Encrypt & Len(Char) & Char
    Next I
    
    If AlphaEncoding Then
      StringToEncrypt = Encrypt
      Encrypt = ""
      For I = 1 To Len(StringToEncrypt)
        Encrypt = Encrypt & Chr(Mid(StringToEncrypt, I, 1) + 147)
      Next I
    End If
    Exit Function

ErrorHandler:
    Encrypt = ""
End Function

Public Function Decrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
  On Error GoTo ErrorHandler
  Dim I         As Integer
  Dim CharCode  As String
  Dim CharPos   As Integer
  Dim Char      As String
  
  If StringToDecrypt = "" Then Exit Function

  If AlphaDecoding Then
    Decrypt = StringToDecrypt
    StringToDecrypt = ""
    For I = 1 To Len(Decrypt)
      StringToDecrypt = StringToDecrypt & (Asc(Mid(Decrypt, I, 1)) - 147)
    Next I
  End If
  
  Decrypt = ""
  
  Do Until StringToDecrypt = ""
    CharPos = Left(StringToDecrypt, 1)
    StringToDecrypt = Mid(StringToDecrypt, 2)
    CharCode = Left(StringToDecrypt, CharPos)
    StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
    Decrypt = Decrypt & Chr(CharCode)
  Loop
  
  Exit Function
    
ErrorHandler:
    Decrypt = ""
End Function

Public Function Get_After_Seperator(ByVal strString As String, ByVal intNthOccurance As Integer, ByVal strSeperator As String) As String
  Dim startOfString As Integer
  Dim itemIndex     As Integer
  Dim notFound      As Boolean
  Dim endOfString   As Integer
  If (intNthOccurance = 0) Then
    If (InStr(strString, strSeperator) > 0) Then
      Get_After_Seperator = Left(strString, InStr(strString, strSeperator) - 1)
    Else
      Get_After_Seperator = strString
    End If
  Else
    startOfString = InStr(strString, strSeperator)
    notFound = 0
    For itemIndex = 1 To intNthOccurance - 1
      startOfString = InStr(startOfString + 1, strString, strSeperator)
      If (startOfString = 0) Then
        notFound = 1
      End If
    Next itemIndex
    startOfString = startOfString + 1

    If (startOfString > Len(strString)) Then
      notFound = 1
    End If
  
    If (notFound = 1) Then
      Get_After_Seperator = "NOT FOUND"
    Else
      endOfString = InStr(startOfString, strString, strSeperator)
      If (endOfString = 0) Then
        endOfString = Len(strString) + 1
      Else
        endOfString = endOfString - 1
      End If
      Get_After_Seperator = Mid$(strString, startOfString, endOfString - startOfString + 1)
    End If
  End If
End Function

