VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncriptPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GeneralKey() As String
       Dim StrGeneralKey As String
       StrGeneralKey = "ABC#" & _
                       "BCD#" & _
                       "CDE#" & _
                       "DEF#" & _
                       "EFG#"
       GeneralKey = StrGeneralKey
End Function

Public Function RandString(strLen As Integer) As String
       On Error Resume Next
       Dim xChar As Integer
       Randomize
       While strLen <> 0
             xChar = Rnd * 255
             If IsAlphaNumeric(xChar) = True Then
                RandString = RandString & Chr(xChar)
                strLen = strLen - 1
             End If
       Wend
End Function

Private Function IsAlphaNumeric(CharCode As Integer) As Boolean
        On Error Resume Next
        If (CharCode >= 65 And CharCode <= 90) Then
           IsAlphaNumeric = True
        ElseIf (CharCode >= 97 And CharCode <= 122) Then
           IsAlphaNumeric = True
        ElseIf (CharCode >= 48 And CharCode <= 57) Then
           IsAlphaNumeric = True
        Else
           IsAlphaNumeric = False
        End If
End Function

Public Function EncryptText(strText As String, ByVal strPwd As String) As String
3      Dim i As Integer, c As Integer
4      Dim strBuff As String
5      If Len(strPwd) Then
6         For i = 1 To Len(strText)
7             c = Asc(Mid$(strText, i, 1))
8             c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
9             strBuff = strBuff & Chr$(c And &HFF)
10        Next i
11     Else
12         strBuff = strText
13     End If
14     EncryptText = strBuff
End Function

Public Function DecryptText(strText As String, ByVal strPwd As String) As String
3      Dim i As Integer, c As Integer
4      Dim strBuff As String
5      If Len(strPwd) Then
6         For i = 1 To Len(strText)
7             c = Asc(Mid$(strText, i, 1))
8             c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
9             strBuff = strBuff & Chr$(c And &HFF)
10        Next i
11     Else
12        strBuff = strText
13     End If
14     DecryptText = strBuff
15     Exit Function
End Function

Public Function DBZEncriptPassword(Password As String) As String
       Dim TEMP
       Dim i As Integer
       Dim ENCPassword As String
       
       i = 5 * Rnd
       TEMP = Split(GeneralKey, "#")
       
       ENCPassword = "": ENCPassword = EncryptText(Password, CStr(TEMP(i - 1)))
       DBZEncriptPassword = ENCPassword & EncryptText(CStr(Format(i, "00")), "ONS")
End Function

Public Function DBZDecriptPassword(Password As String) As String
       Dim TEMP
       Dim i As Integer
       Dim ENCPassword As String
       
       i = Val(DecryptText(Right(Password, 2), "ONS"))
       TEMP = Split(GeneralKey, "#")
       
       ENCPassword = "": ENCPassword = DecryptText(Left(Password, (Len(Password) - 2)), CStr(TEMP(i - 1)))
       DBZDecriptPassword = ENCPassword
End Function

