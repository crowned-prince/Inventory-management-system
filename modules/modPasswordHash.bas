Attribute VB_Name = "modPasswordHash"
Option Explicit
' ============================================================
' modPasswordHash - secure password storage (added in IMS v1.1.0)
'
' Passwords are stored as PBKDF2-HMAC-SHA256 with a unique random 16-byte salt
' and 600,000 iterations. The work is done by the Windows CNG library
' (bcrypt.dll), so no cryptography is written in VBA.
'
' Stored format (plain text in the Password column):
'   PBKDF2-SHA256$<iterations>$<salt hex>$<hash hex>
'
' Old IMS accounts keep their 16-character checksum from modUsers.HashPassword.
' They still log in, and are upgraded to the new format at their next
' successful login (see UpgradeStoredHash).
'
' Needs Excel 2010 or later on Windows (the application already needs Windows).
' If secure hashing is not available, CreateHash raises an error. It NEVER
' falls back to a weaker hash.
' ============================================================
Private Const HASH_TAG As String = "PBKDF2-SHA256"
Private Const CURRENT_ITERATIONS As Long = 600000
Private Const MIN_ITERATIONS As Long = 10000    ' sanity limits for stored values
Private Const MAX_ITERATIONS As Long = 5000000
Private Const SALT_LEN As Long = 16
Private Const KEY_LEN As Long = 32
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = &H8
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = &H2
Private Const CP_UTF8 As Long = 65001

#If VBA7 Then
 Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
  ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, _
  ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
 Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
  ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
 Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
  ByVal hAlgorithm As LongPtr, ByRef pbBuffer As Byte, _
  ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
 #If Win64 Then
  Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" ( _
   ByVal hPrf As LongPtr, ByRef pbPassword As Byte, ByVal cbPassword As Long, _
   ByRef pbSalt As Byte, ByVal cbSalt As Long, ByVal cIterations As LongLong, _
   ByRef pbDerivedKey As Byte, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
 #Else
  ' 32-bit Excel: the 64-bit iteration count is passed as two 32-bit halves (low, high).
  Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" ( _
   ByVal hPrf As LongPtr, ByRef pbPassword As Byte, ByVal cbPassword As Long, _
   ByRef pbSalt As Byte, ByVal cbSalt As Long, ByVal cIterationsLow As Long, _
   ByVal cIterationsHigh As Long, ByRef pbDerivedKey As Byte, _
   ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
 #End If
 Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
  ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, _
  ByVal cchWideChar As Long, ByRef lpMultiByteStr As Byte, ByVal cbMultiByte As Long, _
  ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
 Private Declare PtrSafe Function WideCharToMultiByteSize Lib "kernel32" _
  Alias "WideCharToMultiByte" ( _
  ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, _
  ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
  ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#End If

' ============================================================
' Public interface
' ============================================================

' Returns the text to store in the Password column. Raises an error if secure
' hashing is not available (callers must not store anything in that case).
Public Function CreateHash(ByVal password As String) As String
 Dim salt(0 To SALT_LEN - 1) As Byte
 Dim pw() As Byte
 Dim key() As Byte
 Dim pwLen As Long
 Dim status As Long
 If Len(password) = 0 Then
  Err.Raise vbObjectError + 602, "modPasswordHash", "A password is required."
 End If
 status = BCryptGenRandom(0, salt(0), SALT_LEN, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
 If status <> 0 Then
  Err.Raise vbObjectError + 603, "modPasswordHash", _
   "Secure random numbers are not available on this computer."
 End If
 pw = Utf8Bytes(password, pwLen)
 If Not DeriveKey(pw, pwLen, salt, CURRENT_ITERATIONS, key) Then
  WipeBytes pw
  Err.Raise vbObjectError + 604, "modPasswordHash", _
   "Secure password hashing is not available on this computer."
 End If
 CreateHash = HASH_TAG & "$" & CStr(CURRENT_ITERATIONS) & "$" & _
  BytesToHex(salt) & "$" & BytesToHex(key)
 WipeBytes pw
 WipeBytes key
End Function

' True if the password matches the stored value (new format OR old IMS format).
' Never raises: any problem means "does not match".
Public Function VerifyPassword(ByVal password As String, ByVal stored As String) As Boolean
 Dim parts() As String
 Dim iterations As Long
 Dim salt() As Byte
 Dim expected() As Byte
 Dim actual() As Byte
 Dim pw() As Byte
 Dim pwLen As Long
 Dim ok As Boolean
 On Error GoTo Fail
 If Len(password) = 0 Or Len(stored) = 0 Then Exit Function
 If IsLegacyHash(stored) Then
  VerifyPassword = (StrComp(stored, modUsers.HashPassword(password), vbBinaryCompare) = 0)
  Exit Function
 End If
 parts = Split(stored, "$")
 If UBound(parts) <> 3 Then Exit Function
 If parts(0) <> HASH_TAG Then Exit Function
 If Not IsNumeric(parts(1)) Then Exit Function
 iterations = CLng(parts(1))
 If iterations < MIN_ITERATIONS Or iterations > MAX_ITERATIONS Then Exit Function
 salt = HexToBytes(parts(2), ok)
 If Not ok Then Exit Function
 If (UBound(salt) - LBound(salt) + 1) < 8 Then Exit Function
 expected = HexToBytes(parts(3), ok)
 If Not ok Then Exit Function
 If (UBound(expected) - LBound(expected) + 1) <> KEY_LEN Then Exit Function
 pw = Utf8Bytes(password, pwLen)
 If DeriveKey(pw, pwLen, salt, iterations, actual) Then
  VerifyPassword = ConstantTimeEqual(actual, expected)
 End If
 WipeBytes pw
 WipeBytes actual
 Exit Function
Fail:
 VerifyPassword = False
 WipeBytes pw
End Function

' True for the old 16-character IMS checksum.
Public Function IsLegacyHash(ByVal stored As String) As Boolean
 If Len(stored) <> 16 Then Exit Function
 IsLegacyHash = Not (stored Like "*[!0-9A-Fa-f]*")
End Function

' True when the stored value should be replaced by a current-format hash.
Public Function NeedsRehash(ByVal stored As String) As Boolean
 Dim parts() As String
 If Len(stored) = 0 Then Exit Function
 If IsLegacyHash(stored) Then
  NeedsRehash = True
  Exit Function
 End If
 On Error GoTo Done
 parts = Split(stored, "$")
 If UBound(parts) = 3 Then
  If parts(0) = HASH_TAG And IsNumeric(parts(1)) Then
   NeedsRehash = (CLng(parts(1)) < CURRENT_ITERATIONS)
  End If
 End If
Done:
End Function

' Called after a SUCCESSFUL login: silently upgrades an old hash to the new format.
' If the upgrade fails for any reason the old hash stays valid, so nobody is locked out.
Public Sub UpgradeStoredHash(ByVal userID As String, ByVal password As String, _
 ByVal storedHash As String)
 If Not NeedsRehash(storedHash) Then Exit Sub
 On Error Resume Next
 modUserAdmin.SetUserPassword userID, password
 On Error GoTo 0
End Sub

' Run from the VBA editor (click inside, press F5) to check that secure hashing
' works on this computer. Shows PASS or FAIL for each check.
Public Sub RunPasswordHashSelfTest()
 Dim pw() As Byte, salt() As Byte, key() As Byte
 Dim n As Long, sn As Long
 Dim report As String
 Dim h1 As String, h2 As String
 On Error GoTo Fail
 ' 1. Known-answer test: PBKDF2-HMAC-SHA256("password","salt",4096,32)
 pw = Utf8Bytes("password", n)
 salt = Utf8Bytes("salt", sn)
 If DeriveKey(pw, n, salt, 4096, key) Then
  report = report & Mark(LCase$(BytesToHex(key)) = _
   "c5e478d59288c841aa530db6845c4c8d962893a001ce4e11a4963873aa98134a") & _
   "PBKDF2 known-answer test" & vbCrLf
 Else
  report = report & "FAIL  PBKDF2 known-answer test (bcrypt.dll call failed)" & vbCrLf
 End If
 ' 2. Round trip, wrong password, unique salts
 h1 = CreateHash("Correct Horse 7!")
 h2 = CreateHash("Correct Horse 7!")
 report = report & Mark(VerifyPassword("Correct Horse 7!", h1)) & "Correct password accepted" & vbCrLf
 report = report & Mark(Not VerifyPassword("Correct Horse 8!", h1)) & "Wrong password rejected" & vbCrLf
 report = report & Mark(h1 <> h2) & "Each hash gets its own salt" & vbCrLf
 report = report & Mark(Not NeedsRehash(h1)) & "New hash does not need upgrading" & vbCrLf
 ' 3. Old format still works
 report = report & Mark(VerifyPassword("abc12345", modUsers.HashPassword("abc12345"))) & _
  "Old-format password accepted" & vbCrLf
 report = report & Mark(Not VerifyPassword("abc12346", modUsers.HashPassword("abc12345"))) & _
  "Old-format wrong password rejected" & vbCrLf
 report = report & Mark(NeedsRehash(modUsers.HashPassword("abc12345"))) & _
  "Old format is flagged for upgrade" & vbCrLf
 MsgBox report, vbInformation, "Password hashing self-test"
 Exit Sub
Fail:
 MsgBox "The self-test stopped: error " & Err.Number & vbCrLf & report, vbExclamation, _
  "Password hashing self-test"
End Sub

' ============================================================
' Private helpers
' ============================================================
Private Function Mark(ByVal passed As Boolean) As String
 If passed Then Mark = "PASS  " Else Mark = "FAIL  "
End Function

Private Function DeriveKey(ByRef pw() As Byte, ByVal pwLen As Long, ByRef salt() As Byte, _
 ByVal iterations As Long, ByRef key() As Byte) As Boolean
 Dim hAlg As LongPtr
 Dim algId As String
 Dim status As Long
 algId = "SHA256"
 status = BCryptOpenAlgorithmProvider(hAlg, StrPtr(algId), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG)
 If status <> 0 Then Exit Function
 ReDim key(0 To KEY_LEN - 1)
 #If Win64 Then
  status = BCryptDeriveKeyPBKDF2(hAlg, pw(0), pwLen, salt(0), _
   UBound(salt) - LBound(salt) + 1, CLngLng(iterations), key(0), KEY_LEN, 0)
 #Else
  status = BCryptDeriveKeyPBKDF2(hAlg, pw(0), pwLen, salt(0), _
   UBound(salt) - LBound(salt) + 1, iterations, 0, key(0), KEY_LEN, 0)
 #End If
 BCryptCloseAlgorithmProvider hAlg, 0
 DeriveKey = (status = 0)
End Function

' UTF-8 bytes of the text. byteCount receives the real length.
Private Function Utf8Bytes(ByVal text As String, ByRef byteCount As Long) As Byte()
 Dim buf() As Byte
 Dim n As Long
 byteCount = 0
 If Len(text) = 0 Then
  ReDim buf(0 To 0)
  Utf8Bytes = buf
  Exit Function
 End If
 n = WideCharToMultiByteSize(CP_UTF8, 0, StrPtr(text), Len(text), 0, 0, 0, 0)
 If n <= 0 Then Err.Raise vbObjectError + 605, "modPasswordHash", "Text conversion failed."
 ReDim buf(0 To n - 1)
 If WideCharToMultiByte(CP_UTF8, 0, StrPtr(text), Len(text), buf(0), n, 0, 0) <> n Then
  Err.Raise vbObjectError + 606, "modPasswordHash", "Text conversion failed."
 End If
 byteCount = n
 Utf8Bytes = buf
End Function

Private Function BytesToHex(ByRef b() As Byte) As String
 Dim i As Long
 Dim s As String
 For i = LBound(b) To UBound(b)
  s = s & Right$("0" & Hex$(b(i)), 2)
 Next i
 BytesToHex = s
End Function

Private Function HexToBytes(ByVal hexText As String, ByRef ok As Boolean) As Byte()
 Dim b() As Byte
 Dim i As Long
 ok = False
 If Len(hexText) = 0 Or (Len(hexText) Mod 2) <> 0 Then Exit Function
 If hexText Like "*[!0-9A-Fa-f]*" Then Exit Function
 ReDim b(0 To (Len(hexText) \ 2) - 1)
 For i = 0 To UBound(b)
  b(i) = CByte("&H" & Mid$(hexText, (i * 2) + 1, 2))
 Next i
 ok = True
 HexToBytes = b
End Function

' Compares two byte arrays without stopping at the first difference.
Private Function ConstantTimeEqual(ByRef a() As Byte, ByRef b() As Byte) As Boolean
 Dim diff As Long
 Dim i As Long
 If (UBound(a) - LBound(a)) <> (UBound(b) - LBound(b)) Then Exit Function
 For i = 0 To UBound(a) - LBound(a)
  diff = diff Or (CLng(a(LBound(a) + i)) Xor CLng(b(LBound(b) + i)))
 Next i
 ConstantTimeEqual = (diff = 0)
End Function

' Overwrites secret bytes. Safe to call on an array that was never allocated.
Private Sub WipeBytes(ByRef b() As Byte)
 Dim i As Long
 On Error Resume Next
 For i = LBound(b) To UBound(b)
  b(i) = 0
 Next i
 On Error GoTo 0
End Sub
