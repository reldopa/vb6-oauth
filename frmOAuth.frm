VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOAuth 
   Caption         =   "UserForm1"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   OleObjectBlob   =   "frmOAuth.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Version 1.0
Private fso As New FileSystemObject
Private tokenFile As String
Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private html As New HTMLDocument
Private Type SYSTEMTIME
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type
Private SessionHandle As String
Private TokenExpiration As Date
Private TokenSecret As String
Private Token As String
Public http As WinHttpRequest
Private m_resource As String
Attribute m_resource.VB_VarHelpID = -1
Const EXT_TOKEN As String = ".tok"
Private Const CONSUMER_KEY As String = "<insert here>"    'change
Private Const CONSUMER_SECRET As String = "<insert here>"    'change
Const guid As String = "F317EA682B9141af8D95007A5C73F9CE"

Private Function Base64_HMACSHA1(ByVal sTextToHash As String, ByVal sSharedSecretKey As String)

    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.key = SharedSecretKey

    Dim bytes() As Byte
    bytes = enc.ComputeHash_2((TextToHash))
    Base64_HMACSHA1 = EncodeBase64(bytes)
    Set asc = Nothing
    Set enc = Nothing

End Function

Public Function ClearCredentials() As Boolean
    If fso.FileExists(tokenFile) Then
        fso.DeleteFile
        ClearCredentials = True
    Else
        ClearCredentials = False
    End If
End Function

Function dict2string(dic As Dictionary) As String
    Dim key
    For Each key In dic.Keys
        dict2string = dict2string & "&" & key & "=" & dic.Item(key)


    Next key
    dict2string = Mid(dict2string, 2)

End Function

Public Function doAuth()
    If Not LoadToken Then
        Login
    End If
End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")

    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing

End Function

' /// <summary>
'/// Generate the signature base that is used to produce the signature
'/// </summary>
'/// <param name="url">The full url that needs to be signed including its non OAuth url parameters</param>
'/// <param name="consumerKey">The consumer key</param>
'/// <param name="token">The token, if available. If not available pass null or an empty string</param>
'/// <param name="tokenSecret">The token secret, if available. If not available pass null or an empty string</param>
'/// <param name="httpMethod">The http method used. Must be a valid HTTP method verb (POST,GET,PUT, etc)</param>
'/// <param name="signatureType">The signature type. To use the default values use <see cref="OAuthBase.SignatureTypes">OAuthBase.SignatureTypes</see>.</param>
'/// <returns>The signature base</returns>
Public Function GenerateSignatureBase(URL As MSHTML.HTMLAnchorElement, consumerKey As String, Token As String, httpMethod As String, _
                                      timeStamp As String, nonce As String, hmac As Boolean, ByRef normalizedUrl As String, ByRef normalizedRequestParameters As String) As String

    On Error GoTo GenerateSignatureBase_Error
    If "" = consumerKey Then Err.Raise 5, ("consumerKey")
    If "" = httpMethod Then Err.Raise 5, ("httpMethod")
    normalizedUrl = ""
    normalizedRequestParameters = ""
    Dim parameters As Dictionary
    Set parameters = GetQueryParameters(URL.Search)
    parameters.Add OAUTH_VERSION, "1.0"
    parameters.Add OAUTH_NONCE, nonce
    parameters.Add OAUTH_TIMESTAMP, timeStamp
    parameters.Add OAUTH_SIGNATURE_METHOD, IIf(hmac, "HMAC-SHA1", "PLAINTEXT")
    parameters.Add OAUTH_CONSUMER_KEY, CONSUMER_KEY
    If Token <> "" Then parameters.Add OAUTH_TOKEN, Token
    SortDictionary parameters
    normalizedUrl = URL.Protocol & "//" & URL.hostname
    If (Not ((URL.Protocol = "http:" And URL.Port = 80) Or (URL.Protocol = "https:" And URL.Port = 443))) Then
        normalizedUrl = normalizedUrl & ":" & URL.Port
    End If
    normalizedUrl = normalizedUrl & "/" & URL.pathname
    normalizedRequestParameters = NormalizeRequestParameters(parameters)
    GenerateSignatureBase = UCase$(httpMethod) & "&" _
                          & URLEncode(normalizedUrl) & "&" _
                          & URLEncode(normalizedRequestParameters)

    On Error GoTo 0
    Exit Function

GenerateSignatureBase_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GenerateSignatureBase of Class Module CoAuth"
End Function

Function get_oauth_nonce() As String
    Dim cString As String
    cString = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim nlen As Integer
    nlen = Len(cString)
    Dim cOauth_nonce As String
    cOauth_nonce = ""
    Dim nCount As Integer, nRand As Integer
    For nCount = 1 To 41
        nRand = Int((41 * Rnd) + 1)
        cOauth_nonce = cOauth_nonce + Mid(cString, nRand, 1)
    Next

    get_oauth_nonce = cOauth_nonce

End Function

Function get_oauth_timestamp()
    Dim nOAUTH_TIMESTAMP
    Dim t As SYSTEMTIME
    Call GetSystemTime(t)
    With t
        nOAUTH_TIMESTAMP = DateDiff("s", #1/1/1970#, DateSerial(.Year, .Month, .Day) + TimeSerial(.Hour, .Minute, .Second))
    End With
    get_oauth_timestamp = nOAUTH_TIMESTAMP

End Function

'        /// <summary>
'        /// Internal function to cut out all non oauth query string parameters (all parameters not begining with "oauth_")
'        /// </summary>
'        /// <param name="parameters">The query string part of the Url</param>
'        /// <returns>A list of QueryParameter each containing the parameter name and value</returns>
Private Function GetQueryParameters(parameters As String) As Dictionary

    If Left$(parameters, 1) = "?" Then
        parameters = Mid$(parameters, 2)
    End If

    Dim result As New Dictionary
    Dim temp
    If (parameters <> "") Then
        Dim p As Variant
        p = Split(parameters, "&")
        Dim s
        For Each s In p
            If (s <> "") Then  'And Not StringStartsWith(s, OAuthParameterPrefix, vbTextCompare)) Then
                If InStr(s, "=") > 0 Then
                    temp = Split(s, "=")
                    Call result.Add(temp(0), temp(1))
                Else
                    Call result.Add(s, "")
                End If
            End If
        Next s
    End If

    Set GetQueryParameters = result

End Function

Function HandleToken()
    If http.status = 200 Then
        Dim ts As TextStream
        Set ts = fso.CreateTextFile(tokenFile, True)
        Dim accessDict As Dictionary
        Set accessDict = str2dict(http.responseText)
        Token = URLEncode(accessDict(OAUTH_TOKEN))
        ts.WriteLine Token
        TokenSecret = accessDict(OAUTH_TOKEN_SECRET)
        ts.WriteLine TokenSecret
        SessionHandle = accessDict(OAUTH_SESSION_HANDLE)
        ts.WriteLine SessionHandle
        TokenExpiration = DateAdd("s", accessDict(OAUTH_EXPIRES_IN), Now)
        ts.WriteLine TokenExpiration
        ts.Close
        Set ts = Nothing
    End If
End Function

Private Function InternalSignedRequest(URL As String, hmac As Boolean, Optional method As String = "GET", Optional data As String = "", Optional ignoreExpire As Boolean = False)
    Dim rqURL As String, parameters As String
    Dim a As IHTMLAnchorElement
    If Not ignoreExpire Then
        If TokenExpiration < Now Then RefreshToken
    End If
    Set a = MakeURL(URL)
    Dim base As String
    base = Me.GenerateSignatureBase(a, CONSUMER_KEY, Token, method, get_oauth_timestamp, get_oauth_nonce, hmac, rqURL, parameters)
    Dim sig As String, secrets As String
    secrets = CONSUMER_SECRET & "&" & TokenSecret
    If hmac Then
        sig = Base64_HMACSHA1(base, secrets)
    Else
        sig = URLEncode(secrets)
    End If
    http.SetAutoLogonPolicy AutoLogonPolicy_Always
    http.Open UCase$(method), rqURL & "?" & parameters & "&" & OAUTH_SIGNATURE & "=" & URLEncode(sig), False
    http.send data
    'outResponse = http.responseText
    Set InternalSignedRequest = http

End Function

Function LoadToken() As Boolean
    If fso.FileExists(tokenFile) Then
        Dim ts As TextStream
        Set ts = fso.OpenTextFile(tokenFile, ForReading)
        Token = ts.ReadLine
        TokenSecret = ts.ReadLine
        SessionHandle = ts.ReadLine
        TokenExpiration = CDate(ts.ReadLine)
        ts.Close
        Set ts = Nothing
        LoadToken = True
    Else
        LoadToken = False
    End If

End Function

Public Function Login()
    Dim callback As String
    callback = RequestToken

    Dim verifierURL As String
    WebBrowser1.Navigate callback
    Show    'this will block until DocComp hides the form, at which point we get control and should be able to just
    ' grab the url ourselves


    verifierURL = WebBrowser1.LocationURL
    Dim verifier As Dictionary
    Set verifier = str2dict(Mid$(MakeURL(verifierURL).Search, 2))
    SignedRequest URL_ACCESS_TOKEN & "?" & OAUTH_VERIFIER & "=" & verifier(OAUTH_VERIFIER), True
    HandleToken
End Function

Function MakeURL(URL As String) As IHTMLAnchorElement
    Dim a As IHTMLAnchorElement
    Set a = html.createElement("A")
    a.href = URL
    Set MakeURL = a
End Function

'        /// <summary>
'        /// Normalizes the request parameters according to the spec
'        /// </summary>
'        /// <param name="parameters">The list of parameters already sorted</param>
'        /// <returns>a string representing the normalized parameters</returns>

Friend Function NormalizeRequestParameters(SortedParameters As Dictionary) As _
       String
    Dim sb As String
    Dim i As Integer
    Dim key As String, val As String
    For i = 0 To SortedParameters.Count - 1
        key = SortedParameters.Keys(i)
        val = SortedParameters.Item(key)
        'sb.AppendFormat("{0}={1}", p.Name, p.Value);
        sb = sb & (key) & "=" & (val) & IIf(i < SortedParameters.Count - 1, "&", "")
    Next i

    NormalizeRequestParameters = sb

End Function

Function RefreshToken()
    Set RefreshToken = InternalSignedRequest(URL_ACCESS_TOKEN & "?" & OAUTH_SESSION_HANDLE & "=" & SessionHandle, True, "GET", "", True)
    HandleToken
End Function

Public Function RequestToken() As String
    On Error GoTo RequestToken_Error
    With MakeURL(URL_REQUEST_TOKEN)
        SignedRequest .href & "?" & OAUTH_CALLBACK & "=" & URLEncode(.Protocol & "//" & .Host & "/"), True
    End With
    If http.status <> 200 Then GoTo RequestToken_Error
    Dim response As Scripting.Dictionary
    Set response = str2dict(http.responseText)
    Token = response(OAUTH_TOKEN)
    TokenSecret = response(OAUTH_TOKEN_SECRET)
    TokenExpiration = DateAdd("s", response(OAUTH_EXPIRES_IN), Now)
    RequestToken = response(XOAUTH_REQUEST_AUTH_URL)
    On Error GoTo 0
    Exit Function

RequestToken_Error:

    MsgBox "Error " & http.status & " (" & http.statusText & ") in procedure RequestToken of Class Module CoAuth" & vbCr & http.responseText
    End
End Function

Public Function SignedRequest(URL As String, hmac As Boolean, Optional method As String = "GET", Optional data As String = "")
    Set SignedRequest = InternalSignedRequest(URL, hmac, method, data, False)
End Function

Function SortDictionary(objDict)
    Const dictKey = 1
    Const dictItem = 2

    ' declare our variables
    Const intSort = 1
    Dim strDict()
    Dim objKey
    Dim strKey, strItem
    Dim X, Y, Z

    ' get the dictionary count
    Z = objDict.Count

    ' we need more than one item to warrant sorting
    If Z > 1 Then
        ' create an array to store dictionary information
        ReDim strDict(Z, 2)
        X = 0
        ' populate the string array
        For Each objKey In objDict
            strDict(X, dictKey) = CStr(objKey)
            strDict(X, dictItem) = CStr(objDict(objKey))
            X = X + 1
        Next

        ' perform a a shell sort of the string array
        For X = 0 To (Z - 2)
            For Y = X To (Z - 1)
                If StrComp(strDict(X, intSort), strDict(Y, intSort), vbTextCompare) > 0 Then
                    strKey = strDict(X, dictKey)
                    strItem = strDict(X, dictItem)
                    strDict(X, dictKey) = strDict(Y, dictKey)
                    strDict(X, dictItem) = strDict(Y, dictItem)
                    strDict(Y, dictKey) = strKey
                    strDict(Y, dictItem) = strItem
                End If
            Next
        Next

        ' erase the contents of the dictionary object
        objDict.RemoveAll

        ' repopulate the dictionary with the sorted information
        For X = 0 To (Z - 1)
            objDict.Add strDict(X, dictKey), strDict(X, dictItem)
        Next

    End If

End Function

Function str2dict(str As String) As Dictionary
    Set str2dict = New Dictionary
    Dim kvs As Variant
    kvs = Split(str, "&")
    Dim i As Integer
    Dim kv As Variant
    For i = LBound(kvs) To UBound(kvs)
        kv = Split(kvs(i), "=")
        str2dict.Add kv(0), URLDecode(CStr(kv(1)))
    Next i
End Function

Public Function StringStartsWith(ByVal strValue As String, _
                                 CheckFor As String, Optional CompareType As VbCompareMethod _
                                                   = vbBinaryCompare) As Boolean

'Determines if a string starts with the same characters as
'CheckFor string

'True if starts with CheckFor, false otherwise
'Case sensitive by default.  If you want non-case sensitive, set
'last parameter to vbTextCompare

'Examples:
'MsgBox StringStartsWith("Test", "TE") 'false
'MsgBox StringStartsWith("Test", "TE", vbTextCompare) 'True

    Dim sCompare As String
    Dim lLen As Long

    lLen = Len(CheckFor)
    If lLen > Len(strValue) Then Exit Function
    sCompare = Left(strValue, lLen)
    StringStartsWith = StrComp(sCompare, CheckFor, CompareType) = 0

End Function

Public Function URLDecode(StringToDecode As String) As String

    Dim TempAns As String
    Dim CurChr As Integer

    CurChr = 1

    Do Until CurChr - 1 = Len(StringToDecode)
        Select Case Mid(StringToDecode, CurChr, 1)
        Case "+"
            TempAns = TempAns & " "
        Case "%"
            TempAns = TempAns & Chr(val("&h" & _
                                        Mid(StringToDecode, CurChr + 1, 2)))
            CurChr = CurChr + 2
        Case Else
            TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
        End Select

        CurChr = CurChr + 1
    Loop

    URLDecode = TempAns
End Function

Function URLEncode( _
         StringVal As String, _
         Optional SpaceAsPlus As Boolean = False _
       ) As String

    Dim StringLen As Long: StringLen = Len(StringVal)

    If StringLen > 0 Then
        ReDim result(StringLen) As String
        Dim i As Long, CharCode As Integer
        Dim Char As String, Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For i = 1 To StringLen
            Char = Mid$(StringVal, i, 1)
            CharCode = asc(Char)
            Select Case CharCode
            Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                result(i) = Char
            Case 32
                result(i) = Space
            Case 0 To 15
                result(i) = "%0" & Hex(CharCode)
            Case Else
                result(i) = "%" & Hex(CharCode)
            End Select
        Next i
        URLEncode = Join(result, "")
    End If
End Function

Public Property Get OAUTH_AUTHORIZATION_EXPIRES_IN() As String: OAUTH_AUTHORIZATION_EXPIRES_IN = "oauth_authorization_expires_in": End Property    'Lifetime of the oauth_session_handle in seconds.

Public Property Get OAUTH_CALLBACK() As String: OAUTH_CALLBACK = "oauth_callback": End Property    'Yahoo! redirects Users to this URL after they authorize access to their private data. If your application does not have access to a browser, you must specify the callback as oob (out of bounds).

Public Property Get OAUTH_CALLBACK_CONFIRMED() As String: OAUTH_CALLBACK_CONFIRMED = "oauth_callback_confirmed": End Property    'This parameter confirms that you are using OAuth 1.0 Rev. A. This parameter is always set to true.

Public Property Get OAUTH_CONSUMER_KEY() As String: OAUTH_CONSUMER_KEY = "oauth_consumer_key": End Property    'Consumer Key provided to you when you signed up.

Public Property Get OAUTH_EXPIRES_IN() As String: OAUTH_EXPIRES_IN = "oauth_expires_in": End Property    'The lifetime of the Request Token in seconds. The default number is 3600 seconds, or one hour.

Public Property Get OAUTH_NONCE() As String: OAUTH_NONCE = "oauth_nonce": End Property    'A random string (OAuth Core 1.0 Spec, Section 8)

Public Property Get OAUTH_SESSION_HANDLE() As String: OAUTH_SESSION_HANDLE = "oauth_session_handle": End Property    'The persistent credential used by Yahoo! to identify the Consumer after a User has authorized access to private data. Include this credential in your request to refresh the Access Token once it expires.

Public Property Get OAUTH_SIGNATURE() As String: OAUTH_SIGNATURE = "oauth_signature": End Property    'The concatenated Consumer Secret and Token Secret separated by an "&" character. If you are using the PLAINTEXT signature method, add %26 at the end of the Consumer Secret. If using HMAC-SHA1, refer to OAuth Core 1.0 Spec, Section 9.2. For more information about signing requests, refer to Signing Requests to Yahoo!.

Public Property Get OAUTH_SIGNATURE_METHOD() As String: OAUTH_SIGNATURE_METHOD = "oauth_signature_method": End Property    'The signature method that you use to sign the request. This can be PLAINTEXT or HMAC-SHA1.

Public Property Get OAUTH_TIMESTAMP() As String: OAUTH_TIMESTAMP = "oauth_timestamp": End Property    'Current timestamp of the request. This value must be +-600 seconds of the current time.

Public Property Get OAUTH_TOKEN() As String: OAUTH_TOKEN = "oauth_token": End Property    'The Request Token that Yahoo! returns as a response to the request_token call. The Request Token is required during the User authorization process.

Public Property Get OAUTH_TOKEN_SECRET() As String: OAUTH_TOKEN_SECRET = "oauth_token_secret": End Property    'The secret associated with the Access Token provided in hexstring format.

Public Property Get OAUTH_VERIFIER() As String: OAUTH_VERIFIER = "oauth_verifier": End Property    'The OAuth Verifier is a verification code tied to the Request Token.

Public Property Get OAUTH_VERSION() As String: OAUTH_VERSION = "oauth_version": End Property    'OAuth version (1.0).

Public Property Get OAuthParameterPrefix() As String: OAuthParameterPrefix = "oauth_": End Property '

Public Property Get resource() As String
    resource = m_resource

End Property

Public Property Get URL_ACCESS_TOKEN() As String: URL_ACCESS_TOKEN = "https://api.login.yahoo.com/oauth/v2/get_token": End Property    'change

Public Property Get URL_REQUEST_TOKEN() As String: URL_REQUEST_TOKEN = "https://api.login.yahoo.com/oauth/v2/get_request_token": End Property    'change

Public Property Get XOAUTH_LANG_PREF() As String: XOAUTH_LANG_PREF = "xoauth_lang_pref": End Property    '(optional) The language preference of the User; the default value is EN-US. For further details about this parameter, refer to the OAuth Extension for Specifying User Language Preference.

Public Property Get XOAUTH_REQUEST_AUTH_URL() As String: XOAUTH_REQUEST_AUTH_URL = "xoauth_request_auth_url": End Property    'The URL to the Yahoo! authorization page.

Public Property Get XOAUTH_YAHOO_GUID() As String: XOAUTH_YAHOO_GUID = "xoauth_yahoo_guid": End Property    'The introspective GUID of the currently logged in User. For more information of the GUID, see the Yahoo! Social API Reference.

Public Property Let resource(ByVal vNewValue As String)
    m_resource = vNewValue
End Property

Private Sub UserForm_Initialize()
    Set http = New WinHttpRequest
    http.SetAutoLogonPolicy AutoLogonPolicy_Always
    UserForm_Resize
    tokenFile = fso.GetSpecialFolder(TemporaryFolder) & "\\" & "oauth" & guid & EXT_TOKEN

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then Err.Raise vbError + 1, "frmOauth", "Closed before authentication"
End Sub

Private Sub UserForm_Resize()
    With WebBrowser1
        .Top = 0
        .Left = 0
        .Height = InsideHeight
        .Width = InsideWidth
    End With
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If InStr(1, URL, "oauth_verifier") > 0 Then
        Tag = URL
        Me.Hide
    End If
End Sub

