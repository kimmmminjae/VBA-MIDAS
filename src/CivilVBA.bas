Attribute VB_Name = "CivilVBA"
Option Explicit

'==========================================================
' CivilVBA - MIDAS CIVIL NX / GEN NX Open API toolkit
'==========================================================
'  Sections
'    [1] Config      server / product / timeout / logging
'    [2] Auth        solver detection + MAPI-Key lookup (registry)
'    [3] Transport   WinHttp call site (UTF-8 in/out, status capture, auto retry)
'    [4] Status      status code / error / raw body of the last call
'    [5] JSON        body builders and result readers (JObj / JArr / JVal ...)
'    [6] Raw         ApiRaw - send a finished body to any endpoint
'    [7] Sheet       result table -> sheet, sheet -> Assign payload
'    [8] Helpers     hand written wrappers for the items used most often
'    [9] Helpers     boundaries, connections and item shaped loads
'    [10] Store      collect model data, send it once (ModelCreate)
'    [11] Groups     structure / boundary / load / tendon groups
'    [12] Elements   truss / tension / compression / solid / wall, scale factors, local axis
'    [13] Loads      specified displacement, nodal mass, load to mass, floor / plane loads, temperatures
'    [14] Sections   value / PSC value / composite / tapered, offset, tapered group
'    [15] Time materials   creep / shrinkage, strength, link, change property
'    [16] Stages      construction stage, composite section, time load, creep coefficient, camber
'    [17] Tendons     property, profile, prestress
'    [18] Moving load  code, lanes, vehicles, moving load cases
'    [19] Analysis    controls, settlement, story, force-deformation function, boundary change
'    [20] Dynamic     response spectrum, time history
'    [21] Heat        heat of hydration
'    [22] View        cutting line / plane, selection in NX
'    [23] Graphics    result display, load display, view angle, capture
'    [24] Geometry    nodes / elements by location, selection, group lists
'
'  Companion modules
'    JsonConverter.bas   JSON encode / decode                 (required)
'
'  Model data is collected and sent once - see [10]. Write the model with
'  the helpers, then call ModelCreate.
'
'  References needed (VBE > Tools > References)
'    Microsoft Scripting Runtime      (Dictionary)
'    Microsoft ActiveX Data Objects   (optional, makes UTF-8 handling safe)
'
'  Precondition
'    The target NX application must be running with a model file open,
'    otherwise every request fails.
'
'  Naming rules (kept ASCII on purpose - Korean text in .bas files gets
'  mangled whenever the editor and the file disagree about the code page)
'    Parameters start with p. Locals start with s(tring) i(nteger) d(ouble)
'    b(oolean) o(bject) v(ariant/array). Reserved words and intrinsic
'    function names (Name, Type, Get, Put, Command, Error, Line, Close ...)
'    are never used as procedure names.
'==========================================================


'==========================================================
' [1] Config
'==========================================================

' Defaults. Change at run time with CvSetServer / CvSetRegion / CvSetTimeout.
Private Const DEF_SERVER As String = "moa-engineers.midasit.com"
Private Const DEF_PROGRAM As String = "civil"          ' "civil" or "gen"
Private Const DEF_PORT As String = "443"
Private Const DEF_TIMEOUT_MS As Long = 300000          ' 5 min - analysis is slow

' Kept for older code. New code should call CvBaseUrl().
Public Const DEBUG_LOG As Boolean = False

Private mServer As String
Private mProgram As String
Private mPort As String
Private mTimeoutMs As Long
Private mLogOn As Boolean
Private mConfigReady As Boolean

' --- [2] Auth ---
Private Const REG_ROOT As String = "HKEY_CURRENT_USER\Software\MIDAS\"
Private Const REG_TAIL As String = "\CONNECTION\Key"

Private mSolverName As String
Private mSolverKey As String
Private mSolverCached As Boolean
Private mDeadSolver As String
Private mRetrying As Boolean
Private mManualKey As String

' --- [3] Transport ---
Private mLastVerb As String
Private mLastUrl As String
Private mLastReqBody As String
Private mLastStatus As Long
Private mLastResponse As String
Private mLastError As String
Private mCallCount As Long

' --- [6] Generic call ---
Private mLastWarn As String
' [24] geometry by location - node registry and running ids
Private mGeoXYZ As Object       ' "id" -> Array(x, y, z)
Private mGeoGrid As Object      ' "int(x),int(y),int(z)" -> Collection of ids
Private mGeoNodeMax As Long
Private mGeoElemMax As Long
Private mGeoNodeSeen As Long
Private mGeoElemSeen As Long
Private mGeoLast As Variant     ' end of the last beam / truss (LastLoc)
Private mGrpSeen As Object      ' "list|id" already in a group list ([11])

' Model store - see [10]
Private mStore As Object        ' "NODE" -> Dictionary("1" -> record)
Private mRaw As Object          ' "MATL" -> Collection of finished JSON bodies
Private mOps As Collection      ' Array(verb, path, json) for ope/... calls

Private Sub CvEnsureConfig()
    If mConfigReady Then Exit Sub
    mServer = DEF_SERVER
    mProgram = DEF_PROGRAM
    mPort = DEF_PORT
    mTimeoutMs = DEF_TIMEOUT_MS
    mLogOn = DEBUG_LOG
    mConfigReady = True
End Sub

' e.g. https://moa-engineers.midasit.com:443/civil
Public Function CvBaseUrl() As String
    CvEnsureConfig
    CvBaseUrl = "https://" & mServer & ":" & mPort & "/" & mProgram
End Function

Public Sub CvSetServer(ByVal pServer As String, _
                       Optional ByVal pProgram As String = "", _
                       Optional ByVal pPort As String = "")
    CvEnsureConfig
    If Len(Trim$(pServer)) > 0 Then mServer = Trim$(pServer)
    If Len(Trim$(pProgram)) > 0 Then mProgram = LCase$(Trim$(pProgram))
    If Len(Trim$(pPort)) > 0 Then mPort = Trim$(pPort)
End Sub

' pRegion : GLOBAL / KR / IN / GB / US / CN
Public Sub CvSetRegion(ByVal pRegion As String, Optional ByVal pProgram As String = "")
    CvEnsureConfig

    Select Case UCase$(Trim$(pRegion))
        Case "GLOBAL", "":  mServer = "moa-engineers.midasit.com"
        Case "KR":          mServer = "moa-engineers-kr.midasit.com"
        Case "IN":          mServer = "moa-engineers-in.midasit.com"
        Case "GB", "EU":    mServer = "moa-engineers-gb.midasit.com"
        Case "US":          mServer = "moa-engineers-us.midasit.com"
        Case "CN":          mServer = "moa-engineers.midasit.cn"
        Case Else:          mServer = Trim$(pRegion)
    End Select

    If Len(Trim$(pProgram)) > 0 Then mProgram = LCase$(Trim$(pProgram))
End Sub

Public Sub CvSetTimeout(ByVal pMilliseconds As Long)
    CvEnsureConfig
    If pMilliseconds > 0 Then mTimeoutMs = pMilliseconds
End Sub

Public Function CvTimeout() As Long
    CvEnsureConfig
    CvTimeout = mTimeoutMs
End Function

' Turn Immediate window logging on or off.
Public Sub CvSetLog(ByVal pOn As Boolean)
    CvEnsureConfig
    mLogOn = pOn
End Sub

Public Function CvLogOn() As Boolean
    CvEnsureConfig
    CvLogOn = mLogOn
End Function

Public Function CvProgram() As String
    CvEnsureConfig
    CvProgram = mProgram
End Function


'==========================================================
' [2] Auth - solver detection + MAPI-Key
'==========================================================
'  FES and HYPER-S cannot be told apart by process name (both are CVLw.exe)
'  or by window title, and the registry CONNECTION\Key survives after the
'  program is closed, so both paths often hold a value at the same time.
'  The only reliable signal is the API response: a dead key answers
'  404 "client does not exist".
'
'  Strategy - never make a probe request.
'    - Pick the first solver that has a key in the registry and just use it
'      (zero extra round trips).
'    - If that key is dead the first request comes back 404, and we retry
'      once with the other solver. We only pay when the guess was wrong.
'    - After one success the cached key is reused for the whole session.
'==========================================================

' The first entry wins. Add a row here when a new product or language build
' shows up - reading a missing registry path is harmless.
Private Function CvSolverTable() As Variant
    Dim vMap(1 To 8, 1 To 2) As String

    vMap(1, 1) = "FES":            vMap(1, 2) = "CVLwNX_KR"
    vMap(2, 1) = "HYPER-S":        vMap(2, 2) = "CVLwNX_KR_HYPER_S"
    vMap(3, 1) = "FES-EN":         vMap(3, 2) = "CVLwNX"
    vMap(4, 1) = "HYPER-S-EN":     vMap(4, 2) = "CVLwNX_HYPER_S"
    vMap(5, 1) = "FES-JP":         vMap(5, 2) = "CVLwNX_JP"
    vMap(6, 1) = "GEN":            vMap(6, 2) = "GENwNX_KR"
    vMap(7, 1) = "GEN-EN":         vMap(7, 2) = "GENwNX"
    vMap(8, 1) = "GEN-HYPER-S":    vMap(8, 2) = "GENwNX_KR_HYPER_S"

    CvSolverTable = vMap
End Function

Private Function CvRegPath(ByVal pProductKey As String) As String
    CvRegPath = REG_ROOT & pProductKey & REG_TAIL
End Function

Private Function CvReadReg(ByVal pKeyPath As String) As String
    Dim oShell As Object

    On Error Resume Next
    Set oShell = CreateObject("WScript.Shell")
    CvReadReg = oShell.RegRead(pKeyPath)
    If Err.Number <> 0 Then CvReadReg = ""
    Err.Clear
    On Error GoTo 0
End Function

' Clear the cache. Switching FES <-> HYPER-S is detected automatically,
' so there is normally no reason to call this.
Public Sub ResetSolverCache()
    mSolverName = ""
    mSolverKey = ""
    mSolverCached = False
    mDeadSolver = ""
End Sub

' Ignore the registry and use this key instead. Pass "" to clear.
Public Sub CvSetMapiKey(ByVal pKey As String, Optional ByVal pLabel As String = "MANUAL")
    mManualKey = Trim$(pKey)
    If Len(mManualKey) > 0 Then
        mSolverName = pLabel
        mSolverKey = mManualKey
        mSolverCached = True
    Else
        ResetSolverCache
    End If
End Sub

' Pick a solver that has a registry key and is not already known to be dead.
' No HTTP traffic here.
Private Sub CvSelectSolver()
    Dim vMap As Variant
    Dim iRow As Long
    Dim iPass As Long
    Dim sKey As String

    If Len(mManualKey) > 0 Then
        mSolverKey = mManualKey
        mSolverCached = True
        Exit Sub
    End If

    vMap = CvSolverTable()

    ' pass 0 : skip dead solvers / pass 1 : nothing left, clear the skip list
    For iPass = 0 To 1
        For iRow = LBound(vMap, 1) To UBound(vMap, 1)
            If iPass = 1 Or UCase$(vMap(iRow, 1)) <> UCase$(mDeadSolver) Then
                sKey = CvReadReg(CvRegPath(vMap(iRow, 2)))
                If Len(sKey) > 0 Then
                    mSolverName = vMap(iRow, 1)
                    mSolverKey = sKey
                    mSolverCached = True
                    Exit Sub
                End If
            End If
        Next iRow
        mDeadSolver = ""
    Next iPass

    mSolverName = ""
    mSolverKey = ""
    mSolverCached = False
End Sub

' Name of the solver in use. Picks one from the registry if nothing is cached.
Public Function GetCurrentSolverName() As String
    If Not mSolverCached Then CvSelectSolver
    GetCurrentSolverName = mSolverName
End Function

' MAPI key of the named (or selected) solver. Never falls back to the other one.
Public Function GetMapiKey(Optional ByVal solverName As String = "") As String
    Dim vMap As Variant
    Dim iRow As Long
    Dim sTarget As String

    If Len(Trim$(solverName)) = 0 Then
        If Not mSolverCached Then CvSelectSolver
        GetMapiKey = mSolverKey
        Exit Function
    End If

    sTarget = UCase$(Trim$(solverName))

    If mSolverCached And Len(mSolverKey) > 0 Then
        If sTarget = UCase$(Trim$(mSolverName)) Then
            GetMapiKey = mSolverKey
            Exit Function
        End If
    End If

    vMap = CvSolverTable()
    For iRow = LBound(vMap, 1) To UBound(vMap, 1)
        If UCase$(Trim$(vMap(iRow, 1))) = sTarget Then
            GetMapiKey = CvReadReg(CvRegPath(vMap(iRow, 2)))
            Exit Function
        End If
    Next iRow

    GetMapiKey = ""
End Function

Public Function GetCurrentKeyPath(Optional ByVal solverName As String = "") As String
    Dim vMap As Variant
    Dim iRow As Long
    Dim sTarget As String

    sTarget = Trim$(solverName)
    If Len(sTarget) = 0 Then sTarget = GetCurrentSolverName()
    If Len(sTarget) = 0 Then Exit Function

    vMap = CvSolverTable()
    For iRow = LBound(vMap, 1) To UBound(vMap, 1)
        If UCase$(Trim$(vMap(iRow, 1))) = UCase$(sTarget) Then
            GetCurrentKeyPath = CvRegPath(vMap(iRow, 2))
            Exit Function
        End If
    Next iRow
End Function

' Solvers that currently have a key in the registry.
' Each entry is "name<TAB>first 8 chars<TAB>registry path".
Public Function CvFoundSolvers() As Variant
    Dim vMap As Variant
    Dim vOut() As String
    Dim iRow As Long
    Dim iCnt As Long
    Dim sKey As String

    vMap = CvSolverTable()
    ReDim vOut(0 To UBound(vMap, 1) - 1)

    For iRow = LBound(vMap, 1) To UBound(vMap, 1)
        sKey = CvReadReg(CvRegPath(vMap(iRow, 2)))
        If Len(sKey) > 0 Then
            vOut(iCnt) = vMap(iRow, 1) & vbTab & Left$(sKey, 8) & "..." & _
                         vbTab & CvRegPath(vMap(iRow, 2))
            iCnt = iCnt + 1
        End If
    Next iRow

    If iCnt = 0 Then
        CvFoundSolvers = Array()
    Else
        ReDim Preserve vOut(0 To iCnt - 1)
        CvFoundSolvers = vOut
    End If
End Function

' Check whether this MAPI key is attached to a live NX (manual check).
Public Function IsMapiKeyAlive(ByVal mapiKey As String) As Boolean
    Dim oHttp As Object

    On Error GoTo Fail

    Set oHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    oHttp.SetTimeouts 5000, 5000, 5000, 5000
    oHttp.Open "GET", CvBaseUrl() & "/db/UNIT", False
    oHttp.setRequestHeader "MAPI-Key", mapiKey
    oHttp.setRequestHeader "Content-Type", "application/json"
    oHttp.Send

    IsMapiKeyAlive = (oHttp.Status = 200)
    Exit Function
Fail:
    IsMapiKeyAlive = False
End Function

Public Function GetRunningMidasProcesses() As Collection
    Dim oOut As Collection
    Dim oWmi As Object
    Dim oItems As Object
    Dim oOne As Object

    Set oOut = New Collection

    On Error Resume Next
    Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
    Set oItems = oWmi.ExecQuery("SELECT Name, ProcessId FROM Win32_Process", , 48)

    If Not oItems Is Nothing Then
        For Each oOne In oItems
            If InStr(LCase$(oOne.Name), "cvlw") > 0 Or _
               InStr(LCase$(oOne.Name), "genw") > 0 Or _
               InStr(LCase$(oOne.Name), "midas") > 0 Then
                oOut.Add oOne.Name & " (PID " & oOne.ProcessId & ")"
            End If
        Next oOne
    End If

    On Error GoTo 0
    Set GetRunningMidasProcesses = oOut
End Function


'==========================================================
' [3] Transport
'==========================================================

' Send the body as UTF-8 bytes so non-ASCII names survive the round trip.
Private Function CvUtf8Bytes(ByVal pText As String) As Variant
    Dim oStm As Object

    On Error GoTo Fallback

    Set oStm = CreateObject("ADODB.Stream")
    oStm.Type = 2                       ' adTypeText
    oStm.Charset = "utf-8"
    oStm.Open
    oStm.WriteText pText
    oStm.Position = 0
    oStm.Type = 1                       ' adTypeBinary
    oStm.Position = 3                   ' skip the 3 byte UTF-8 BOM
    CvUtf8Bytes = oStm.Read
    oStm.Close
    Exit Function

Fallback:
    CvUtf8Bytes = pText
End Function

' Read the response body as UTF-8, falling back to responseText.
Private Function CvUtf8Text(ByVal pHttp As Object) As String
    Dim oStm As Object

    On Error GoTo Fallback

    Set oStm = CreateObject("ADODB.Stream")
    oStm.Type = 1                       ' adTypeBinary
    oStm.Open
    oStm.Write pHttp.responseBody
    oStm.Position = 0
    oStm.Type = 2                       ' adTypeText
    oStm.Charset = "utf-8"
    CvUtf8Text = oStm.ReadText
    oStm.Close
    Exit Function

Fallback:
    On Error Resume Next
    CvUtf8Text = pHttp.responseText
End Function

' Accepts "/db/NODE" and "db/NODE" alike.
Private Function CvNormalizePath(ByVal pPath As String) As String
    Dim sTmp As String

    sTmp = Trim$(pPath)
    sTmp = Replace(sTmp, "\", "/")
    Do While Left$(sTmp, 1) = "/"
        sTmp = Mid$(sTmp, 2)
    Loop

    CvNormalizePath = "/" & sTmp
End Function

'----------------------------------------------------------
' The single point every HTTP call goes through.
'   pVerb   : GET / POST / PUT / DELETE
'   pPath   : "/db/NODE"
'   pBody   : JSON text. Empty means no body is sent.
'   pSolver : force a solver. Leave blank for auto pick + 404 auto retry.
' Returns the response body. If the call itself fails it returns an
' {"error":...} shaped string instead of raising.
'----------------------------------------------------------
Public Function CvHttp(ByVal pVerb As String, _
                       ByVal pPath As String, _
                       Optional ByVal pBody As String = "", _
                       Optional ByVal pSolver As String = "") As String

    Dim oHttp As Object
    Dim sKey As String
    Dim sSolver As String
    Dim sUrl As String
    Dim sVerb As String

    CvEnsureConfig

    sVerb = UCase$(Trim$(pVerb))
    If Len(sVerb) = 0 Then sVerb = "GET"

    If Len(Trim$(pSolver)) > 0 Then
        sSolver = Trim$(pSolver)
    Else
        sSolver = GetCurrentSolverName()
    End If

    sKey = GetMapiKey(sSolver)
    sUrl = CvBaseUrl() & CvNormalizePath(pPath)

    mLastVerb = sVerb
    mLastUrl = sUrl
    mLastReqBody = pBody
    mLastStatus = 0
    mLastResponse = ""
    mLastError = ""
    mCallCount = mCallCount + 1

    If Len(sKey) = 0 Then
        mLastError = "No MAPI-Key found. Check that NX is running and that a " & _
                     "key has been issued in its API settings."
        CvHttp = "{""error"":""" & mLastError & """}"
        If mLogOn Then Debug.Print "[CivilVBA] " & mLastError
        Exit Function
    End If

    On Error GoTo Fail

    Set oHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    oHttp.SetTimeouts mTimeoutMs, mTimeoutMs, mTimeoutMs, mTimeoutMs
    oHttp.Open sVerb, sUrl, False
    oHttp.setRequestHeader "MAPI-Key", sKey
    oHttp.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    oHttp.setRequestHeader "Accept", "application/json"

    ' Ignore certificate errors, for networks with a TLS inspecting proxy.
    ' Option is a VBA keyword, so reach it through CallByName instead of a dot.
    ' If it is not supported we simply carry on.
    On Error Resume Next
    CallByName oHttp, "Option", VbLet, 4, 13056
    On Error GoTo Fail

    If Len(pBody) = 0 Or sVerb = "GET" Or sVerb = "DELETE" Then
        oHttp.Send
    Else
        oHttp.Send CvUtf8Bytes(pBody)
    End If

    mLastStatus = oHttp.Status
    mLastResponse = CvUtf8Text(oHttp)
    CvHttp = mLastResponse

    If mLogOn Then
        Debug.Print "[CivilVBA] " & sVerb & " " & CvNormalizePath(pPath) & _
                    " -> " & mLastStatus & " (" & Len(mLastResponse) & " bytes)"
    End If

    ' The key we picked was dead (that solver is not running): switch to the
    ' other one and retry exactly once.
    If mLastStatus = 404 And Not mRetrying And Len(Trim$(pSolver)) = 0 Then
        If InStr(mLastResponse, "client does not exist") > 0 Then
            mDeadSolver = sSolver
            mSolverCached = False
            mRetrying = True
            CvHttp = CvHttp(sVerb, pPath, pBody, pSolver)
            mRetrying = False
            Exit Function
        End If
    End If

    If mLastStatus < 200 Or mLastStatus >= 300 Then
        mLastError = "HTTP " & mLastStatus & " - " & Left$(mLastResponse, 500)
    End If

    Exit Function

Fail:
    mLastStatus = -1
    mLastError = "Transport failure (" & Err.Number & ") " & Err.Description & _
                 " / " & sVerb & " " & sUrl
    mLastResponse = ""
    CvHttp = "{""error"":""" & Replace(mLastError, """", "'") & """}"
    If mLogOn Then Debug.Print "[CivilVBA] " & mLastError
End Function

'---------- Original names kept, so existing code still runs ----------

Public Function CallAPI(ByVal METHOD As String, ByVal endpoint As String, _
                        Optional ByVal jsonBody As String = "", _
                        Optional ByVal solverName As String = "") As String
    CallAPI = CvHttp(METHOD, endpoint, jsonBody, solverName)
End Function

Public Function CallGet(ByVal endpoint As String, _
                        Optional ByVal solverName As String = "") As String
    CallGet = CvHttp("GET", endpoint, "", solverName)
End Function

Public Function CallPost(ByVal endpoint As String, ByVal jsonBody As String, _
                         Optional ByVal solverName As String = "") As String
    CallPost = CvHttp("POST", endpoint, jsonBody, solverName)
End Function

Public Function CallPut(ByVal endpoint As String, ByVal jsonBody As String, _
                        Optional ByVal solverName As String = "") As String
    CallPut = CvHttp("PUT", endpoint, jsonBody, solverName)
End Function

Public Function CallDelete(ByVal endpoint As String, _
                           Optional ByVal solverName As String = "") As String
    CallDelete = CvHttp("DELETE", endpoint, "", solverName)
End Function


'==========================================================
' [4] Status of the last call
'==========================================================

Public Function CvLastStatus() As Long
    CvLastStatus = mLastStatus
End Function

Public Function CvLastVerb() As String
    CvLastVerb = mLastVerb
End Function

Public Function CvLastUrl() As String
    CvLastUrl = mLastUrl
End Function

Public Function CvLastRequestBody() As String
    CvLastRequestBody = mLastReqBody
End Function

Public Function CvLastResponse() As String
    CvLastResponse = mLastResponse
End Function

Public Function CvLastError() As String
    CvLastError = mLastError
End Function

Public Function CvCallCount() As Long
    CvCallCount = mCallCount
End Function

' Did the last call finish with a 2xx status?
Public Function CvIsOk() As Boolean
    CvIsOk = (mLastStatus >= 200 And mLastStatus < 300)
End Function

' Pull the error message out of a response body. Empty when there is none.
Public Function CvErrorText(Optional ByVal pJson As String = "") As String
    Dim sRaw As String
    Dim oRoot As Object
    Dim vKey As Variant

    sRaw = pJson
    If Len(sRaw) = 0 Then sRaw = mLastResponse
    If Len(sRaw) = 0 Then
        CvErrorText = mLastError
        Exit Function
    End If

    Set oRoot = JParse(sRaw)
    If oRoot Is Nothing Then
        If Not CvIsOk() Then CvErrorText = mLastError
        Exit Function
    End If

    If TypeName(oRoot) = "Dictionary" Then
        For Each vKey In Array("error", "Error", "ERROR", "message", "MESSAGE")
            If oRoot.Exists(vKey) Then
                CvErrorText = JText(oRoot.Item(vKey))
                Exit Function
            End If
        Next vKey
    End If

    If Not CvIsOk() Then CvErrorText = mLastError
End Function

' Dump a summary of the last call to the Immediate window.
Public Sub CvPrintLast()
    Debug.Print "--- CivilVBA : last call ---"
    Debug.Print " request  : " & mLastVerb & " " & mLastUrl
    Debug.Print " body     : " & Left$(mLastReqBody, 400)
    Debug.Print " status   : " & mLastStatus & IIf(CvIsOk(), "  (ok)", "  (failed)")
    Debug.Print " response : " & Left$(mLastResponse, 800)
    If Len(mLastError) > 0 Then Debug.Print " error    : " & mLastError
End Sub


'==========================================================
' [5] JSON builders and readers
'==========================================================
'  Long lines are painful in VBA and a single statement may only be split
'  across 25 lines, so never hand write request bodies as text. Build them:
'
'    Set oArg = JObj("TABLE_TYPE", "REACTIONG", _
'                    "UNIT", JObj("FORCE", "KN", "DIST", "M"), _
'                    "LOAD_CASE_NAMES", JArr("Dead Load", "Live Load"))
'
'  Reach deeper with JPut - missing intermediate objects are created:
'    JPut oArg, "STYLES/FORMAT", "FIXED"
'
'  Read back with JVal - dictionaries by key, arrays by 1 based index:
'    dX = JVal(oRes, "NODE/12/X", 0)
'==========================================================

' A new empty Dictionary.
Public Function JNew() As Object
    Set JNew = CreateObject("Scripting.Dictionary")
End Function

' Store one key. Works whether the value is an object or not.
Public Sub JSet(ByVal pObj As Object, ByVal pKey As String, ByVal pValue As Variant)
    If pObj Is Nothing Then Exit Sub

    If IsObject(pValue) Then
        Set pObj.Item(pKey) = pValue
    Else
        pObj.Item(pKey) = pValue
    End If
End Sub

' Build a Dictionary from alternating key, value pairs. No args = empty object.
Public Function JObj(ParamArray pPairs() As Variant) As Object
    Dim oOut As Object
    Dim iIdx As Long

    Set oOut = JNew()

    If UBound(pPairs) >= LBound(pPairs) + 1 Then
        For iIdx = LBound(pPairs) To UBound(pPairs) - 1 Step 2
            JSet oOut, CStr(pPairs(iIdx)), pPairs(iIdx + 1)
        Next iIdx
    End If

    Set JObj = oOut
End Function

' Build a Collection that serialises as a JSON array.
Public Function JArr(ParamArray pItems() As Variant) As Collection
    Dim oOut As Collection
    Dim iIdx As Long

    Set oOut = New Collection

    If UBound(pItems) >= LBound(pItems) Then
        For iIdx = LBound(pItems) To UBound(pItems)
            oOut.Add pItems(iIdx)
        Next iIdx
    End If

    Set JArr = oOut
End Function

' Turn a VBA array (or a single value) into a Collection.
Public Function JArrFrom(ByVal pSource As Variant) As Collection
    Dim oOut As Collection
    Dim vOne As Variant

    Set oOut = New Collection

    If IsObject(pSource) Then
        If TypeName(pSource) = "Collection" Then
            Set JArrFrom = pSource
            Exit Function
        End If
    End If

    If IsArray(pSource) Then
        For Each vOne In pSource
            oOut.Add vOne
        Next vOne
    ElseIf Not IsEmpty(pSource) Then
        oOut.Add pSource
    End If

    Set JArrFrom = oOut
End Function

' Copy a value that may or may not be an object (used by JVal).
Private Sub JAssign(ByRef pTarget As Variant, ByVal pSource As Variant)
    If IsObject(pSource) Then
        Set pTarget = pSource
    Else
        pTarget = pSource
    End If
End Sub

' Store a value at path "A/B/C", creating intermediate Dictionaries as needed.
Public Sub JPut(ByVal pRoot As Object, ByVal pPath As String, ByVal pValue As Variant)
    Dim vSeg As Variant
    Dim iIdx As Long
    Dim oCur As Object
    Dim sKey As String
    Dim bMake As Boolean

    If pRoot Is Nothing Then Exit Sub

    vSeg = Split(Replace(Replace(pPath, ".", "/"), "\", "/"), "/")
    Set oCur = pRoot

    For iIdx = LBound(vSeg) To UBound(vSeg) - 1
        sKey = Trim$(CStr(vSeg(iIdx)))
        If Len(sKey) > 0 Then
            bMake = False
            If Not oCur.Exists(sKey) Then
                bMake = True
            ElseIf Not IsObject(oCur.Item(sKey)) Then
                bMake = True
            End If

            If bMake Then Set oCur.Item(sKey) = JNew()
            Set oCur = oCur.Item(sKey)
        End If
    Next iIdx

    JSet oCur, Trim$(CStr(vSeg(UBound(vSeg)))), pValue
End Sub

' Read a value at a path such as "NODE/12/X".
'   Dictionary segments match by key, Collection segments are 1 based,
'   plain arrays use their own subscripts.
'   Returns pDefault when the path does not exist.
Public Function JVal(ByVal pAny As Variant, ByVal pPath As String, _
                     Optional ByVal pDefault As Variant = Empty) As Variant
    Dim vSeg As Variant
    Dim vCur As Variant
    Dim vNext As Variant
    Dim iIdx As Long
    Dim iPos As Long
    Dim sKey As String

    JAssign JVal, pDefault
    JAssign vCur, pAny

    If Len(Trim$(pPath)) = 0 Then
        JAssign JVal, pAny
        Exit Function
    End If

    vSeg = Split(Replace(Replace(pPath, ".", "/"), "\", "/"), "/")

    For iIdx = LBound(vSeg) To UBound(vSeg)
        sKey = Trim$(CStr(vSeg(iIdx)))
        If Len(sKey) > 0 Then

            If IsObject(vCur) Then
                Select Case TypeName(vCur)
                    Case "Dictionary"
                        If Not vCur.Exists(sKey) Then Exit Function
                        JAssign vNext, vCur.Item(sKey)

                    Case "Collection"
                        If Not IsNumeric(sKey) Then Exit Function
                        iPos = CLng(sKey)
                        If iPos < 1 Then Exit Function
                        If iPos > vCur.Count Then Exit Function
                        JAssign vNext, vCur.Item(iPos)

                    Case Else
                        Exit Function
                End Select

            ElseIf IsArray(vCur) Then
                If Not IsNumeric(sKey) Then Exit Function
                iPos = CLng(sKey)
                If iPos < LBound(vCur) Then Exit Function
                If iPos > UBound(vCur) Then Exit Function
                JAssign vNext, vCur(iPos)

            Else
                Exit Function
            End If

            JAssign vCur, vNext
        End If
    Next iIdx

    JAssign JVal, vCur
End Function

' Parse text into a Dictionary / Collection. Returns Nothing on failure.
Public Function JParse(ByVal pJson As String) As Object
    On Error GoTo Fail

    If Len(Trim$(pJson)) = 0 Then Exit Function
    Set JParse = JsonConverter.ParseJson(pJson)
    Exit Function

Fail:
    Set JParse = Nothing
End Function

' Serialise anything to JSON text. Non objects are just converted to string.
Public Function JText(ByVal pAny As Variant) As String
    On Error GoTo Fail

    If IsObject(pAny) Then
        JText = JsonConverter.ConvertToJson(pAny)
    ElseIf IsArray(pAny) Then
        JText = JsonConverter.ConvertToJson(pAny)
    ElseIf IsNull(pAny) Then
        JText = "null"
    ElseIf IsEmpty(pAny) Then
        JText = ""
    Else
        JText = CStr(pAny)
    End If
    Exit Function

Fail:
    JText = ""
End Function

' Same, indented by 2 spaces for reading.
Public Function JPretty(ByVal pAny As Variant) As String
    On Error GoTo Fail
    JPretty = JsonConverter.ConvertToJson(pAny, 2)
    Exit Function
Fail:
    JPretty = JText(pAny)
End Function

' Keys of a Dictionary, "1".."N" for a Collection, empty array otherwise.
Public Function JKeys(ByVal pAny As Variant) As Variant
    Dim vOut() As String
    Dim iIdx As Long

    If IsObject(pAny) Then
        Select Case TypeName(pAny)
            Case "Dictionary"
                If pAny.Count = 0 Then
                    JKeys = Array()
                Else
                    JKeys = pAny.Keys
                End If
                Exit Function

            Case "Collection"
                If pAny.Count = 0 Then
                    JKeys = Array()
                    Exit Function
                End If
                ReDim vOut(0 To pAny.Count - 1)
                For iIdx = 1 To pAny.Count
                    vOut(iIdx - 1) = CStr(iIdx)
                Next iIdx
                JKeys = vOut
                Exit Function
        End Select
    End If

    JKeys = Array()
End Function

Public Function JCount(ByVal pAny As Variant) As Long
    If IsObject(pAny) Then
        Select Case TypeName(pAny)
            Case "Dictionary", "Collection"
                JCount = pAny.Count
        End Select
    ElseIf IsArray(pAny) Then
        On Error Resume Next
        JCount = UBound(pAny) - LBound(pAny) + 1
        On Error GoTo 0
    End If
End Function

Public Function JHas(ByVal pAny As Variant, ByVal pKey As String) As Boolean
    If Not IsObject(pAny) Then Exit Function
    If TypeName(pAny) <> "Dictionary" Then Exit Function
    JHas = pAny.Exists(pKey)
End Function

' Read as a number (pDefault when missing or not numeric).
Public Function JNum(ByVal pAny As Variant, ByVal pPath As String, _
                     Optional ByVal pDefault As Double = 0) As Double
    Dim vRaw As Variant

    vRaw = JVal(pAny, pPath, Empty)
    If IsEmpty(vRaw) Then
        JNum = pDefault
    ElseIf IsNumeric(vRaw) Then
        JNum = CDbl(vRaw)
    Else
        JNum = pDefault
    End If
End Function

' Read as text.
Public Function JStr(ByVal pAny As Variant, ByVal pPath As String, _
                     Optional ByVal pDefault As String = "") As String
    Dim vRaw As Variant

    vRaw = JVal(pAny, pPath, Empty)
    If IsEmpty(vRaw) Then
        JStr = pDefault
    Else
        JStr = JText(vRaw)
    End If
End Function

' Read as a boolean.
Public Function JBool(ByVal pAny As Variant, ByVal pPath As String, _
                      Optional ByVal pDefault As Boolean = False) As Boolean
    Dim vRaw As Variant

    vRaw = JVal(pAny, pPath, Empty)
    If IsEmpty(vRaw) Then
        JBool = pDefault
    ElseIf VarType(vRaw) = vbBoolean Then
        JBool = CBool(vRaw)
    ElseIf IsNumeric(vRaw) Then
        JBool = (CDbl(vRaw) <> 0)
    Else
        JBool = (UCase$(CStr(vRaw)) = "TRUE")
    End If
End Function


'==========================================================
' [6] Raw call
'==========================================================
'  ApiRaw sends a body exactly as given, to any endpoint, right away.
'    ApiRaw "/db/NODE", "GET"
'    ApiRaw "/post/TABLE", "POST", JText(JObj("Argument", JObj("TABLE_TYPE", "REACTIONG")))
'  Model data should go through the helpers and ModelCreate instead ([10]).
'==========================================================

' Anything the last call disagreed with the catalog about (unsupported
' method, unknown uri). Empty when there was nothing to report.
Public Function CvLastWarning() As String
    CvLastWarning = mLastWarn
End Function

Private Function ApiLeafOf(ByVal pUri As String) As String
    Dim vSeg As Variant

    vSeg = Split(pUri, "/")
    ApiLeafOf = vSeg(UBound(vSeg))
End Function

' Is this already a finished JSON string?
Private Function ApiLooksLikeJson(ByVal pAny As Variant) As Boolean
    Dim sTmp As String

    If IsObject(pAny) Then Exit Function
    If IsArray(pAny) Then Exit Function
    If VarType(pAny) <> vbString Then Exit Function

    sTmp = Trim$(CStr(pAny))
    If Len(sTmp) < 2 Then Exit Function

    ApiLooksLikeJson = (Left$(sTmp, 1) = "{" Or Left$(sTmp, 1) = "[")
End Function

' Send a body verbatim, with no convention applied (escape hatch).
Public Function ApiRaw(ByVal pUri As String, ByVal pVerb As String, _
                       Optional ByVal pRawJson As String = "") As String
    ApiRaw = CvHttp(pVerb, pUri, pRawJson)
End Function


'==========================================================
' [7] Sheet helpers
'==========================================================

' Write a { "table name": { "HEAD": [...], "DATA": [[...]] } } response
' (as returned by post/TABLE) onto a sheet.
' Returns the number of data rows written, or -1 when no table was found.
Public Function TableToSheet(ByVal pJson As String, ByVal pTarget As Range, _
                             Optional ByVal pTableName As String = "", _
                             Optional ByVal pWriteHead As Boolean = True) As Long
    Dim oRoot As Object
    Dim oTable As Object
    Dim vKeys As Variant
    Dim vKey As Variant
    Dim vHead As Variant
    Dim vRow As Variant
    Dim vCell As Variant
    Dim iCol As Long
    Dim iRow As Long

    TableToSheet = -1
    If pTarget Is Nothing Then Exit Function

    Set oRoot = JParse(pJson)
    If oRoot Is Nothing Then Exit Function
    If TypeName(oRoot) <> "Dictionary" Then Exit Function

    ' Locate the table: by name when given, else the first child with HEAD.
    If Len(pTableName) > 0 Then
        If oRoot.Exists(pTableName) Then Set oTable = oRoot.Item(pTableName)
    Else
        vKeys = oRoot.Keys
        For Each vKey In vKeys
            If IsObject(oRoot.Item(vKey)) Then
                If TypeName(oRoot.Item(vKey)) = "Dictionary" Then
                    If oRoot.Item(vKey).Exists("HEAD") Then
                        Set oTable = oRoot.Item(vKey)
                        Exit For
                    End If
                End If
            End If
        Next vKey
    End If

    If oTable Is Nothing Then Exit Function
    If Not oTable.Exists("DATA") Then Exit Function

    iRow = 0

    If pWriteHead And oTable.Exists("HEAD") Then
        iCol = 0
        Set vHead = oTable.Item("HEAD")
        For Each vCell In vHead
            pTarget.Offset(0, iCol).Value = CvCellSafe(vCell)
            iCol = iCol + 1
        Next vCell
        iRow = 1
    End If

    For Each vRow In oTable.Item("DATA")
        iCol = 0
        For Each vCell In vRow
            pTarget.Offset(iRow, iCol).Value = CvCellSafe(vCell)
            iCol = iCol + 1
        Next vCell
        iRow = iRow + 1
    Next vRow

    If pWriteHead Then
        TableToSheet = iRow - 1
    Else
        TableToSheet = iRow
    End If
End Function

' Null cannot be written to a cell, so turn it into an empty string.
Private Function CvCellSafe(ByVal pValue As Variant) As Variant
    If IsObject(pValue) Then
        CvCellSafe = JText(pValue)
    ElseIf IsNull(pValue) Then
        CvCellSafe = ""
    ElseIf IsArray(pValue) Then
        CvCellSafe = JText(pValue)
    Else
        CvCellSafe = pValue
    End If
End Function

' Spread a db item over a sheet, first column holding the id.
'   pFields : "X,Y,Z" to choose columns. Blank uses every scalar field of the
'             first record.
' Returns the number of data rows written.
Public Function ItemsToSheet(ByVal pUri As String, ByVal pTarget As Range, _
                             Optional ByVal pFields As String = "") As Long
    Dim oAll As Object
    Dim oOne As Object
    Dim vIds As Variant
    Dim vId As Variant
    Dim vFields As Variant
    Dim vKey As Variant
    Dim sList As String
    Dim iRow As Long
    Dim iCol As Long

    If pTarget Is Nothing Then Exit Function

    Set oAll = CvReadItems(pUri)
    If oAll Is Nothing Then Exit Function
    If TypeName(oAll) <> "Dictionary" Then Exit Function
    If oAll.Count = 0 Then Exit Function

    vIds = oAll.Keys

    ' Decide the column list.
    If Len(Trim$(pFields)) > 0 Then
        vFields = Split(Replace(pFields, " ", ""), ",")
    Else
        Set oOne = Nothing
        If IsObject(oAll.Item(vIds(0))) Then Set oOne = oAll.Item(vIds(0))
        If oOne Is Nothing Then Exit Function

        sList = ""
        For Each vKey In oOne.Keys
            If Not IsObject(oOne.Item(vKey)) Then
                sList = sList & "," & CStr(vKey)
            End If
        Next vKey
        If Len(sList) = 0 Then Exit Function
        vFields = Split(Mid$(sList, 2), ",")
    End If

    ' Header row.
    pTarget.Offset(0, 0).Value = "ID"
    For iCol = LBound(vFields) To UBound(vFields)
        pTarget.Offset(0, iCol - LBound(vFields) + 1).Value = vFields(iCol)
    Next iCol

    ' Data rows.
    iRow = 1
    For Each vId In vIds
        pTarget.Offset(iRow, 0).Value = vId
        For iCol = LBound(vFields) To UBound(vFields)
            pTarget.Offset(iRow, iCol - LBound(vFields) + 1).Value = _
                JStr(oAll.Item(vId), CStr(vFields(iCol)), "")
        Next iCol
        iRow = iRow + 1
    Next vId

    ItemsToSheet = iRow - 1
End Function

' Turn a sheet range into an Assign payload.
'   Row 1    = header (first cell is the id column, the rest are field names)
'   Row 2..n = data. Rows with a blank id are skipped.
'   A cell whose text starts with "[" or "{" is parsed as JSON.
Public Function SheetToAssign(ByVal pSource As Range) As Object
    Dim oOut As Object
    Dim oOne As Object
    Dim vGrid As Variant
    Dim iRow As Long
    Dim iCol As Long
    Dim sId As String
    Dim sField As String
    Dim vRaw As Variant

    Set oOut = JNew()
    If pSource Is Nothing Then
        Set SheetToAssign = oOut
        Exit Function
    End If

    vGrid = pSource.Value
    If Not IsArray(vGrid) Then
        Set SheetToAssign = oOut
        Exit Function
    End If

    For iRow = LBound(vGrid, 1) + 1 To UBound(vGrid, 1)
        sId = Trim$(CStr(vGrid(iRow, LBound(vGrid, 2)) & ""))
        If Len(sId) > 0 Then
            Set oOne = JNew()

            For iCol = LBound(vGrid, 2) + 1 To UBound(vGrid, 2)
                sField = Trim$(CStr(vGrid(LBound(vGrid, 1), iCol) & ""))
                If Len(sField) > 0 Then
                    vRaw = vGrid(iRow, iCol)
                    JSet oOne, sField, CvCoerceCell(vRaw)
                End If
            Next iCol

            JSet oOut, sId, oOne
        End If
    Next iRow

    Set SheetToAssign = oOut
End Function

' Convert one cell value into something JSON friendly.
Private Function CvCoerceCell(ByVal pRaw As Variant) As Variant
    Dim sTxt As String
    Dim oTmp As Object

    If IsEmpty(pRaw) Then
        CvCoerceCell = ""
        Exit Function
    End If

    If IsNumeric(pRaw) And VarType(pRaw) <> vbString Then
        CvCoerceCell = pRaw
        Exit Function
    End If

    If VarType(pRaw) = vbBoolean Then
        CvCoerceCell = CBool(pRaw)
        Exit Function
    End If

    sTxt = Trim$(CStr(pRaw))

    Select Case UCase$(sTxt)
        Case "TRUE":  CvCoerceCell = True:  Exit Function
        Case "FALSE": CvCoerceCell = False: Exit Function
    End Select

    If Left$(sTxt, 1) = "[" Or Left$(sTxt, 1) = "{" Then
        Set oTmp = JParse(sTxt)
        If Not oTmp Is Nothing Then
            Set CvCoerceCell = oTmp
            Exit Function
        End If
    End If

    CvCoerceCell = sTxt
End Function

' Collect one column (or row) into a Collection, skipping blanks.
Public Function RangeToList(ByVal pSource As Range, _
                            Optional ByVal pAsNumber As Boolean = False) As Collection
    Dim oOut As Collection
    Dim oCell As Range
    Dim sTxt As String

    Set oOut = New Collection
    If pSource Is Nothing Then
        Set RangeToList = oOut
        Exit Function
    End If

    For Each oCell In pSource.Cells
        sTxt = Trim$(CStr(oCell.Value & ""))
        If Len(sTxt) > 0 Then
            If pAsNumber And IsNumeric(sTxt) Then
                oOut.Add CDbl(sTxt)
            Else
                oOut.Add sTxt
            End If
        End If
    Next oCell

    Set RangeToList = oOut
End Function

' 1,2,3,7,8,9 -> "1to3 7to9", the range notation MIDAS tables expect.
Public Function IdsToRangeText(ByVal pIds As Variant) As String
    Dim vOne As Variant
    Dim lPrev As Long
    Dim lStart As Long
    Dim bFirst As Boolean
    Dim sOut As String

    bFirst = True

    For Each vOne In pIds
        If IsNumeric(vOne) Then
            If bFirst Then
                lStart = CLng(vOne)
                lPrev = lStart
                bFirst = False
            ElseIf CLng(vOne) = lPrev + 1 Then
                lPrev = CLng(vOne)
            Else
                sOut = sOut & " " & CvRangePart(lStart, lPrev)
                lStart = CLng(vOne)
                lPrev = lStart
            End If
        End If
    Next vOne

    If Not bFirst Then sOut = sOut & " " & CvRangePart(lStart, lPrev)
    IdsToRangeText = Trim$(sOut)
End Function

Private Function CvRangePart(ByVal pFrom As Long, ByVal pTo As Long) As String
    If pFrom = pTo Then
        CvRangePart = CStr(pFrom)
    Else
        CvRangePart = CStr(pFrom) & "to" & CStr(pTo)
    End If
End Function


'==========================================================
' [8] Helpers - hand written wrappers for the items used most often
'==========================================================
'  These are the original CivilVBA helper functions: same names, same
'  parameter order. Two things changed:
'    - helpers that write model data add it to the store ([10]) instead of
'      sending it; call ModelCreate when the model is written
'    - parameters are ByVal and ids are Long, so any numeric variable can
'      be passed (no ByRef type mismatch) and ids above 32767 work
'  Anything not covered here can be sent with ApiRaw ([6]) or the raw calls
'  in [3] (CallGet / CallPost / CallPut / CallDelete).
'==========================================================

Function NewFile() As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", New Dictionary
    NewFile = CallPost("/doc/NEW", JsonConverter.ConvertToJson(body))
End Function

Function OpenFile(ByVal filePath As String) As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", filePath
    OpenFile = CallPost("/doc/OPEN", JsonConverter.ConvertToJson(body))
End Function

Function CloseFile() As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", New Dictionary
    CloseFile = CallPost("/doc/CLOSE", JsonConverter.ConvertToJson(body))
End Function

Function SaveFile() As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", New Dictionary
    ' send what is still in the store first ([10])
    Dim sPending As String
    sPending = CvFlushPending()
    If Len(sPending) > 0 Then
        SaveFile = "{""error"":""" & Replace(sPending, """", "'") & """}"
        Exit Function
    End If
    SaveFile = CallPost("/doc/SAVE", JsonConverter.ConvertToJson(body))
End Function

Function SaveFileAs(ByVal filePath As String) As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", filePath
    ' send what is still in the store first ([10])
    Dim sPending As String
    sPending = CvFlushPending()
    If Len(sPending) > 0 Then
        SaveFileAs = "{""error"":""" & Replace(sPending, """", "'") & """}"
        Exit Function
    End If
    SaveFileAs = CallPost("/doc/SAVEAS", JsonConverter.ConvertToJson(body))
End Function

Function SaveStageAs(ByVal filePath As String) As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", filePath
    ' send what is still in the store first ([10])
    Dim sPending As String
    sPending = CvFlushPending()
    If Len(sPending) > 0 Then
        SaveStageAs = "{""error"":""" & Replace(sPending, """", "'") & """}"
        Exit Function
    End If
    SaveStageAs = CallPost("/doc/STAGAS", JsonConverter.ConvertToJson(body))
End Function

Function ImportJson(ByVal filePath As String) As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", filePath
    ImportJson = CallPost("/doc/IMPORT", JsonConverter.ConvertToJson(body))
End Function

Function ImportMct(ByVal filePath As String) As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", filePath
    ImportMct = CallPost("/doc/IMPORTMXT", JsonConverter.ConvertToJson(body))
End Function

Function ExportJson(ByVal filePath As String) As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", filePath
    ' send what is still in the store first ([10])
    Dim sPending As String
    sPending = CvFlushPending()
    If Len(sPending) > 0 Then
        ExportJson = "{""error"":""" & Replace(sPending, """", "'") & """}"
        Exit Function
    End If
    ExportJson = CallPost("/doc/EXPORT", JsonConverter.ConvertToJson(body))
End Function

Function ExportMct(ByVal filePath As String) As String
    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", filePath
    ' send what is still in the store first ([10])
    Dim sPending As String
    sPending = CvFlushPending()
    If Len(sPending) > 0 Then
        ExportMct = "{""error"":""" & Replace(sPending, """", "'") & """}"
        Exit Function
    End If
    ExportMct = CallPost("/doc/EXPORTMXT", JsonConverter.ConvertToJson(body))
End Function

Function Unit(Optional ByVal FORCE As String = "", Optional ByVal DIST As String = "", _
              Optional ByVal HEAT As String = "", Optional ByVal TEMPER As String = "") As String

    If FORCE = "" And DIST = "" And HEAT = "" And TEMPER = "" Then
        Unit = CallGet("/db/UNIT")
        Exit Function
    End If

    Dim unitData As Object, assignData As Object, body As Object

    Set unitData = New Dictionary
    If FORCE <> "" Then unitData.Add "FORCE", UCase(FORCE)
    If DIST <> "" Then unitData.Add "DIST", UCase(DIST)
    If HEAT <> "" Then unitData.Add "HEAT", UCase(HEAT)
    If TEMPER <> "" Then unitData.Add "TEMPER", UCase(TEMPER)

    Set assignData = New Dictionary
    assignData.Add "1", unitData

    Set body = New Dictionary
    body.Add "Assign", assignData

    Unit = CvQueue("/db/UNIT", body)

End Function

Function StructureType(Optional ByVal STYP As Variant = Null, _
                       Optional ByVal MASS As Variant = Null, _
                       Optional ByVal bMASSOFFSET As Variant = Null, _
                       Optional ByVal bSELFWEIGHT As Variant = Null, _
                       Optional ByVal SMASS As Variant = Null, _
                       Optional ByVal GRAV As Variant = Null, _
                       Optional ByVal TEMP As Variant = Null, _
                       Optional ByVal bALIGNBEAM As Variant = Null, _
                       Optional ByVal bALIGNSLAB As Variant = Null, _
                       Optional ByVal bROTRIGID As Variant = Null) As String

    If IsNull(STYP) And IsNull(MASS) And IsNull(bMASSOFFSET) And IsNull(bSELFWEIGHT) And _
       IsNull(SMASS) And IsNull(GRAV) And IsNull(TEMP) And IsNull(bALIGNBEAM) And _
       IsNull(bALIGNSLAB) And IsNull(bROTRIGID) Then
        StructureType = CallGet("/db/STYP")
        Exit Function
    End If

    Dim item As Object, assignData As Object, body As Object

    Set item = New Dictionary
    If Not IsNull(STYP) Then item.Add "STYP", STYP
    If Not IsNull(MASS) Then item.Add "MASS", MASS
    If Not IsNull(bMASSOFFSET) Then item.Add "bMASSOFFSET", bMASSOFFSET
    If Not IsNull(bSELFWEIGHT) Then item.Add "bSELFWEIGHT", bSELFWEIGHT
    If Not IsNull(SMASS) Then item.Add "SMASS", SMASS
    If Not IsNull(GRAV) Then item.Add "GRAV", GRAV
    If Not IsNull(TEMP) Then item.Add "TEMP", TEMP
    If Not IsNull(bALIGNBEAM) Then item.Add "bALIGNBEAM", bALIGNBEAM
    If Not IsNull(bALIGNSLAB) Then item.Add "bALIGNSLAB", bALIGNSLAB
    If Not IsNull(bROTRIGID) Then item.Add "bROTRIGID", bROTRIGID

    Set assignData = New Dictionary
    assignData.Add "1", item

    Set body = New Dictionary
    body.Add "Assign", assignData

    StructureType = CvQueue("/db/STYP", body)

End Function


' ------------------------------------------
' Material Properties
' ------------------------------------------
' Material(matlId, matlType, NAME, STANDARD, DB, [CODE], [STANDARD2], [DB2], [CODE2], [DAMP_RAT], [HE_SPEC], [HE_COND])
' MaterialUser(matlId, matlType, NAME, ELAST, POISN, THERMAL, DEN, MASS, [SHEAR], [DAMP_RAT], [HE_SPEC], [HE_COND])
' ------------------------------------------

Function Material(ByVal matlId As Long, ByVal matlType As String, ByVal NAME As String, _
                  ByVal STANDARD As String, ByVal DB As String, Optional ByVal CODE As String = "", _
                  Optional ByVal STANDARD2 As String = "", Optional ByVal DB2 As String = "", _
                  Optional ByVal CODE2 As String = "", Optional ByVal DAMP_RAT As Variant = Null, _
                  Optional ByVal HE_SPEC As Double = 0, Optional ByVal HE_COND As Double = 0) As String

    If matlId = 0 Then
        Material = CallGet("/db/MATL")
        Exit Function
    End If

    Dim dampRat As Variant
    ' 0.05 for every material type
    If IsNull(DAMP_RAT) Then
        dampRat = 0.05
    Else
        dampRat = DAMP_RAT
    End If

    Dim param1 As Object
    Set param1 = New Dictionary
    param1.Add "P_TYPE", 1
    param1.Add "STANDARD", STANDARD
    param1.Add "CODE", CODE
    param1.Add "DB", DB

    Dim paramArr As Variant

    If STANDARD2 <> "" Then
        Dim param2 As Object
        Set param2 = New Dictionary
        param2.Add "P_TYPE", 1
        param2.Add "STANDARD", STANDARD2
        param2.Add "CODE", CODE2
        param2.Add "DB", DB2

        Dim arr2(1) As Object
        Set arr2(0) = param1
        Set arr2(1) = param2
        paramArr = arr2
    Else
        Dim arr1(0) As Object
        Set arr1(0) = param1
        paramArr = arr1
    End If

    Dim item As Object
    Set item = New Dictionary
    item.Add "TYPE", matlType
    item.Add "NAME", NAME
    item.Add "HE_SPEC", HE_SPEC
    item.Add "HE_COND", HE_COND
    If Not IsNull(dampRat) Then item.Add "DAMP_RAT", dampRat
    item.Add "PARAM", paramArr

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(matlId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    Material = CvQueue("/db/MATL", body)

End Function

' db/MATL - User Defined (Isotropic / Orthotropic)
Function MaterialUser(ByVal matlId As Long, ByVal matlType As String, ByVal NAME As String, _
                      ByVal ELAST As Variant, ByVal POISN As Variant, ByVal THERMAL As Variant, _
                      ByVal DEN As Double, ByVal MASS As Double, _
                      Optional ByVal SHEAR As Variant = Null, _
                      Optional ByVal DAMP_RAT As Double = 0.05, Optional ByVal HE_SPEC As Double = 0, _
                      Optional ByVal HE_COND As Double = 0) As String

    If matlId = 0 Then
        MaterialUser = CallGet("/db/MATL")
        Exit Function
    End If

    Dim param As Object
    Set param = New Dictionary

    If IsArray(ELAST) Then
        param.Add "P_TYPE", 3
        param.Add "ELAST_M", ELAST
        param.Add "POISN_M", POISN
        param.Add "THERMAL_M", THERMAL
        param.Add "SHEAR_M", SHEAR
    Else
        param.Add "P_TYPE", 2
        param.Add "ELAST", ELAST
        param.Add "POISN", POISN
        param.Add "THERMAL", THERMAL
    End If

    param.Add "DEN", DEN
    param.Add "MASS", MASS

    Dim paramArr(0) As Object
    Set paramArr(0) = param

    Dim item As Object
    Set item = New Dictionary
    item.Add "TYPE", matlType
    item.Add "NAME", NAME
    item.Add "HE_SPEC", HE_SPEC
    item.Add "HE_COND", HE_COND
    item.Add "DAMP_RAT", DAMP_RAT
    item.Add "PARAM", paramArr

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(matlId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    MaterialUser = CvQueue("/db/MATL", body)

End Function

' ------------------------------------------
' Section Properties - db/SECT, SECTTYPE "DBUSER"
' ------------------------------------------
' The DB/User tab of the Section dialog has two input modes and this block
' covers both for every shape the product ships.
'
'   user dimensions (DATATYPE 2) - one function per shape:
'     SectionSolidRectangle(id, name, H, B)                        SB
'     SectionSolidRound(id, name, D)                               SR
'     SectionHSection(id, name, H, B1, tw, tf1, [B2,tf2,r1,r2])    H
'     SectionTee(id, name, H, B, tw, tf)                           T
'     SectionBox(id, name, H, B, tw, tf1, [C, tf2])                B
'     SectionPipe(id, name, D, t)                                  P
'     SectionChannel(id, name, H, B1, tw, tf1, [B2,tf2])           C
'     SectionAngle(id, name, H, B, tw, tf)                         L
'     SectionDoubleAngle(id, name, H, B, tw, tf, [C])              2L
'     SectionDoubleChannel(id, name, H, B, tw, tf, [C])            2C
'     SectionColdChannel(id, name, H, B, tw, r, [d])               CC
'     SectionZSection(id, name, H, B, t, r, [d], [lipAngle])       Z
'     SectionStarAngle(id, name, H, B, tw, tf, [C])                CL
'     SectionUpright(id, name, v1 ... v10)                         UP
'     SectionURib(id, name, H, B1, B2, t, [r])                     URIB
'     SectionInvertedTee(id, name, H, B1, B2, tw, tf)              UDT
'     SectionOctagon(id, name, H, B, rY, rZ, t)                    OCT
'     SectionSolidOctagon(id, name, H, B, rY, rZ)                  SOCT
'     SectionRoundOctagon(id, name, H, B, rY, rZ, tw, tf,
'                         [twInner], [cells])                      ROCT
'     SectionTrack(id, name, H, B, t)                              TRK
'     SectionSolidTrack(id, name, H, B)                            STRK
'     SectionHalfTrack(id, name, H, B)                             HTRK
'     SectionPipeStiffener(id, name, D, t, hStif, tStif, [nStif])  PSTF
'     SectionBoxStiffener(id, name, H, B, tf, tw, sWeb, hWeb, tWeb,
'                         sFlange, hFlange, tFlange,
'                         [nWeb], [nFlange])                       BSTF
'
'   Every dimension order above was checked against the running product
'   with ope/SECTPROP, not guessed.  Lengths are in the model's own unit,
'   so set db/UNIT first (or call SetUnit).
'
'   picked from a steel database (DATATYPE 1):
'     SectionDb(id, name, shape, dbName, specName)
'       dbName is the database - KS, KS21, AISC, AISC10(SI), JIS, JIS2K,
'       BS, BS4-93, EN_10365-2017, GB-YB ... 17 of them.
'       specName is the entry inside it, e.g. "H 200x200x8/12".
'       Use Ope_SECT_DBNAME / Ope_SECT_NAME to list what the product has
'       rather than inventing a name.
'
'   anything else, or a shape added by a later version:
'     SectionUser(id, name, shape, vSize)   vSize is SECT_I.vSIZE itself
'
' Every one of these takes the same optional tail:
'   offsetPt "CC" (default), useShearDeform True, useWarping False
'
' Passing sectId = 0 reads the whole item instead of writing, matching the
' other helpers in this module.
' ------------------------------------------

' Any shape, dimensions passed as the raw vSIZE array. Always exact, and
' keeps working if a later version adds a shape this module does not name.
' vSize accepts an Array, a Collection, a worksheet Range or a single number
' and is padded to the 10 slots the API expects.
Function SectionUser(ByVal sectId As Long, ByVal sectName As String, ByVal shape As String, _
                     Optional ByVal vSize As Variant, Optional ByVal offsetPt As String = "CC", _
                     Optional ByVal useShearDeform As Boolean = True, _
                     Optional ByVal useWarping As Boolean = False) As String

    Dim slot(0 To 9) As Double
    Dim vOne As Variant
    Dim i As Long
    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionUser = CallGet("/db/SECT")
        Exit Function
    End If

    If Not IsMissing(vSize) Then
        If IsObject(vSize) Then
            For Each vOne In vSize
                If i > 9 Then Exit For
                If IsNumeric(vOne) Then slot(i) = CDbl(vOne)
                i = i + 1
            Next vOne
        ElseIf IsArray(vSize) Then
            For Each vOne In vSize
                If i > 9 Then Exit For
                If IsNumeric(vOne) Then slot(i) = CDbl(vOne)
                i = i + 1
            Next vOne
        ElseIf IsNumeric(vSize) Then
            slot(0) = CDbl(vSize)
        End If
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", slot

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", UCase$(Trim$(shape))
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionUser = CvQueue("/db/SECT", body)
End Function

' Picked from a steel database instead of typed in.
' Ope_SECT_DBNAME lists the databases, Ope_SECT_NAME the entries in one.
Function SectionDb(ByVal sectId As Long, ByVal sectName As String, ByVal shape As String, _
                   ByVal dbName As String, ByVal specName As String, _
                   Optional ByVal offsetPt As String = "CC", _
                   Optional ByVal useShearDeform As Boolean = True, _
                   Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionDb = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "DB_NAME", dbName
    sectI.Add "SECT_NAME", specName

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", UCase$(Trim$(shape))
    sectBefore.Add "DATATYPE", 1
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionDb = CvQueue("/db/SECT", body)
End Function

' SB : Solid Rectangle.  H, B
' H = depth, B = width.
' (verified with ope/SECTPROP)
Function SectionSolidRectangle(ByVal sectId As Long, ByVal NAME As String, ByVal h As Double, _
                               ByVal B As Double, Optional ByVal offsetPt As String = "CC", _
                               Optional ByVal useShearDeform As Boolean = True, _
                               Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionSolidRectangle = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, 0, 0, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "SB"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", NAME
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionSolidRectangle = CvQueue("/db/SECT", body)
End Function

' SR : Solid Round.  D
' D = diameter.
' (verified with ope/SECTPROP)
Function SectionSolidRound(ByVal sectId As Long, ByVal NAME As String, ByVal d As Double, _
                           Optional ByVal offsetPt As String = "CC", _
                           Optional ByVal useShearDeform As Boolean = True, _
                           Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionSolidRound = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(d, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "SR"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", NAME
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionSolidRound = CvQueue("/db/SECT", body)
End Function

' H : H-Section.  H, B1, tw, tf1, [B2], [tf2], [r1], [r2]
' B2 and tf2 default to B1 and tf1, giving a symmetric section.
' (verified with ope/SECTPROP)
Function SectionHSection(ByVal sectId As Long, ByVal NAME As String, ByVal h As Double, _
                         ByVal B1 As Double, ByVal tw As Double, ByVal tf1 As Double, _
                         Optional ByVal B2 As Variant = Null, Optional ByVal tf2 As Variant = Null, _
                         Optional ByVal r1 As Double = 0, Optional ByVal r2 As Double = 0, _
                         Optional ByVal offsetPt As String = "CC", _
                         Optional ByVal useShearDeform As Boolean = True, _
                         Optional ByVal useWarping As Boolean = False) As String

    Dim act_B2 As Double
    Dim act_tf2 As Double
    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionHSection = CallGet("/db/SECT")
        Exit Function
    End If

    If IsNull(B2) Then act_B2 = B1 Else act_B2 = B2
    If IsNull(tf2) Then act_tf2 = tf1 Else act_tf2 = tf2

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B1, tw, tf1, act_B2, act_tf2, r1, r2, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "H"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", NAME
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionHSection = CvQueue("/db/SECT", body)
End Function

' T : T-Section.  H, B, tw, tf
' Flange on top, web down.
' (verified with ope/SECTPROP)
Function SectionTee(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                    ByVal B As Double, ByVal tw As Double, ByVal tf As Double, _
                    Optional ByVal offsetPt As String = "CC", _
                    Optional ByVal useShearDeform As Boolean = True, _
                    Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionTee = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, tw, tf, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "T"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionTee = CvQueue("/db/SECT", body)
End Function

' B : Box.  H, B, tw, tf1, [c], [tf2]
' c is the inner flange spacing, tf2 the bottom flange.
' (verified with ope/SECTPROP)
Function SectionBox(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                    ByVal B As Double, ByVal tw As Double, ByVal tf1 As Double, _
                    Optional ByVal c As Double = 0, Optional ByVal tf2 As Double = 0, _
                    Optional ByVal offsetPt As String = "CC", _
                    Optional ByVal useShearDeform As Boolean = True, _
                    Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionBox = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, tw, tf1, c, tf2, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "B"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionBox = CvQueue("/db/SECT", body)
End Function

' P : Pipe.  D, t
' D = outer diameter, t = wall thickness.
' (verified with ope/SECTPROP)
Function SectionPipe(ByVal sectId As Long, ByVal sectName As String, ByVal d As Double, _
                     ByVal t As Double, Optional ByVal offsetPt As String = "CC", _
                     Optional ByVal useShearDeform As Boolean = True, _
                     Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionPipe = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(d, t, 0, 0, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "P"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionPipe = CvQueue("/db/SECT", body)
End Function

' C : Channel.  H, B1, tw, tf1, [B2], [tf2], [r1], [r2]
' Same dimension order as H.
' (verified with ope/SECTPROP)
Function SectionChannel(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                        ByVal B1 As Double, ByVal tw As Double, ByVal tf1 As Double, _
                        Optional ByVal B2 As Variant = Null, Optional ByVal tf2 As Variant = Null, _
                        Optional ByVal r1 As Double = 0, Optional ByVal r2 As Double = 0, _
                        Optional ByVal offsetPt As String = "CC", _
                        Optional ByVal useShearDeform As Boolean = True, _
                        Optional ByVal useWarping As Boolean = False) As String

    Dim act_B2 As Double
    Dim act_tf2 As Double
    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionChannel = CallGet("/db/SECT")
        Exit Function
    End If

    If IsNull(B2) Then act_B2 = B1 Else act_B2 = B2
    If IsNull(tf2) Then act_tf2 = tf1 Else act_tf2 = tf2

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B1, tw, tf1, act_B2, act_tf2, r1, r2, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "C"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionChannel = CvQueue("/db/SECT", body)
End Function

' L : Angle.  H, B, tw, tf
' Equal leg angles use the same value for H and B, and for tw and tf.
' (verified with ope/SECTPROP)
Function SectionAngle(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                      ByVal B As Double, ByVal tw As Double, ByVal tf As Double, _
                      Optional ByVal offsetPt As String = "CC", _
                      Optional ByVal useShearDeform As Boolean = True, _
                      Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionAngle = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, tw, tf, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "L"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionAngle = CvQueue("/db/SECT", body)
End Function

' 2L : Double Angle.  H, B, tw, tf, [c]
' c is the gap between the two angles.
' (verified with ope/SECTPROP)
Function SectionDoubleAngle(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                            ByVal B As Double, ByVal tw As Double, ByVal tf As Double, _
                            Optional ByVal c As Double = 0, _
                            Optional ByVal offsetPt As String = "CC", _
                            Optional ByVal useShearDeform As Boolean = True, _
                            Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionDoubleAngle = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, tw, tf, c, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "2L"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionDoubleAngle = CvQueue("/db/SECT", body)
End Function

' 2C : Double Channel.  H, B1, tw, tf1, [B2], [tf2], [c]
' c is the gap between the two channels.
' (verified with ope/SECTPROP)
Function SectionDoubleChannel(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                              ByVal B1 As Double, ByVal tw As Double, ByVal tf1 As Double, _
                              Optional ByVal B2 As Variant = Null, _
                              Optional ByVal tf2 As Variant = Null, Optional ByVal c As Double = 0, _
                              Optional ByVal offsetPt As String = "CC", _
                              Optional ByVal useShearDeform As Boolean = True, _
                              Optional ByVal useWarping As Boolean = False) As String

    Dim act_B2 As Double
    Dim act_tf2 As Double
    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionDoubleChannel = CallGet("/db/SECT")
        Exit Function
    End If

    If IsNull(B2) Then act_B2 = B1 Else act_B2 = B2
    If IsNull(tf2) Then act_tf2 = tf1 Else act_tf2 = tf2

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B1, tw, tf1, act_B2, act_tf2, c, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "2C"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionDoubleChannel = CvQueue("/db/SECT", body)
End Function

' CC : Cold Formed Channel.  H, B, tw, r, [d]
' tw is the sheet thickness, r the corner radius, d the lip.
' (verified with ope/SECTPROP and against the Section Data dialog)
Function SectionColdChannel(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                            ByVal B As Double, ByVal tw As Double, ByVal r As Double, _
                            Optional ByVal d As Double = 0, _
                            Optional ByVal offsetPt As String = "CC", _
                            Optional ByVal useShearDeform As Boolean = True, _
                            Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionColdChannel = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, tw, r, d, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "CC"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionColdChannel = CvQueue("/db/SECT", body)
End Function

' UP : Upright (rack post).  10 dimensions
' The profile is a lipped, multiply folded post; the product takes
' all ten slots and rejects most partial sets, so pass what the
' Section dialog shows.  A working set is
'   0.81, 0.61, 0.007, 0.421, 0.312, 0.113, 0.124, 0.325, 0.16, 0.037
' Measured: v1 = overall depth H, v2 = overall width B,
' v3 = sheet thickness.  v4 and v5 must stay inside v1/v2 or the
' shape is rejected; v6..v10 are the fold and lip dimensions.
' (v1, v2, v3 verified with ope/SECTPROP; v4..v10 taken from the
'  Section Properties DB/User manual example)
Function SectionUpright(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                        ByVal B As Double, ByVal t As Double, ByVal v4 As Double, _
                        ByVal v5 As Double, ByVal v6 As Double, ByVal v7 As Double, _
                        ByVal v8 As Double, ByVal v9 As Double, ByVal v10 As Double, _
                        Optional ByVal offsetPt As String = "CC", _
                        Optional ByVal useShearDeform As Boolean = True, _
                        Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionUpright = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, t, v4, v5, v6, v7, v8, v9, v10)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "UP"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionUpright = CvQueue("/db/SECT", body)
End Function

' URIB : U-Rib.  H, B1, B2, t, [r]
' H  = height, B1 = top opening width, B2 = bottom width,
' t  = plate thickness, r = bend radius.
' B2 sits inside B1, so it changes the area but not the
' overall width.
' (verified with ope/SECTPROP)
Function SectionURib(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                     ByVal b1 As Double, ByVal b2 As Double, ByVal t As Double, _
                     Optional ByVal r As Double = 0, Optional ByVal offsetPt As String = "CC", _
                     Optional ByVal useShearDeform As Boolean = True, _
                     Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionURib = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, b1, b2, t, r, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "URIB"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionURib = CvQueue("/db/SECT", body)
End Function

' Z : Z-Section (cold formed purlin).  H, B, t, r, [d], [lipAngle]
' H = depth, B = flange width, t = sheet thickness,
' r = bend radius, d = lip length, lipAngle = lip angle in degrees
' (50 on the AISI purlins; 0 lays the lip flat).
' (verified with ope/SECTPROP)
Function SectionZSection(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                         ByVal B As Double, ByVal t As Double, ByVal r As Double, _
                         Optional ByVal d As Double = 0, Optional ByVal lipAngle As Double = 0, _
                         Optional ByVal offsetPt As String = "CC", _
                         Optional ByVal useShearDeform As Boolean = True, _
                         Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionZSection = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, t, r, d, lipAngle, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "Z"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionZSection = CvQueue("/db/SECT", body)
End Function

' CL : Star-battened Angle (cross shaped pair of angles).
'   H, B, tw, tf, [c]
' H = leg depth, B = leg width, tw = vertical leg thickness,
' tf = horizontal leg thickness, c = gap between the two angles.
' Overall size comes out near 2H and 2B because the angles are
' battened back to back.
' (verified with ope/SECTPROP)
Function SectionStarAngle(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                          ByVal B As Double, ByVal tw As Double, ByVal tf As Double, _
                          Optional ByVal c As Double = 0, Optional ByVal offsetPt As String = "CC", _
                          Optional ByVal useShearDeform As Boolean = True, _
                          Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionStarAngle = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, tw, tf, c, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "CL"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionStarAngle = CvQueue("/db/SECT", body)
End Function

' OCT : Octagon, hollow.  H, B, rY, rZ, t
' H = depth, B = width, rY and rZ = the horizontal and vertical
' corner cuts, t = wall thickness.
' (verified with ope/SECTPROP)
Function SectionOctagon(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                        ByVal B As Double, ByVal rY As Double, ByVal rZ As Double, _
                        ByVal t As Double, Optional ByVal offsetPt As String = "CC", _
                        Optional ByVal useShearDeform As Boolean = True, _
                        Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionOctagon = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, rY, rZ, t, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "OCT"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionOctagon = CvQueue("/db/SECT", body)
End Function

' SOCT : Octagon, solid.  H, B, rY, rZ
' H = depth, B = width, rY and rZ = the horizontal and vertical
' corner cuts.
' (verified with ope/SECTPROP)
Function SectionSolidOctagon(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                             ByVal B As Double, ByVal rY As Double, ByVal rZ As Double, _
                             Optional ByVal offsetPt As String = "CC", _
                             Optional ByVal useShearDeform As Boolean = True, _
                             Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionSolidOctagon = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, rY, rZ, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "SOCT"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionSolidOctagon = CvQueue("/db/SECT", body)
End Function

' ROCT : Round Octagon, hollow, one or more cells.
'   H, B, rY, rZ, tw, tf, [twInner], [cells]
' H = depth, B = width, rY and rZ = the corner cuts,
' tw = outer web thickness, tf = top and bottom plate thickness,
' twInner = thickness of the interior webs,
' cells = number of cells (1 leaves no interior web, 3 adds two).
' (verified with ope/SECTPROP - cells maps to CELL_SHAPE)
Function SectionRoundOctagon(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                             ByVal B As Double, ByVal rY As Double, ByVal rZ As Double, _
                             ByVal tw As Double, ByVal tf As Double, _
                             Optional ByVal twInner As Double = 0, _
                             Optional ByVal cells As Long = 1, _
                             Optional ByVal offsetPt As String = "CC", _
                             Optional ByVal useShearDeform As Boolean = True, _
                             Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionRoundOctagon = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, rY, rZ, tw, tf, twInner, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "ROCT"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI
    sectBefore.Add "CELL_SHAPE", cells

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionRoundOctagon = CvQueue("/db/SECT", body)
End Function

' TRK : Track (stadium / obround), hollow.  H, B, t
' H = depth, B = width over the round ends, t = wall thickness.
' (verified with ope/SECTPROP)
Function SectionTrack(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                      ByVal B As Double, ByVal t As Double, _
                      Optional ByVal offsetPt As String = "CC", _
                      Optional ByVal useShearDeform As Boolean = True, _
                      Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionTrack = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, t, 0, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "TRK"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionTrack = CvQueue("/db/SECT", body)
End Function

' STRK : Track (stadium / obround), solid.  H, B
' H = depth, B = width over the round ends.
' (verified with ope/SECTPROP)
Function SectionSolidTrack(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                           ByVal B As Double, Optional ByVal offsetPt As String = "CC", _
                           Optional ByVal useShearDeform As Boolean = True, _
                           Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionSolidTrack = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, 0, 0, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "STRK"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionSolidTrack = CvQueue("/db/SECT", body)
End Function

' HTRK : Half Track (flat bottom, round top), solid.  H, B
' H = depth, B = width.
' (verified with ope/SECTPROP)
Function SectionHalfTrack(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                          ByVal B As Double, Optional ByVal offsetPt As String = "CC", _
                          Optional ByVal useShearDeform As Boolean = True, _
                          Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionHalfTrack = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, 0, 0, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "HTRK"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionHalfTrack = CvQueue("/db/SECT", body)
End Function

' UDT : Inverted T-Section.  H, B1, B2, tw, tf
' H = depth, B1 = flange width on one side of the web,
' B2 = flange width on the other side, tw = web thickness,
' tf = flange thickness.  Overall width is B1 + B2 + tw, so the
' two sides can differ.
' (verified with ope/SECTPROP)
Function SectionInvertedTee(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                            ByVal b1 As Double, ByVal b2 As Double, ByVal tw As Double, _
                            ByVal tf As Double, Optional ByVal offsetPt As String = "CC", _
                            Optional ByVal useShearDeform As Boolean = True, _
                            Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionInvertedTee = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, b1, b2, tw, tf, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "UDT"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionInvertedTee = CvQueue("/db/SECT", body)
End Function

' PSTF : Pipe with longitudinal stiffeners.
'   D, t, hStif, tStif, [nStif]
' D = outside diameter, t = wall thickness,
' hStif and tStif = height and thickness of one stiffener,
' nStif = how many stiffeners are spaced around the pipe.
' nStif must be 1 or more; the product rejects 0.
' (verified with ope/SECTPROP - nStif maps to CELL_SHAPE)
Function SectionPipeStiffener(ByVal sectId As Long, ByVal sectName As String, ByVal d As Double, _
                              ByVal t As Double, ByVal hStif As Double, ByVal tStif As Double, _
                              Optional ByVal nStif As Long = 1, _
                              Optional ByVal offsetPt As String = "CC", _
                              Optional ByVal useShearDeform As Boolean = True, _
                              Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionPipeStiffener = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(d, t, hStif, tStif, 0, 0, 0, 0, 0, 0)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "PSTF"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI
    sectBefore.Add "CELL_SHAPE", nStif

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionPipeStiffener = CvQueue("/db/SECT", body)
End Function

' BSTF : Box with longitudinal stiffeners.
'   H, B, tf, tw, sWeb, hWeb, tWeb, sFlange, hFlange, tFlange,
'   [nWeb], [nFlange]
' H  = depth, B = width,
' tf = top and bottom plate thickness - note this comes BEFORE tw,
'      unlike SectionBox,
' tw = web thickness,
' sWeb, hWeb, tWeb = spacing, height and thickness of one web
'      stiffener,
' sFlange, hFlange, tFlange = the same for a flange stiffener,
' nWeb    = stiffeners on each web    (CELL_SHAPE),
' nFlange = stiffeners on each flange (CELL_TYPE).
' The stiffener dimensions cannot be left at 0 - the product
' rejects the section.
' (verified with ope/SECTPROP)
Function SectionBoxStiffener(ByVal sectId As Long, ByVal sectName As String, ByVal h As Double, _
                             ByVal B As Double, ByVal tf As Double, ByVal tw As Double, _
                             ByVal sWeb As Double, ByVal hWeb As Double, ByVal tWeb As Double, _
                             ByVal sFlange As Double, ByVal hFlange As Double, _
                             ByVal tFlange As Double, Optional ByVal nWeb As Long = 1, _
                             Optional ByVal nFlange As Long = 1, _
                             Optional ByVal offsetPt As String = "CC", _
                             Optional ByVal useShearDeform As Boolean = True, _
                             Optional ByVal useWarping As Boolean = False) As String

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionBoxStiffener = CallGet("/db/SECT")
        Exit Function
    End If

    Set sectI = New Dictionary
    sectI.Add "vSIZE", Array(h, B, tf, tw, sWeb, hWeb, tWeb, sFlange, hFlange, tFlange)

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", "BSTF"
    sectBefore.Add "DATATYPE", 2
    sectBefore.Add "SECT_I", sectI
    sectBefore.Add "CELL_SHAPE", nWeb
    sectBefore.Add "CELL_TYPE", nFlange

    Set item = New Dictionary
    item.Add "SECTTYPE", "DBUSER"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionBoxStiffener = CvQueue("/db/SECT", body)
End Function

' ------------------------------------------
' Section Properties - db/SECT, SECTTYPE "PSC"
' ------------------------------------------
' A PSC section is described by four dimension arrays instead of one vSIZE.
' They are the four groups of the Section Data dialog:
'
'   aSizes -> vSIZE_PSC_A   outer shape at the I end
'   bSizes -> vSIZE_PSC_B   inner shape (cell) at the I end
'   cSizes -> vSIZE_PSC_C   outer shape at the J end
'   dSizes -> vSIZE_PSC_D   inner shape (cell) at the J end
'   sWidth -> SWIDTH        the symmetric width the dialog shows on top
'
' How many numbers each array wants depends on the shape:
'
'   shape  function                 A   B   C   D   JOINT  options
'   1CEL   SectionPsc1Cell          6   6  10   8     8    -
'   2CEL   SectionPsc2Cell          6   6  10   8     8    -
'   3CEL   SectionPsc3Cell         10  12  13  13    13    -
'   NCEL   SectionPscNCell          5   8   -   -     -    opt1 CHAMFER,
'                                                          opt2 cell count
'   NCE2   SectionPscNCell2        11  18  11  18     2    opt1 CIRCLE|POLYGON,
'                                                          opt2 cell count
'   PSCM   SectionPscMid           10   6   9   6     9    opt1/opt2
'                                                          NONE|CIRCLE|POLYGON
'   PSCI   SectionPscI             10   7   9   7     9    -
'   PSCH   SectionPscHalf           6   6  10   7     8    opt1 LEFT|RIGHT,
'                                                          opt2 NONE|CIRCLE|POLYGON
'   PSCT   SectionPscTee            8   8   7   8     9    -
'   PSCB   SectionPscPlate         10   6   8   5     2    opt1 HALF|1CELL|2CELL,
'                                                          opt2 LEFT|RIGHT|
'                                                               CIRCLE|POLYGON
'
' Short arrays are padded with zeros and long ones are trimmed, so passing a
' worksheet Range of the right height is enough.  aSizes .. dSizes accept an
' Array, a Collection, a Range or a single number.
'
' JOINT turns the individual joints of the outline on and off.  The wrappers
' default it to all-on, which is what the dialog gives you; pass your own
' array of True/False to match a section that has some of them cleared.
'
' SectionPsc is the generic one and takes the shape code itself, so a shape
' a later version adds still works.
'
' Every dimension group above was created and read back with ope/SECTPROP on
' the running product.  Lengths are in the model's own unit.
' ------------------------------------------

' Turns whatever the caller passed into an array of exactly n Doubles.
' Design check fields written for PSC sections, with their defaults.
Private Sub PscCheckFields(ByVal pBefore As Object)
    pBefore.Add "WARPING_CHK_AUTO_I", True
    pBefore.Add "WARPING_CHK_AUTO_J", True
    pBefore.Add "SHEAR_CHK", True
    pBefore.Add "WARPING_CHK_POS_I", Array(CvNums(Empty, 6), CvNums(Empty, 6))
    pBefore.Add "WARPING_CHK_POS_J", Array(CvNums(Empty, 6), CvNums(Empty, 6))
    pBefore.Add "USE_AUTO_SHEAR_CHK_POS", Array(Array(True, False, True), Array(False, False, False))
    pBefore.Add "USE_WEB_THICK_SHEAR", Array(Array(True, True, True), Array(False, False, False))
    pBefore.Add "SHEAR_CHK_POS", Array(CvNums(Empty, 3), CvNums(Empty, 3))
    pBefore.Add "USE_WEB_THICK", Array(True, False)
    pBefore.Add "WEB_THICK", Array(0, 0)
End Sub

Private Function PscNums(ByVal src As Variant, ByVal n As Long) As Variant
    Dim out() As Double
    Dim one As Variant
    Dim i As Long

    If n < 1 Then
        PscNums = Array()
        Exit Function
    End If

    ReDim out(0 To n - 1)

    If IsMissing(src) Then
        PscNums = out
        Exit Function
    End If
    If Not IsObject(src) Then
        If IsEmpty(src) Then
            PscNums = out
            Exit Function
        End If
    End If

    If IsObject(src) Or IsArray(src) Then
        For Each one In src
            If i > n - 1 Then Exit For
            If IsNumeric(one) Then out(i) = CDbl(one)
            i = i + 1
        Next one
    ElseIf IsNumeric(src) Then
        out(0) = CDbl(src)
    End If

    PscNums = out
End Function

' All joints on - what the dialog gives you for a fresh section.
Private Function PscAllOn(ByVal n As Long) As Variant
    Dim out() As Boolean
    Dim i As Long
    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = True
    Next i
    PscAllOn = out
End Function

' Turns whatever the caller passed into an array of exactly n Booleans.
' Missing means "all joints on", which is the dialog default.
Private Function PscFlags(ByVal src As Variant, ByVal n As Long) As Variant
    Dim out() As Boolean
    Dim one As Variant
    Dim i As Long

    If n < 1 Then
        PscFlags = Array()
        Exit Function
    End If

    ReDim out(0 To n - 1)

    If IsMissing(src) Then
        PscFlags = PscAllOn(n)
        Exit Function
    End If
    If Not IsObject(src) Then
        If IsEmpty(src) Then
            PscFlags = PscAllOn(n)
            Exit Function
        End If
    End If

    If IsObject(src) Or IsArray(src) Then
        For Each one In src
            If i > n - 1 Then Exit For
            If Not IsEmpty(one) Then out(i) = CBool(one)
            i = i + 1
        Next one
    Else
        out(0) = CBool(src)
    End If

    PscFlags = out
End Function

' The one that does the work.  Every wrapper below calls it.
Function SectionPsc(ByVal sectId As Long, ByVal sectName As String, ByVal shape As String, _
                    Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                    Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                    Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                    Optional ByVal pscOpt1 As String = "", Optional ByVal pscOpt2 As String = "", _
                    Optional ByVal nA As Long = 0, Optional ByVal nB As Long = 0, _
                    Optional ByVal nC As Long = 0, Optional ByVal nD As Long = 0, _
                    Optional ByVal nJoint As Long = 0, _
                    Optional ByVal useSymmetric As Boolean = False, _
                    Optional ByVal useSmallHole As Boolean = False, _
                    Optional ByVal offsetPt As String = "CC", _
                    Optional ByVal useShearDeform As Boolean = True, _
                    Optional ByVal useWarping As Boolean = False, _
                    Optional ByVal checkFields As Boolean = False) As String
    ' checkFields: the 1 cell / 2 cell / I section shape -
    '   design check fields with their defaults, joints off unless given,
    '   SWIDTH only when not 0, USE_SYMMETRIC only for PSCI, no USE_SMALL_HOLE

    Dim sectI As Object
    Dim sectBefore As Object
    Dim item As Object
    Dim assignData As Object
    Dim body As Object

    If sectId = 0 Then
        SectionPsc = CallGet("/db/SECT")
        Exit Function
    End If

    If nA = 0 Then nA = 10
    If nB = 0 Then nB = 10
    If nC = 0 Then nC = 10
    If nD = 0 Then nD = 10

    Set sectI = New Dictionary
    sectI.Add "vSIZE_PSC_A", PscNums(aSizes, nA)
    sectI.Add "vSIZE_PSC_B", PscNums(bSizes, nB)
    If nC > 0 Then sectI.Add "vSIZE_PSC_C", PscNums(cSizes, nC)
    If nD > 0 Then sectI.Add "vSIZE_PSC_D", PscNums(dSizes, nD)
    If Not checkFields Or sWidth <> 0 Then sectI.Add "SWIDTH", sWidth

    Set sectBefore = New Dictionary
    sectBefore.Add "OFFSET_PT", offsetPt
    sectBefore.Add "OFFSET_CENTER", 0
    sectBefore.Add "USER_OFFSET_REF", 0
    sectBefore.Add "HORZ_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_YI", 0
    sectBefore.Add "USERDEF_OFFSET_YJ", 0
    sectBefore.Add "VERT_OFFSET_OPT", 0
    sectBefore.Add "USERDEF_OFFSET_ZI", 0
    sectBefore.Add "USERDEF_OFFSET_ZJ", 0
    sectBefore.Add "USE_SHEAR_DEFORM", useShearDeform
    sectBefore.Add "USE_WARPING_EFFECT", useWarping
    sectBefore.Add "SHAPE", UCase$(Trim$(shape))
    sectBefore.Add "SECT_I", sectI
    If checkFields Then
        PscCheckFields sectBefore
        If UCase$(Trim$(shape)) = "PSCI" Then sectBefore.Add "USE_SYMMETRIC", useSymmetric
        If nJoint > 0 Then
            If IsMissing(jointFlags) Then
                sectBefore.Add "JOINT", CvBools(Empty, nJoint)
            Else
                sectBefore.Add "JOINT", PscFlags(jointFlags, nJoint)
            End If
        End If
    Else
        If nJoint > 0 Then sectBefore.Add "JOINT", PscFlags(jointFlags, nJoint)
        If Len(pscOpt1) > 0 Then sectBefore.Add "PSC_OPT1", pscOpt1
        If Len(pscOpt2) > 0 Then sectBefore.Add "PSC_OPT2", pscOpt2
        sectBefore.Add "USE_SYMMETRIC", useSymmetric
        sectBefore.Add "USE_SMALL_HOLE", useSmallHole
    End If

    Set item = New Dictionary
    item.Add "SECTTYPE", "PSC"
    item.Add "SECT_NAME", sectName
    item.Add "SECT_BEFORE", sectBefore

    Set assignData = New Dictionary
    assignData.Add CStr(sectId), item

    Set body = New Dictionary
    body.Add "Assign", assignData

    SectionPsc = CvQueue("/db/SECT", body)
End Function

' 1CEL : PSC single cell box.  A 6, B 6, C 10, D 8, JOINT 8
' A = outer shape at I, B = cell at I, C = outer at J, D = cell at J.
Function SectionPsc1Cell(ByVal sectId As Long, ByVal sectName As String, _
                         Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                         Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                         Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                         Optional ByVal offsetPt As String = "CC", _
                         Optional ByVal useShearDeform As Boolean = True, _
                         Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPsc1Cell = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPsc1Cell = SectionPsc(sectId, sectName, "1CEL", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      "", "", _
                      6, 6, 10, 8, 8, _
                      True, False, offsetPt, useShearDeform, useWarping, True)
End Function

' 2CEL : PSC two cell box.  A 6, B 6, C 10, D 8, JOINT 8
' Same groups as 1CEL; the second cell comes from the D group.
Function SectionPsc2Cell(ByVal sectId As Long, ByVal sectName As String, _
                         Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                         Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                         Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                         Optional ByVal offsetPt As String = "CC", _
                         Optional ByVal useShearDeform As Boolean = True, _
                         Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPsc2Cell = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPsc2Cell = SectionPsc(sectId, sectName, "2CEL", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      "", "", _
                      6, 6, 10, 8, 8, _
                      True, False, offsetPt, useShearDeform, useWarping, True)
End Function

' 3CEL : PSC three cell box.  A 10, B 12, C 13, D 13, JOINT 13
Function SectionPsc3Cell(ByVal sectId As Long, ByVal sectName As String, _
                         Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                         Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                         Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                         Optional ByVal offsetPt As String = "CC", _
                         Optional ByVal useShearDeform As Boolean = True, _
                         Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPsc3Cell = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPsc3Cell = SectionPsc(sectId, sectName, "3CEL", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      "", "", _
                      10, 12, 13, 13, 13, _
                      True, False, offsetPt, useShearDeform, useWarping)
End Function

' NCEL : PSC n cell box.  A 5, B 8, no C/D, no JOINT
' pscOpt1 is the corner treatment (CHAMFER),
' pscOpt2 is the cell count as text - "4" means four cells.
Function SectionPscNCell(ByVal sectId As Long, ByVal sectName As String, _
                         Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                         Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                         Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                         Optional ByVal pscOpt1 As String = "CHAMFER", _
                         Optional ByVal pscOpt2 As String = "4", _
                         Optional ByVal offsetPt As String = "CC", _
                         Optional ByVal useShearDeform As Boolean = True, _
                         Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPscNCell = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPscNCell = SectionPsc(sectId, sectName, "NCEL", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      pscOpt1, pscOpt2, _
                      5, 8, 0, 0, 0, _
                      True, False, offsetPt, useShearDeform, useWarping)
End Function

' NCE2 : PSC n cell box, type 2.  A 11, B 18, C 11, D 18, JOINT 2
' pscOpt1 is the cell shape, CIRCLE or POLYGON,
' pscOpt2 is the cell count as text.
Function SectionPscNCell2(ByVal sectId As Long, ByVal sectName As String, _
                          Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                          Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                          Optional ByVal sWidth As Double = 0, _
                          Optional ByVal jointFlags As Variant, _
                          Optional ByVal pscOpt1 As String = "POLYGON", _
                          Optional ByVal pscOpt2 As String = "2", _
                          Optional ByVal offsetPt As String = "CC", _
                          Optional ByVal useShearDeform As Boolean = True, _
                          Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPscNCell2 = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPscNCell2 = SectionPsc(sectId, sectName, "NCE2", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      pscOpt1, pscOpt2, _
                      11, 18, 11, 18, 2, _
                      True, False, offsetPt, useShearDeform, useWarping)
End Function

' PSCM : PSC mid girder.  A 10, B 6, C 9, D 6, JOINT 9
' pscOpt1 and pscOpt2 pick the upper and lower haunch:
' NONE, CIRCLE or POLYGON.  They change the overall width, so the
' A/C groups have to match the choice.
Function SectionPscMid(ByVal sectId As Long, ByVal sectName As String, _
                       Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                       Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                       Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                       Optional ByVal pscOpt1 As String = "NONE", _
                       Optional ByVal pscOpt2 As String = "NONE", _
                       Optional ByVal offsetPt As String = "CC", _
                       Optional ByVal useShearDeform As Boolean = True, _
                       Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPscMid = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPscMid = SectionPsc(sectId, sectName, "PSCM", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      pscOpt1, pscOpt2, _
                      10, 6, 9, 6, 9, _
                      False, False, offsetPt, useShearDeform, useWarping)
End Function

' PSCI : PSC I girder.  A 10, B 7, C 9, D 7, JOINT 9
Function SectionPscI(ByVal sectId As Long, ByVal sectName As String, _
                     Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                     Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                     Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                     Optional ByVal offsetPt As String = "CC", _
                     Optional ByVal useShearDeform As Boolean = True, _
                     Optional ByVal useWarping As Boolean = False, _
                     Optional ByVal bSymm As Boolean = True) As String
    ' bSymm: the right side (C, D, JR) is the left side
    Dim oSym As Object

    If sectId = 0 Then
        SectionPscI = CallGet("/db/SECT")
        Exit Function
    End If

    If bSymm Then
        Set oSym = CvPscISizes(True, aSizes, bSizes, Empty, Empty)
        cSizes = oSym.Item("vSIZE_PSC_C")
        dSizes = oSym.Item("vSIZE_PSC_D")
        jointFlags = CvPscIJoints(True, jointFlags)
    End If

    SectionPscI = SectionPsc(sectId, sectName, "PSCI", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      "", "", _
                      10, 7, 9, 7, 9, _
                      bSymm, False, offsetPt, useShearDeform, useWarping, True)
End Function

' PSCH : PSC half section.  A 6, B 6, C 10, D 7, JOINT 8
' pscOpt1 is the side that is kept, LEFT or RIGHT,
' pscOpt2 is the cell shape, NONE, CIRCLE or POLYGON.
Function SectionPscHalf(ByVal sectId As Long, ByVal sectName As String, _
                        Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                        Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                        Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                        Optional ByVal pscOpt1 As String = "LEFT", _
                        Optional ByVal pscOpt2 As String = "NONE", _
                        Optional ByVal offsetPt As String = "CC", _
                        Optional ByVal useShearDeform As Boolean = True, _
                        Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPscHalf = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPscHalf = SectionPsc(sectId, sectName, "PSCH", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      pscOpt1, pscOpt2, _
                      6, 6, 10, 7, 8, _
                      True, False, offsetPt, useShearDeform, useWarping)
End Function

' PSCT : PSC tee girder.  A 8, B 8, C 7, D 8, JOINT 9
Function SectionPscTee(ByVal sectId As Long, ByVal sectName As String, _
                       Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                       Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                       Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                       Optional ByVal offsetPt As String = "CC", _
                       Optional ByVal useShearDeform As Boolean = True, _
                       Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPscTee = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPscTee = SectionPsc(sectId, sectName, "PSCT", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      "", "", _
                      8, 8, 7, 8, 9, _
                      False, False, offsetPt, useShearDeform, useWarping)
End Function

' PSCB : PSC plate / slab girder.  A 10, B 6, C 8, D 5, JOINT 2
' pscOpt1 is HALF, 1CELL or 2CELL,
' pscOpt2 is LEFT or RIGHT for HALF, CIRCLE or POLYGON otherwise.
Function SectionPscPlate(ByVal sectId As Long, ByVal sectName As String, _
                         Optional ByVal aSizes As Variant, Optional ByVal bSizes As Variant, _
                         Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                         Optional ByVal sWidth As Double = 0, Optional ByVal jointFlags As Variant, _
                         Optional ByVal pscOpt1 As String = "1CELL", _
                         Optional ByVal pscOpt2 As String = "CIRCLE", _
                         Optional ByVal offsetPt As String = "CC", _
                         Optional ByVal useShearDeform As Boolean = True, _
                         Optional ByVal useWarping As Boolean = False) As String

    If sectId = 0 Then
        SectionPscPlate = CallGet("/db/SECT")
        Exit Function
    End If

    SectionPscPlate = SectionPsc(sectId, sectName, "PSCB", _
                      aSizes, bSizes, cSizes, dSizes, sWidth, jointFlags, _
                      pscOpt1, pscOpt2, _
                      10, 6, 8, 5, 2, _
                      False, False, offsetPt, useShearDeform, useWarping)
End Function

' ------------------------------------------
' Node
' ------------------------------------------
' Node(nodeId, [X], [Y], [Z])
' NodeRange(startId, count, X1, Y1, Z1, dX, dY, dZ)
' ------------------------------------------

Function Node(ByVal nodeId As Long, Optional ByVal x As Double = 0, Optional ByVal Y As Double = 0, _
              Optional ByVal Z As Double = 0, _
              Optional ByVal GROUP As String = "") As String

    If nodeId = 0 Then
        Node = CallGet("/db/NODE")
        Exit Function
    End If

    Dim item As Object
    Set item = New Dictionary
    item.Add "X", x
    item.Add "Y", Y
    item.Add "Z", Z

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(nodeId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    Node = CvQueue("/db/NODE", body)
    CvGroupJoin GROUP, nodeId, Null   ' [11]

End Function

Function NodeRange(ByVal startId As Long, ByVal count As Long, ByVal X1 As Double, _
                   ByVal Y1 As Double, ByVal Z1 As Double, ByVal X2 As Double, ByVal Y2 As Double, _
                   ByVal Z2 As Double, Optional ByVal tolerance As Double = 0.0001) As String

    Dim dx As Double, dy As Double, dz As Double
    dx = (X2 - X1) / (count - 1)
    dy = (Y2 - Y1) / (count - 1)
    dz = (Z2 - Z1) / (count - 1)

    ' nodes in NX + nodes still waiting in the store
    Dim existingData As Object
    Set existingData = CvAllNodes()

    Dim assignData As Object
    Set assignData = New Dictionary

    Dim nextId As Integer
    nextId = startId

    Dim i As Integer
    For i = 0 To count - 1
        Dim tx As Double, ty As Double, tz As Double
        tx = X1 + dx * i
        ty = Y1 + dy * i
        tz = Z1 + dz * i

        Dim isDuplicate As Boolean
        isDuplicate = False

        Dim key As Variant
        For Each key In existingData.Keys
            Dim ex As Double, ey As Double, ez As Double
            ex = existingData(key)("X")
            ey = existingData(key)("Y")
            ez = existingData(key)("Z")

            If Abs(ex - tx) <= tolerance And Abs(ey - ty) <= tolerance And Abs(ez - tz) <= tolerance Then
                isDuplicate = True
                Exit For
            End If
        Next key

        If Not isDuplicate Then
            Do While existingData.Exists(CStr(nextId))
                nextId = nextId + 1
            Loop

            Dim item As Object
            Set item = New Dictionary
            item.Add "X", tx
            item.Add "Y", ty
            item.Add "Z", tz

            assignData.Add CStr(nextId), item
            nextId = nextId + 1
        End If
    Next i

    If assignData.count = 0 Then
        NodeRange = "{""message"":""no new nodes""}"
        Exit Function
    End If

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    NodeRange = CvQueue("/db/NODE", body)

End Function

' ------------------------------------------
' ------------------------------------------
' NodeRectangle(startId, plane, d1, d2, [tolerance])
' ------------------------------------------

Function NodeRectangle(ByVal startId As Long, ByVal plane As String, ByVal d1 As Double, _
                       ByVal d2 As Double, Optional ByVal tolerance As Double = 0.0001) As String

    Dim corners(3, 2) As Double

    Select Case UCase(plane)
        Case "XY"
            corners(0, 0) = 0: corners(0, 1) = 0: corners(0, 2) = 0
            corners(1, 0) = d1: corners(1, 1) = 0: corners(1, 2) = 0
            corners(2, 0) = d1: corners(2, 1) = d2: corners(2, 2) = 0
            corners(3, 0) = 0: corners(3, 1) = d2: corners(3, 2) = 0
        Case "XZ"
            corners(0, 0) = 0: corners(0, 1) = 0: corners(0, 2) = 0
            corners(1, 0) = d1: corners(1, 1) = 0: corners(1, 2) = 0
            corners(2, 0) = d1: corners(2, 1) = 0: corners(2, 2) = d2
            corners(3, 0) = 0: corners(3, 1) = 0: corners(3, 2) = d2
        Case "YZ"
            corners(0, 0) = 0: corners(0, 1) = 0: corners(0, 2) = 0
            corners(1, 0) = 0: corners(1, 1) = d1: corners(1, 2) = 0
            corners(2, 0) = 0: corners(2, 1) = d1: corners(2, 2) = d2
            corners(3, 0) = 0: corners(3, 1) = 0: corners(3, 2) = d2
        Case Else
            NodeRectangle = "{""error"":""plane must be XY, XZ, or YZ""}"
            Exit Function
    End Select

    ' nodes in NX + nodes still waiting in the store
    Dim existingData As Object
    Set existingData = CvAllNodes()

    Dim assignData As Object
    Set assignData = New Dictionary

    Dim nextId As Integer
    nextId = startId

    Dim i As Integer
    For i = 0 To 3
        Dim tx As Double, ty As Double, tz As Double
        tx = corners(i, 0): ty = corners(i, 1): tz = corners(i, 2)

        Dim isDuplicate As Boolean
        isDuplicate = False

        Dim key As Variant
        For Each key In existingData.Keys
            Dim ex As Double, ey As Double, ez As Double
            ex = existingData(key)("X")
            ey = existingData(key)("Y")
            ez = existingData(key)("Z")

            If Abs(ex - tx) <= tolerance And Abs(ey - ty) <= tolerance And Abs(ez - tz) <= tolerance Then
                isDuplicate = True
                Exit For
            End If
        Next key

        If Not isDuplicate Then
            Do While existingData.Exists(CStr(nextId))
                nextId = nextId + 1
            Loop

            Dim item As Object
            Set item = New Dictionary
            item.Add "X", tx
            item.Add "Y", ty
            item.Add "Z", tz

            assignData.Add CStr(nextId), item
            nextId = nextId + 1
        End If
    Next i

    If assignData.count = 0 Then
        NodeRectangle = "{""message"":""no new nodes""}"
        Exit Function
    End If

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    NodeRectangle = CvQueue("/db/NODE", body)

End Function

' ------------------------------------------
' Element
' ------------------------------------------
' Beam(elemId, matlId, sectId, ni, nj, [ANGLE])
' BeamRange(startId, count, startNi, matlId, sectId, [ANGLE])
' ------------------------------------------

Function Beam(ByVal elemId As Long, ByVal matlId As Long, ByVal sectId As Long, ByVal ni As Long, _
              ByVal nj As Long, Optional ByVal ANGLE As Double = 0, _
              Optional ByVal GROUP As String = "") As String

    If elemId = 0 Then
        Beam = CallGet("/db/ELEM")
        Exit Function
    End If

    Dim item As Object
    Set item = New Dictionary
    item.Add "TYPE", "BEAM"
    item.Add "MATL", matlId
    item.Add "SECT", sectId
    item.Add "NODE", Array(ni, nj)
    item.Add "ANGLE", ANGLE

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(elemId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    Beam = CvQueue("/db/ELEM", body)
    CvGroupJoin GROUP, Array(ni, nj), elemId   ' [11]

End Function

Function BeamRange(ByVal startId As Long, ByVal ni As Long, ByVal nj As Long, ByVal matlId As Long, _
                   ByVal sectId As Long, Optional ByVal ANGLE As Double = 0, _
                   Optional ByVal tolerance As Double = 0.0001) As String

    ' nodes in NX + nodes still waiting in the store
    Dim nodeData As Object
    Set nodeData = CvAllNodes()

    Dim niX As Double, niY As Double, niZ As Double
    Dim njX As Double, njY As Double, njZ As Double
    niX = nodeData(CStr(ni))("X"): niY = nodeData(CStr(ni))("Y"): niZ = nodeData(CStr(ni))("Z")
    njX = nodeData(CStr(nj))("X"): njY = nodeData(CStr(nj))("Y"): njZ = nodeData(CStr(nj))("Z")

    Dim dx As Double, dy As Double, dz As Double, lenSq As Double
    dx = njX - niX: dy = njY - niY: dz = njZ - niZ
    lenSq = dx * dx + dy * dy + dz * dz

    Dim onLineIds() As String
    Dim onLineT() As Double
    Dim n As Integer
    n = 0
    ReDim onLineIds(nodeData.count - 1)
    ReDim onLineT(nodeData.count - 1)

    Dim key As Variant
    For Each key In nodeData.Keys
        Dim px As Double, py As Double, pz As Double
        px = nodeData(key)("X"): py = nodeData(key)("Y"): pz = nodeData(key)("Z")

        Dim t As Double
        t = ((px - niX) * dx + (py - niY) * dy + (pz - niZ) * dz) / lenSq

        If t >= -tolerance And t <= 1 + tolerance Then
            Dim projX As Double, projY As Double, projZ As Double
            projX = niX + t * dx: projY = niY + t * dy: projZ = niZ + t * dz
            Dim distSq As Double
            distSq = (px - projX) ^ 2 + (py - projY) ^ 2 + (pz - projZ) ^ 2

            If distSq <= tolerance ^ 2 Then
                onLineIds(n) = CStr(key)
                onLineT(n) = t
                n = n + 1
            End If
        End If
    Next key

    Dim i As Integer, j As Integer
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If onLineT(j) > onLineT(j + 1) Then
                Dim tmpT As Double, tmpId As String
                tmpT = onLineT(j): onLineT(j) = onLineT(j + 1): onLineT(j + 1) = tmpT
                tmpId = onLineIds(j): onLineIds(j) = onLineIds(j + 1): onLineIds(j + 1) = tmpId
            End If
        Next j
    Next i

    Dim assignData As Object
    Set assignData = New Dictionary

    For i = 0 To n - 2
        Dim item As Object
        Set item = New Dictionary
        item.Add "TYPE", "BEAM"
        item.Add "MATL", matlId
        item.Add "SECT", sectId
        item.Add "NODE", Array(CInt(onLineIds(i)), CInt(onLineIds(i + 1)))
        item.Add "ANGLE", ANGLE

        assignData.Add CStr(startId + i), item
    Next i

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    BeamRange = CvQueue("/db/ELEM", body)

End Function

' ------------------------------------------
' ------------------------------------------
' BeamDivide(startId, ni, nj, matlId, sectId, divMethod, divValue, [ANGLE])
' ------------------------------------------

Function BeamDivide(ByVal startId As Long, ByVal ni As Long, ByVal nj As Long, _
                    ByVal matlId As Long, ByVal sectId As Long, ByVal divMethod As String, _
                    ByVal divValue As Variant, Optional ByVal ANGLE As Double = 0) As String

    Dim createResult As String
    createResult = Beam(startId, matlId, sectId, ni, nj, ANGLE)

    Dim divOption As Object
    Set divOption = New Dictionary

    Dim divideObj As Object
    Set divideObj = New Dictionary
    divideObj.Add "ELEM_TYPE", "Frame"

    If LCase(divMethod) = "equal" Then
        Dim equalOpt As Object
        Set equalOpt = New Dictionary
        equalOpt.Add "NUM_X", divValue
        divOption.Add "EQUAL_OPTION", equalOpt
        divideObj.Add "DIV_METHOD", "Equal"
    Else
        Dim unequalOpt As Object
        Set unequalOpt = New Dictionary
        unequalOpt.Add "DIST_X", CStr(divValue)
        divOption.Add "UNEQUAL_OPTION", unequalOpt
        divideObj.Add "DIV_METHOD", "Unequal"
    End If

    divideObj.Add "OPTION", divOption

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "TARGETS", Array(startId)
    arg.Add "DIVIDE", divideObj

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    BeamDivide = CvQueue("/ope/DIVIDEELEM", body)

End Function

' ------------------------------------------
' ------------------------------------------
' GetPlateElement(elemId)
' ------------------------------------------

Function GetPlateElement(ByVal elemId As Long) As String

    GetPlateElement = CallGet("/db/ELEM/" & elemId)

End Function

' ------------------------------------------
' ------------------------------------------
' Plate(elemId, matlId, thikId, nodeIds, [ANGLE], [STYPE])
' ------------------------------------------

' STYPE: Thick=1, Thin=2, Thick+Drilling DOF=3, Thin+Drilling DOF=4
Function Plate(ByVal elemId As Long, ByVal matlId As Long, ByVal thikId As Long, _
               ByVal nodeIds As Variant, Optional ByVal ANGLE As Double = 0, _
               Optional ByVal STYPE As Long = 2, _
               Optional ByVal GROUP As String = "") As String

    If elemId = 0 Then
        Plate = CallGet("/db/ELEM")
        Exit Function
    End If

    Dim item As Object
    Set item = New Dictionary
    item.Add "TYPE", "PLATE"
    item.Add "MATL", matlId
    item.Add "SECT", thikId
    item.Add "NODE", nodeIds
    item.Add "ANGLE", ANGLE
    item.Add "STYPE", STYPE

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(elemId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    Plate = CvQueue("/db/ELEM", body)
    CvGroupJoin GROUP, nodeIds, elemId   ' [11]

End Function

' ------------------------------------------
' Support
' ------------------------------------------
' Support(nodeIds, [CONSTRAINT], [GROUP_NAME])
' ------------------------------------------

Function Support(ByVal nodeIds As Variant, Optional ByVal CONSTRAINT As String = "", _
                 Optional ByVal GROUP_NAME As String = "") As String

    If Not IsArray(nodeIds) Then
        If nodeIds = 0 Then
            Support = CallGet("/db/CONS")
            Exit Function
        End If
    End If

    Dim idArr As Variant
    If IsArray(nodeIds) Then
        idArr = nodeIds
    Else
        idArr = Array(nodeIds)
    End If
    CONSTRAINT = CvConstraint(CONSTRAINT)     ' "fix" / "pin" / "roller" / "111" -> 7 digits ([24])

    Dim assignData As Object
    Set assignData = New Dictionary

    Dim i As Integer
    For i = LBound(idArr) To UBound(idArr)
        Dim conItem As Object
        Set conItem = New Dictionary
        conItem.Add "ID", idArr(i)
        conItem.Add "CONSTRAINT", CONSTRAINT
        conItem.Add "GROUP_NAME", GROUP_NAME

        Dim items(0) As Object
        Set items(0) = conItem

        Dim item As Object
        Set item = New Dictionary
        item.Add "ITEMS", items

        assignData.Add CStr(idArr(i)), item
    Next i

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    Support = CvQueue("/db/CONS", body)

End Function

' ------------------------------------------
' Surface Spring
' ------------------------------------------
' SurfaceSpringLinear(elems, WIDTH, Kx, Ky, Kz, [GROUP_NAME])
' SurfaceSpringComp(elems, WIDTH, SUBGRADE, DIR, [GROUP_NAME])
' ------------------------------------------

Function SurfaceSpringLinear(ByVal elems As Variant, ByVal width As Double, ByVal Kx As Double, _
                             ByVal Ky As Double, ByVal Kz As Double, _
                             Optional ByVal GROUP_NAME As String = "") As String

    Dim stiff(2) As Double
    stiff(0) = Kx: stiff(1) = Ky: stiff(2) = Kz

    Dim boundary As Object
    Set boundary = New Dictionary
    boundary.Add "TYPE", "LINEAR"
    boundary.Add "STIFF", stiff
    boundary.Add "bDAMP", False

    SurfaceSpringLinear = SurfaceSpringCall(elems, width, GROUP_NAME, boundary)

End Function

Function SurfaceSpringComp(ByVal elems As Variant, ByVal width As Double, ByVal SUBGRADE As Double, _
                           ByVal DIR As Long, Optional ByVal GROUP_NAME As String = "") As String

    Dim boundary As Object
    Set boundary = New Dictionary
    boundary.Add "TYPE", "COMP"
    boundary.Add "DIR", DIR
    boundary.Add "SUBGRADE", SUBGRADE

    SurfaceSpringComp = SurfaceSpringCall(elems, width, GROUP_NAME, boundary)

End Function

Private Function SurfaceSpringCall(elems As Variant, width As Double, _
                                   GROUP_NAME As String, boundary As Object) As String

    Dim nodeElems As Object
    Set nodeElems = New Dictionary
    nodeElems.Add "KEYS", elems

    Dim elementObj As Object
    Set elementObj = New Dictionary
    elementObj.Add "TYPE", "FRAME"
    elementObj.Add "WIDTH", width

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "CONVERT_TO", "POINT_SPRING"
    arg.Add "NODE_ELEMS", nodeElems
    arg.Add "ELEMENT", elementObj
    arg.Add "BOUNDARY", boundary
    If GROUP_NAME <> "" Then arg.Add "GROUP_NAME", GROUP_NAME

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    SurfaceSpringCall = CvQueue("/ope/SSPS", body)

End Function

' ------------------------------------------
' Beam End Offset
' ------------------------------------------
' BeamOffsetAsymmetric(elemId, RGDYi, RGDZi, RGDYj, RGDZj, [GROUP_NAME])
' BeamOffsetSymmetric(elemId, [RGDi], [RGDj], [GROUP_NAME])
' BeamOffsetGlobal(elemId, [RGDXi], [RGDYi], [RGDZi], [RGDXj], [RGDYj], [RGDZj], [GROUP_NAME])
' ------------------------------------------

Function BeamOffsetAsymmetric(ByVal elemId As Variant, ByVal RGDYi As Double, _
                              ByVal RGDZi As Double, ByVal RGDYj As Double, ByVal RGDZj As Double, _
                              Optional ByVal GROUP_NAME As String = "") As String

    If elemId = 0 Then
        BeamOffsetAsymmetric = CallGet("/db/OFFS")
        Exit Function
    End If

    Dim conItem As Object
    Set conItem = New Dictionary
    conItem.Add "ID", CInt(elemId)
    conItem.Add "TYPE", "ELEMENT"
    conItem.Add "RGDYi", RGDYi
    conItem.Add "RGDZi", RGDZi
    conItem.Add "RGDYj", RGDYj
    conItem.Add "RGDZj", RGDZj
    If GROUP_NAME <> "" Then conItem.Add "GROUP_NAME", GROUP_NAME

    Dim items(0) As Object
    Set items(0) = conItem

    Dim item As Object
    Set item = New Dictionary
    item.Add "ITEMS", items

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(elemId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    BeamOffsetAsymmetric = CvQueue("/db/OFFS", body)

End Function

Function BeamOffsetSymmetric(ByVal elemId As Variant, Optional ByVal RGDi As Double = 0, _
                             Optional ByVal RGDj As Double = 0, _
                             Optional ByVal GROUP_NAME As String = "") As String

    BeamOffsetSymmetric = BeamOffsetAsymmetric(elemId, RGDYi:=RGDi, RGDZi:=RGDi, RGDYj:=RGDj, RGDZj:=RGDj, GROUP_NAME:=GROUP_NAME)

End Function

Function BeamOffsetGlobal(ByVal elemId As Variant, Optional ByVal RGDXi As Double = 0, _
                          Optional ByVal RGDYi As Double = 0, Optional ByVal RGDZi As Double = 0, _
                          Optional ByVal RGDXj As Double = 0, Optional ByVal RGDYj As Double = 0, _
                          Optional ByVal RGDZj As Double = 0, _
                          Optional ByVal GROUP_NAME As String = "") As String

    If elemId = 0 Then
        BeamOffsetGlobal = CallGet("/db/OFFS")
        Exit Function
    End If

    Dim conItem As Object
    Set conItem = New Dictionary
    conItem.Add "ID", CInt(elemId)
    conItem.Add "TYPE", "GLOBAL"
    conItem.Add "RGDXi", RGDXi
    conItem.Add "RGDYi", RGDYi
    conItem.Add "RGDZi", RGDZi
    conItem.Add "RGDXj", RGDXj
    conItem.Add "RGDYj", RGDYj
    conItem.Add "RGDZj", RGDZj
    If GROUP_NAME <> "" Then conItem.Add "GROUP_NAME", GROUP_NAME

    Dim items(0) As Object
    Set items(0) = conItem

    Dim item As Object
    Set item = New Dictionary
    item.Add "ITEMS", items

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(elemId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    BeamOffsetGlobal = CvQueue("/db/OFFS", body)

End Function

' ------------------------------------------
' Boundary Group
' ------------------------------------------
' BoundaryGroup(bngrId, [NAME])
' ------------------------------------------

Function BoundaryGroup(ByVal bngrId As Long, Optional ByVal NAME As String = "") As String

    If bngrId = 0 Then
        BoundaryGroup = CallGet("/db/BNGR")
        Exit Function
    End If

    Dim item As Object
    Set item = New Dictionary
    item.Add "NAME", NAME
    item.Add "AUTOTYPE", 0

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(bngrId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    BoundaryGroup = CvQueue("/db/BNGR", body)

End Function


' ------------------------------------------
' Static Load Case
' ------------------------------------------
' LoadCase(lcIds, NAMEs, TYPEs, [DESCs])
' ------------------------------------------

Function LoadCase(ByVal lcIds As Variant, Optional ByVal NAMEs As Variant = Null, _
                  Optional ByVal TYPEs As Variant = Null, Optional ByVal DESCs As Variant = Null) As String
    ' One case or several: LoadCase 1, "DL", "D"  /  LoadCase Array(1, 2), Array("DL", "LL"), Array("D", "L")
    ' Record: NO (= id) / NAME / TYPE / DESC
    Dim vIds As Variant
    Dim item As Object
    Dim assignData As Object
    Dim body As Object
    Dim i As Long

    If Not IsArray(lcIds) Then
        If lcIds = 0 Then
            LoadCase = CallGet("/db/STLD")
            Exit Function
        End If
    End If

    vIds = CvIds(lcIds)
    Set assignData = New Dictionary
    For i = LBound(vIds) To UBound(vIds)
        Set item = New Dictionary
        item.Add "NO", CLng(vIds(i))
        item.Add "NAME", CStr(CvPick(NAMEs, i - LBound(vIds), ""))
        item.Add "TYPE", CStr(CvPick(TYPEs, i - LBound(vIds), ""))
        item.Add "DESC", CStr(CvPick(DESCs, i - LBound(vIds), ""))
        assignData.Add CStr(vIds(i)), item
    Next i

    Set body = New Dictionary
    body.Add "Assign", assignData

    LoadCase = CvQueue("/db/STLD", body)

End Function

' n-th value of an argument that may be a single value or an array (0 based).
' A single value counts for every n. Null / missing gives pDefault.
Private Function CvPick(ByVal pAny As Variant, ByVal n As Long, ByVal pDefault As Variant) As Variant
    If IsMissing(pAny) Then
        CvPick = pDefault
    ElseIf IsNull(pAny) Then
        CvPick = pDefault
    ElseIf IsArray(pAny) Then
        If n + LBound(pAny) <= UBound(pAny) Then CvPick = pAny(n + LBound(pAny)) Else CvPick = pDefault
    Else
        CvPick = pAny
    End If
End Function

' ------------------------------------------
' Self-Weight
' ------------------------------------------
' SelfWeight(bodfId, LCNAME, [Fx], [Fy], [Fz], [GROUP_NAME])
' ------------------------------------------

Function SelfWeight(ByVal bodfId As Long, Optional ByVal LCNAME As String = "", _
                    Optional ByVal Fx As Double = 0, Optional ByVal Fy As Double = 0, _
                    Optional ByVal Fz As Double = -1, Optional ByVal GROUP_NAME As String = "") As String

    If bodfId = 0 Then
        SelfWeight = CallGet("/db/BODF")
        Exit Function
    End If

    Dim item As Object
    Set item = New Dictionary
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "FV", Array(Fx, Fy, Fz)

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(bodfId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    SelfWeight = CvQueue("/db/BODF", body)

End Function

' ------------------------------------------
' Beam Load
' ------------------------------------------
' BeamLoad(elemId, LCNAME, D, P, [TYPE], [DIRECTION], [GROUP_NAME], [USE_ECCEN], [ECCEN_TYPE], [ECCEN_DIR], [I_END], [J_END], [USE_J_END])
' ------------------------------------------

' db/BMLD - Element Beam Loads (CMD: BEAM)
Function BeamLoad(ByVal elemId As Variant, ByVal LCNAME As String, ByVal d As Variant, _
                  ByVal P As Variant, Optional ByVal loadType As String = "UNILOAD", _
                  Optional ByVal DIRECTION As String = "GZ", _
                  Optional ByVal GROUP_NAME As String = "", _
                  Optional ByVal USE_ECCEN As Boolean = False, _
                  Optional ByVal ECCEN_TYPE As Long = 0, Optional ByVal ECCEN_DIR As String = "", _
                  Optional ByVal I_END As Double = 0, Optional ByVal J_END As Double = 0, _
                  Optional ByVal USE_J_END As Boolean = False) As String

    Dim vOne As Variant
    If IsArray(elemId) Then
        For Each vOne In elemId
            BeamLoad CLng(vOne), LCNAME, d, P, loadType, DIRECTION, GROUP_NAME, USE_ECCEN, _
                     ECCEN_TYPE, ECCEN_DIR, I_END, J_END, USE_J_END
        Next vOne
        BeamLoad = ""
        Exit Function
    End If
    If elemId = 0 Then
        BeamLoad = CallGet("/db/BMLD")
        Exit Function
    End If

    Dim d4(3) As Double, p4(3) As Double
    Dim i As Integer
    For i = 0 To 3
        If i <= UBound(d) Then d4(i) = d(i) Else d4(i) = 0
        If i <= UBound(P) Then p4(i) = P(i) Else p4(i) = 0
    Next i

    Dim conItem As Object
    Set conItem = New Dictionary
    conItem.Add "ID", CLng(elemId)
    conItem.Add "LCNAME", LCNAME
    conItem.Add "GROUP_NAME", GROUP_NAME
    conItem.Add "CMD", "BEAM"
    conItem.Add "TYPE", UCase(loadType)
    conItem.Add "DIRECTION", UCase(DIRECTION)
    conItem.Add "USE_PROJECTION", False
    conItem.Add "USE_ECCEN", USE_ECCEN
    If USE_ECCEN Then
        conItem.Add "ECCEN_TYPE", ECCEN_TYPE
        conItem.Add "ECCEN_DIR", ECCEN_DIR
        conItem.Add "I_END", I_END
        conItem.Add "USE_J_END", USE_J_END
        conItem.Add "J_END", J_END
    End If
    conItem.Add "D", d4
    conItem.Add "P", p4

    Dim items(0) As Object
    Set items(0) = conItem

    Dim item As Object
    Set item = New Dictionary
    item.Add "ITEMS", items

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(elemId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    BeamLoad = CvQueue("/db/BMLD", body)

End Function

' ------------------------------------------
' Line Beam Load
' ------------------------------------------
' LineBeamLoad(nodes, LCNAME, LoadType, [D], [P], [DIR], [DistanceType], [GROUP_NAME], [elems],
'              [USE_PROJECTION], [A], [B], [C],
'              [USE_ECCEN], [ECCEN_TYPE], [ECCEN_DIR], [ECCEN_I_END], [ECCEN_J_END], [USE_ECCEN_J_END],
'              [USE_ADD_H], [ADDH_I_END], [ADDH_J_END], [USE_ADDH_J_END],
'              [USE_COPY], [COPY_AXIS], [COPY_DIST])
' ------------------------------------------

' LoadType: CONLOAD, CONMOMENT, UNILOAD, UNIMOMENT, TRALOAD, TRAMOMENT, UNIPRESSURE, TRAPRESSURE, CURVED
Function LineBeamLoad(ByVal nodes As Variant, ByVal LCNAME As String, ByVal loadType As String, _
                      Optional ByVal d As Variant = Null, Optional ByVal P As Variant = Null, _
                      Optional ByVal DIR As String = "GZ", Optional ByVal DistanceType As Long = 0, _
                      Optional ByVal GROUP_NAME As String = "", _
                      Optional ByVal elems As Variant = Null, _
                      Optional ByVal USE_PROJECTION As Variant = Null, _
                      Optional ByVal A As Double = 0, Optional ByVal B As Double = 0, _
                      Optional ByVal c As Double = 0, Optional ByVal USE_ECCEN As Boolean = False, _
                      Optional ByVal ECCEN_TYPE As Long = 0, _
                      Optional ByVal ECCEN_DIR As String = "", _
                      Optional ByVal ECCEN_I_END As Double = 0, _
                      Optional ByVal ECCEN_J_END As Double = 0, _
                      Optional ByVal USE_ECCEN_J_END As Boolean = False, _
                      Optional ByVal USE_ADD_H As Boolean = False, _
                      Optional ByVal ADDH_I_END As Double = 0, _
                      Optional ByVal ADDH_J_END As Double = 0, _
                      Optional ByVal USE_ADDH_J_END As Boolean = False, _
                      Optional ByVal USE_COPY As Boolean = False, _
                      Optional ByVal COPY_AXIS As String = "", _
                      Optional ByVal COPY_DIST As String = "") As String

    Dim methodVal As Integer
    If IsNull(elems) Then
        methodVal = 0
    Else
        methodVal = 1
    End If

    Dim Target As Object
    Set Target = New Dictionary
    Target.Add "METHOD", methodVal
    Target.Add "NODE", nodes
    If Not IsNull(elems) Then Target.Add "ELEM", elems

    Dim load As Object
    Set load = New Dictionary
    load.Add "DIR", UCase(DIR)

    Dim useProj As Boolean
    If IsNull(USE_PROJECTION) Then
        useProj = (methodVal = 1)
    Else
        useProj = USE_PROJECTION
    End If
    load.Add "USE_PROJECTION", useProj
    load.Add "TYPE", DistanceType

    If UCase(loadType) = "CURVED" Then
        load.Add "A", A
        load.Add "B", B
        load.Add "C", c
    Else
        load.Add "D", d
        load.Add "P", P
    End If

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "LCNAME", LCNAME
    If GROUP_NAME <> "" Then arg.Add "GROUP_NAME", GROUP_NAME
    arg.Add "TYPE", UCase(loadType)
    arg.Add "TARGET", Target

    If USE_ECCEN Then
        Dim eccen As Object
        Set eccen = New Dictionary
        eccen.Add "USE", True
        eccen.Add "TYPE", ECCEN_TYPE
        eccen.Add "DIR", ECCEN_DIR
        eccen.Add "I_END", ECCEN_I_END
        eccen.Add "USE_J_END", USE_ECCEN_J_END
        eccen.Add "J_END", ECCEN_J_END
        arg.Add "ECCEN", eccen
    End If

    If USE_ADD_H Then
        Dim addH As Object
        Set addH = New Dictionary
        addH.Add "USE", True
        addH.Add "I_END", ADDH_I_END
        addH.Add "USE_J_END", USE_ADDH_J_END
        addH.Add "J_END", ADDH_J_END
        arg.Add "ADD_H", addH
    End If

    arg.Add "LOAD", load

    If USE_COPY Then
        Dim copyObj As Object
        Set copyObj = New Dictionary
        copyObj.Add "USE", True
        copyObj.Add "AXIS", COPY_AXIS
        copyObj.Add "DIST", COPY_DIST
        arg.Add "COPY", copyObj
    End If

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    LineBeamLoad = CvQueue("/ope/LINEBMLD", body)

End Function

' ------------------------------------------
' Load Combination
' ------------------------------------------
' LoadCombination(lcId, NAME, lcNames, factors, [analTypes], [iType], [ACTIVE], [DESC], [endpoint])
' ------------------------------------------

Function LoadCombination(ByVal lcId As Long, Optional ByVal NAME As String = "", _
                         Optional ByVal lcNames As Variant = Null, _
                         Optional ByVal factors As Variant = Null, _
                         Optional ByVal analTypes As Variant = Null, _
                         Optional ByVal iType As Long = 0, _
                         Optional ByVal ACTIVE As String = "ACTIVE", _
                         Optional ByVal DESC As String = "", _
                         Optional ByVal endpoint As String = "/db/LCOM-GEN") As String

    If lcId = 0 Then
        LoadCombination = CallGet(endpoint)
        Exit Function
    End If

    Dim vComb() As Object
    ReDim vComb(UBound(lcNames))

    Dim i As Integer
    For i = LBound(lcNames) To UBound(lcNames)
        Dim combItem As Object
        Set combItem = New Dictionary

        If IsNull(analTypes) Then
            combItem.Add "ANAL", "ST"
        Else
            combItem.Add "ANAL", analTypes(i)
        End If
        combItem.Add "LCNAME", lcNames(i)
        combItem.Add "FACTOR", factors(i)

        Set vComb(i) = combItem
    Next i

    Dim item As Object
    Set item = New Dictionary
    item.Add "NAME", NAME
    item.Add "ACTIVE", ACTIVE
    item.Add "iTYPE", iType
    item.Add "DESC", DESC
    item.Add "vCOMB", vComb

    Dim assignData As Object
    Set assignData = New Dictionary
    assignData.Add CStr(lcId), item

    Dim body As Object
    Set body = New Dictionary
    body.Add "Assign", assignData

    LoadCombination = CvQueue(endpoint, body)

End Function

' ------------------------------------------
' Using Load Combination
' ------------------------------------------
' UseLoadCombination(lcomType, lcomNames, POSITION, [PREFIX], [excludeLoads])
' UseLoadCombinationSelected(lcomType, lcomNames, POSITION, [PREFIX], [includeLoads])
' ------------------------------------------

Private Function LoadKeys() As Variant
    LoadKeys = Array("SELF_WEIGHT", "NODAL_BODY_FROCE", "NODAL_LOAD", "SPECIFIED_DISPLACEMENT", _
                      "BEAM_LOAD", "FLOOR_LOAD", "FINISHING_MATERIAL_LOAD", "PRESSURE_LOAD", _
                      "PLANE_LOAD", "SYSTEM_TEMPERATURE", "NODAL_TEMPERATURE", "ELEMENT_TEMPERATURE", _
                      "TEMPERATURE_GRADIENT", "BEAM_SECTION_TEMPERATURE", "PRESTRESS_LOAD", _
                      "PRETENSION_LOAD", "TENDON_PRESTRESS_LOAD")
End Function

Private Function BuildLoads(defaultVal As Boolean, Optional overrideNames As Variant = Null, Optional flipVal As Boolean = False) As Object

    Dim loads As Object
    Set loads = New Dictionary

    Dim key As Variant
    For Each key In LoadKeys()
        loads.Add key, defaultVal
    Next key

    If Not IsNull(overrideNames) Then
        Dim n As Variant
        For Each n In overrideNames
            loads(n) = flipVal
        Next n
    End If

    Set BuildLoads = loads

End Function

Private Function CallUSLC(lcomType As String, lcomNames As Variant, position As String, PREFIX As String, loads As Object) As String

    Dim lcomList() As Object
    ReDim lcomList(UBound(lcomNames))

    Dim i As Integer
    For i = LBound(lcomNames) To UBound(lcomNames)
        Dim lcomItem As Object
        Set lcomItem = New Dictionary
        lcomItem.Add "TYPE", lcomType
        lcomItem.Add "NAME", lcomNames(i)
        Set lcomList(i) = lcomItem
    Next i

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "PREFIX", PREFIX
    arg.Add "POSITION", position
    arg.Add "LCOM_LIST", lcomList
    arg.Add "LOADS", loads

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    CallUSLC = CvQueue("/ope/USLC", body)

End Function

Function UseLoadCombination(ByVal lcomType As String, ByVal lcomNames As Variant, _
                            ByVal position As String, Optional ByVal PREFIX As String = "N", _
                            Optional ByVal excludeLoads As Variant = Null) As String

    Dim loads As Object
    Set loads = BuildLoads(defaultVal:=True, overrideNames:=excludeLoads, flipVal:=False)

    UseLoadCombination = CallUSLC(lcomType, lcomNames, position, PREFIX, loads)

End Function

Function UseLoadCombinationSelected(ByVal lcomType As String, ByVal lcomNames As Variant, _
                                    ByVal position As String, _
                                    Optional ByVal PREFIX As String = "N", _
                                    Optional ByVal includeLoads As Variant = Null) As String

    Dim loads As Object
    Set loads = BuildLoads(defaultVal:=False, overrideNames:=includeLoads, flipVal:=True)

    UseLoadCombinationSelected = CallUSLC(lcomType, lcomNames, position, PREFIX, loads)

End Function


' ------------------------------------------
' Perform Analysis
' ------------------------------------------
' RunAnalysis([analysisType])
' ------------------------------------------

Function RunAnalysis(Optional ByVal analysisType As String = "") As String

    Dim body As Object
    Set body = New Dictionary

    If analysisType <> "" Then
        Dim arg As Object
        Set arg = New Dictionary
        arg.Add "TYPE", analysisType
        body.Add "Argument", arg
    End If

    ' send what is still in the store first ([10])
    Dim sPending As String
    sPending = CvFlushPending()
    If Len(sPending) > 0 Then
        RunAnalysis = "{""error"":""" & Replace(sPending, """", "'") & """}"
        Exit Function
    End If
    RunAnalysis = CallPost("/doc/ANAL", JsonConverter.ConvertToJson(body))

End Function

' ------------------------------------------
' Reaction Table
' ------------------------------------------
' ReactionTable(reactionType, nodeIds, loadCaseNames, startCell, [tableName])
' ------------------------------------------

' reactionType: "Global", "Local", "SurfaceSpring"
Function ReactionTable(ByVal reactionType As String, ByVal nodeIds As Variant, _
                       ByVal loadCaseNames As Variant, ByVal startCell As Range, _
                       Optional ByVal tableName As String = "Reaction") As String

    Dim tableType As String
    Select Case UCase(reactionType)
        Case "GLOBAL": tableType = "REACTIONG"
        Case "LOCAL": tableType = "REACTIONL"
        Case "SURFACESPRING": tableType = "REACTIONSURFACESPRING"
    End Select

    Dim nodeElems As Object
    Set nodeElems = New Dictionary
    nodeElems.Add "KEYS", nodeIds

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "TABLE_NAME", tableName
    arg.Add "TABLE_TYPE", tableType
    arg.Add "NODE_ELEMS", nodeElems
    arg.Add "LOAD_CASE_NAMES", loadCaseNames

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    Dim resultJson As String
    resultJson = CallPost("/post/TABLE", JsonConverter.ConvertToJson(body))

    Dim resultData As Object
    Set resultData = JsonConverter.ParseJson(resultJson)(tableName)

    Dim col As Integer
    col = 0
    Dim h As Variant
    For Each h In resultData("HEAD")
        startCell.Offset(0, col).Value = h
        col = col + 1
    Next h

    Dim r As Integer
    r = 1
    Dim rowData As Variant, cellVal As Variant
    For Each rowData In resultData("DATA")
        col = 0
        For Each cellVal In rowData
            startCell.Offset(r, col).Value = cellVal
            col = col + 1
        Next cellVal
        r = r + 1
    Next rowData

    ReactionTable = resultJson

End Function

' ------------------------------------------
' Displacement Table
' ------------------------------------------
' DisplacementTable(nodeIds, loadCaseNames, startCell, [tableName])
' ------------------------------------------

Function DisplacementTable(ByVal nodeIds As Variant, ByVal loadCaseNames As Variant, _
                           ByVal startCell As Range, _
                           Optional ByVal tableName As String = "Displacements") As String

    Dim nodeElems As Object
    Set nodeElems = New Dictionary
    nodeElems.Add "KEYS", nodeIds

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "TABLE_NAME", tableName
    arg.Add "TABLE_TYPE", "DISPLACEMENTG"
    arg.Add "NODE_ELEMS", nodeElems
    arg.Add "LOAD_CASE_NAMES", loadCaseNames

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    Dim resultJson As String
    resultJson = CallPost("/post/TABLE", JsonConverter.ConvertToJson(body))

    Dim resultData As Object
    Set resultData = JsonConverter.ParseJson(resultJson)(tableName)

    Dim col As Integer
    col = 0
    Dim h As Variant
    For Each h In resultData("HEAD")
        startCell.Offset(0, col).Value = h
        col = col + 1
    Next h

    Dim r As Integer
    r = 1
    Dim rowData As Variant, cellVal As Variant
    For Each rowData In resultData("DATA")
        col = 0
        For Each cellVal In rowData
            startCell.Offset(r, col).Value = cellVal
            col = col + 1
        Next cellVal
        r = r + 1
    Next rowData

    DisplacementTable = resultJson

End Function

' ------------------------------------------
' Beam Force Table
' ------------------------------------------
' BeamForceTable(elemIds, loadCaseNames, startCell, [tableName])
' ------------------------------------------

Function BeamForceTable(ByVal elemIds As Variant, ByVal loadCaseNames As Variant, _
                        ByVal startCell As Range, Optional ByVal tableName As String = "BeamForce") As String

    Dim nodeElems As Object
    Set nodeElems = New Dictionary
    nodeElems.Add "KEYS", elemIds

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "TABLE_NAME", tableName
    arg.Add "TABLE_TYPE", "BEAMFORCE"
    arg.Add "COMPONENTS", Array("Elem", "Load", "Part", "Axial", "Shear-y", "Shear-z", "Torsion", "Moment-y", "Moment-z", "Bi-Moment", "T-Moment", "W-Moment")
    arg.Add "NODE_ELEMS", nodeElems
    arg.Add "LOAD_CASE_NAMES", loadCaseNames
    arg.Add "PARTS", Array("PartI", "PartJ")

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    Dim resultJson As String
    resultJson = CallPost("/post/TABLE", JsonConverter.ConvertToJson(body))

    Dim resultData As Object
    Set resultData = JsonConverter.ParseJson(resultJson)(tableName)

    Dim col As Integer
    col = 0
    Dim h As Variant
    For Each h In resultData("HEAD")
        startCell.Offset(0, col).Value = h
        col = col + 1
    Next h

    Dim r As Integer
    r = 1
    Dim rowData As Variant, cellVal As Variant
    For Each rowData In resultData("DATA")
        col = 0
        For Each cellVal In rowData
            startCell.Offset(r, col).Value = cellVal
            col = col + 1
        Next cellVal
        r = r + 1
    Next rowData

    BeamForceTable = resultJson

End Function

' ------------------------------------------
' Plate Force Table (Unit Length : Local)
' ------------------------------------------
' PlateForceLocalTable(loadCaseNames, startCell, [component], [elemIds], [tableName])
' ------------------------------------------

Function PlateForceLocalTable(ByVal loadCaseNames As Variant, ByVal startCell As Range, _
                              Optional ByVal component As String = "Mxx", _
                              Optional ByVal elemIds As Variant = Null, _
                              Optional ByVal tableName As String = "PlateForce") As String

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "TABLE_NAME", tableName
    arg.Add "TABLE_TYPE", "PLATEFORCEUL"
    arg.Add "COMPONENTS", Array("Elem", "Load", "Node", component)

    If Not IsNull(elemIds) Then
        Dim nodeElems As Object
        Set nodeElems = New Dictionary
        nodeElems.Add "KEYS", elemIds
        arg.Add "NODE_ELEMS", nodeElems
    End If

    arg.Add "LOAD_CASE_NAMES", loadCaseNames
    arg.Add "AVERAGE_NODAL_RESULT", True

    Dim nodeFlag As Object
    Set nodeFlag = New Dictionary
    nodeFlag.Add "CENTER", False
    nodeFlag.Add "NODES", True
    arg.Add "NODE_FLAG", nodeFlag

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    Dim resultJson As String
    resultJson = CallPost("/post/TABLE", JsonConverter.ConvertToJson(body))

    Dim resultData As Object
    Set resultData = JsonConverter.ParseJson(resultJson)(tableName)

    Dim col As Integer
    col = 0
    Dim h As Variant
    For Each h In resultData("HEAD")
        startCell.Offset(0, col).Value = h
        col = col + 1
    Next h

    Dim r As Integer
    r = 1
    Dim rowData As Variant, cellVal As Variant
    For Each rowData In resultData("DATA")
        col = 0
        For Each cellVal In rowData
            startCell.Offset(r, col).Value = cellVal
            col = col + 1
        Next cellVal
        r = r + 1
    Next rowData

    PlateForceLocalTable = resultJson

End Function

' ------------------------------------------
' ------------------------------------------
' GetTableExtreme(resultJson, tableName, columnName, [findMax], [idColumn])
' ------------------------------------------

Function GetTableExtreme(ByVal resultJson As String, ByVal tableName As String, _
                         ByVal columnName As String, Optional ByVal findMax As Boolean = True, _
                         Optional ByVal idColumn As String = "Elem") As Object

    Dim resultData As Object
    Set resultData = JsonConverter.ParseJson(resultJson)(tableName)

    Dim head As Object
    Set head = resultData("HEAD")

    Dim colIndex As Integer, idIndex As Integer
    colIndex = -1: idIndex = -1

    Dim i As Integer
    i = 0
    Dim h As Variant
    For Each h In head
        If h = columnName Then colIndex = i
        If h = idColumn Then idIndex = i
        i = i + 1
    Next h

    Dim result As Object
    Set result = New Dictionary

    If colIndex = -1 Then
        result.Add "Error", "Column not found: " & columnName
        Set GetTableExtreme = result
        Exit Function
    End If

    Dim extremeVal As Double
    Dim extremeId As Variant
    Dim first As Boolean
    first = True

    Dim rowData As Variant
    For Each rowData In resultData("DATA")

        Dim rowArr() As Variant
        Dim n As Integer
        n = 0
        ReDim rowArr(rowData.count - 1)

        Dim cellVal As Variant
        For Each cellVal In rowData
            rowArr(n) = cellVal
            n = n + 1
        Next cellVal

        Dim v As Double
        v = CDbl(rowArr(colIndex))

        If first Then
            extremeVal = v
            If idIndex <> -1 Then extremeId = rowArr(idIndex)
            first = False
        Else
            If (findMax And v > extremeVal) Or (Not findMax And v < extremeVal) Then
                extremeVal = v
                If idIndex <> -1 Then extremeId = rowArr(idIndex)
            End If
        End If

    Next rowData

    result.Add "Value", extremeVal
    result.Add "ID", extremeId

    Set GetTableExtreme = result

End Function


Function DisplacementGlobalTable(ByVal loadCaseNames As Variant, ByVal startCell As Range, _
                                 Optional ByVal component As String = "DZ", _
                                 Optional ByVal nodeIds As Variant = Null, _
                                 Optional ByVal tableName As String = "Displacements(Global)") As String

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "TABLE_NAME", tableName
    arg.Add "TABLE_TYPE", "DISPLACEMENTG"
    arg.Add "COMPONENTS", Array("Node", "Load", component)

    If Not IsNull(nodeIds) Then
        Dim nodeElems As Object
        Set nodeElems = New Dictionary
        nodeElems.Add "KEYS", nodeIds
        arg.Add "NODE_ELEMS", nodeElems
    End If

    arg.Add "LOAD_CASE_NAMES", loadCaseNames

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    Dim resultJson As String
    resultJson = CallPost("/post/TABLE", JsonConverter.ConvertToJson(body))

    Dim resultData As Object
    Set resultData = JsonConverter.ParseJson(resultJson)(tableName)

    Dim col As Integer
    col = 0
    Dim h As Variant
    For Each h In resultData("HEAD")
        startCell.Offset(0, col).Value = h
        col = col + 1
    Next h

    Dim r As Integer
    r = 1
    Dim rowData As Variant, cellVal As Variant
    For Each rowData In resultData("DATA")
        col = 0
        For Each cellVal In rowData
            startCell.Offset(r, col).Value = cellVal
            col = col + 1
        Next cellVal
        r = r + 1
    Next rowData

    DisplacementGlobalTable = resultJson

End Function

' ------------------------------------------
' Load Summary Table
' ------------------------------------------
' LoadSummaryTable(axis, startCell, [tableName])
' ------------------------------------------

Function LoadSummaryTable(ByVal axis As String, ByVal startCell As Range, _
                          Optional ByVal tableName As String = "Example") As String

    Dim body As Object
    Set body = New Dictionary

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "TABLE_NAME", tableName
    arg.Add "TABLE_TYPE", "LOAD_SUMMARY_" & UCase(axis)
    body.Add "Argument", arg

    Dim resultJson As String
    resultJson = CallPost("/post/TABLE", JsonConverter.ConvertToJson(body))

    Dim resultData As Object
    Set resultData = JsonConverter.ParseJson(resultJson)(tableName)

    Dim col As Integer
    col = 0
    Dim h As Variant
    For Each h In resultData("HEAD")
        startCell.Offset(0, col).Value = h
        col = col + 1
    Next h

    Dim r As Integer
    r = 1
    Dim rowData As Variant, cellVal As Variant
    For Each rowData In resultData("DATA")
        col = 0
        For Each cellVal In rowData
            startCell.Offset(r, col).Value = cellVal
            col = col + 1
        Next cellVal
        r = r + 1
    Next rowData

    LoadSummaryTable = resultJson

End Function


'==========================================================
' [9] Helpers - boundaries, connections and item shaped loads
'==========================================================
'  Every endpoint in this section stores its data the same way
'
'      Assign : { "<key>" : { "ITEMS" : [ {ID:1, ...}, {ID:2, ...} ] } }
'
'  One key - a node or an element - can carry several entries, normally
'  one per load case. A PUT replaces the whole record, so a second call
'  for the same key would silently drop what the first one wrote. That
'  is not a corner case: across the 47 tutorial models 45% of the PRES
'  records hold more than one item, 22% of GTMP, 15% of NSPR, 13% of
'  ETMP and 12% of CNLD.
'
'  So the helpers here never overwrite: an item for a key that already
'  holds some is appended inside the store ([10]) and the ITEMS IDs are
'  renumbered 1, 2, 3 ... when the record is built. bAppend is still
'  accepted so older code compiles; it no longer changes anything.
'
'  ID inside ITEMS is a slot number, not a position. Deleting an item
'  leaves its number behind, so a record can hold ID 2, 3, 4 with no 1.
'  Across the tutorial models that is 23% of them, and no record anywhere
'  has a repeated or out of order ID. That is why appending takes the
'  highest ID and adds one instead of counting the items.
'
'  ELNK, THIK and STMP are flat records with no ITEMS array, so they
'  have no bAppend. Those three follow the section [8] habit of
'  answering a GET when the id is 0.
'
'  Ids may be a single number or an array:  15  or  Array(15, 16, 17)
'==========================================================


'----------------------------------------------------------
' Shared plumbing
'----------------------------------------------------------

' A single number or an array, always out as an array.
Private Function CvIds(ByVal pIds As Variant) As Variant
    If IsArray(pIds) Then
        CvIds = pIds
    Else
        CvIds = Array(pIds)
    End If
End Function

' A shallow copy, so one template can be written under many keys.
Private Function CvDictClone(ByVal pSrc As Object) As Object
    Dim out As Object
    Dim k As Variant

    Set out = New Dictionary
    For Each k In pSrc.Keys
        out.Add k, pSrc.Item(k)
    Next k

    Set CvDictClone = out
End Function

' Exactly pCount doubles, whatever came in. Missing entries are 0.
'   CvNums(-18.59, 5)              -> -18.59 0 0 0 0
'   CvNums(Array(0, -9.9, -9.9), 5) ->  0 -9.9 -9.9 0 0
Private Function CvNums(ByVal pAny As Variant, ByVal pCount As Long) As Variant
    Dim out() As Double
    Dim i As Long
    Dim n As Long

    ReDim out(pCount - 1)

    If IsMissing(pAny) Then
        CvNums = out
        Exit Function
    End If

    If IsArray(pAny) Then
        n = UBound(pAny) - LBound(pAny) + 1
        For i = 0 To pCount - 1
            If i < n Then out(i) = CDbl(pAny(LBound(pAny) + i))
        Next i
    Else
        out(0) = CDbl(pAny)
    End If

    CvNums = out
End Function

' Exactly pCount booleans. Missing or omitted means all False.
Private Function CvBools(ByVal pAny As Variant, ByVal pCount As Long) As Variant
    Dim out() As Boolean
    Dim i As Long
    Dim n As Long

    ReDim out(pCount - 1)

    If Not IsMissing(pAny) Then
        If IsArray(pAny) Then
            n = UBound(pAny) - LBound(pAny) + 1
            For i = 0 To pCount - 1
                If i < n Then out(i) = CBool(pAny(LBound(pAny) + i))
            Next i
        End If
    End If

    CvBools = out
End Function

' Write one item under every key. The engine behind every ITEMS shaped
' helper in this section. The item goes into the store ([10]); items for a
' key that already holds some are appended and renumbered there.
' pAppend is kept so older code still compiles - appending is now always on.
Private Function CvItemsPut(ByVal pUri As String, ByVal pIds As Variant, _
                            ByVal pItem As Object, ByVal pAppend As Boolean) As String
    Dim vIds As Variant
    Dim rec As Object
    Dim col As Collection
    Dim i As Long

    vIds = CvIds(pIds)
    For i = LBound(vIds) To UBound(vIds)
        Set col = New Collection
        col.Add CvDictClone(pItem)
        Set rec = New Dictionary
        rec.Add "ITEMS", col
        StorePut pUri, vIds(i), rec
    Next i
    mLastStatus = 200
    mLastError = ""
    mLastResponse = ""
    CvItemsPut = ""
End Function

' Write one flat record - no ITEMS array.
Private Function CvRecordPut(ByVal pUri As String, ByVal pKey As Variant, _
                             ByVal pRec As Object) As String
    StorePut pUri, pKey, pRec
    mLastStatus = 200
    mLastError = ""
    mLastResponse = ""
    CvRecordPut = ""
End Function


' ------------------------------------------
' Nodal Load                      db/CNLD
' ------------------------------------------
' NodalLoad(nodeIds, LCNAME, [FX], [FY], [FZ], [MX], [MY], [MZ],
'           [GROUP_NAME], [bAppend])
'
'   NodalLoad 12, "WX", FX:=100000
'   NodalLoad Array(2, 3), "DL", FZ:=-50, bAppend:=True
' ------------------------------------------

Function NodalLoad(ByVal nodeIds As Variant, ByVal LCNAME As String, _
                   Optional ByVal FX As Double = 0, Optional ByVal FY As Double = 0, _
                   Optional ByVal FZ As Double = 0, Optional ByVal MX As Double = 0, _
                   Optional ByVal MY As Double = 0, Optional ByVal MZ As Double = 0, _
                   Optional ByVal GROUP_NAME As String = "", _
                   Optional ByVal bAppend As Boolean = False) As String

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "FX", FX
    item.Add "FY", FY
    item.Add "FZ", FZ
    item.Add "MX", MX
    item.Add "MY", MY
    item.Add "MZ", MZ

    NodalLoad = CvItemsPut("db/CNLD", nodeIds, item, bAppend)

End Function


' ------------------------------------------
' Pressure Load                   db/PRES
' ------------------------------------------
' PressureLoad(elemIds, LCNAME, P, [DIRECTION], [CMD], [ELEM_TYPE],
'              [GROUP_NAME], [bAppend])
' PressureEdgeLoad(elemIds, LCNAME, P, EDGE_FACE, [DIRECTION],
'              [GROUP_NAME], [bAppend])
'
'   CMD  "PRES"   one pressure over the face, P is a single number
'        "HYDRO"  a varying pressure, P is up to 5 numbers, corner by
'                 corner - this is how earth and water pressure on a
'                 culvert or a tunnel wall is stored
'   DIRECTION  GX GY GZ (global) or LZ (normal to the face)
'   EDGE_FACE  which edge of the element carries the load, 1 to 4
'
'   PressureLoad 11, "DL", -18.59
'   PressureLoad Array(20, 21), "EP", Array(0, -9.92, -9.92, -16.92, -16.92), _
'                DIRECTION:="LZ", CMD:="HYDRO", bAppend:=True
' ------------------------------------------

Function PressureLoad(ByVal elemIds As Variant, ByVal LCNAME As String, ByVal P As Variant, _
                      Optional ByVal DIRECTION As String = "GZ", _
                      Optional ByVal CMD As String = "PRES", _
                      Optional ByVal ELEM_TYPE As String = "PLATE", _
                      Optional ByVal GROUP_NAME As String = "", _
                      Optional ByVal bAppend As Boolean = False, _
                      Optional ByVal VECTORS As Variant = Null) As String

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "CMD", UCase$(CMD)
    item.Add "ELEM_TYPE", UCase$(ELEM_TYPE)
    item.Add "FACE_EDGE_TYPE", "FACE"
    item.Add "DIRECTION", UCase$(DIRECTION)
    ' direction vector (used when DIRECTION = "VECTOR"), always sent, default (1, 0, 0)
    If IsNull(VECTORS) Then item.Add "VECTORS", Array(1, 0, 0) Else item.Add "VECTORS", CvNums(VECTORS, 3)

    ' HYDRO records carry no projection flag.

    item.Add "FORCES", CvNums(P, 5)

    PressureLoad = CvItemsPut("db/PRES", elemIds, item, bAppend)

End Function

Function PressureEdgeLoad(ByVal elemIds As Variant, ByVal LCNAME As String, ByVal P As Variant, _
                          ByVal EDGE_FACE As Long, Optional ByVal DIRECTION As String = "GZ", _
                          Optional ByVal GROUP_NAME As String = "", _
                          Optional ByVal bAppend As Boolean = False) As String

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "CMD", "PRES"
    item.Add "ELEM_TYPE", "PLATE"
    item.Add "FACE_EDGE_TYPE", "EDGE"
    item.Add "DIRECTION", UCase$(DIRECTION)
    item.Add "OPT_PROJECTION", False
    item.Add "EDGE_LOADS", CvNums(P, 3)
    item.Add "EDGE_FACE", EDGE_FACE

    PressureEdgeLoad = CvItemsPut("db/PRES", elemIds, item, bAppend)

End Function


' ------------------------------------------
' Point Spring                    db/NSPR
' ------------------------------------------
' NodeSpringLinear(nodeIds, SDR, [F_S], [GROUP_NAME], [bAppend])
' NodeSpringComp(nodeIds, STIFF, [DIR], [DV], [GROUP_NAME], [bAppend])
'
'   SDR    six stiffnesses, SDx SDy SDz SRx SRy SRz
'   F_S    six flags, True fixes that direction instead of springing it
'   DIR    compression only direction code, 6 is -Z (the usual soil case)
'   DV     the direction vector when DIR needs one, three numbers
'
'   NodeSpringLinear 1, Array(180000, 180000, 180000, 90000, 90000, 90000)
'   NodeSpringComp Array(16, 17, 18), 8059.93, 6
'
' A spring under a slab is usually easier to make with
' SurfaceSpringComp, which spreads one subgrade modulus over the
' elements. Use these two when the stiffness is already per node.
' ------------------------------------------

Function NodeSpringLinear(ByVal nodeIds As Variant, ByVal SDR As Variant, _
                          Optional ByVal F_S As Variant, Optional ByVal GROUP_NAME As String = "", _
                          Optional ByVal bAppend As Boolean = False) As String

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "TYPE", "LINEAR"
    item.Add "SDR", CvNums(SDR, 6)
    item.Add "F_S", CvBools(F_S, 6)
    item.Add "DAMPING", False
    item.Add "GROUP_NAME", GROUP_NAME

    NodeSpringLinear = CvItemsPut("db/NSPR", nodeIds, item, bAppend)

End Function

Function NodeSpringComp(ByVal nodeIds As Variant, ByVal STIFF As Double, _
                        Optional ByVal DIR As Long = 6, Optional ByVal DV As Variant, _
                        Optional ByVal GROUP_NAME As String = "", _
                        Optional ByVal bAppend As Boolean = False) As String

    Dim vDir As Variant

    If IsMissing(DV) Then
        vDir = Array(0#, 0#, -1#)
    Else
        vDir = CvNums(DV, 3)
    End If

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "TYPE", "COMP"
    item.Add "DIR", DIR
    item.Add "DV", vDir
    item.Add "STIFF", STIFF
    item.Add "GROUP_NAME", GROUP_NAME

    NodeSpringComp = CvItemsPut("db/NSPR", nodeIds, item, bAppend)

End Function


' ------------------------------------------
' Rigid Link                      db/RIGD
' ------------------------------------------
' RigidLink(masterNode, slaveNodes, [DOF], [GROUP_NAME], [bAppend])
'
'   DOF  six digits, one per direction, 1 means tied. 111111 is the
'        full rigid body link and is what nearly every model uses.
'
'   RigidLink 3, Array(6, 9, 48, 52)
' ------------------------------------------

Function RigidLink(ByVal masterNode As Variant, ByVal slaveNodes As Variant, _
                   Optional ByVal DOF As Long = 111111, Optional ByVal GROUP_NAME As String = "", _
                   Optional ByVal bAppend As Boolean = False) As String

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "DOF", DOF
    item.Add "S_NODE", CvIds(slaveNodes)

    RigidLink = CvItemsPut("db/RIGD", masterNode, item, bAppend)

End Function


' ------------------------------------------
' Beam End Release                db/FRLS
' ------------------------------------------
' BeamRelease(elemIds, FLAG_I, FLAG_J, [GROUP_NAME], [bAppend])
'
'   FLAG  seven characters, Fx Fy Fz Mx My Mz Wa, 1 means released.
'         "0000110" frees My and Mz - the usual pin.
'         "0000000" is fully fixed, which is the same as no release.
'
'   BeamRelease 14, "0000110", "0000110"
' ------------------------------------------

Function BeamRelease(ByVal elemIds As Variant, ByVal FLAG_I As String, ByVal FLAG_J As String, _
                     Optional ByVal GROUP_NAME As String = "", _
                     Optional ByVal bAppend As Boolean = False) As String

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "bVALUE", False
    item.Add "FLAG_I", FLAG_I
    item.Add "VALUE_I", CvNums(Empty, 7)
    item.Add "FLAG_J", FLAG_J
    item.Add "VALUE_J", CvNums(Empty, 7)

    BeamRelease = CvItemsPut("db/FRLS", elemIds, item, bAppend)

End Function


' ------------------------------------------
' Element Temperature             db/ETMP
' ------------------------------------------
' ElemTemp(elemIds, LCNAME, TEMP, [GROUP_NAME], [bAppend])
'
'   A uniform temperature change on the elements themselves.
'
'   ElemTemp Array(1, 2, 3), "T(+15)", 15
'   ElemTemp Array(1, 2, 3), "T(-15)", -15, bAppend:=True
' ------------------------------------------

Function ElemTemp(ByVal elemIds As Variant, ByVal LCNAME As String, ByVal TEMP As Double, _
                  Optional ByVal GROUP_NAME As String = "", _
                  Optional ByVal bAppend As Boolean = False) As String

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "TEMP", TEMP

    ElemTemp = CvItemsPut("db/ETMP", elemIds, item, bAppend)

End Function


' ------------------------------------------
' Temperature Gradient            db/GTMP
' ------------------------------------------
' TempGradient(elemIds, LCNAME, TZ, [ETYPE], [HZ], [TY], [HY],
'              [GROUP_NAME], [bAppend])
'
'   ETYPE  1 beam, 2 plate. TY and HY only exist for a beam.
'   TZ     temperature difference across the depth
'   HZ     the depth the difference acts over. Leave it 0 and the
'          section depth is used instead (USE_HZ goes True), which is
'          what every tutorial model does.
'
'   TempGradient 27, "top to bottom", 10, ETYPE:=2
' ------------------------------------------

Function TempGradient(ByVal elemIds As Variant, ByVal LCNAME As String, ByVal TZ As Double, _
                      Optional ByVal ETYPE As Long = 1, Optional ByVal HZ As Double = 0, _
                      Optional ByVal TY As Double = 0, Optional ByVal HY As Double = 0, _
                      Optional ByVal GROUP_NAME As String = "", _
                      Optional ByVal bAppend As Boolean = False) As String

    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "TYPE", ETYPE
    item.Add "TZ", TZ
    item.Add "USE_HZ", (HZ = 0)
    If HZ <> 0 Then item.Add "HZ", HZ

    If ETYPE = 1 Then
        item.Add "TY", TY
        item.Add "USE_HY", (HY = 0)
        If HY <> 0 Then item.Add "HY", HY
    End If

    TempGradient = CvItemsPut("db/GTMP", elemIds, item, bAppend)

End Function


' ------------------------------------------
' System Temperature              db/STMP
' ------------------------------------------
' SystemTemp(stmpId, TEMPER, LCNAME, [GROUP_NAME])
'
'   One temperature for the whole model, per load case. Shrinkage is
'   usually entered this way too.
'
'   SystemTemp 1, -25, "T(-25)"
'   SystemTemp 0                       ' read them all back
' ------------------------------------------

Function SystemTemp(ByVal stmpId As Long, Optional ByVal TEMPER As Double = 0, _
                    Optional ByVal LCNAME As String = "", Optional ByVal GROUP_NAME As String = "") As String

    If stmpId = 0 Then
        SystemTemp = CallGet("/db/STMP")
        Exit Function
    End If

    Dim rec As Object
    Set rec = New Dictionary
    rec.Add "TEMPER", TEMPER
    rec.Add "LCNAME", LCNAME
    rec.Add "GROUP_NAME", GROUP_NAME

    SystemTemp = CvRecordPut("db/STMP", stmpId, rec)

End Function


' ------------------------------------------
' Elastic Link                    db/ELNK
' ------------------------------------------
' ElasticLink(linkId, node1, node2, [LINK], [SDR], [ANGLE], [bSHEAR],
'             [BNGR_NAME])
'
'   LINK  "RIGID" no stiffness needed - the usual way to tie a girder
'                 to a bearing
'         "GEN"   general, SDR carries six stiffnesses
'         "COMP"  compression only, SDR carries six stiffnesses
'   SDR   SDx SDy SDz SRx SRy SRz
'   bSHEAR  put the shear spring at the link ends rather than at the
'         centre. Leave it out and it follows what the models do:
'         True for COMP, False otherwise. In the 47 tutorials every
'         one of the 537 COMP links has it on, every RIGID link has it
'         off, and GEN is split 423 on / 181 off - so say it out loud
'         for a GEN link.
'
'   ElasticLink 1, 27, 83
'   ElasticLink 2, 82, 84, "GEN", Array(9.8E+11, 9.8E+11, 9.8E+11, 0, 0, 0)
'   ElasticLink 0                      ' read them all back
' ------------------------------------------

Function ElasticLink(ByVal linkId As Long, Optional ByVal node1 As Long = 0, _
                     Optional ByVal node2 As Long = 0, Optional ByVal LINK As String = "RIGID", _
                     Optional ByVal SDR As Variant, Optional ByVal ANGLE As Double = 0, _
                     Optional ByVal bSHEAR As Variant, Optional ByVal BNGR_NAME As String = "") As String

    If linkId = 0 Then
        ElasticLink = CallGet("/db/ELNK")
        Exit Function
    End If

    Dim bShr As Boolean
    If IsMissing(bSHEAR) Then
        bShr = (UCase$(LINK) = "COMP")
    Else
        bShr = CBool(bSHEAR)
    End If

    Dim nodes(1) As Long
    nodes(0) = node1
    nodes(1) = node2

    Dim dr(1) As Double
    dr(0) = 0.5
    dr(1) = 0.5

    Dim rec As Object
    Set rec = New Dictionary
    rec.Add "NODE", nodes
    rec.Add "ANGLE", ANGLE
    rec.Add "LINK", UCase$(LINK)

    If UCase$(LINK) = "GEN" Then rec.Add "R_S", CvBools(Empty, 6)
    If UCase$(LINK) <> "RIGID" Then rec.Add "SDR", CvNums(SDR, 6)

    If UCase$(LINK) <> "RIGID" Then
        rec.Add "bSHEAR", bShr
        rec.Add "DR", dr
    End If
    rec.Add "BNGR_NAME", BNGR_NAME

    ElasticLink = CvRecordPut("db/ELNK", linkId, rec)

End Function


' ------------------------------------------
' Thickness                       db/THIK
' ------------------------------------------
' PlateThickness(thikId, T_IN, [NAME], [O_VALUE])
'
'   The thickness a plate element points at through its SECT field,
'   the same way a beam points at a section.
'
'   PlateThickness 1, 0.3
'   PlateThickness 0                   ' read them all back
' ------------------------------------------

Function PlateThickness(ByVal thikId As Long, Optional ByVal T_IN As Double = 0, _
                        Optional ByVal NAME As String = "", Optional ByVal O_VALUE As Double = 0, _
                        Optional ByVal T_OUT As Double = -1, Optional ByVal OFF_TYPE As String = "rat") As String
    ' Plate thickness: in / out thickness and offset
    '   T_OUT  -1 = same as T_IN (bINOUT False)
    '   O_VALUE offset value, OFF_TYPE "rat" (ratio) / "val" (value) - no offset when O_VALUE = 0
    Dim rec As Object

    If thikId = 0 Then
        PlateThickness = CallGet("/db/THIK")
        Exit Function
    End If

    Set rec = New Dictionary
    If NAME = "" Then rec.Add "NAME", CStr(T_IN) Else rec.Add "NAME", NAME
    rec.Add "TYPE", "VALUE"
    rec.Add "bINOUT", (T_OUT <> -1)
    rec.Add "T_IN", T_IN
    If T_OUT = -1 Then rec.Add "T_OUT", T_IN Else rec.Add "T_OUT", T_OUT
    If O_VALUE = 0 Then
        rec.Add "OFFSET", 0
    ElseIf LCase$(OFF_TYPE) = "rat" Then
        rec.Add "OFFSET", 1
    Else
        rec.Add "OFFSET", 2
    End If
    rec.Add "O_VALUE", O_VALUE
    PlateThickness = CvRecordPut("db/THIK", thikId, rec)
End Function


'==========================================================
' [10] Model store - collect first, send once
'==========================================================
'  Every helper that writes model data (nodes, elements, materials,
'  sections, supports, loads, groups ...) no longer calls NX right away.
'  It adds its records to a store inside this module. ModelCreate then
'  sends everything in one go, item by item, in dependency order
'  (materials, sections, nodes, elements, ...), with PUT so a record that
'  already exists is overwritten.
'
'      Node 1, 0, 0, 0
'      Node 2, 10, 0, 0
'      Beam 1, 1, 1, 1, 2
'      Support 1, "1110000"
'      sErr = ModelCreate()       ' "" when every item went through
'
'  Why: one request per item instead of one per call - a model with a few
'  thousand nodes goes out in a handful of requests - and the items always
'  reach NX in dependency order whatever order the code wrote them in.
'
'  Records that share a key are merged the way NX stores them:
'    - flat records (NODE, ELEM, SECT ...)  the later call replaces the record
'    - ITEMS records (CONS, CNLD, BMLD, PRES, NSPR, ETMP, GTMP ...)
'      the new items are appended and ITEMS(i).ID is renumbered 1, 2, 3 ...
'
'  Calls on the live model (ope/... : divide element, line beam load,
'  surface spring, load combination usage) are queued and run after the
'  model data, in the order they were written.
'
'  Still sent right away: files (NewFile, OpenFile ...), results (post/TABLE),
'  raw calls (CallGet / CallPost / CallPut / CallDelete / ApiRaw) and every
'  read (id 0 / GET). RunAnalysis, SaveFile, SaveFileAs, SaveStageAs,
'  ExportJson and ExportMct send whatever is still in the store first.
'
'  ModelCreate empties the store when it is done (pKeep:=True keeps it),
'  so pressing a button twice does not send everything twice.
'==========================================================

' (mStore / mRaw / mOps are declared at the top of the module)

' Order the items are sent in, each after the items it depends on.
' Names not listed go last, in the order they were first written.
Private Function CvStoreOrder() As Variant
    CvStoreOrder = Split( _
        "UNIT STYP MVCD MATL TDMT TDME TMAT EDMP SECT THIK NODE ELEM TSGR SKEW " & _
        "GRUP BNGR LDGR TDGR STOR ESSF WSSF CONS ELNK RIGD MLFC NSPR FRLS OFFS " & _
        "STLD BODF CNLD BMLD PRES FBLD FBLA PLCB PNLD PNLA LTOM SDSP NMAS " & _
        "STMP ETMP GTMP NTMP BTMP TDNT TDNA TDPL STAG CSCS TMLD CRPC CMCS " & _
        "LLANID LLANCH LLAN MVHL MVLDID MVLDEU MVLD " & _
        "ACTL PDEL BUCK EIGV SMCT HHCT SPFC SPLC THFC THIS THGA THSL THNL THMS " & _
        "HPCE HSPT ETFC CCFC HECB HSFC HAHS HSTG HHND SMPT SMLC BCCT " & _
        "LCOM-GEN LCOM-CONC LCOM-STEEL LCOM-SRC LCOM-STLCOMP LCOM-SEISMIC LCOM-FDN CUTL CLWP", " ")
End Function

Private Sub CvStoreInit()
    If mStore Is Nothing Then Set mStore = New Dictionary
    If mRaw Is Nothing Then Set mRaw = New Dictionary
    If mOps Is Nothing Then Set mOps = New Collection
End Sub

Private Function CvStoreName(ByVal pEndpoint As String) As String
    Dim s As String
    s = Trim$(Replace(pEndpoint, "\", "/"))
    Do While Left$(s, 1) = "/"
        s = Mid$(s, 2)
    Loop
    If LCase$(Left$(s, 3)) = "db/" Then s = Mid$(s, 4)
    CvStoreName = UCase$(s)
End Function

Private Function CvStoreTable(ByVal pName As String) As Object
    CvStoreInit
    If Not mStore.Exists(pName) Then mStore.Add pName, New Dictionary
    Set CvStoreTable = mStore.Item(pName)
End Function

' Put one record under one key. ITEMS records are appended to what the key
' already holds; anything else replaces it.
Public Sub StorePut(ByVal pEndpoint As String, ByVal pKey As Variant, ByVal pRecord As Object)
    Dim oTable As Object
    Dim oOld As Object
    Dim oItems As Collection
    Dim v As Variant
    Dim sKey As String
    Dim nextId As Long

    Set oTable = CvStoreTable(CvStoreName(pEndpoint))
    sKey = Trim$(CStr(pKey))
    CvAutoRecord pEndpoint, pRecord          ' [11] groups / load cases the moment they are named

    If Not pRecord.Exists("ITEMS") Then
        If oTable.Exists(sKey) Then oTable.Remove sKey
        oTable.Add sKey, pRecord
        Exit Sub
    End If

    Set oItems = New Collection
    nextId = 0
    If oTable.Exists(sKey) Then
        Set oOld = oTable.Item(sKey)
        If oOld.Exists("ITEMS") Then
            For Each v In oOld.Item("ITEMS")
                oItems.Add v
                nextId = nextId + 1
            Next v
        End If
        oTable.Remove sKey
    End If
    For Each v In pRecord.Item("ITEMS")
        If IsObject(v) Then
            nextId = nextId + 1
            If TypeName(v) = "Dictionary" Then v.Item("ID") = nextId
            oItems.Add v
        End If
    Next v
    Set pRecord.Item("ITEMS") = oItems
    oTable.Add sKey, pRecord
End Sub

' Put a whole { id : record } block - what JNew / JSet build.
'   StoreAssign "NODE", oAll
Public Sub StoreAssign(ByVal pEndpoint As String, ByVal pAssign As Object)
    Dim vKey As Variant
    If pAssign Is Nothing Then Exit Sub
    For Each vKey In pAssign.Keys
        If IsObject(pAssign.Item(vKey)) Then StorePut pEndpoint, vKey, pAssign.Item(vKey)
    Next vKey
End Sub

' Put a finished request body (a JSON string with "Assign"), sent as is
' before the records of the same item.
Public Sub StoreRaw(ByVal pEndpoint As String, ByVal pJson As String)
    Dim sName As String
    CvStoreInit
    sName = CvStoreName(pEndpoint)
    If Not mRaw.Exists(sName) Then mRaw.Add sName, New Collection
    mRaw.Item(sName).Add pJson
    CvStoreTable sName
End Sub

' The record stored under a key, or Nothing.
Public Function StoreGet(ByVal pEndpoint As String, ByVal pKey As Variant) As Object
    Dim oTable As Object
    Set oTable = CvStoreTable(CvStoreName(pEndpoint))
    If oTable.Exists(Trim$(CStr(pKey))) Then Set StoreGet = oTable.Item(Trim$(CStr(pKey)))
End Function

' Number of records waiting for an item. ModelCount() = every item.
Public Function ModelCount(Optional ByVal pEndpoint As String = "") As Long
    Dim vName As Variant
    CvStoreInit
    If Len(pEndpoint) > 0 Then
        ModelCount = CvStoreTable(CvStoreName(pEndpoint)).Count
        Exit Function
    End If
    For Each vName In mStore.Keys
        ModelCount = ModelCount + mStore.Item(vName).Count
    Next vName
    ModelCount = ModelCount + mOps.Count
    For Each vName In mRaw.Keys
        ModelCount = ModelCount + mRaw.Item(vName).Count
    Next vName
End Function

Public Sub ModelClear()
    Set mStore = New Dictionary
    Set mRaw = New Dictionary
    Set mOps = New Collection
    Set mGrpSeen = Nothing
    CvGeoReset                               ' [24]
End Sub

' Send everything in the store. Returns "" when every request went through,
' otherwise one line per failed item: "ELEM: <NX error>". Items after a
' failure are still sent.
Public Function ModelCreate(Optional ByVal pKeep As Boolean = False) As String
    Dim vName As Variant
    Dim vOrder As Variant
    Dim oDone As Object
    Dim sErr As String
    Dim v As Variant
    Dim i As Long

    CvStoreInit
    CvAutoGroupsAndCases           ' [11]
    Set oDone = New Dictionary
    vOrder = CvStoreOrder()

    For i = LBound(vOrder) To UBound(vOrder)
        If mStore.Exists(vOrder(i)) Then
            sErr = sErr & CvSendStored(CStr(vOrder(i)))
            oDone.Add vOrder(i), True
        End If
    Next i
    For Each vName In mStore.Keys
        If Not oDone.Exists(vName) Then sErr = sErr & CvSendStored(CStr(vName))
    Next vName

    For Each v In mOps
        CvHttp CStr(v(0)), CStr(v(1)), CStr(v(2))
        If Not CvIsOk() Then sErr = sErr & CStr(v(1)) & ": " & CvErrorText() & vbLf
    Next v

    If Not pKeep Then ModelClear
    If Len(sErr) > 0 Then
        sErr = Left$(sErr, Len(sErr) - 1)
        mLastStatus = -2
        mLastError = sErr
        mLastResponse = ""
    Else
        mLastStatus = 200
        mLastError = ""
    End If
    ModelCreate = sErr
End Function

' Send one item only (it is taken out of the store).
Public Function ModelCreateItem(ByVal pEndpoint As String) As String
    Dim sName As String
    Dim sErr As String

    CvStoreInit
    sName = CvStoreName(pEndpoint)
    If Not mStore.Exists(sName) Then Exit Function
    sErr = CvSendStored(sName)
    mStore.Remove sName
    If mRaw.Exists(sName) Then mRaw.Remove sName
    If Len(sErr) > 0 Then sErr = Left$(sErr, Len(sErr) - 1)
    ModelCreateItem = sErr
End Function

' PUT one item: finished bodies first, then the records - split into chunks
' when the body would pass 7 MB.
Private Function CvSendStored(ByVal pName As String) As String
    Const MAX_BYTES As Double = 7340032#
    Dim oTable As Object
    Dim oPart As Object
    Dim body As Object
    Dim sJson As String
    Dim sPath As String
    Dim sErr As String
    Dim vKeys As Variant
    Dim v As Variant
    Dim nParts As Long
    Dim nPer As Long
    Dim i As Long
    Dim k As Long
    Dim kEnd As Long

    sPath = "/db/" & pName
    If mRaw.Exists(pName) Then
        For Each v In mRaw.Item(pName)
            CvHttp "PUT", sPath, CStr(v)
            If Not CvIsOk() Then sErr = sErr & pName & ": " & CvErrorText() & vbLf
        Next v
    End If

    Set oTable = mStore.Item(pName)
    If pName = "GRUP" Then Set oTable = CvVisibleGroups(oTable)   ' [11] "#name" groups stay local
    If oTable.Count = 0 Then
        CvSendStored = sErr
        Exit Function
    End If

    Set body = New Dictionary
    body.Add "Assign", oTable
    sJson = JsonConverter.ConvertToJson(body)
    nParts = Int(Len(sJson) / MAX_BYTES) + 1

    If nParts = 1 Then
        CvHttp "PUT", sPath, sJson
        If Not CvIsOk() Then sErr = sErr & pName & ": " & CvErrorText() & vbLf
        CvSendStored = sErr
        Exit Function
    End If

    vKeys = oTable.Keys
    nPer = -Int(-(oTable.Count / nParts))
    For i = 0 To nParts - 1
        Set oPart = New Dictionary
        kEnd = (i + 1) * nPer - 1
        If kEnd > oTable.Count - 1 Then kEnd = oTable.Count - 1
        For k = i * nPer To kEnd
            oPart.Add vKeys(k), oTable.Item(vKeys(k))
        Next k
        If oPart.Count > 0 Then
            Set body = New Dictionary
            body.Add "Assign", oPart
            CvHttp "PUT", sPath, JsonConverter.ConvertToJson(body)
            If Not CvIsOk() Then sErr = sErr & pName & ": " & CvErrorText() & vbLf
        End If
    Next i
    CvSendStored = sErr
End Function

' Where every helper hands over its body. db/... with Assign goes into the
' store; anything else (ope/...) is queued to run after the model data.
' Leaves CvIsOk() True - the call itself cannot fail any more.
Private Function CvQueue(ByVal pPath As String, ByVal pBody As Object) As String
    Dim sPath As String

    CvStoreInit
    sPath = CvNormalizePath(pPath)
    If LCase$(Left$(sPath, 4)) = "/db/" And pBody.Exists("Assign") Then
        StoreAssign sPath, pBody.Item("Assign")
    Else
        mOps.Add Array("POST", sPath, JsonConverter.ConvertToJson(pBody))
    End If
    mLastStatus = 200
    mLastError = ""
    mLastResponse = ""
    CvQueue = ""
End Function

' Before a call that needs the model in NX (analysis, save, export):
' send what is still waiting. "" when nothing failed.
Private Function CvFlushPending() As String
    If ModelCount() = 0 Then Exit Function
    CvFlushPending = ModelCreate()
End Function

' Nodes as NX has them, with the store laid over the top - for the helpers
' that look nodes up by position (NodeRange, NodeRectangle, BeamRange).
Private Function CvAllNodes() As Object
    Dim oAll As Object
    Dim oNx As Object
    Dim vKey As Variant

    Set oAll = New Dictionary
    On Error Resume Next
    Set oNx = JsonConverter.ParseJson(CallGet("/db/NODE"))("NODE")
    On Error GoTo 0
    If Not oNx Is Nothing Then
        For Each vKey In oNx.Keys
            Set oAll.Item(CStr(vKey)) = oNx.Item(vKey)
        Next vKey
    End If
    With CvStoreTable("NODE")
        For Each vKey In .Keys
            If oAll.Exists(CStr(vKey)) Then oAll.Remove CStr(vKey)
            oAll.Add CStr(vKey), .Item(vKey)
        Next vKey
    End With
    Set CvAllNodes = oAll
End Function

' GET /db/XXXX -> { id : record }, or Nothing.
Private Function CvReadItems(ByVal pUri As String) As Object
    Dim oRoot As Object
    Dim sName As String
    Dim vKeys As Variant

    sName = CvStoreName(pUri)
    Set oRoot = JParse(CallGet("/db/" & sName))
    If oRoot Is Nothing Then Exit Function
    If TypeName(oRoot) <> "Dictionary" Then Exit Function
    If oRoot.Exists(sName) Then
        If IsObject(oRoot.Item(sName)) Then Set CvReadItems = oRoot.Item(sName)
    ElseIf oRoot.Count = 1 Then
        vKeys = oRoot.Keys
        If IsObject(oRoot.Item(vKeys(0))) Then Set CvReadItems = oRoot.Item(vKeys(0))
    End If
End Function


'==========================================================
' [11] Groups - structure / boundary / load / tendon
'==========================================================
'  How groups work:
'    - StructureGroup "Girder", nodes, elements   (several calls with the
'      same name add to the same group)
'    - Node / Beam / Plate ... take an optional group: the node, or the
'      element and its nodes, join that structure group
'    - a GROUP_NAME given to a support, spring, link, release or stiffness
'      scale factor makes the boundary group if it does not exist yet;
'      one given to a load or temperature makes the load group
'    - a load (self weight, nodal, beam, pressure, specified displacement,
'      plane load) whose load case was never defined gets a load case of
'      type D with that name
'  The last two happen the moment the load / support is written
'  (so ids follow the call order), and once more in ModelCreate
'  for anything written with StorePut / StoreAssign.
'  Group ids are given in the order the groups are first met.
'  A structure group whose name starts with "#" is a working group: it can
'  be filled and read (NodesInGroup / ElemsInGroup, [24]) but is not sent.
'==========================================================

'   StructureGroup "Girder", Array(1, 2, 3), Array(1, 2)
'   StructureGroup "Pier"                       ' empty group, filled later
Public Function StructureGroup(ByVal NAME As String, Optional ByVal nodeIds As Variant = Null, _
                               Optional ByVal elemIds As Variant = Null) As String
    CvGroupJoin NAME, nodeIds, elemIds
    StructureGroup = ""
End Function

'   LoadGroup "Dead"
Public Function LoadGroup(ByVal NAME As String) As String
    CvGroupRecord "LDGR", NAME
    LoadGroup = ""
End Function

'   TendonGroup "Web"
Public Function TendonGroup(ByVal NAME As String) As String
    CvGroupRecord "TDGR", NAME
    TendonGroup = ""
End Function

' The group record with that name - made with the next free id if missing.
Private Function CvGroupRecord(ByVal pEndpoint As String, ByVal pName As String) As Object
    Dim oTable As Object
    Dim oRec As Object
    Dim vKey As Variant
    Dim nextId As Long

    Set oTable = CvStoreTable(CvStoreName(pEndpoint))
    For Each vKey In oTable.Keys
        If CStr(oTable.Item(vKey).Item("NAME")) = pName Then
            Set CvGroupRecord = oTable.Item(vKey)
            Exit Function
        End If
        If CLng(vKey) > nextId Then nextId = CLng(vKey)
    Next vKey

    Set oRec = New Dictionary
    oRec.Add "NAME", pName
    Select Case CvStoreName(pEndpoint)
        Case "GRUP"
            oRec.Add "P_TYPE", 0
            oRec.Add "N_LIST", New Collection
            oRec.Add "E_LIST", New Collection
        Case "BNGR"
            oRec.Add "AUTOTYPE", 0
    End Select
    oTable.Add CStr(nextId + 1), oRec
    Set CvGroupRecord = oRec
End Function

' Add nodes / elements to a structure group (ids already there are skipped).
' pName "" does nothing. Several names: "A,B".
Private Sub CvGroupJoin(ByVal pName As String, ByVal pNodes As Variant, ByVal pElems As Variant)
    Dim vName As Variant
    Dim oRec As Object

    If Len(Trim$(pName)) = 0 Then Exit Sub
    For Each vName In Split(pName, ",")
        If Len(Trim$(vName)) > 0 Then
            Set oRec = CvGroupRecord("GRUP", Trim$(vName))
            CvListAdd oRec.Item("N_LIST"), pNodes
            CvListAdd oRec.Item("E_LIST"), pElems
        End If
    Next vName
End Sub

Private Sub CvListAdd(ByVal pList As Collection, ByVal pIds As Variant)
    Dim vIds As Variant
    Dim v As Variant
    Dim i As Long
    Dim oSeen As Object
    Dim sKey As String

    If IsMissing(pIds) Then Exit Sub
    If IsNull(pIds) Then Exit Sub
    If IsEmpty(pIds) Then Exit Sub
    ' ids already in the list, kept beside it (rebuilt when the list changed elsewhere)
    If mGrpSeen Is Nothing Then Set mGrpSeen = New Dictionary
    sKey = CStr(ObjPtr(pList))
    If Not mGrpSeen.Exists(sKey) Then mGrpSeen.Add sKey, New Dictionary
    Set oSeen = mGrpSeen.Item(sKey)
    If oSeen.Count <> pList.Count Then
        oSeen.RemoveAll
        For Each v In pList
            oSeen.Item(CStr(v)) = True
        Next v
    End If
    vIds = CvIds(pIds)
    For i = LBound(vIds) To UBound(vIds)
        v = CLng(vIds(i))
        If Not oSeen.Exists(CStr(v)) Then
            pList.Add v
            oSeen.Add CStr(v), True
        End If
    Next i
End Sub

' The structure groups to send - working groups ("#name") left out.
Private Function CvVisibleGroups(ByVal pTable As Object) As Object
    Dim oOut As Object
    Dim vKey As Variant
    Set oOut = New Dictionary
    For Each vKey In pTable.Keys
        If Left$(CStr(pTable.Item(vKey).Item("NAME")), 1) <> "#" Then oOut.Add vKey, pTable.Item(vKey)
    Next vKey
    Set CvVisibleGroups = oOut
End Function

' Groups and load cases a record names, made as soon as it is stored.
Private Sub CvAutoRecord(ByVal pEndpoint As String, ByVal pRec As Object)
    Dim sName As String
    Dim bBnd As Boolean
    Dim bCase As Boolean
    Dim v As Variant

    sName = " " & CvStoreName(pEndpoint) & " "
    bBnd = InStr(CvAutoBoundary(), sName) > 0
    If Not bBnd And InStr(CvAutoLoads(), sName) = 0 Then Exit Sub
    bCase = InStr(CvAutoCases(), sName) > 0
    CvAutoOne CvStoreName(pEndpoint), pRec, bBnd, bCase
    If pRec.Exists("ITEMS") Then
        For Each v In pRec.Item("ITEMS")
            If IsObject(v) Then CvAutoOne CvStoreName(pEndpoint), v, bBnd, bCase
        Next v
    End If
End Sub

Private Function CvAutoBoundary() As String
    CvAutoBoundary = " CONS NSPR ELNK RIGD FRLS OFFS MLFC ESSF WSSF SSPS HECB "
End Function

Private Function CvAutoLoads() As String
    CvAutoLoads = " BODF CNLD BMLD PRES STMP ETMP GTMP NTMP BTMP SDSP NMAS FBLA PNLA LTOM PLCB "
End Function

Private Function CvAutoCases() As String
    CvAutoCases = " BODF CNLD BMLD PRES SDSP PNLA "
End Function

' Group names used by the stored records (GROUP_NAME on the record or on
' its ITEMS, BNGR_NAME on elastic links) - make the missing groups.
' Load names (LCNAME) of the load items - make the missing load cases (D).
Private Sub CvAutoGroupsAndCases()
    Dim BOUNDARY As String
    Dim LOADS As String
    Dim AUTO_CASE As String
    Dim vName As Variant
    Dim vKey As Variant
    Dim oRec As Object
    Dim v As Variant
    Dim sName As String

    BOUNDARY = CvAutoBoundary()
    LOADS = CvAutoLoads()
    AUTO_CASE = CvAutoCases()
    CvStoreInit
    For Each vName In mStore.Keys
        sName = " " & CStr(vName) & " "
        If InStr(BOUNDARY, sName) > 0 Or InStr(LOADS, sName) > 0 Then
            For Each vKey In mStore.Item(vName).Keys
                Set oRec = mStore.Item(vName).Item(vKey)
                CvAutoOne CStr(vName), oRec, InStr(BOUNDARY, sName) > 0, InStr(AUTO_CASE, sName) > 0
                If oRec.Exists("ITEMS") Then
                    For Each v In oRec.Item("ITEMS")
                        If IsObject(v) Then CvAutoOne CStr(vName), v, InStr(BOUNDARY, sName) > 0, InStr(AUTO_CASE, sName) > 0
                    Next v
                End If
            Next vKey
        End If
    Next vName
End Sub

Private Sub CvAutoOne(ByVal pEndpoint As String, ByVal pRec As Object, ByVal pBoundary As Boolean, ByVal pCase As Boolean)
    Dim sGroup As String

    If pRec.Exists("GROUP_NAME") Then sGroup = CStr(pRec.Item("GROUP_NAME"))
    If pRec.Exists("BNGR_NAME") Then sGroup = CStr(pRec.Item("BNGR_NAME"))
    If pRec.Exists("LOAD_GROUP") Then sGroup = CStr(pRec.Item("LOAD_GROUP"))
    If Len(sGroup) > 0 Then
        If pBoundary Then CvGroupRecord "BNGR", sGroup Else CvGroupRecord "LDGR", sGroup
    End If
    If pCase And pRec.Exists("LCNAME") Then CvAutoCase CStr(pRec.Item("LCNAME"))
End Sub

Private Sub CvAutoCase(ByVal pName As String, Optional ByVal pType As String = "D")
    Dim oTable As Object
    Dim oRec As Object
    Dim vKey As Variant
    Dim nextId As Long

    If Len(pName) = 0 Then Exit Sub
    Set oTable = CvStoreTable("STLD")
    For Each vKey In oTable.Keys
        If CStr(oTable.Item(vKey).Item("NAME")) = pName Then Exit Sub
        If CLng(vKey) > nextId Then nextId = CLng(vKey)
    Next vKey
    Set oRec = New Dictionary
    oRec.Add "NO", nextId + 1
    oRec.Add "NAME", pName
    oRec.Add "TYPE", pType
    oRec.Add "DESC", ""
    oTable.Add CStr(nextId + 1), oRec
End Sub


'==========================================================
' [12] Helpers - more element types, stiffness scale factors, node local axis
'==========================================================
'  Truss / tension / compression / solid / wall elements, stiffness scale
'  factors, node local axis. Optional values are written only when given.
'  Every element helper takes an optional structure group (see [11]).
'==========================================================

'   Truss 10, 1, 2, 3, 4
Public Function Truss(ByVal elemId As Long, ByVal matlId As Long, ByVal sectId As Long, _
                      ByVal ni As Long, ByVal nj As Long, Optional ByVal ANGLE As Double = 0, _
                      Optional ByVal GROUP As String = "") As String
    Dim rec As Object
    Set rec = CvElemRecord("TRUSS", matlId, sectId, Array(ni, nj))
    rec.Add "ANGLE", ANGLE
    Truss = CvElemPut(elemId, rec, GROUP)
End Function

' Tension only element (TENSTR). STYPE 1 truss, 2 hook, 3 cable.
'   Tension 11, 1, 2, 3, 4, 1, TENS:=10
'   Tension 12, 1, 2, 3, 4, 3, NON_LEN:=1.5, CABLE:=1
Public Function Tension(ByVal elemId As Long, ByVal matlId As Long, ByVal sectId As Long, _
                        ByVal ni As Long, ByVal nj As Long, ByVal STYPE As Long, _
                        Optional ByVal ANGLE As Double = 0, Optional ByVal TENS As Variant = Null, _
                        Optional ByVal T_LIMIT As Variant = Null, Optional ByVal NON_LEN As Variant = Null, _
                        Optional ByVal CABLE As Variant = Null, Optional ByVal GROUP As String = "") As String
    Dim rec As Object
    Set rec = CvElemRecord("TENSTR", matlId, sectId, Array(ni, nj))
    rec.Add "ANGLE", ANGLE
    rec.Add "STYPE", STYPE
    If Not IsNull(CABLE) Then rec.Add "CABLE", CABLE
    If Not IsNull(NON_LEN) Then rec.Add "NON_LEN", NON_LEN
    If Not IsNull(TENS) Then rec.Add "TENS", TENS
    If Not IsNull(T_LIMIT) Then rec.Add "T_LIMIT", T_LIMIT
    Tension = CvElemPut(elemId, rec, GROUP)
End Function

' Compression only element (COMPTR). STYPE 1 truss, 2 gap.
Public Function Compression(ByVal elemId As Long, ByVal matlId As Long, ByVal sectId As Long, _
                            ByVal ni As Long, ByVal nj As Long, ByVal STYPE As Long, _
                            Optional ByVal ANGLE As Double = 0, Optional ByVal TENS As Variant = Null, _
                            Optional ByVal T_LIMIT As Variant = Null, Optional ByVal NON_LEN As Variant = Null, _
                            Optional ByVal GROUP As String = "") As String
    Dim rec As Object
    Set rec = CvElemRecord("COMPTR", matlId, sectId, Array(ni, nj))
    rec.Add "ANGLE", ANGLE
    rec.Add "STYPE", STYPE
    If Not IsNull(TENS) Then rec.Add "TENS", TENS
    If Not IsNull(T_LIMIT) Then rec.Add "T_LIMIT", T_LIMIT
    If Not IsNull(NON_LEN) Then rec.Add "NON_LEN", NON_LEN
    Compression = CvElemPut(elemId, rec, GROUP)
End Function

' Solid - 4 (tetra), 6 (penta) or 8 (hexa) nodes.
'   Solid 20, 1, Array(1, 2, 3, 4, 5, 6, 7, 8)
Public Function Solid(ByVal elemId As Long, ByVal matlId As Long, ByVal nodeIds As Variant, _
                      Optional ByVal GROUP As String = "") As String
    Solid = CvElemPut(elemId, CvElemRecord("SOLID", matlId, 0, nodeIds), GROUP)
End Function

' Wall - 4 nodes. STYPE 1 membrane, 2 plate. WALL_ID wall number, W_TYPE 0 / 1.
'   Wall 30, 1, 1, Array(1, 2, 6, 5)
Public Function Wall(ByVal elemId As Long, ByVal matlId As Long, ByVal sectId As Long, _
                     ByVal nodeIds As Variant, Optional ByVal STYPE As Long = 2, _
                     Optional ByVal WALL_ID As Long = 1, Optional ByVal W_TYPE As Long = 0, _
                     Optional ByVal GROUP As String = "") As String
    Dim rec As Object
    Set rec = CvElemRecord("WALL", matlId, sectId, nodeIds)
    rec.Add "STYPE", STYPE
    rec.Add "WALL", WALL_ID
    rec.Add "W_TYPE", W_TYPE
    rec.Add "W_CON", 0
    Wall = CvElemPut(elemId, rec, GROUP)
End Function

Private Function CvElemRecord(ByVal pType As String, ByVal pMatl As Long, ByVal pSect As Long, _
                              ByVal pNodes As Variant) As Object
    Dim rec As Object
    Dim vNodes As Variant
    Dim col As Collection
    Dim i As Long

    vNodes = CvIds(pNodes)
    Set col = New Collection
    For i = LBound(vNodes) To UBound(vNodes)
        col.Add CLng(vNodes(i))
    Next i
    Set rec = New Dictionary
    rec.Add "TYPE", pType
    rec.Add "MATL", pMatl
    rec.Add "SECT", pSect
    rec.Add "NODE", col
    Set CvElemRecord = rec
End Function

Private Function CvElemPut(ByVal pElemId As Long, ByVal pRec As Object, ByVal pGroup As String) As String
    Dim vNodes() As Variant
    Dim v As Variant
    Dim i As Long

    StorePut "ELEM", pElemId, pRec
    If Len(pGroup) > 0 Then
        ReDim vNodes(0 To pRec.Item("NODE").Count - 1)
        i = 0
        For Each v In pRec.Item("NODE")
            vNodes(i) = v
            i = i + 1
        Next v
        CvGroupJoin pGroup, vNodes, pElemId
    End If
    mLastStatus = 200
    mLastError = ""
    CvElemPut = ""
End Function

' Beam stiffness scale factors (ESSF) - one item per call, appended per element.
'   StiffnessScaleFactor Array(1, 2), AREA_SF:=0.5, IYY_SF:=0.8
Public Function StiffnessScaleFactor(ByVal elemIds As Variant, Optional ByVal AREA_SF As Double = 1, _
                                     Optional ByVal ASY_SF As Double = 1, Optional ByVal ASZ_SF As Double = 1, _
                                     Optional ByVal IXX_SF As Double = 1, Optional ByVal IYY_SF As Double = 1, _
                                     Optional ByVal IZZ_SF As Double = 1, Optional ByVal WGT_SF As Double = 1, _
                                     Optional ByVal GROUP_NAME As String = "") As String
    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "AREA_SF", AREA_SF
    item.Add "ASY_SF", ASY_SF
    item.Add "ASZ_SF", ASZ_SF
    item.Add "IXX_SF", IXX_SF
    item.Add "IYY_SF", IYY_SF
    item.Add "IZZ_SF", IZZ_SF
    item.Add "WGT_SF", WGT_SF
    item.Add "GROUP_NAME", GROUP_NAME
    StiffnessScaleFactor = CvItemsPut("db/ESSF", elemIds, item, True)
End Function

' Wall stiffness scale factors (WSSF). The out of plane factors are written
' only when given (field names from the API schema).
'   WallScaleFactor 30, 0.7, 0.5
Public Function WallScaleFactor(ByVal elemIds As Variant, Optional ByVal SHEAR As Double = 1, _
                                Optional ByVal BENDING As Double = 1, Optional ByVal AXIAL As Variant = Null, _
                                Optional ByVal OUT_TORSION As Variant = Null, Optional ByVal OUT_SHEAR As Variant = Null, _
                                Optional ByVal OUT_BENDING As Variant = Null, _
                                Optional ByVal GROUP_NAME As String = "") As String
    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "SHEAR", SHEAR
    item.Add "BENDING", BENDING
    If Not IsNull(AXIAL) Then item.Add "AXIAL", AXIAL
    If Not IsNull(OUT_TORSION) Then item.Add "OUT_TORSION", OUT_TORSION
    If Not IsNull(OUT_SHEAR) Then item.Add "OUT_SHEAR", OUT_SHEAR
    If Not IsNull(OUT_BENDING) Then item.Add "OUT_BENDING", OUT_BENDING
    item.Add "GROUP_NAME", GROUP_NAME
    WallScaleFactor = CvItemsPut("db/WSSF", elemIds, item, True)
End Function

' Node local axis (SKEW).
'   NodeLocalAxis 5, "Z", 30                       ' about one axis - keeps the other two angles
'   NodeLocalAxis 7, "XYZ", Array(0, 0, 45)
'   NodeLocalAxis 9, "Vector", Array(1, 0, 0, 0, 0, 1)   ' local X then local Y
Public Function NodeLocalAxis(ByVal nodeIds As Variant, ByVal METHOD As String, ByVal VALUES As Variant) As String
    Dim vIds As Variant
    Dim vVal As Variant
    Dim rec As Object
    Dim old As Object
    Dim i As Long
    Dim sAxis As String

    vIds = CvIds(nodeIds)
    sAxis = UCase$(Trim$(METHOD))
    For i = LBound(vIds) To UBound(vIds)
        Set rec = New Dictionary
        If sAxis = "VECTOR" Then
            vVal = CvNums(VALUES, 6)
            rec.Add "iMETHOD", 3
            rec.Add "V1X", vVal(0): rec.Add "V1Y", vVal(1): rec.Add "V1Z", vVal(2)
            rec.Add "V2X", vVal(3): rec.Add "V2Y", vVal(4): rec.Add "V2Z", vVal(5)
        Else
            rec.Add "iMETHOD", 1
            rec.Add "ANGLE_X", 0
            rec.Add "ANGLE_Y", 0
            rec.Add "ANGLE_Z", 0
            Set old = StoreGet("SKEW", vIds(i))
            If Not old Is Nothing Then
                If old.Item("iMETHOD") = 1 Then
                    rec.Item("ANGLE_X") = old.Item("ANGLE_X")
                    rec.Item("ANGLE_Y") = old.Item("ANGLE_Y")
                    rec.Item("ANGLE_Z") = old.Item("ANGLE_Z")
                End If
            End If
            Select Case sAxis
                Case "X": rec.Item("ANGLE_X") = CDbl(VALUES)
                Case "Y": rec.Item("ANGLE_Y") = CDbl(VALUES)
                Case "Z": rec.Item("ANGLE_Z") = CDbl(VALUES)
                Case Else
                    vVal = CvNums(VALUES, 3)
                    rec.Item("ANGLE_X") = vVal(0)
                    rec.Item("ANGLE_Y") = vVal(1)
                    rec.Item("ANGLE_Z") = vVal(2)
            End Select
        End If
        StorePut "SKEW", vIds(i), rec
    Next i
    mLastStatus = 200
    mLastError = ""
    NodeLocalAxis = ""
End Function


'==========================================================
' [13] Helpers - more loads and temperatures
'==========================================================
'  Specified displacement, nodal mass, load to mass, floor load, plane
'  load, pre-composite section, nodal / beam section temperature.
'  Ids (floor / plane load, load to mass) are given first here, as everywhere else in this module.
'==========================================================

' Specified displacement (SDSP). VALUES = Dx Dy Dz Rx Ry Rz - a component
' that is not 0 is switched on.
'   SpecifiedDisp Array(1, 2), "SD", Array(0, 0, -0.01, 0, 0, 0)
Public Function SpecifiedDisp(ByVal nodeIds As Variant, ByVal LCNAME As String, ByVal VALUES As Variant, _
                              Optional ByVal GROUP_NAME As String = "") As String
    Dim item As Object
    Dim one As Object
    Dim col As Collection
    Dim v As Variant
    Dim i As Long

    v = CvNums(VALUES, 6)
    Set col = New Collection
    For i = 0 To 5
        Set one = New Dictionary
        one.Add "OPT_FLAG", (v(i) <> 0)
        one.Add "DISPLACEMENT", v(i)
        col.Add one
    Next i
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "VALUES", col
    SpecifiedDisp = CvItemsPut("db/SDSP", nodeIds, item, True)
End Function

' Nodal mass (NMAS).
'   NodalMass Array(1, 2), 1.5, 1.5, 0.5
Public Function NodalMass(ByVal nodeIds As Variant, Optional ByVal mX As Double = 0, Optional ByVal mY As Double = 0, _
                          Optional ByVal mZ As Double = 0, Optional ByVal rmX As Double = 0, _
                          Optional ByVal rmY As Double = 0, Optional ByVal rmZ As Double = 0) As String
    Dim vIds As Variant
    Dim rec As Object
    Dim i As Long

    vIds = CvIds(nodeIds)
    For i = LBound(vIds) To UBound(vIds)
        Set rec = New Dictionary
        rec.Add "mX", mX
        rec.Add "mY", mY
        rec.Add "mZ", mZ
        rec.Add "rmX", rmX
        rec.Add "rmY", rmY
        rec.Add "rmZ", rmZ
        StorePut "NMAS", vIds(i), rec
    Next i
    mLastStatus = 200
    NodalMass = ""
End Function

' Loads to masses (LTOM). DIR X / Y / Z / XY / YZ / XZ / XYZ. Factors default 1.
'   LoadToMass 1, "XY", Array("DL", "SDL"), Array(1, 1)
Public Function LoadToMass(ByVal ltomId As Long, ByVal DIR As String, ByVal lcNames As Variant, _
                           Optional ByVal factors As Variant = Null, Optional ByVal bNODAL As Boolean = True, _
                           Optional ByVal bBEAM As Boolean = True, Optional ByVal bFLOOR As Boolean = True, _
                           Optional ByVal bPRES As Boolean = True, Optional ByVal GRAV As Double = 9.806) As String
    Dim rec As Object
    Dim one As Object
    Dim col As Collection
    Dim vNames As Variant
    Dim i As Long

    vNames = CvIds(lcNames)
    Set col = New Collection
    For i = LBound(vNames) To UBound(vNames)
        Set one = New Dictionary
        one.Add "LCNAME", CStr(vNames(i))
        one.Add "FACTOR", CDbl(CvPick(factors, i - LBound(vNames), 1#))
        col.Add one
    Next i
    Set rec = New Dictionary
    rec.Add "DIR", UCase$(DIR)
    rec.Add "bNODAL", bNODAL
    rec.Add "bBEAM", bBEAM
    rec.Add "bFLOOR", bFLOOR
    rec.Add "bPRES", bPRES
    rec.Add "GRAV", GRAV
    rec.Add "vLC", col
    StorePut "LTOM", ltomId, rec
    mLastStatus = 200
    LoadToMass = ""
End Function

' Floor load type (FBLD) - one load per load case.
'   FloorLoadDefine 1, "Office", Array("DL", "LL"), Array(-5, -3)
Public Function FloorLoadDefine(ByVal fbldId As Long, ByVal NAME As String, ByVal lcNames As Variant, _
                                ByVal loads As Variant, Optional ByVal DESC As String = "", _
                                Optional ByVal bSUB_BEAM_WEIGHT As Variant = False) As String
    Dim rec As Object
    Dim one As Object
    Dim col As Collection
    Dim vNames As Variant
    Dim i As Long

    vNames = CvIds(lcNames)
    Set col = New Collection
    For i = LBound(vNames) To UBound(vNames)
        Set one = New Dictionary
        one.Add "LCNAME", CStr(vNames(i))
        one.Add "FLOOR_LOAD", CDbl(CvPick(loads, i - LBound(vNames), 0#))
        one.Add "OPT_SUB_BEAM_WEIGHT", CBool(CvPick(bSUB_BEAM_WEIGHT, i - LBound(vNames), False))
        col.Add one
    Next i
    Set rec = New Dictionary
    rec.Add "NAME", NAME
    rec.Add "DESC", DESC
    rec.Add "ITEM", col
    StorePut "FBLD", fbldId, rec
    mLastStatus = 200
    FloorLoadDefine = ""
End Function

' Floor load assignment (FBLA). DIST_TYPE 1 one way, 2 two way, 3 polygon
' centroid, 4 polygon length.
'   FloorLoadAssign 1, "Office", Array(1, 2, 3, 4)
Public Function FloorLoadAssign(ByVal fblaId As Long, ByVal FLOOR_NAME As String, ByVal nodeIds As Variant, _
                                Optional ByVal DIST_TYPE As Long = 2, Optional ByVal DIR As String = "GZ", _
                                Optional ByVal GROUP_NAME As String = "", Optional ByVal LOAD_ANGLE As Double = 0, _
                                Optional ByVal SUB_BEAM_NUM As Long = 0, Optional ByVal SUB_BEAM_ANGLE As Double = 0, _
                                Optional ByVal UNIT_SELF_WEIGHT As Double = 0, Optional ByVal bPROJECTION As Boolean = False, _
                                Optional ByVal bEXCLUDE_INNER As Boolean = False, _
                                Optional ByVal bALLOW_POLYGON As Boolean = False) As String
    Dim rec As Object
    Dim col As Collection
    Dim vIds As Variant
    Dim i As Long

    vIds = CvIds(nodeIds)
    Set col = New Collection
    For i = LBound(vIds) To UBound(vIds)
        col.Add CLng(vIds(i))
    Next i
    Set rec = New Dictionary
    rec.Add "FLOOR_LOAD_TYPE_NAME", FLOOR_NAME
    rec.Add "FLOOR_DIST_TYPE", DIST_TYPE
    rec.Add "DIR", UCase$(DIR)
    rec.Add "OPT_PROJECTION", bPROJECTION
    rec.Add "DESC", ""
    rec.Add "GROUP_NAME", GROUP_NAME
    rec.Add "NODES", col
    If DIST_TYPE = 1 Or DIST_TYPE = 2 Then
        rec.Add "SUB_BEAM_NUM", SUB_BEAM_NUM
        rec.Add "SUB_BEAM_ANGLE", SUB_BEAM_ANGLE
        rec.Add "UNIT_SELF_WEIGHT", UNIT_SELF_WEIGHT
        rec.Add "OPT_EXCLUDE_INNER_ELEM_AREA", bEXCLUDE_INNER
    End If
    If DIST_TYPE = 1 Then rec.Add "LOAD_ANGLE", LOAD_ANGLE
    If DIST_TYPE = 2 Then rec.Add "OPT_ALLOW_POLYGON_TYPE_UNIT_AREA", bALLOW_POLYGON
    StorePut "FBLA", fblaId, rec
    mLastStatus = 200
    FloorLoadAssign = ""
End Function

' Plane load type (PNLD). LTYPE "POINT" / "LINE" / "AREA". X, Y, F are
' arrays of the same length: POINT any number of points, LINE 2, AREA 3 or 4.
'   PlaneLoadDefine 1, "Wheel", "POINT", Array(0), Array(0), Array(10)
'   PlaneLoadDefine 2, "Strip", "LINE", Array(0, 2), Array(0, 0), Array(10, 20), COPY_X:=Array(3, 3)
Public Function PlaneLoadDefine(ByVal pnldId As Long, ByVal NAME As String, ByVal LTYPE As String, _
                                ByVal X As Variant, ByVal Y As Variant, ByVal F As Variant, _
                                Optional ByVal COPY_X As Variant = Null, Optional ByVal COPY_Y As Variant = Null, _
                                Optional ByVal DESC As String = "") As String
    Dim rec As Object
    Dim one As Object
    Dim col As Collection
    Dim vX As Variant
    Dim vY As Variant
    Dim vF As Variant
    Dim i As Long
    Dim n As Long

    vX = CvIds(X)
    vY = CvIds(Y)
    vF = CvIds(F)
    n = UBound(vX) - LBound(vX) + 1
    Set rec = New Dictionary
    rec.Add "NAME", NAME
    rec.Add "DESC", DESC
    rec.Add "LTYPE", UCase$(LTYPE)
    rec.Add "COPY_X", CvNumList(COPY_X)
    rec.Add "COPY_Y", CvNumList(COPY_Y)
    rec.Add "SEQ", pnldId
    Select Case UCase$(LTYPE)
        Case "POINT"
            Set col = New Collection
            For i = 0 To n - 1
                Set one = New Dictionary
                one.Add "X", vX(LBound(vX) + i)
                one.Add "Y", vY(LBound(vY) + i)
                one.Add "F", vF(LBound(vF) + i)
                col.Add one
            Next i
            rec.Add "POINTLOAD", col
        Case "LINE"
            Set one = New Dictionary
            one.Add "bUNIFORM", False
            one.Add "X", CvNumList(X)
            one.Add "Y", CvNumList(Y)
            one.Add "F", CvNumList(F)
            rec.Add "LINELOAD", one
        Case Else
            Set one = New Dictionary
            one.Add "bUNIFORM", False
            one.Add "b3PNT", (n = 3)
            one.Add "X", CvNumList(X)
            one.Add "Y", CvNumList(Y)
            one.Add "LOAD", CvNumList(F)
            rec.Add "AREALOAD", one
    End Select
    StorePut "PNLD", pnldId, rec
    mLastStatus = 200
    PlaneLoadDefine = ""
End Function

' Plane load assignment (PNLA) - the loading plane through ORIGIN with the
' given X axis and a point (direction) in the XY plane.
'   PlaneLoadAssign 1, "DL", 1
Public Function PlaneLoadAssign(ByVal pnlaId As Long, ByVal LCNAME As String, ByVal PNLD_KEY As Long, _
                                Optional ByVal LOAD_GROUP As String = "", Optional ByVal ORIGIN As Variant = Null, _
                                Optional ByVal AXIS_X As Variant = Null, Optional ByVal AXIS_Y As Variant = Null, _
                                Optional ByVal TOL As Double = 0.001, Optional ByVal SELECT_TYPE As String = "ON_PLANE", _
                                Optional ByVal LOAD_DIR As String = "NORMAL_PLANE", _
                                Optional ByVal PROJECT_TYPE As String = "NO", Optional ByVal DESC As String = "") As String
    Dim rec As Object
    Set rec = New Dictionary
    rec.Add "LCNAME", LCNAME
    rec.Add "LOAD_GROUP", LOAD_GROUP
    rec.Add "PNLD_KEY", PNLD_KEY
    rec.Add "ELEM_TYPE", "PLATE"
    If IsNull(ORIGIN) Then rec.Add "POINT_ORIGIN", Array(0, 0, 0) Else rec.Add "POINT_ORIGIN", CvNums(ORIGIN, 3)
    If IsNull(AXIS_X) Then rec.Add "AXIS_X", Array(1, 0, 0) Else rec.Add "AXIS_X", CvNums(AXIS_X, 3)
    If IsNull(AXIS_Y) Then rec.Add "AXIS_Y", Array(0, 1, 0) Else rec.Add "AXIS_Y", CvNums(AXIS_Y, 3)
    rec.Add "TOL", TOL
    rec.Add "SELECT_TYPE", SELECT_TYPE
    rec.Add "LOAD_DIR", LOAD_DIR
    rec.Add "PROJECT_TYPE", PROJECT_TYPE
    rec.Add "DESC", DESC
    StorePut "PNLA", pnlaId, rec
    mLastStatus = 200
    PlaneLoadAssign = ""
End Function

' Load cases applied before the composite section is active (PLCB).
' Several calls add to the same list.
'   PreCompositeLoad Array("DL", "SDL")
Public Function PreCompositeLoad(ByVal lcNames As Variant) As String
    Dim rec As Object
    Dim col As Collection
    Dim vNames As Variant
    Dim v As Variant
    Dim w As Variant
    Dim i As Long
    Dim found As Boolean

    Set rec = StoreGet("PLCB", 1)
    If rec Is Nothing Then
        Set rec = New Dictionary
        rec.Add "LCNAME_ITEM", New Collection
        StorePut "PLCB", 1, rec
    End If
    Set col = rec.Item("LCNAME_ITEM")
    vNames = CvIds(lcNames)
    For i = LBound(vNames) To UBound(vNames)
        v = CStr(vNames(i))
        found = False
        For Each w In col
            If w = v Then found = True
        Next w
        If Not found Then col.Add v
    Next i
    mLastStatus = 200
    PreCompositeLoad = ""
End Function

' Nodal temperature (NTMP).
'   NodalTemp Array(1, 2), "TU", 25
Public Function NodalTemp(ByVal nodeIds As Variant, ByVal LCNAME As String, ByVal TEMPER As Double, _
                          Optional ByVal GROUP_NAME As String = "") As String
    Dim item As Object
    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "TEMPER", TEMPER
    NodalTemp = CvItemsPut("db/NTMP", nodeIds, item, True)
End Function

' Beam section temperature (BTMP). VALUES: one row Array(h1, h2, t1, t2) or
' several rows Array(Array(...), Array(...)). bPSC rows may use "Z1" "Z2" "Z3"
' for h1 / h2. TYPE "Element" (material of the element) or "Input" (ELAST,
' THERMAL given). REF "Centroid" / "Top" / "Bot".
'   BeamSectionTemp 3, "TB", Array(0.1, 0.2, 3, 12.4)
'   BeamSectionTemp 3, "TB", Array(Array("Z1", "Z2", 17.8, 4), Array(0.15, 0.4, 4, 0)), bPSC:=True
Public Function BeamSectionTemp(ByVal elemIds As Variant, ByVal LCNAME As String, ByVal VALUES As Variant, _
                                Optional ByVal bPSC As Boolean = False, Optional ByVal TYPE_ As String = "Element", _
                                Optional ByVal DIR As String = "LZ", Optional ByVal REF As String = "Centroid", _
                                Optional ByVal VAL_B As Double = 0, Optional ByVal ELAST As Variant = Null, _
                                Optional ByVal THERMAL As Variant = Null, Optional ByVal GROUP_NAME As String = "") As String
    Dim item As Object
    Dim row As Object
    Dim col As Collection
    Dim vRows As Variant
    Dim r As Variant
    Dim i As Long

    ' one flat row -> a list of one row
    If IsArray(VALUES(LBound(VALUES))) Then vRows = VALUES Else vRows = Array(VALUES)
    Set col = New Collection
    For i = LBound(vRows) To UBound(vRows)
        r = vRows(i)
        Set row = New Dictionary
        row.Add "TYPE", UCase$(TYPE_)
        If Not IsNull(ELAST) Then row.Add "ELAST", ELAST
        If Not IsNull(THERMAL) Then row.Add "THERMAL", THERMAL
        If bPSC Then
            If UCase$(REF) = "BOT" Then row.Add "REF", 1 Else row.Add "REF", 0
            row.Add "OPT_B", CLng(VAL_B)
            row.Add "VAL_B", 0
            CvTempH row, "1", CvPick(r, 0, 0)
            CvTempH row, "2", CvPick(r, 1, 0)
        Else
            row.Add "VAL_B", VAL_B
            row.Add "VAL_H1", CDbl(CvPick(r, 0, 0))
            row.Add "VAL_H2", CDbl(CvPick(r, 1, 0))
        End If
        row.Add "VAL_T1", CDbl(CvPick(r, 2, 0))
        row.Add "VAL_T2", CDbl(CvPick(r, 3, 0))
        col.Add row
    Next i

    Set item = New Dictionary
    item.Add "ID", 0
    item.Add "LCNAME", LCNAME
    item.Add "GROUP_NAME", GROUP_NAME
    item.Add "DIR", UCase$(DIR)
    item.Add "REF", REF
    item.Add "NUM", col.Count
    item.Add "bPSC", bPSC
    item.Add "vSECTTMP", col
    BeamSectionTemp = CvItemsPut("db/BTMP", elemIds, item, True)
End Function

' PSC height: "Z1" "Z2" "Z3" -> OPT 0 / 1 / 2 and value 0, a number -> OPT 3 and the value.
Private Sub CvTempH(ByVal pRow As Object, ByVal pWhich As String, ByVal pValue As Variant)
    Dim s As String
    s = UCase$(Trim$(CStr(pValue)))
    If s = "Z1" Or s = "Z2" Or s = "Z3" Then
        pRow.Add "OPT_H" & pWhich, CLng(Mid$(s, 2)) - 1
        pRow.Add "VAL_H" & pWhich, 0
    Else
        pRow.Add "OPT_H" & pWhich, 3
        pRow.Add "VAL_H" & pWhich, CDbl(pValue)
    End If
End Sub

' A list of numbers as a Collection (empty when Null / missing).
Private Function CvNumList(ByVal pAny As Variant) As Collection
    Dim col As Collection
    Dim v As Variant
    Dim i As Long

    Set col = New Collection
    If Not IsMissing(pAny) Then
        If Not IsNull(pAny) Then
            v = CvIds(pAny)
            For i = LBound(v) To UBound(v)
                col.Add v(i)
            Next i
        End If
    End If
    Set CvNumList = col
End Function


'==========================================================
' [14] Sections - value / PSC value / composite / tapered, offset, tapered group
'==========================================================
'  Value, PSC value, composite and tapered sections, tapered groups, offset:
'    - every function here takes offsetPt / useShearDeform / useWarping
'      like the other section functions; SectionOffset sets the rest of
'      the offset (user distances, J end) on a section already defined
'    - polygons (PSC value sections) are a list of points:
'        Array(Array(0, 0), Array(1, 0), Array(1, 1), Array(0, 1))
'      or a 2-column worksheet range. Do not repeat the first point.
'      The outer polygon is turned counter-clockwise and the holes
'      clockwise. Several holes: Array(hole1, hole2).
'    - PSC I shaped composites: aSizes = H1 HL1 HL2 HL21 HL22 HL3 HL4 HL41 HL42 HL5,
'      bSizes = BL1 BL2 BL21 BL22 BL4 BL41 BL42, cSizes / dSizes the right side
'      (HR.. / BR..), jointFlags = J1 JL1..JL4 JR1..JR4.  bSymm (default)
'      copies the left side to the right side.
'    - material ratios of composites: girder (or steel) / slab
'      E, density, Poisson girder, Poisson slab, thermal
'==========================================================

'   SectionValue 1, "V1", "SB", Array(0.5, 0.3), 0.15, 0.002, 0.003, 0.001
' Stiffness values left out are calculated by NX from the shape.
Public Function SectionValue(ByVal sectId As Long, ByVal NAME As String, ByVal SHAPE As String, _
                             ByVal vSize As Variant, _
                             Optional ByVal AREA As Variant, Optional ByVal IXX As Variant, _
                             Optional ByVal IYY As Variant, Optional ByVal IZZ As Variant, _
                             Optional ByVal offsetPt As String = "CC", _
                             Optional ByVal useShearDeform As Boolean = True, _
                             Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object
    Dim oSect As Object
    Dim oStiff As Object

    Set oStiff = New Dictionary
    If Not IsMissing(AREA) Then oStiff.Add "AREA", AREA
    If Not IsMissing(IXX) Then oStiff.Add "RXX", IXX
    If Not IsMissing(IYY) Then oStiff.Add "RYY", IYY
    If Not IsMissing(IZZ) Then oStiff.Add "RZZ", IZZ
    Set oSect = New Dictionary
    oSect.Add "vSIZE", CvSizes(vSize)
    oSect.Add "STIFF", oStiff

    Set oItem = CvSectItem("VALUE", NAME)
    oItem.Add "CALC_OPT", True
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", UCase$(Trim$(SHAPE))
    oBefore.Add "SECT_I", oSect
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    SectionValue = CvSectPut(sectId, oItem)
End Function

'   SectionPscValue 1, "P1", Array(Array(0, 0), Array(2, 0), Array(2, 1.5), Array(0, 1.5))
' T1 T2 BT HT: design dimensions, Z1 Z2 Z3: shear check positions.
Public Function SectionPscValue(ByVal sectId As Long, ByVal NAME As String, ByVal outerPts As Variant, _
                                Optional ByVal innerPts As Variant, _
                                Optional ByVal T1 As Double = 0.1, Optional ByVal T2 As Double = 0.1, _
                                Optional ByVal BT As Double = 0.1, Optional ByVal HT As Double = 0.1, _
                                Optional ByVal Z1 As Double = 0, Optional ByVal Z2 As Double = 0, _
                                Optional ByVal Z3 As Double = 0, Optional ByVal thkTorsion As Double = 0, _
                                Optional ByVal offsetPt As String = "CC", _
                                Optional ByVal useShearDeform As Boolean = True, _
                                Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object
    Dim oSect As Object

    Set oSect = New Dictionary
    oSect.Add "SECT_NAME", ""
    oSect.Add "vSIZE", Array(HT, BT, T1, T2)
    CvPolygonsAdd oSect, outerPts, innerPts

    Set oItem = CvSectItem("PSC", NAME)
    oItem.Add "CALC_OPT", True
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", "VALU"
    oBefore.Add "SECT_I", oSect
    oBefore.Add "SHEAR_CHK", True
    oBefore.Add "SHEAR_CHK_POS", Array(Array(Z1, Z2, Z3), Array(0, 0, 0))
    oBefore.Add "USE_AUTO_QY", Array(Array(True, True, True), Array(False, False, False))
    oBefore.Add "WEB_THICK", Array(thkTorsion, 0)
    oBefore.Add "USE_WEB_THICK_SHEAR", Array(Array(True, True, True), Array(False, False, False))
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    SectionPscValue = CvSectPut(sectId, oItem)
End Function

'   SectionCompositeI 1, "CI", 2.5, 0.25, 0.05, 1.8, 0.5, 0.03, 0.016, 0.6, 0.04, 7, 3.2, 0.3, 0.2, 1.2
' Steel I girder + slab (Type 1). Bc tc Hh: slab width, slab thickness, haunch.
Public Function SectionCompositeI(ByVal sectId As Long, ByVal NAME As String, _
                                  ByVal Bc As Double, ByVal tc As Double, ByVal Hh As Double, _
                                  ByVal Hw As Double, ByVal B1 As Double, ByVal tf1 As Double, _
                                  ByVal tw As Double, ByVal B2 As Double, ByVal tf2 As Double, _
                                  Optional ByVal EsEc As Double = 0, Optional ByVal DsDc As Double = 0, _
                                  Optional ByVal Ps As Double = 0, Optional ByVal Pc As Double = 0, _
                                  Optional ByVal TsTc As Double = 0, _
                                  Optional ByVal MultiModulus As Boolean = False, _
                                  Optional ByVal CreepEratio As Double = 0, _
                                  Optional ByVal ShrinkEratio As Double = 0, _
                                  Optional ByVal offsetPt As String = "CC", _
                                  Optional ByVal useShearDeform As Boolean = True, _
                                  Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object

    Set oItem = CvSectItem("COMPOSITE", NAME)
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", "I"
    oBefore.Add "SECT_I", CvOneKey("vSIZE", Array(Hw, tw, B1, tf1, B2, tf2))
    CvCompMatl oBefore, EsEc, DsDc, Ps, Pc, TsTc, MultiModulus, CreepEratio, ShrinkEratio
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    oItem.Add "SECT_AFTER", CvOneKey("SLAB", Array(Bc, tc, Hh))
    SectionCompositeI = CvSectPut(sectId, oItem)
End Function

' Steel tub girder + slab (Type 1).
Public Function SectionCompositeTub(ByVal sectId As Long, ByVal NAME As String, _
                                    ByVal Bc As Double, ByVal tc As Double, ByVal Hh As Double, _
                                    ByVal Hw As Double, ByVal B1 As Double, ByVal Bf1 As Double, _
                                    ByVal tf1 As Double, ByVal Bf3 As Double, ByVal tw As Double, _
                                    ByVal B2 As Double, ByVal Bf2 As Double, ByVal tf2 As Double, _
                                    ByVal tfp As Double, _
                                    Optional ByVal EsEc As Double = 0, Optional ByVal DsDc As Double = 0, _
                                    Optional ByVal Ps As Double = 0, Optional ByVal Pc As Double = 0, _
                                    Optional ByVal TsTc As Double = 0, _
                                    Optional ByVal MultiModulus As Boolean = False, _
                                    Optional ByVal CreepEratio As Double = 0, _
                                    Optional ByVal ShrinkEratio As Double = 0, _
                                    Optional ByVal offsetPt As String = "CC", _
                                    Optional ByVal useShearDeform As Boolean = True, _
                                    Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object

    Set oItem = CvSectItem("COMPOSITE", NAME)
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", "Tub"
    oBefore.Add "SECT_I", CvOneKey("vSIZE", CvTubSize(Array(Hw, B1, Bf1, tf1, Bf3, tw, B2, Bf2, tf2, tfp)))
    CvCompMatl oBefore, EsEc, DsDc, Ps, Pc, TsTc, MultiModulus, CreepEratio, ShrinkEratio
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    oItem.Add "SECT_AFTER", CvOneKey("SLAB", Array(Bc, tc, Hh))
    SectionCompositeTub = CvSectPut(sectId, oItem)
End Function

'   SectionCompositePscI 1, "CP", 2, 0.2, 0.05, Array(1.8, 0.2), Array(0.3)
' PSC I girder + slab. Sizes and joints as in the head of this section.
Public Function SectionCompositePscI(ByVal sectId As Long, ByVal NAME As String, _
                                     ByVal Bc As Double, ByVal tc As Double, ByVal Hh As Double, _
                                     ByVal aSizes As Variant, ByVal bSizes As Variant, _
                                     Optional ByVal cSizes As Variant, Optional ByVal dSizes As Variant, _
                                     Optional ByVal bSymm As Boolean = True, _
                                     Optional ByVal jointFlags As Variant, _
                                     Optional ByVal EgdEsb As Double = 0, Optional ByVal DgdDsb As Double = 0, _
                                     Optional ByVal Pgd As Double = 0, Optional ByVal Psb As Double = 0, _
                                     Optional ByVal TgdTsb As Double = 0, _
                                     Optional ByVal MultiModulus As Boolean = False, _
                                     Optional ByVal CreepEratio As Double = 0, _
                                     Optional ByVal ShrinkEratio As Double = 0, _
                                     Optional ByVal offsetPt As String = "CC", _
                                     Optional ByVal useShearDeform As Boolean = True, _
                                     Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object

    Set oItem = CvSectItem("COMPOSITE", NAME)
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", "CI"
    oBefore.Add "SECT_I", CvPscISizes(bSymm, aSizes, bSizes, cSizes, dSizes)
    PscCheckFields oBefore
    oBefore.Add "JOINT", CvPscIJoints(bSymm, jointFlags)
    oBefore.Add "USE_SYMMETRIC", bSymm
    CvCompMatl oBefore, EgdEsb, DgdDsb, Pgd, Psb, TgdTsb, MultiModulus, CreepEratio, ShrinkEratio
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    oItem.Add "SECT_AFTER", CvOneKey("SLAB", Array(Bc, tc, Hh))
    SectionCompositePscI = CvSectPut(sectId, oItem)
End Function

' PSC polygon girder + slab.
Public Function SectionCompositePscValue(ByVal sectId As Long, ByVal NAME As String, _
                                         ByVal Bc As Double, ByVal tc As Double, ByVal Hh As Double, _
                                         ByVal outerPts As Variant, Optional ByVal innerPts As Variant, _
                                         Optional ByVal EgEs As Double = 1, Optional ByVal DgDs As Double = 1, _
                                         Optional ByVal Pg As Double = 0.2, Optional ByVal Ps As Double = 0.2, _
                                         Optional ByVal TgTs As Double = 1, _
                                         Optional ByVal MultiModulus As Boolean = False, _
                                         Optional ByVal CreepEratio As Double = 0, _
                                         Optional ByVal ShrinkEratio As Double = 0, _
                                         Optional ByVal offsetPt As String = "CC", _
                                         Optional ByVal useShearDeform As Boolean = True, _
                                         Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object
    Dim oSect As Object
    Dim oAfter As Object
    Dim oAfterI As Object

    Set oSect = CvOneKey("vSIZE", Array(0.1, 0.1, 0.1, 0.1))
    CvPolygonsAdd oSect, outerPts, innerPts

    Set oItem = CvSectItem("COMPOSITE", NAME)
    oItem.Add "CALC_OPT", True
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", "PC"
    oBefore.Add "SECT_I", oSect
    oBefore.Add "SHEAR_CHK", True
    oBefore.Add "SHEAR_CHK_POS", Array(Array(0.1, 0, 0.1), Array(0, 0, 0))
    oBefore.Add "USE_AUTO_QY", Array(Array(True, True, True), Array(False, False, False))
    oBefore.Add "WEB_THICK", Array(0, 0)
    oBefore.Add "USE_WEB_THICK_SHEAR", Array(Array(True, True, True), Array(False, False, False))
    CvCompMatl oBefore, EgEs, DgDs, Pg, Ps, TgTs, MultiModulus, CreepEratio, ShrinkEratio
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping

    Set oAfterI = CvOneKey("vSIZE", Array(Bc, Hh))
    oAfterI.Add "BUILT_FLAG", 1
    Set oAfter = New Dictionary
    oAfter.Add "SECT_I", oAfterI
    oAfter.Add "SECT_J", CvOneKey("vSIZE", Array(Bc, tc, Hh))
    oItem.Add "SECT_AFTER", oAfter
    SectionCompositePscValue = CvSectPut(sectId, oItem)
End Function

'   SectionTaperedUser 1, "TD", "SB", Array(0.5, 0.3), Array(0.8, 0.3)
' A standard shape whose dimensions go from sizeI at the I end to sizeJ.
Public Function SectionTaperedUser(ByVal sectId As Long, ByVal NAME As String, ByVal SHAPE As String, _
                                   ByVal sizeI As Variant, ByVal sizeJ As Variant, _
                                   Optional ByVal offsetPt As String = "CC", _
                                   Optional ByVal useShearDeform As Boolean = True, _
                                   Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object

    Set oItem = CvSectItem("TAPERED", NAME)
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", UCase$(Trim$(SHAPE))
    oBefore.Add "TYPE", 2
    oBefore.Add "SECT_I", CvOneKey("vSIZE", CvSizes(sizeI))
    oBefore.Add "SECT_J", CvOneKey("vSIZE", CvSizes(sizeJ))
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    SectionTaperedUser = CvSectPut(sectId, oItem)
End Function

'   SectionTaperedPscCell 1, "TP", "1CEL", Array(0.2), Array(1.5), , , Array(0.3), Array(1.5)
' PSC 1 / 2 cell box, I end and J end. Groups as SectionPsc1Cell:
' A = HO1 HO2 HO21 HO22 HO3 HO31, B = BO1 BO11 BO12 BO2 BO21 BO3,
' C = HI1 HI2 HI21 HI22 HI3 HI31 HI4 HI41 HI42 HI5, D = BI1 BI11 BI12 BI21 BI3 BI31 BI32 BI4,
' jointFlags = JO1 JO2 JO3 JI1..JI5 (both ends).
Public Function SectionTaperedPscCell(ByVal sectId As Long, ByVal NAME As String, ByVal SHAPE As String, _
                                      ByVal aI As Variant, ByVal bI As Variant, _
                                      ByVal cI As Variant, ByVal dI As Variant, _
                                      ByVal aJ As Variant, ByVal bJ As Variant, _
                                      ByVal cJ As Variant, ByVal dJ As Variant, _
                                      Optional ByVal jointFlags As Variant, _
                                      Optional ByVal offsetPt As String = "CC", _
                                      Optional ByVal useShearDeform As Boolean = True, _
                                      Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object

    Set oItem = CvSectItem("TAPERED", NAME)
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", UCase$(Trim$(SHAPE))
    oBefore.Add "TYPE", 11
    oBefore.Add "SECT_I", CvCellEnd(aI, bI, cI, dI)
    oBefore.Add "SECT_J", CvCellEnd(aJ, bJ, cJ, dJ)
    oBefore.Add "Y_VAR", 1
    oBefore.Add "Z_VAR", 1
    oBefore.Add "WARPING_CHK_AUTO_I", True
    oBefore.Add "WARPING_CHK_AUTO_J", True
    oBefore.Add "SHEAR_CHK", False
    oBefore.Add "WARPING_CHK_POS_I", Array(CvNums(Empty, 6), CvNums(Empty, 6))
    oBefore.Add "WARPING_CHK_POS_J", Array(CvNums(Empty, 6), CvNums(Empty, 6))
    oBefore.Add "USE_WEB_THICK_SHEAR", Array(Array(True, True, True), Array(True, True, True))
    oBefore.Add "WEB_THICK_SHEAR", Array(CvNums(Empty, 3), CvNums(Empty, 3))
    oBefore.Add "USE_WEB_THICK", Array(True, True)
    oBefore.Add "WEB_THICK", Array(0, 0)
    oBefore.Add "USE_SYMMETRIC", False
    oBefore.Add "USE_SMALL_HOLE", False
    oBefore.Add "USE_USER_DEF_MESHSIZE", False
    oBefore.Add "USE_USER_INTPUT_STIFF", False
    oBefore.Add "PSC_OPT1", ""
    oBefore.Add "PSC_OPT2", ""
    oBefore.Add "JOINT", CvJoints(jointFlags, 8)
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    SectionTaperedPscCell = CvSectPut(sectId, oItem)
End Function

' Steel tub + slab, tub dimensions from sizeI to sizeJ:
' Hw B1 Bf1 tf1 Bf3 tw B2 Bf2 tf2 tfp (the order of SectionCompositeTub).
Public Function SectionTaperedCompositeTub(ByVal sectId As Long, ByVal NAME As String, _
                                           ByVal Bc As Double, ByVal tc As Double, ByVal Hh As Double, _
                                           ByVal sizeI As Variant, ByVal sizeJ As Variant, _
                                           Optional ByVal EsEc As Double = 0, Optional ByVal DsDc As Double = 0, _
                                           Optional ByVal Ps As Double = 0, Optional ByVal Pc As Double = 0, _
                                           Optional ByVal TsTc As Double = 0, _
                                           Optional ByVal MultiModulus As Boolean = False, _
                                           Optional ByVal CreepEratio As Double = 0, _
                                           Optional ByVal ShrinkEratio As Double = 0, _
                                           Optional ByVal offsetPt As String = "CC", _
                                           Optional ByVal useShearDeform As Boolean = True, _
                                           Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object

    Set oItem = CvSectItem("TAPERED", NAME)
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", "CP_T"
    oBefore.Add "TYPE", 16
    oBefore.Add "SECT_I", CvOneKey("vSIZE", CvTubSize(PscNums(sizeI, 10)))
    oBefore.Add "Y_VAR", 1
    oBefore.Add "Z_VAR", 1
    CvCompMatl oBefore, EsEc, DsDc, Ps, Pc, TsTc, MultiModulus, CreepEratio, ShrinkEratio
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    oItem.Add "SECT_AFTER", CvOneKey("SLAB", Array(Bc, tc, Hh))
    oItem.Add "COMPOSITE_J", CvOneKey("vSIZE", CvTubSize(PscNums(sizeJ, 10)))
    SectionTaperedCompositeTub = CvSectPut(sectId, oItem)
End Function

' PSC polygon section, I end and J end (same number of holes at both ends).
' dgnI / dgnJ = HT BT T1 T2, shearPosI / shearPosJ = Z1 Z2 Z3.
Public Function SectionTaperedPscValue(ByVal sectId As Long, ByVal NAME As String, _
                                       ByVal outerI As Variant, ByVal outerJ As Variant, _
                                       Optional ByVal innerI As Variant, Optional ByVal innerJ As Variant, _
                                       Optional ByVal dgnI As Variant, Optional ByVal dgnJ As Variant, _
                                       Optional ByVal shearPosI As Variant, Optional ByVal shearPosJ As Variant, _
                                       Optional ByVal thkTorsionI As Double = 0, _
                                       Optional ByVal thkTorsionJ As Double = 0, _
                                       Optional ByVal offsetPt As String = "CC", _
                                       Optional ByVal useShearDeform As Boolean = True, _
                                       Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object
    Dim oI As Object
    Dim oJ As Object

    If IsMissing(dgnI) Then dgnI = Array(0.1, 0.1, 0.1, 0.1)
    If IsMissing(dgnJ) Then dgnJ = Array(0.1, 0.1, 0.1, 0.1)
    Set oI = CvOneKey("vSIZE", PscNums(dgnI, 4))
    CvPolygonsAdd oI, outerI, innerI
    Set oJ = CvOneKey("vSIZE", PscNums(dgnJ, 4))
    CvPolygonsAdd oJ, outerJ, innerJ

    Set oItem = CvSectItem("TAPERED", NAME)
    oItem.Add "CALC_OPT", True
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", "VALU"
    oBefore.Add "SECT_I", oI
    oBefore.Add "SECT_J", oJ
    oBefore.Add "Y_VAR", 1
    oBefore.Add "Z_VAR", 1
    oBefore.Add "SHEAR_CHK", True
    oBefore.Add "SHEAR_CHK_POS", Array(PscNums(shearPosI, 3), PscNums(shearPosJ, 3))
    oBefore.Add "USE_AUTO_QY", Array(Array(True, True, True), Array(True, True, True))
    oBefore.Add "WEB_THICK", Array(thkTorsionI, thkTorsionJ)
    oBefore.Add "USE_WEB_THICK_SHEAR", Array(Array(True, True, True), Array(True, True, True))
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping
    SectionTaperedPscValue = CvSectPut(sectId, oItem)
End Function

' PSC I girder + slab, girder from the I end sizes to the J end sizes.
' bSymm copies the left side to the right side at both ends.
Public Function SectionTaperedCompositePscI(ByVal sectId As Long, ByVal NAME As String, _
                                            ByVal Bc As Double, ByVal tc As Double, ByVal Hh As Double, _
                                            ByVal aI As Variant, ByVal bI As Variant, _
                                            ByVal aJ As Variant, ByVal bJ As Variant, _
                                            Optional ByVal cI As Variant, Optional ByVal dI As Variant, _
                                            Optional ByVal cJ As Variant, Optional ByVal dJ As Variant, _
                                            Optional ByVal bSymm As Boolean = True, _
                                            Optional ByVal jointFlags As Variant, _
                                            Optional ByVal EgdEsb As Double = 0, Optional ByVal DgdDsb As Double = 0, _
                                            Optional ByVal Pgd As Double = 0, Optional ByVal Psb As Double = 0, _
                                            Optional ByVal TgdTsb As Double = 0, _
                                            Optional ByVal MultiModulus As Boolean = False, _
                                            Optional ByVal CreepEratio As Double = 0, _
                                            Optional ByVal ShrinkEratio As Double = 0, _
                                            Optional ByVal offsetPt As String = "CC", _
                                            Optional ByVal useShearDeform As Boolean = True, _
                                            Optional ByVal useWarping As Boolean = False) As String
    Dim oItem As Object
    Dim oBefore As Object
    Dim oAfter As Object

    Set oItem = CvSectItem("TAPERED", NAME)
    Set oBefore = oItem.Item("SECT_BEFORE")
    oBefore.Add "SHAPE", "CPCI"
    oBefore.Add "TYPE", 12
    oBefore.Add "SECT_I", CvPscISizes(bSymm, aI, bI, cI, dI)
    oBefore.Add "Y_VAR", 1
    oBefore.Add "Z_VAR", 1
    PscCheckFields oBefore
    oBefore.Add "JOINT", CvPscIJoints(bSymm, jointFlags)
    oBefore.Add "USE_SYMMETRIC", bSymm
    CvCompMatl oBefore, EgdEsb, DgdDsb, Pgd, Psb, TgdTsb, MultiModulus, CreepEratio, ShrinkEratio
    CvSectTail oBefore, offsetPt, useShearDeform, useWarping

    Set oAfter = CvOneKey("SLAB", Array(Bc, tc, Hh))
    oAfter.Add "SECT_I", CvOneKey("BUILT_FLAG", 1)
    oItem.Add "SECT_AFTER", oAfter
    oItem.Add "COMPOSITE_J", CvPscISizes(bSymm, aJ, bJ, cJ, dJ)
    SectionTaperedCompositePscI = CvSectPut(sectId, oItem)
End Function

'   SectionOffset 1, "LT", HOffset:=0.1, HOffOpt:=1, VOffset:=0.2, VOffOpt:=1
' The offset of a section defined above (call it after the section).
' CenterLocation 0 centroid / 1 centre of section, HOffOpt / VOffOpt
' 0 extreme fiber / 1 user, UsrOffOpt 0 centroid / 1 extreme fiber.
' HOffset_J / VOffset_J (tapered) default to the I end values.
Public Function SectionOffset(ByVal sectId As Long, Optional ByVal OffsetPoint As String = "CC", _
                              Optional ByVal CenterLocation As Long = 0, _
                              Optional ByVal HOffset As Double = 0, Optional ByVal HOffOpt As Long = 0, _
                              Optional ByVal VOffset As Double = 0, Optional ByVal VOffOpt As Long = 0, _
                              Optional ByVal UsrOffOpt As Long = 0, _
                              Optional ByVal HOffset_J As Variant, Optional ByVal VOffset_J As Variant) As String
    Dim oRec As Object
    Dim oBefore As Object

    Set oRec = StoreGet("/db/SECT", sectId)
    If oRec Is Nothing Then
        SectionOffset = "SECT " & sectId & ": define the section before its offset"
        Exit Function
    End If
    If IsMissing(HOffset_J) Then HOffset_J = HOffset
    If IsMissing(VOffset_J) Then VOffset_J = VOffset
    Set oBefore = oRec.Item("SECT_BEFORE")
    oBefore.Item("OFFSET_PT") = OffsetPoint
    oBefore.Item("OFFSET_CENTER") = CenterLocation
    oBefore.Item("USER_OFFSET_REF") = UsrOffOpt
    oBefore.Item("HORZ_OFFSET_OPT") = HOffOpt
    oBefore.Item("USERDEF_OFFSET_YI") = HOffset
    oBefore.Item("USERDEF_OFFSET_YJ") = CDbl(HOffset_J)
    oBefore.Item("VERT_OFFSET_OPT") = VOffOpt
    oBefore.Item("USERDEF_OFFSET_ZI") = VOffset
    oBefore.Item("USERDEF_OFFSET_ZJ") = CDbl(VOffset_J)
    SectionOffset = ""
End Function

'   TaperedGroup 1, "TG1", Array(1, 2, 3), "POLY", "LINEAR", 2.5
' How a tapered section varies along the elements: "LINEAR" or "POLY".
' The exponent / symmetric plane ("i" or "j") / distance are used for POLY only.
Public Function TaperedGroup(ByVal groupId As Long, ByVal NAME As String, ByVal elemIds As Variant, _
                             Optional ByVal ZVAR As String = "LINEAR", Optional ByVal YVAR As String = "LINEAR", _
                             Optional ByVal ZEXP As Double = 2, Optional ByVal ZFROM As String = "i", _
                             Optional ByVal ZDIST As Double = 0, _
                             Optional ByVal YEXP As Double = 2, Optional ByVal YFROM As String = "i", _
                             Optional ByVal YDIST As Double = 0) As String
    Dim oRec As Object
    Dim oList As Collection
    Dim v As Variant
    Dim i As Long
    Dim oAssign As Object
    Dim oBody As Object

    Set oList = New Collection
    v = CvIds(elemIds)
    For i = LBound(v) To UBound(v)
        oList.Add CLng(v(i))
    Next i

    Set oRec = New Dictionary
    oRec.Add "NAME", NAME
    oRec.Add "ELEMLIST", oList
    oRec.Add "ZVAR", UCase$(ZVAR)
    oRec.Add "YVAR", UCase$(YVAR)
    oRec.Add "ZFROM", ZFROM
    oRec.Add "YFROM", YFROM
    If UCase$(ZVAR) = "POLY" Then
        oRec.Add "ZEXP", ZEXP
        oRec.Add "ZDIST", ZDIST
    End If
    If UCase$(YVAR) = "POLY" Then
        oRec.Add "YEXP", YEXP
        oRec.Add "YDIST", YDIST
    End If

    Set oAssign = New Dictionary
    oAssign.Add CStr(groupId), oRec
    Set oBody = New Dictionary
    oBody.Add "Assign", oAssign
    TaperedGroup = CvQueue("/db/TSGR", oBody)
End Function

' ---- helpers ----

Private Function CvSectItem(ByVal pType As String, ByVal pName As String) As Object
    Dim oItem As Object
    Set oItem = New Dictionary
    oItem.Add "SECTTYPE", pType
    oItem.Add "SECT_NAME", pName
    oItem.Add "SECT_BEFORE", New Dictionary
    Set CvSectItem = oItem
End Function

' Offset (no user distances - SectionOffset sets those) and the two flags.
Private Sub CvSectTail(ByVal pBefore As Object, ByVal pOffsetPt As String, _
                       ByVal pShear As Boolean, ByVal pWarp As Boolean)
    pBefore.Add "OFFSET_PT", pOffsetPt
    pBefore.Add "OFFSET_CENTER", 0
    pBefore.Add "USER_OFFSET_REF", 0
    pBefore.Add "HORZ_OFFSET_OPT", 0
    pBefore.Add "USERDEF_OFFSET_YI", 0
    pBefore.Add "USERDEF_OFFSET_YJ", 0
    pBefore.Add "VERT_OFFSET_OPT", 0
    pBefore.Add "USERDEF_OFFSET_ZI", 0
    pBefore.Add "USERDEF_OFFSET_ZJ", 0
    pBefore.Add "USE_SHEAR_DEFORM", pShear
    pBefore.Add "USE_WARPING_EFFECT", pWarp
End Sub

Private Function CvSectPut(ByVal pId As Long, ByVal pItem As Object) As String
    Dim oAssign As Object
    Dim oBody As Object
    Set oAssign = New Dictionary
    oAssign.Add CStr(pId), pItem
    Set oBody = New Dictionary
    oBody.Add "Assign", oAssign
    CvSectPut = CvQueue("/db/SECT", oBody)
End Function

Private Sub CvCompMatl(ByVal pBefore As Object, ByVal pE As Double, ByVal pD As Double, _
                       ByVal pPs As Double, ByVal pPc As Double, ByVal pT As Double, _
                       ByVal pMulti As Boolean, ByVal pCreep As Double, ByVal pShrink As Double)
    pBefore.Add "MATL_ELAST", pE
    pBefore.Add "MATL_DENS", pD
    pBefore.Add "MATL_POIS_S", pPs
    pBefore.Add "MATL_POIS_C", pPc
    pBefore.Add "MATL_THERMAL", pT
    pBefore.Add "USE_MULTI_ELAST", pMulti
    pBefore.Add "LONGTERM_ESEC", pCreep
    pBefore.Add "SHRINK_ESEC", pShrink
End Sub

Private Function CvOneKey(ByVal pKey As String, ByVal pValue As Variant) As Object
    Dim o As Object
    Set o = New Dictionary
    o.Add pKey, pValue
    Set CvOneKey = o
End Function

' The numbers as given, no padding: Array, worksheet Range or one number.
Private Function CvSizes(ByVal pAny As Variant) As Variant
    Dim out() As Double
    Dim v As Variant
    Dim n As Long

    If IsMissing(pAny) Then
        CvSizes = Array()
        Exit Function
    End If
    If IsObject(pAny) Or IsArray(pAny) Then
        For Each v In pAny
            If Not IsEmpty(v) Then n = n + 1
        Next v
        If n = 0 Then
            CvSizes = Array()
            Exit Function
        End If
        ReDim out(0 To n - 1)
        n = 0
        For Each v In pAny
            If Not IsEmpty(v) Then
                out(n) = CDbl(v)
                n = n + 1
            End If
        Next v
    Else
        ReDim out(0 To 0)
        out(0) = CDbl(pAny)
    End If
    CvSizes = out
End Function

' Hw B1 Bf1 tf1 Bf3 tw B2 Bf2 tf2 tfp -> the vSIZE order of a tub.
Private Function CvTubSize(ByVal p As Variant) As Variant
    Dim b As Long
    b = LBound(p)
    CvTubSize = Array(p(b), p(b + 5), p(b + 1), p(b + 2), p(b + 3), _
                      p(b + 6), p(b + 7), p(b + 8), p(b + 4), p(b + 9))
End Function

' Missing joints are all off; given ones are padded with off.
Private Function CvJoints(ByVal pFlags As Variant, ByVal n As Long) As Variant
    Dim out() As Boolean
    Dim v As Variant
    Dim i As Long

    ReDim out(0 To n - 1)
    If Not IsMissing(pFlags) Then
        If IsObject(pFlags) Or IsArray(pFlags) Then
            For Each v In pFlags
                If i > n - 1 Then Exit For
                If Not IsEmpty(v) Then out(i) = CBool(v)
                i = i + 1
            Next v
        ElseIf Not IsEmpty(pFlags) Then
            out(0) = CBool(pFlags)
        End If
    End If
    CvJoints = out
End Function

' One end of a PSC 1 / 2 cell box. S_WIDTH is HO1.
Private Function CvCellEnd(ByVal pA As Variant, ByVal pB As Variant, _
                           ByVal pC As Variant, ByVal pD As Variant) As Object
    Dim o As Object
    Dim vA As Variant
    vA = PscNums(pA, 6)
    Set o = New Dictionary
    o.Add "vSIZE_PSC_A", vA
    o.Add "vSIZE_PSC_B", PscNums(pB, 6)
    o.Add "vSIZE_PSC_C", PscNums(pC, 10)
    o.Add "vSIZE_PSC_D", PscNums(pD, 8)
    o.Add "S_WIDTH", vA(0)
    Set CvCellEnd = o
End Function

' A 10, B 7, C 9, D 7 of a PSC I shape. Symmetric: C = A without H1, D = B.
Private Function CvPscISizes(ByVal pSymm As Boolean, ByVal pA As Variant, ByVal pB As Variant, _
                             ByVal pC As Variant, ByVal pD As Variant) As Object
    Dim o As Object
    Dim vA As Variant
    Dim vC() As Double
    Dim i As Long

    vA = PscNums(pA, 10)
    Set o = New Dictionary
    o.Add "vSIZE_PSC_A", vA
    o.Add "vSIZE_PSC_B", PscNums(pB, 7)
    If pSymm Then
        ReDim vC(0 To 8)
        For i = 0 To 8
            vC(i) = vA(i + 1)
        Next i
        o.Add "vSIZE_PSC_C", vC
        o.Add "vSIZE_PSC_D", PscNums(pB, 7)
    Else
        o.Add "vSIZE_PSC_C", PscNums(pC, 9)
        o.Add "vSIZE_PSC_D", PscNums(pD, 7)
    End If
    Set CvPscISizes = o
End Function

' J1 JL1..JL4 JR1..JR4. Symmetric: JR = JL.
Private Function CvPscIJoints(ByVal pSymm As Boolean, ByVal pFlags As Variant) As Variant
    Dim v As Variant
    Dim i As Long
    v = CvJoints(pFlags, 9)
    If pSymm Then
        For i = 1 To 4
            v(i + 4) = v(i)
        Next i
    End If
    CvPscIJoints = v
End Function

' OUTER_POLYGON (and INNER_POLYGON when there are holes) into a SECT_I / SECT_J.
Private Sub CvPolygonsAdd(ByVal pSect As Object, ByVal pOuter As Variant, ByVal pInner As Variant)
    Dim oList As Collection
    Dim vHoles As Variant
    Dim i As Long

    Set oList = New Collection
    oList.Add CvOneKey("VERTEX", CvPolygon(pOuter, False))
    pSect.Add "OUTER_POLYGON", oList

    If IsMissing(pInner) Then Exit Sub
    If IsEmpty(pInner) Then Exit Sub
    If IsArray(pInner) Then
        If UBound(pInner) < LBound(pInner) Then Exit Sub
    End If
    Set oList = New Collection
    If CvIsPolygonList(pInner) Then
        vHoles = pInner
        For i = LBound(vHoles) To UBound(vHoles)
            oList.Add CvOneKey("VERTEX", CvPolygon(vHoles(i), True))
        Next i
    Else
        oList.Add CvOneKey("VERTEX", CvPolygon(pInner, True))
    End If
    pSect.Add "INNER_POLYGON", oList
End Sub

' Array(Array(Array(x, y), ...), ...) - several polygons.
Private Function CvIsPolygonList(ByVal p As Variant) As Boolean
    Dim vFirst As Variant
    If IsObject(p) Or Not IsArray(p) Then Exit Function
    If CvIs2D(p) Then Exit Function
    vFirst = p(LBound(p))
    If Not IsArray(vFirst) Then Exit Function
    If CvIs2D(vFirst) Then
        CvIsPolygonList = True
    Else
        CvIsPolygonList = IsArray(vFirst(LBound(vFirst)))
    End If
End Function

Private Function CvIs2D(ByVal p As Variant) As Boolean
    Dim n As Long
    On Error GoTo NotTwo
    n = UBound(p, 2)
    CvIs2D = True
    Exit Function
NotTwo:
    CvIs2D = False
End Function

' Points -> list of {X, Y}, turned counter-clockwise (pCW: clockwise).
Private Function CvPolygon(ByVal pPts As Variant, ByVal pCW As Boolean) As Collection
    Dim vPts As Variant
    Dim x() As Double
    Dim y() As Double
    Dim n As Long
    Dim i As Long
    Dim r As Long
    Dim cx As Double
    Dim cy As Double
    Dim d As Double
    Dim bRev As Boolean
    Dim oPt As Object
    Dim col As Collection

    If IsObject(pPts) Then vPts = pPts.Value Else vPts = pPts
    If CvIs2D(vPts) Then
        n = UBound(vPts, 1) - LBound(vPts, 1) + 1
        ReDim x(0 To n - 1)
        ReDim y(0 To n - 1)
        For i = 0 To n - 1
            r = LBound(vPts, 1) + i
            x(i) = CDbl(vPts(r, LBound(vPts, 2)))
            y(i) = CDbl(vPts(r, LBound(vPts, 2) + 1))
        Next i
    Else
        n = UBound(vPts) - LBound(vPts) + 1
        ReDim x(0 To n - 1)
        ReDim y(0 To n - 1)
        For i = 0 To n - 1
            x(i) = CDbl(vPts(LBound(vPts) + i)(LBound(vPts(LBound(vPts) + i))))
            y(i) = CDbl(vPts(LBound(vPts) + i)(LBound(vPts(LBound(vPts) + i)) + 1))
        Next i
    End If

    For i = 0 To n - 1
        cx = cx + x(i)
        cy = cy + y(i)
    Next i
    cx = cx / n
    cy = cy / n
    For i = 0 To n - 2
        d = d + (x(i) - cx) * (y(i + 1) - cy) - (y(i) - cy) * (x(i + 1) - cx)
    Next i
    bRev = (d < 0)
    If pCW Then bRev = Not bRev

    Set col = New Collection
    For i = 0 To n - 1
        If bRev Then r = n - 1 - i Else r = i
        Set oPt = New Dictionary
        oPt.Add "X", x(r)
        oPt.Add "Y", y(r)
        col.Add oPt
    Next i
    Set CvPolygon = col
End Function


'==========================================================
' [15] Time dependent materials - creep / shrinkage, strength, link, change property
'==========================================================
'  Creep / shrinkage, strength development, material link, change
'  property. Creep / shrinkage functions start with Creep, the
'  compressive strength (development) functions with Strength; both take
'  the id first, then the name the link uses.
'    CreepKorean 1, "C40", "KDS_2016", 40000, 70, 3, 1.2
'    StrengthKDS 1, "C40", 40000
'    TimeMaterialLink 1, "C40", "C40"         ' material 1 uses both
'  They are sent right after the materials.
'==========================================================

' ---- creep / shrinkage (/db/TDMT) ----

' codeYear 2000 / 2011 / 2020. typeCement SL / NR / RS, typeAggregate
' Basalt / Quartzite / Limestone / Sandstone (2020 only).
Public Function CreepIRC(ByVal tdmtId As Long, ByVal NAME As String, Optional ByVal codeYear As Long = 2011, _
                         Optional ByVal fck As Double = 0, Optional ByVal notionalSize As Double = 1, _
                         Optional ByVal humidity As Double = 70, Optional ByVal ageShrinkage As Long = 3, _
                         Optional ByVal typeCement As String = "NR", _
                         Optional ByVal typeAggregate As String = "Basalt") As String
    Dim oRec As Object
    ' the API takes SL and RS the other way round
    Select Case UCase$(typeCement)
        Case "SL": typeCement = "RS"
        Case "RS": typeCement = "SL"
    End Select
    Set oRec = CvTdRec("NAME", NAME, "CODE", "INDIA_IRC_18_2000", "STR", fck, "HU", humidity, _
                       "AGE", ageShrinkage, "MSIZE", notionalSize)
    If codeYear = 2020 Then
        oRec.Item("CODE") = "INDIA_IRC_112_2020"
        oRec.Add "CTYPE", typeCement
        oRec.Add "TYPEOFAFFR", CvAggIndex(typeAggregate)
    ElseIf codeYear = 2011 Then
        oRec.Item("CODE") = "INDIA_IRC_112_2011"
        oRec.Add "CTYPE", typeCement
    End If
    CreepIRC = CvTdPut("/db/TDMT", tdmtId, oRec)
End Function

' codeYear 2010 / 1990 / 1978. typeAggregate (2010 only) 0 Basalt .. 3 Sandstone.
Public Function CreepCEB(ByVal tdmtId As Long, ByVal NAME As String, Optional ByVal codeYear As Long = 2010, _
                         Optional ByVal fck As Double = 0, Optional ByVal notionalSize As Double = 1, _
                         Optional ByVal humidity As Double = 70, Optional ByVal ageShrinkage As Long = 3, _
                         Optional ByVal typeCement As String = "RS", _
                         Optional ByVal typeAggregate As Long = 0) As String
    Dim oRec As Object
    Dim sCode As String
    Select Case codeYear
        Case 1990: sCode = "CEB"
        Case 1978: sCode = "CEB_FIP_1978"
        Case Else: sCode = "CEB_FIP_2010"
    End Select
    Set oRec = CvTdRec("NAME", NAME, "CODE", sCode, "STR", fck, "HU", humidity, _
                       "AGE", ageShrinkage, "MSIZE", notionalSize, "CTYPE", typeCement)
    If codeYear = 2010 Then oRec.Add "TYPEOFAFFR", typeAggregate
    CreepCEB = CvTdPut("/db/TDMT", tdmtId, oRec)
End Function

' materialType CODE (cement content, slump, fine aggregate %, air) or
' USER (creep coefficient, shrinkage strain E-6). curing MOIST / STEAM.
Public Function CreepACI(ByVal tdmtId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                         Optional ByVal humidity As Double = 70, Optional ByVal ageShrinkage As Long = 3, _
                         Optional ByVal volSurface As Double = 1.2, Optional ByVal cfactA As Double = 4, _
                         Optional ByVal cfactB As Double = 0.85, Optional ByVal curing As String = "MOIST", _
                         Optional ByVal materialType As String = "CODE", _
                         Optional ByVal cementContent As Double = 24, Optional ByVal slump As Double = 1.1, _
                         Optional ByVal fineAggPercent As Double = 12, Optional ByVal airContent As Double = 13, _
                         Optional ByVal creepCoeff As Double = 1.4, _
                         Optional ByVal shrinkStrain As Double = 500) As String
    Dim oRec As Object
    Set oRec = CvTdRec("NAME", NAME, "CODE", "ACI", "STR", fck, "HU", humidity, "AGE", ageShrinkage, _
                       "VOL", volSurface, "CFACTA", cfactA, "CFACTB", cfactB, _
                       "TYPE", materialType, "CMETHOD", curing)
    If materialType = "CODE" Then
        oRec.Add "CEMCONTENT", cementContent
        oRec.Add "SLUMP", slump
        oRec.Add "FAPERCENT", fineAggPercent
        oRec.Add "AIRCONTENT", airContent
    ElseIf materialType = "USER" Then
        oRec.Add "CREEPCOEFF", creepCoeff
        oRec.Add "SHRINKSTRAIN", shrinkStrain
    End If
    CreepACI = CvTdPut("/db/TDMT", tdmtId, oRec)
End Function

Public Function CreepAASHTO(ByVal tdmtId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                            Optional ByVal humidity As Double = 70, Optional ByVal ageShrinkage As Long = 3, _
                            Optional ByVal volSurface As Double = 1.2, _
                            Optional ByVal bExpose As Boolean = False) As String
    CreepAASHTO = CvTdPut("/db/TDMT", tdmtId, _
                          CvTdRec("NAME", NAME, "CODE", "AASHTO", "STR", fck, "HU", humidity, _
                                  "AGE", ageShrinkage, "VOL", volSurface, "bEXPOSE", bExpose))
End Function

' typeCement "Class S" / "Class N" / "Class R". tCode 0 EN 1992-1, 1 EN 1992-2 (bridge).
Public Function CreepEuropean(ByVal tdmtId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                              Optional ByVal humidity As Double = 70, Optional ByVal ageShrinkage As Long = 3, _
                              Optional ByVal notionalSize As Double = 1.2, _
                              Optional ByVal typeCement As String = "Class N", _
                              Optional ByVal tCode As Long = 0, Optional ByVal bSilica As Boolean = False) As String
    Dim oRec As Object
    Set oRec = CvTdRec("NAME", NAME, "CODE", "EUROPEAN", "STR", fck, "HU", humidity, "AGE", ageShrinkage, _
                       "MSIZE", notionalSize, "CTYPE", typeCement, "TCODE", tCode)
    If tCode = 1 Then oRec.Add "bSILICA", bSilica
    CreepEuropean = CvTdPut("/db/TDMT", tdmtId, oRec)
End Function

Public Function CreepRussian(ByVal tdmtId As Long, ByVal NAME As String, ByVal fck As Double, _
                             ByVal humidity As Double, ByVal moduleExposed As Double, ByVal ageConcrete As Long, _
                             ByVal waterContent As Double, ByVal maxAggregate As Double, ByVal airContent As Double, _
                             ByVal cementPaste As Double, Optional ByVal curing As Long = 0, _
                             Optional ByVal cementType As Long = 1, Optional ByVal fastCreep As Boolean = False, _
                             Optional ByVal concreteType As Long = 0) As String
    CreepRussian = CvTdPut("/db/TDMT", tdmtId, _
                           CvTdRec("NAME", NAME, "CODE", "RUSSIAN", "STR", fck, "HU", humidity, _
                                   "M", moduleExposed, "AGE", ageConcrete, "CMETH", curing, _
                                   "iCTYPE", cementType, "CREEP", fastCreep, "CONCT", concreteType, _
                                   "W", waterContent, "MAXS", maxAggregate, "A", airContent, "PZ", cementPaste))
End Function

' standard AS_5100_5_2017 / AS_5100_5_2016 / AS_RTA_5100_5_2011 / AS_3600_2009 / NEWZEALAND.
' dryingType: AS 0 800, 1 900, 2 1000, 3 user; NZ 0..9 by place, 10 user.
Public Function CreepASNZ(ByVal tdmtId As Long, ByVal NAME As String, ByVal standard As String, _
                          ByVal fck As Double, ByVal ageConcrete As Long, ByVal thickness As Double, _
                          Optional ByVal dryingType As Long = 0, Optional ByVal userStrain As Double = 0, _
                          Optional ByVal humidityFactor As Double = 0.72, _
                          Optional ByVal exposure As Long = 0) As String
    Dim oRec As Object
    Dim vEps As Variant
    Dim bNZ As Boolean

    bNZ = (standard = "NEWZEALAND")
    vEps = Null
    If (Not bNZ And dryingType = 3) Or (bNZ And dryingType = 10) Then
        vEps = userStrain
    ElseIf Not bNZ Then
        If dryingType >= 0 And dryingType <= 2 Then vEps = Array(800#, 900#, 1000#)(dryingType)
    Else
        If dryingType >= 0 And dryingType <= 9 Then
            vEps = Array(1500#, 1460#, 1315#, 1080#, 1000#, 990#, 950#, 775#, 735#, 570#)(dryingType)
        End If
    End If
    Set oRec = CvTdRec("NAME", NAME, "CODE", standard, "STR", fck, "THIK", thickness, _
                       "AGE", ageConcrete, "iEPS_DRY", dryingType, "EPS_DRY", vEps)
    If bNZ Then oRec.Add "FS", humidityFactor Else oRec.Add "EXPOSURE", exposure
    CreepASNZ = CvTdPut("/db/TDMT", tdmtId, oRec)
End Function

' standard CHINESE / JTG / CHINA_JTG3362_2018. humidityType CU / RH.
Public Function CreepChinese(ByVal tdmtId As Long, ByVal NAME As String, ByVal standard As String, _
                             ByVal fck As Double, ByVal humidity As Double, ByVal ageConcrete As Long, _
                             ByVal notionalSize As Double, Optional ByVal humidityType As String = "RH", _
                             Optional ByVal cementCoeff As Double = 5, _
                             Optional ByVal flyAsh As Double = 20) As String
    Dim oRec As Object
    Set oRec = CvTdRec("NAME", NAME, "CODE", standard, "STR", fck, "HU", humidity, "AGE", ageConcrete, _
                       "MSIZE", notionalSize, "HTYPE", humidityType)
    If InStr(standard, "JTG") > 0 Then oRec.Add "BSC", cementCoeff
    If standard = "CHINA_JTG3362_2018" Then oRec.Add "FLYASH", flyAsh
    CreepChinese = CvTdPut("/db/TDMT", tdmtId, oRec)
End Function

' standard KDS_2016 / KSI_USD12 / KSCE_2010 / KS. cementType SL / NR / RS.
' density is used by KDS_2016 only.
Public Function CreepKorean(ByVal tdmtId As Long, ByVal NAME As String, ByVal standard As String, _
                            ByVal fck As Double, ByVal humidity As Double, ByVal ageConcrete As Long, _
                            ByVal notionalSize As Double, Optional ByVal cementType As String = "NR", _
                            Optional ByVal density As Double = 240) As String
    Dim oRec As Object
    Set oRec = CvTdRec("NAME", NAME, "CODE", standard, "STR", fck, "HU", humidity, "AGE", ageConcrete, _
                       "MSIZE", notionalSize, "CTYPE", cementType)
    If standard = "KDS_2016" Then oRec.Add "DENSITY", density
    CreepKorean = CvTdPut("/db/TDMT", tdmtId, oRec)
End Function

Public Function CreepPCA(ByVal tdmtId As Long, ByVal NAME As String, ByVal fck As Double, _
                         ByVal humidity As Double, ByVal ultimateCreep As Double, ByVal volSurface As Double, _
                         ByVal reinforcementRatio As Double, ByVal steelModulus As Double, _
                         ByVal ultimateShrinkage As Double) As String
    CreepPCA = CvTdPut("/db/TDMT", tdmtId, _
                       CvTdRec("NAME", NAME, "CODE", "PCA", "STR", fck, "HU", humidity, _
                               "UCS", ultimateCreep, "VOL", volSurface, "RR", reinforcementRatio, _
                               "MOD", steelModulus, "USS", ultimateShrinkage))
End Function

' standard JSCE_12 / JSCE_07 / JSCE.
Public Function CreepJapan(ByVal tdmtId As Long, ByVal NAME As String, ByVal standard As String, _
                           ByVal humidity As Double, ByVal ageConcrete As Long, ByVal volSurface As Double, _
                           ByVal cementContent As Double, ByVal waterContent As Double, _
                           Optional ByVal fck As Double = 30000, Optional ByVal impactFactor As Double = 1, _
                           Optional ByVal ageSolidification As Long = 5, Optional ByVal alphaFactor As Long = 11, _
                           Optional ByVal bAutogenous As Boolean = True, Optional ByVal gammaFactor As Long = 1, _
                           Optional ByVal aFactor As Double = 0.1, Optional ByVal bFactor As Double = 0.7, _
                           Optional ByVal bGeneral As Boolean = True) As String
    Dim oRec As Object
    Set oRec = CvTdRec("NAME", NAME, "CODE", standard, "HU", humidity, "AGE", ageConcrete, _
                       "VOL", volSurface, "CEMCONTENT", cementContent, "WATERCONTENT", waterContent)
    If standard <> "JSCE" Then oRec.Add "STR", fck
    If standard = "JSCE_12" Then
        oRec.Add "IPFACT", impactFactor
        oRec.Add "AGESOL", ageSolidification
    End If
    If standard = "JSCE_07" Then
        oRec.Add "ALPHAFACT", alphaFactor
        oRec.Add "bAUTO", bAutogenous
        If bAutogenous Then
            oRec.Add "GAMMAFACT", gammaFactor
            oRec.Add "AFACT", aFactor
            oRec.Add "BFACT", bFactor
        End If
        oRec.Add "bGEN", bGeneral
    End If
    CreepJapan = CvTdPut("/db/TDMT", tdmtId, oRec)
End Function

' Japanese standard (CODE JAPAN). method JSCE / AIJ, humidityType RH / CU, cementType RH / NC.
Public Function CreepJapanese(ByVal tdmtId As Long, ByVal NAME As String, ByVal fck As Double, _
                              ByVal humidity As Double, ByVal ageConcrete As Long, ByVal notionalSize As Double, _
                              Optional ByVal method As String = "JSCE", Optional ByVal humidityType As String = "RH", _
                              Optional ByVal cementType As String = "NC", _
                              Optional ByVal envCoeff As Long = 1) As String
    CreepJapanese = CvTdPut("/db/TDMT", tdmtId, _
                            CvTdRec("NAME", NAME, "CODE", "JAPAN", "STR", fck, "HU", humidity, _
                                    "HTYPE", humidityType, "AGE", ageConcrete, "MSIZE", notionalSize, _
                                    "CTYPE", cementType, "CM", method, "LAMBDA", envCoeff))
End Function

' Functions defined by the user (by their names).
Public Function CreepUser(ByVal tdmtId As Long, ByVal NAME As String, ByVal shrinkFunc As String, _
                          ByVal creepFunc As String, ByVal creepAge As Long) As String
    Dim oAges As Collection
    Set oAges = New Collection
    oAges.Add CvTdRec("NAME", creepFunc, "AGE", creepAge)
    CreepUser = CvTdPut("/db/TDMT", tdmtId, _
                        CvTdRec("NAME", NAME, "CODE", "USER_DEFINED", "SSFNAME", shrinkFunc, "vCREEP_AGE", oAges))
End Function

' ---- compressive strength development (/db/TDME) ----

' codeYear 2000 / 2011 / 2020. cementType / aggregateType for 2011 and 2020.
Public Function StrengthIRC(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal codeYear As Long = 2020, _
                            Optional ByVal fckDelta As Double = 0, Optional ByVal cementType As Long = 1, _
                            Optional ByVal aggregateType As Long = 0) As String
    Dim oRec As Object
    Dim sCode As String
    Select Case codeYear
        Case 2011: sCode = "INDIA(IRC:112-2011)"
        Case 2000: sCode = "INDIA(IRC:18-2000)"
        Case Else: sCode = "INDIA(IRC:112-2020)"
    End Select
    Set oRec = CvStrRec(NAME, sCode, fckDelta)
    If codeYear = 2020 Or codeYear = 2011 Then
        oRec.Add "iCTYPE", cementType
        oRec.Add "nAGGRE", aggregateType
    End If
    StrengthIRC = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

Public Function StrengthACI(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                            Optional ByVal factorA As Double = 1, Optional ByVal factorB As Double = 2) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "ACI", fck)
    oRec.Add "A", factorA
    oRec.Add "B", factorB
    StrengthACI = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

' codeYear 2010 / 1990 / 1978.
Public Function StrengthCEB(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal codeYear As Long = 2010, _
                            Optional ByVal fck As Double = 0, Optional ByVal cementType As Long = 1, _
                            Optional ByVal aggregateType As Long = 0) As String
    Dim oRec As Object
    Dim sCode As String
    Select Case codeYear
        Case 1978: sCode = "CEB-FIP(1978)"
        Case 1990: sCode = "CEB-FIP(1990)"
        Case Else: sCode = "CEB-FIP(2010)"
    End Select
    Set oRec = CvStrRec(NAME, sCode, fck)
    If codeYear = 1990 Or codeYear = 2010 Then oRec.Add "iCTYPE", cementType
    If codeYear = 2010 Then oRec.Add "nAGGRE", aggregateType
    StrengthCEB = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

Public Function StrengthOhzagi(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                               Optional ByVal cementType As Long = 2) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "Ohzagi", fck)
    oRec.Add "iCTYPE", cementType
    StrengthOhzagi = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

Public Function StrengthEuropean(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                                 Optional ByVal cementType As Long = 2) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "EUROPEAN", fck)
    oRec.Add "iCTYPE", cementType
    StrengthEuropean = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

Public Function StrengthRussian(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                                Optional ByVal cementType As Long = 1, Optional ByVal curing As Long = 1, _
                                Optional ByVal concreteType As Long = 1, Optional ByVal maxAggregate As Double = 0.02, _
                                Optional ByVal cementContent As Double = 0.25) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "RUSSIAN", fck)
    oRec.Add "iCTYPE", cementType
    oRec.Add "CMETH", curing
    oRec.Add "CTYPE", concreteType
    oRec.Add "MAXS", maxAggregate
    oRec.Add "PZ", cementContent
    StrengthRussian = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

' standard AS5100.5-2017 / AS5100.5-2016 / AS/RTA5100.5-2011 / AS3600-2009.
Public Function StrengthAS(ByVal tdmeId As Long, ByVal NAME As String, _
                           Optional ByVal standard As String = "AS5100.5-2017", _
                           Optional ByVal fck As Double = 0) As String
    StrengthAS = CvTdPut("/db/TDME", tdmeId, CvStrRec(NAME, standard, fck))
End Function

Public Function StrengthGilbertRanzi(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                                     Optional ByVal cementType As Long = 1, _
                                     Optional ByVal density As Double = 230) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "GILBERT AND RANZI", fck)
    oRec.Add "iCTYPE", cementType
    oRec.Add "DENSITY", density
    StrengthGilbertRanzi = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

' useConcreteData False: factors a, b, d are written.
Public Function StrengthJapanHydration(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                                       Optional ByVal cementType As Long = 1, _
                                       Optional ByVal useConcreteData As Boolean = True, _
                                       Optional ByVal tensileFactor As Double = 3, _
                                       Optional ByVal factorA As Double = 4.5, Optional ByVal factorB As Double = 0.95, _
                                       Optional ByVal factorD As Double = 1.11) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "Japan(hydration)", fck)
    oRec.Add "iCTYPE", cementType
    oRec.Add "bUSE", useConcreteData
    oRec.Add "TENS_STRN_FACTOR", tensileFactor
    If Not useConcreteData Then
        oRec.Add "A", factorA
        oRec.Add "B", factorB
        oRec.Add "D", factorD
    End If
    StrengthJapanHydration = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

Public Function StrengthJapanElastic(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                                     Optional ByVal elasticCementType As Long = 0) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "Japan(elastic)", fck)
    oRec.Add "iECTYPE", elasticCementType
    StrengthJapanElastic = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

Public Function StrengthKDS(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                            Optional ByVal cementType As Long = 1, Optional ByVal density As Double = 230) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "KDS-2016", fck)
    oRec.Add "iCTYPE", cementType
    oRec.Add "DENSITY", density
    StrengthKDS = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

Public Function StrengthKCI(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                            Optional ByVal cementType As Long = 1) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "KCI-USD12", fck)
    oRec.Add "iCTYPE", cementType
    StrengthKCI = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

Public Function StrengthKorean(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal fck As Double = 0, _
                               Optional ByVal factorA As Double = 1, Optional ByVal factorB As Double = 2) As String
    Dim oRec As Object
    Set oRec = CvStrRec(NAME, "KoreanStandard", fck)
    oRec.Add "A", factorA
    oRec.Add "B", factorB
    StrengthKorean = CvTdPut("/db/TDME", tdmeId, oRec)
End Function

'   StrengthUser 1, "U1", 1, Array(Array(0, 0, 0, 0), Array(28, 30000, 1000, 3000000))
' timeData rows: TIME COMP TENS ELAST (array of rows or a 4-column range).
' Left out: two rows (0 / 1000 days).
Public Function StrengthUser(ByVal tdmeId As Long, ByVal NAME As String, Optional ByVal scaleFactor As Double = 1, _
                             Optional ByVal timeData As Variant) As String
    Dim oRows As Collection
    Dim vData As Variant
    Dim i As Long
    Dim r As Long
    Dim c As Long

    If IsMissing(timeData) Then timeData = Array(Array(0, 0, 0, 0), Array(1000, 30000, 1000, 3000000))
    If IsObject(timeData) Then vData = timeData.Value Else vData = timeData
    Set oRows = New Collection
    If CvIs2D(vData) Then
        c = LBound(vData, 2)
        For r = LBound(vData, 1) To UBound(vData, 1)
            oRows.Add CvTdRec("TIME", vData(r, c), "COMP", vData(r, c + 1), _
                              "TENS", vData(r, c + 2), "ELAST", vData(r, c + 3))
        Next r
    Else
        For i = LBound(vData) To UBound(vData)
            c = LBound(vData(i))
            oRows.Add CvTdRec("TIME", vData(i)(c), "COMP", vData(i)(c + 1), _
                              "TENS", vData(i)(c + 2), "ELAST", vData(i)(c + 3))
        Next i
    End If
    StrengthUser = CvTdPut("/db/TDME", tdmeId, _
                           CvTdRec("NAME", NAME, "TYPE", "USER", "SCALE", scaleFactor, "aDATA", oRows))
End Function

' ---- link and change property ----

'   TimeMaterialLink 1, "C40", "C40"
' Material matId uses the creep / shrinkage and the strength with those names.
Public Function TimeMaterialLink(ByVal matId As Long, Optional ByVal creepName As String = "", _
                                 Optional ByVal strengthName As String = "") As String
    TimeMaterialLink = CvTdPut("/db/TMAT", matId, _
                               CvTdRec("TDMT_NAME", creepName, "TDME_NAME", strengthName))
End Function

'   ChangeProperty Array(1, 2), notionalSize:=0.5
' Element dependent notional size (NSM) or volume / surface ratio (VSR).
Public Function ChangeProperty(ByVal elemIds As Variant, Optional ByVal notionalSize As Double = 0, _
                               Optional ByVal volSurface As Double = 0) As String
    Dim v As Variant
    Dim i As Long
    Dim sType As String
    Dim dValue As Double

    sType = "NSM"
    If notionalSize <> 0 Then dValue = notionalSize
    If volSurface <> 0 Then
        sType = "VSR"
        dValue = volSurface
    End If
    v = CvIds(elemIds)
    For i = LBound(v) To UBound(v)
        CvTdPut "/db/EDMP", CLng(v(i)), CvTdRec("TYPE", sType, "H_VS", dValue)
    Next i
    ChangeProperty = ""
End Function

' ---- helpers ----

' A record from key, value, key, value ...
Private Function CvTdRec(ParamArray pKeyValues() As Variant) As Object
    Dim o As Object
    Dim i As Long
    Set o = New Dictionary
    For i = LBound(pKeyValues) To UBound(pKeyValues) - 1 Step 2
        o.Add CStr(pKeyValues(i)), pKeyValues(i + 1)
    Next i
    Set CvTdRec = o
End Function

Private Function CvStrRec(ByVal pName As String, ByVal pCode As String, ByVal pStrength As Double) As Object
    Set CvStrRec = CvTdRec("NAME", pName, "TYPE", "CODE", "CODENAME", pCode, "STRENGTH", pStrength)
End Function

Private Function CvTdPut(ByVal pEndpoint As String, ByVal pId As Long, ByVal pRec As Object) As String
    Dim oAssign As Object
    Dim oBody As Object
    Set oAssign = New Dictionary
    oAssign.Add CStr(pId), pRec
    Set oBody = New Dictionary
    oBody.Add "Assign", oAssign
    CvTdPut = CvQueue(pEndpoint, oBody)
End Function

Private Function CvAggIndex(ByVal pName As String) As Long
    Select Case pName
        Case "Quartzite": CvAggIndex = 1
        Case "Limestone": CvAggIndex = 2
        Case "Sandstone": CvAggIndex = 3
        Case Else: CvAggIndex = 0
    End Select
End Function


'==========================================================
' [16] Construction stages - stage, composite section, time load, creep coefficient, camber
'==========================================================
'  Stages, composite sections, time loads, creep coefficients, camber.
'  The groups a stage names must exist (StructureGroup,
'  boundary / load groups of the supports and loads, LoadGroup).
'    Stage 1, "CS1", 10, actElem:=Array("Pier", "Girder"), actAge:=7, _
'          actBndr:="Supports", actLoad:="SW"
'    Stage 2, "CS2", 20, deactElem:="Temp", redist:=100, deactLoad:="SW"
'==========================================================

' Every group list takes one name or an Array of names. Ages / days /
' positions / redistributions: one value for all or one per group.
'   actAge   age of the activated elements (default 0)
'   redist   force redistribution % of the removed elements (default 0)
'   bndrPos  "DEFORMED" (default) or "ORIGINAL"
'   loadDay / deactDay  "FIRST" (default), "LAST" or a day
' svResult / svStep: save results of the stage / of every step.
' loadIn + nl: apply the load in nl increments. addStep: extra step days.
Public Function Stage(ByVal stageId As Long, ByVal NAME As String, Optional ByVal DURATION As Double = 0, _
                      Optional ByVal actElem As Variant, Optional ByVal actAge As Variant, _
                      Optional ByVal deactElem As Variant, Optional ByVal redist As Variant, _
                      Optional ByVal actBndr As Variant, Optional ByVal bndrPos As Variant, _
                      Optional ByVal deactBndr As Variant, _
                      Optional ByVal actLoad As Variant, Optional ByVal loadDay As Variant, _
                      Optional ByVal deactLoad As Variant, Optional ByVal deactDay As Variant, _
                      Optional ByVal svResult As Boolean = True, Optional ByVal svStep As Boolean = False, _
                      Optional ByVal loadIn As Boolean = False, Optional ByVal nl As Long = 5, _
                      Optional ByVal addStep As Variant) As String
    Dim oRec As Object
    Dim col As Collection
    Dim vNames As Variant
    Dim i As Long

    Set oRec = New Dictionary
    oRec.Add "NAME", NAME
    oRec.Add "DURATION", DURATION
    oRec.Add "bSV_RSLT", svResult
    oRec.Add "bSV_STEP", svStep
    oRec.Add "bLOAD_STEP", loadIn
    oRec.Add "NO", stageId
    If loadIn Then oRec.Add "INCRE_STEP", nl
    oRec.Add "ADD_STEP", CvNumList(addStep)

    If CvHasNames(actElem) Then
        vNames = CvIds(actElem)
        Set col = New Collection
        For i = LBound(vNames) To UBound(vNames)
            col.Add CvTdRec("GRUP_NAME", vNames(i), "AGE", CvPick(actAge, i - LBound(vNames), 0))
        Next i
        oRec.Add "ACT_ELEM", col
    End If
    If CvHasNames(deactElem) Then
        vNames = CvIds(deactElem)
        Set col = New Collection
        For i = LBound(vNames) To UBound(vNames)
            col.Add CvTdRec("GRUP_NAME", vNames(i), "REDIST", CvPick(redist, i - LBound(vNames), 0))
        Next i
        oRec.Add "DACT_ELEM", col
    End If
    If CvHasNames(actBndr) Then
        vNames = CvIds(actBndr)
        Set col = New Collection
        For i = LBound(vNames) To UBound(vNames)
            col.Add CvTdRec("BNGR_NAME", vNames(i), "POS", CvPick(bndrPos, i - LBound(vNames), "DEFORMED"))
        Next i
        oRec.Add "ACT_BNGR", col
    End If
    If CvHasNames(deactBndr) Then oRec.Add "DACT_BNGR", CvNumList(deactBndr)
    If CvHasNames(actLoad) Then
        vNames = CvIds(actLoad)
        Set col = New Collection
        For i = LBound(vNames) To UBound(vNames)
            col.Add CvTdRec("LOAD_NAME", vNames(i), "DAY", CStr(CvPick(loadDay, i - LBound(vNames), "FIRST")))
        Next i
        oRec.Add "ACT_LOAD", col
    End If
    If CvHasNames(deactLoad) Then
        vNames = CvIds(deactLoad)
        Set col = New Collection
        For i = LBound(vNames) To UBound(vNames)
            col.Add CvTdRec("LOAD_NAME", vNames(i), "DAY", CStr(CvPick(deactDay, i - LBound(vNames), "FIRST")))
        Next i
        oRec.Add "DACT_LOAD", col
    End If

    Stage = CvTdPut("/db/STAG", stageId, oRec)
End Function

'   StageComposite 1, "CS1", 1, Array(Array(1, "ELEM"), Array(2, "MATL", "2", "CS2", 5))
' Composite section built in stages. Each part row:
'   PART, MTYPE ("ELEM" / "MATL"), MAT, CSTAGE, AGE, PARTINFO_H ("AUTO"),
'   PARTINFO_VS, PARTINFO_M, AREA, ASY, ASZ, IXX, IYY, IZZ, WAREA, IW
' (only the first two are needed; the rest have defaults).
' compType GENERAL / USER / NORMAL, bTAP for a tapered section.
Public Function StageComposite(ByVal cscsId As Long, ByVal activationStage As String, _
                               ByVal sectId As Long, ByVal parts As Variant, _
                               Optional ByVal compType As String = "GENERAL", _
                               Optional ByVal bTAP As Boolean = False) As String
    Dim keys As Variant
    Dim defs As Variant
    Dim vParts As Variant
    Dim vRow As Variant
    Dim col As Collection
    Dim oPart As Object
    Dim i As Long
    Dim k As Long
    Dim n As Long

    keys = Array("PART", "MTYPE", "MAT", "CSTAGE", "AGE", "PARTINFO_H", "PARTINFO_VS", "PARTINFO_M", _
                 "AREA", "ASY", "ASZ", "IXX", "IYY", "IZZ", "WAREA", "IW")
    defs = Array(Empty, Empty, "", "", 0, "AUTO", 0, 0, 1, 1, 1, 1, 1, 1, 1, 1)
    If IsObject(parts) Then vParts = parts.Value Else vParts = parts

    Set col = New Collection
    If CvIs2D(vParts) Then
        For i = LBound(vParts, 1) To UBound(vParts, 1)
            Set oPart = New Dictionary
            n = UBound(vParts, 2) - LBound(vParts, 2) + 1
            For k = 0 To 15
                If k < n Then
                    If IsEmpty(vParts(i, LBound(vParts, 2) + k)) Then
                        oPart.Add keys(k), defs(k)
                    Else
                        oPart.Add keys(k), vParts(i, LBound(vParts, 2) + k)
                    End If
                Else
                    oPart.Add keys(k), defs(k)
                End If
            Next k
            col.Add oPart
        Next i
    Else
        For i = LBound(vParts) To UBound(vParts)
            vRow = vParts(i)
            n = UBound(vRow) - LBound(vRow) + 1
            Set oPart = New Dictionary
            For k = 0 To 15
                If k < n Then oPart.Add keys(k), vRow(LBound(vRow) + k) Else oPart.Add keys(k), defs(k)
            Next k
            col.Add oPart
        Next i
    End If

    StageComposite = CvTdPut("/db/CSCS", cscsId, _
                             CvTdRec("SEC", sectId, "ASTAGE", activationStage, "TYPE", compType, _
                                     "bTAP", bTAP, "vPARTINFO", col))
End Function

'   StageTimeLoad Array(10, 11), 35, "DL2"
' Time load (days) on elements.
Public Function StageTimeLoad(ByVal elemIds As Variant, ByVal DAY As Double, _
                              Optional ByVal GROUP_NAME As String = "") As String
    StageTimeLoad = CvItemsPut("/db/TMLD", elemIds, CvTdRec("ID", 1, "GROUP_NAME", GROUP_NAME, "DAY", DAY), True)
End Function

'   StageCreepCoeff 25, 1.2, "SDL"
' Creep coefficient for the construction stage analysis.
Public Function StageCreepCoeff(ByVal elemIds As Variant, ByVal CREEP As Double, _
                                Optional ByVal GROUP_NAME As String = "") As String
    StageCreepCoeff = CvItemsPut("/db/CRPC", elemIds, CvTdRec("ID", 1, "GROUP_NAME", GROUP_NAME, "CREEP", CREEP), True)
End Function

'   StageCamber 25, 0.17, 0.1
' User camber and deformation at nodes.
Public Function StageCamber(ByVal nodeIds As Variant, ByVal camber As Double, ByVal deform As Double) As String
    Dim v As Variant
    Dim i As Long
    v = CvIds(nodeIds)
    For i = LBound(v) To UBound(v)
        CvTdPut "/db/CMCS", CLng(v(i)), CvTdRec("DEFORM", deform, "USER", camber)
    Next i
    StageCamber = ""
End Function

' A group name or a list of them was given.
Private Function CvHasNames(ByVal p As Variant) As Boolean
    If IsMissing(p) Then Exit Function
    If IsNull(p) Or IsEmpty(p) Then Exit Function
    If IsArray(p) Then
        CvHasNames = (UBound(p) >= LBound(p))
    Else
        CvHasNames = (Len(CStr(p)) > 0)
    End If
End Function


'==========================================================
' [17] Tendons - property, profile, prestress
'==========================================================
'  Tendon property, relaxation, profile, prestress.
'    TendonProperty 1, "Strand", 2, 3, 0.00139, 0.1, "CEBFIP_2010", 1860000, 1600000, _
'                   0.2, 0.002, relaxValue:=2.5, relaxClass:=2
'    TendonProfile 1, "T1", 1, Array(1, 2, 3), _
'                  Array(Array(0, 0, -0.5), Array(15, 0, -1.2), Array(30, 0, -0.5))
'    TendonPrestress "T1", "PS", stress:=1300000
'==========================================================

' tdnType 1 (or "Internal - Pre"), 2 ("Internal - Post"), 3 ("External").
' relax and what relaxValue means:
'   "CEBFIP_2010"  rho (relaxClass: relaxation class)   "CEBFIP_1990" / "CEBFIP_1978"  rho
'   "European"     relaxation class                      "IRC_18" / "IRC_112"           factor
'   "Magura"       factor 10 or 45                       "Null" (default)               no relaxation
' unintAngle: unintentional angular displacement instead of the wobble factor
' (CEB-FIP and European only).
Public Function TendonProperty(ByVal tdntId As Long, ByVal NAME As String, ByVal tdnType As Variant, _
                               ByVal matId As Long, ByVal area As Double, ByVal ductDia As Double, _
                               Optional ByVal relax As String = "Null", _
                               Optional ByVal ultStrength As Double = 0, Optional ByVal yieldStrength As Double = 0, _
                               Optional ByVal curvFric As Double = 0, Optional ByVal wobbleFric As Double = 0, _
                               Optional ByVal relaxValue As Double = 0, Optional ByVal relaxClass As Long = 0, _
                               Optional ByVal unintAngle As Double = 0, Optional ByVal extMomMag As Double = 0, _
                               Optional ByVal anchSlipBegin As Double = 0, Optional ByVal anchSlipEnd As Double = 0, _
                               Optional ByVal bBonded As Boolean = True) As String
    Dim oRec As Object
    Dim sType As String
    Dim sTens As String
    Dim nRm As Long
    Dim bAngle As Boolean

    Select Case CStr(tdnType)
        Case "2", "Internal - Post": sType = "INTERNAL": sTens = "POST"
        Case "3", "External": sType = "EXTERNAL": sTens = "POST"
        Case Else: sType = "INTERNAL": sTens = "PRE"
    End Select
    Set oRec = CvTdRec("NAME", NAME, "TYPE", sType, "LT", sTens, "MATL", matId, "AREA", area, _
                       "D_AREA", ductDia, "ASB", anchSlipBegin, "ASE", anchSlipEnd, _
                       "bBONDED", bBonded, "ALPHA", extMomMag)

    Select Case UCase$(Replace(relax, "-", "_"))
        Case "CEBFIP_2010": nRm = 9: bAngle = True
        Case "CEBFIP_1978": nRm = 1: bAngle = True
        Case "CEBFIP_1990": nRm = 8: bAngle = True
        Case "EUROPEAN": nRm = 5: bAngle = True
        Case "IRC_18": nRm = 4
        Case "IRC_112": nRm = 7
        Case "MAGURA"
            nRm = 0
            If relaxValue <> 10 And relaxValue <> 45 Then relaxValue = 45
        Case Else: nRm = -1
    End Select
    If nRm < 0 Then
        oRec.Add "RM", 0
        oRec.Add "RV", 0
    Else
        oRec.Add "RM", nRm
        oRec.Add "RV", relaxValue
        If nRm = 9 Then oRec.Add "TDMFK", relaxClass
    End If
    oRec.Add "US", ultStrength
    oRec.Add "YS", yieldStrength
    oRec.Add "FF", curvFric
    oRec.Add "WF", wobbleFric
    If nRm < 0 Then oRec.Add "bRELAX", False
    If bAngle And unintAngle <> 0 Then
        oRec.Add "W_TYPE", 1
        oRec.Add "W_ANGLE", unintAngle
    End If
    TendonProperty = CvTdPut("/db/TDNT", tdntId, oRec)
End Function

' points: profile points along the tendon.
'   INPUT "3D": Array(Array(x, y, z), ...) - a 4th value fixes the curve there:
'               the radius (CURVE "ROUND") or Array(ry, rz) (CURVE "SPLINE").
'   INPUT "2D": points = x-y points, pointsXZ = x-z points, Array(x, y[, R]).
' refAxis "ELEMENT" (insPointEnd, insPointElem, xAxisDirElem, offsetY / Z),
'         "STRAIGHT" (insPoint, xAxisDirStraight X / Y / VECTOR, xAxisVec,
'         gradRotAxis, gradRotAngle) or "CURVE" (insPoint, radiusCenter,
'         curveOffset, curveDir CW / CCW, gradRotAxis, gradRotAngle).
' transLenOpt "USER" (transLenBegin / End) or "AUTO" (by the property:
' post-tensioned AUTO1, pre-tensioned AUTO2). nTypical > 0: typical tendon.
Public Function TendonProfile(ByVal tdnaId As Long, ByVal NAME As String, ByVal propId As Long, _
                              ByVal elemIds As Variant, ByVal points As Variant, _
                              Optional ByVal INPUT_ As String = "3D", Optional ByVal CURVE As String = "SPLINE", _
                              Optional ByVal pointsXZ As Variant, Optional ByVal groupId As Long = 0, _
                              Optional ByVal stLenBegin As Double = 0, Optional ByVal stLenEnd As Double = 0, _
                              Optional ByVal nTypical As Long = 0, Optional ByVal transLenOpt As String = "USER", _
                              Optional ByVal transLenBegin As Double = 0, Optional ByVal transLenEnd As Double = 0, _
                              Optional ByVal debondBegin As Double = 0, Optional ByVal debondEnd As Double = 0, _
                              Optional ByVal refAxis As String = "ELEMENT", _
                              Optional ByVal insPointEnd As String = "END-I", Optional ByVal insPointElem As Long = 0, _
                              Optional ByVal xAxisDirElem As String = "I-J", Optional ByVal xAxisRotAngle As Double = 0, _
                              Optional ByVal projection As Boolean = True, _
                              Optional ByVal offsetY As Double = 0, Optional ByVal offsetZ As Double = 0, _
                              Optional ByVal insPoint As Variant, Optional ByVal xAxisDirStraight As String = "X", _
                              Optional ByVal xAxisVec As Variant, Optional ByVal gradRotAxis As String = "X", _
                              Optional ByVal gradRotAngle As Double = 0, Optional ByVal radiusCenter As Variant, _
                              Optional ByVal curveOffset As Double = 0, Optional ByVal curveDir As String = "CW") As String
    Dim oRec As Object
    Dim oProp As Object
    Dim vElems As Variant
    Dim bRound As Boolean

    If INPUT_ <> "2D" Then INPUT_ = "3D"
    If CURVE <> "SPLINE" And CURVE <> "ROUND" Then CURVE = "ROUND"
    bRound = (CURVE = "ROUND")
    Select Case transLenOpt
        Case "USER", "AUTO1", "AUTO2"
        Case "AUTO"
            transLenOpt = "AUTO1"
            Set oProp = StoreGet("/db/TDNT", propId)
            If Not oProp Is Nothing Then
                If oProp.Item("LT") <> "POST" Then transLenOpt = "AUTO2"
            End If
        Case Else: transLenOpt = "USER"
    End Select
    If refAxis <> "STRAIGHT" And refAxis <> "CURVE" Then refAxis = "ELEMENT"
    vElems = CvIds(elemIds)
    If insPointElem = 0 Then insPointElem = CLng(vElems(LBound(vElems)))

    Set oRec = CvTdRec("NAME", NAME, "TDN_PROP", propId, "ELEM", CvNumList(elemIds), _
                       "BELENG", stLenBegin, "ELENG", stLenEnd, "CURVE", CURVE, "INPUT", INPUT_, _
                       "TDN_GRUP", groupId, "LENG_OPT", transLenOpt, "BLEN", transLenBegin, "ELEN", transLenEnd, _
                       "bTP", (nTypical > 0), "CNT", nTypical, "DeBondBLEN", debondBegin, _
                       "DeBondELEN", debondEnd, "SHAPE", refAxis)
    Select Case refAxis
        Case "ELEMENT"
            If insPointEnd <> "END-J" Then insPointEnd = "END-I"
            If xAxisDirElem <> "J-I" Then xAxisDirElem = "I-J"
            oRec.Add "INS_PT", insPointEnd
            oRec.Add "INS_ELEM", insPointElem
            oRec.Add "AXIS_IJ", xAxisDirElem
            oRec.Add "XAR_ANGLE", xAxisRotAngle
            oRec.Add "bPJ", projection
            oRec.Add "OFF_YZ", Array(offsetY, offsetZ)
        Case "STRAIGHT"
            If xAxisDirStraight <> "Y" And xAxisDirStraight <> "VECTOR" Then xAxisDirStraight = "X"
            oRec.Add "IP", PscNums(insPoint, 3)
            oRec.Add "AXIS", xAxisDirStraight
            oRec.Add "VEC", PscNums(xAxisVec, 2)
            oRec.Add "XAR_ANGLE", xAxisRotAngle
            oRec.Add "bPJ", projection
            oRec.Add "GR_AXIS", IIf(gradRotAxis = "Y", "Y", "X")
            oRec.Add "GR_ANGLE", gradRotAngle
        Case "CURVE"
            oRec.Add "IP", PscNums(insPoint, 3)
            oRec.Add "RC", PscNums(radiusCenter, 2)
            oRec.Add "OFFSET", curveOffset
            oRec.Add "DIR", IIf(curveDir = "CCW", "CCW", "CW")
            oRec.Add "XAR_ANGLE", xAxisRotAngle
            oRec.Add "bPJ", projection
            oRec.Add "GR_AXIS", IIf(gradRotAxis = "Y", "Y", "X")
            oRec.Add "GR_ANGLE", gradRotAngle
    End Select
    If INPUT_ = "3D" Then
        oRec.Add "PROF", CvTendonPoints(points, 3, bRound)
    Else
        oRec.Add "PROFY", CvTendonPoints(points, 2, bRound)
        oRec.Add "PROFZ", CvTendonPoints(pointsXZ, 2, bRound)
    End If
    TendonProfile = CvTdPut("/db/TDNA", tdnaId, oRec)
End Function

'   TendonPrestress "T1", "PS", "PS_G", "STRESS", "BOTH", 1300000, 1300000
' Prestress on the tendon profile with that name (define the profile first).
' A load case of type PS and the load group are made when missing.
' TYPE "STRESS" / "FORCE", ORDER (jacking) "BEGIN" / "END" / "BOTH".
Public Function TendonPrestress(ByVal profileName As String, ByVal LCNAME As String, _
                                Optional ByVal GROUP_NAME As String = "", Optional ByVal TYPE_ As String = "STRESS", _
                                Optional ByVal ORDER As String = "BEGIN", Optional ByVal jackBegin As Double = 0, _
                                Optional ByVal jackEnd As Double = 0, Optional ByVal grouting As Long = 0) As String
    Dim oTable As Object
    Dim vKey As Variant
    Dim sKey As String

    Set oTable = CvStoreTable("TDNA")
    For Each vKey In oTable.Keys
        If CStr(oTable.Item(vKey).Item("NAME")) = profileName Then
            sKey = CStr(vKey)
            Exit For
        End If
    Next vKey
    If Len(sKey) = 0 Then
        TendonPrestress = "TDPL: tendon profile """ & profileName & """ is not defined"
        Exit Function
    End If

    CvAutoCase LCNAME, "PS"
    If Len(GROUP_NAME) > 0 Then CvGroupRecord "LDGR", GROUP_NAME
    If TYPE_ <> "FORCE" Then TYPE_ = "STRESS"
    If ORDER <> "END" And ORDER <> "BOTH" Then ORDER = "BEGIN"
    TendonPrestress = CvItemsPut("/db/TDPL", CLng(sKey), _
                                 CvTdRec("ID", 1, "LCNAME", LCNAME, "GROUP_NAME", GROUP_NAME, _
                                         "TENDON_NAME", profileName, "TYPE", TYPE_, "ORDER", ORDER, _
                                         "BEGIN", jackBegin, "END", jackEnd, "GROUTING", grouting), True)
End Function

' Profile points -> PROF / PROFY / PROFZ entries.
Private Function CvTendonPoints(ByVal pPts As Variant, ByVal pDim As Long, ByVal pRound As Boolean) As Collection
    Dim col As Collection
    Dim oPt As Object
    Dim vRow As Variant
    Dim vPt As Variant
    Dim i As Long
    Dim b As Long
    Dim n As Long
    Dim bFix As Boolean

    Set col = New Collection
    If IsMissing(pPts) Then
        Set CvTendonPoints = col
        Exit Function
    End If
    For i = LBound(pPts) To UBound(pPts)
        vRow = pPts(i)
        b = LBound(vRow)
        n = UBound(vRow) - b + 1
        If pDim = 3 Then vPt = Array(vRow(b), vRow(b + 1), vRow(b + 2)) Else vPt = Array(vRow(b), vRow(b + 1))
        bFix = (n > pDim)
        Set oPt = New Dictionary
        oPt.Add "PT", vPt
        oPt.Add "bFIX", bFix
        If pRound Then
            If bFix Then oPt.Add "RADIUS", vRow(b + pDim) Else oPt.Add "RADIUS", 0
        ElseIf pDim = 3 Then
            If bFix Then oPt.Add "R", vRow(b + pDim) Else oPt.Add "R", Array(0, 0)
        Else
            If bFix Then oPt.Add "R", vRow(b + pDim) Else oPt.Add "R", 0
        End If
        col.Add oPt
    Next i
    Set CvTendonPoints = col
End Function


'==========================================================
' [18] Moving load - code, lanes, vehicles, moving load cases
'==========================================================
'  Moving load code, lanes, vehicles, moving load cases.
'    MovingLane 1, "KOREA", "L1", 0, 1.8, Array(1, 2, 3, 4)
'    MovingCase 1, "KOREA", "MV1", 0, Array(Array("VL", "DB-24", 1, 1, 2, Array("L1")))
'  A lane sets the moving load code (MovingCode) to its own code.
'  The code is sent first, before the rest of the model.
'==========================================================

' code: "KSCE-LSD15", "KOREA", "AASHTO STANDARD", "AASHTO LRFD", "AASHTO LRFD(PENDOT)",
' "CHINA", "INDIA", "TAIWAN", "CANADA", "BS", "EUROCODE", "AUSTRALIA", "POLAND",
' "RUSSIA", "SOUTH AFRICA"
Public Function MovingCode(ByVal code As String) As String
    MovingCode = CvTdPut("/db/MVCD", 1, CvTdRec("CODE", code))
End Function

'   MovingLane 1, "KOREA", "L1", 0, 1.8, Array(1, 2, 3, 4), 0.3
' A line lane on elements (in order). factor is the impact factor
' (INDIA, KOREA, TAIWAN, AASHTO STANDARD), scale factor (CHINA),
' centrifugal force (AASHTO LRFD) or eccentricity of the vertical load
' (EUROCODE). span: INDIA / CHINA. width: lane width (3 for the codes whose
' lane has no width). optWidth > 0: auto lane with that width.
' GROUP_NAME given: the load is distributed cross (to that structure group).
' direction "BOTH" / "FORWARD" / "BACKWARD".
Public Function MovingLane(ByVal laneId As Long, ByVal code As String, ByVal NAME As String, _
                           ByVal ecc As Double, ByVal wheelSpace As Double, ByVal elemIds As Variant, _
                           Optional ByVal factor As Double = 0, Optional ByVal span As Double = 0, _
                           Optional ByVal width As Variant, Optional ByVal optWidth As Double = 0, _
                           Optional ByVal GROUP_NAME As String = "", Optional ByVal direction As String = "BOTH", _
                           Optional ByVal skewStart As Double = 0, Optional ByVal skewEnd As Double = 0) As String
    Dim oCommon As Object
    Dim oItem As Object
    Dim col As Collection
    Dim v As Variant
    Dim i As Long
    Dim bStart As Boolean
    Dim sEndpoint As String

    code = UCase$(code)
    If IsMissing(width) Then
        Select Case code
            Case "KOREA", "TAIWAN", "AASHTO STANDARD", "AASHTO LRFD", "AASHTO LRFD(PENDOT)", "CANADA", "KSCE-LSD15"
                width = 3
            Case Else
                width = 0
        End Select
    End If
    MovingCode code

    Set oCommon = CvTdRec("LL_NAME", NAME, "LOAD_DIST", IIf(Len(GROUP_NAME) > 0, "CROSS", "LANE"), _
                          "GROUP_NAME", GROUP_NAME, "SKEW_START", skewStart, "SKEW_END", skewEnd, _
                          "MOVING", direction, "WHEEL_SPACE", wheelSpace, "WIDTH", width, _
                          "OPT_AUTO_LANE", (optWidth > 0), "ALLOW_WIDTH", optWidth)
    If code = "CHINA" Then oCommon.Remove "WIDTH"

    Set col = New Collection
    v = CvIds(elemIds)
    For i = LBound(v) To UBound(v)
        bStart = (i = LBound(v))
        Set oItem = CvTdRec("ELEM", v(i), "ECC", ecc)
        Select Case code
            Case "INDIA"
                oItem.Add "SPAN", span
                oItem.Add "IMPACT_SPAN", IIf(span > 0, 1, 0)
                oItem.Add "IMPACT_FACTOR", factor
            Case "CHINA"
                oItem.Add "SPAN", span
                oItem.Add "SPAN_START", bStart
                oItem.Add "SCALE_FACTOR", factor
            Case "KOREA", "TAIWAN", "AASHTO STANDARD"
                oItem.Add "FACT", factor
                oItem.Add "SPAN_START", bStart
            Case "AASHTO LRFD(PENDOT)", "AUSTRALIA", "POLAND"
                oItem.Add "SPAN_START", bStart
            Case "AASHTO LRFD"
                oItem.Add "CENT_F", factor
                oItem.Add "SPAN_START", bStart
            Case "EUROCODE"
                oItem.Add "ECCEN_VERT_LOAD", factor
        End Select
        col.Add oItem
    Next i

    Select Case code
        Case "INDIA": sEndpoint = "/db/LLANID"
        Case "CHINA": sEndpoint = "/db/LLANCH"
        Case Else: sEndpoint = "/db/LLAN"
    End Select
    MovingLane = CvTdPut(sEndpoint, laneId, CvTdRec("COMMON", oCommon, "LANE_ITEMS", col))
End Function

'   MovingVehicleIndia 1, "70R", "IRC", "Class 70R"
'   MovingVehicleIndia 2, "BG", "IRS", "BG-1676", 11
' standardCode "IRC", "Footway", "IRS", "Fatigue". IRS vehicleType BG-1676,
' MG-1000, NG-762, HML, FTB with vehicleNo (1 based) picking the vehicle.
Public Function MovingVehicleIndia(ByVal vehId As Long, ByVal NAME As String, ByVal standardCode As String, _
                                   ByVal vehicleType As String, Optional ByVal vehicleNo As Long = 0) As String
    Dim oRec As Object
    Dim oIn As Object
    Dim sStd As String
    Dim sTypeName As String
    Dim vList As Variant
    Dim sSel As String

    Select Case standardCode
        Case "IRC", "Footway": sStd = "IRC:6-2000"
        Case "IRS": sStd = "IRS: BRIDGE RULES"
        Case "Fatigue": sStd = "IRC:6-2014"
        Case Else
            MovingVehicleIndia = "MVHL: standardCode must be IRC, IRS, Footway or Fatigue"
            Exit Function
    End Select

    sTypeName = vehicleType
    If standardCode = "IRS" Then
        vList = CvIrsVehicles(vehicleType, sTypeName)
        If IsEmpty(vList) Then
            MovingVehicleIndia = "MVHL: IRS vehicleType must be BG-1676, MG-1000, NG-762, HML or FTB"
            Exit Function
        End If
        If vehicleNo < 1 Or vehicleNo > UBound(vList) + 1 Then
            MovingVehicleIndia = "MVHL: vehicleNo must be 1 to " & (UBound(vList) + 1)
            Exit Function
        End If
        sSel = vList(vehicleNo - 1)
        Set oIn = CvIrsDefaults(vehicleType, sSel)
        oIn.Add "SEL_VEHICLE", sSel
    ElseIf standardCode = "IRC" And vehicleType = "Footway" Then
        Set oIn = CvTdRec("FOOTWAY", 4.903325, "FOOTWAY_WIDTH", 3)
    End If

    Set oRec = CvTdRec("MVLD_CODE", 7, "VEHICLE_LOAD_NAME", NAME, "VEHICLE_LOAD_NUM", 1, _
                       "STANDARD_CODE", sStd, "VEHICLE_TYPE_NAME", sTypeName)
    If Not oIn Is Nothing Then oRec.Add "VEH_IN", oIn
    MovingVehicleIndia = CvTdPut("/db/MVHL", vehId, oRec)
End Function

'   MovingVehicleEurocode 1, "LM1", "RoadBridge", "Load Model 1"
'   MovingVehicleEurocode 2, "LM3", "RoadBridge", "Load Model 3", 2
' standardCode "RoadBridge", "FTB", "RoadBridgeFatigue", "RailTraffic".
' vehicleNo (1 based) for the types with a vehicle list (Load Model 3,
' Load Model 3 (UK NA), HSLM A1 ~ HSLM A10).
Public Function MovingVehicleEurocode(ByVal vehId As Long, ByVal NAME As String, ByVal standardCode As String, _
                                      ByVal vehicleType As String, Optional ByVal vehicleNo As Long = 0) As String
    Dim oEuro As Object
    Dim nSub As Long
    Dim vSel As Variant

    Select Case standardCode
        Case "RoadBridge": nSub = 19
        Case "FTB": nSub = 20
        Case "RoadBridgeFatigue": nSub = 21
        Case "RailTraffic": nSub = 23
        Case Else
            MovingVehicleEurocode = "MVHL: standardCode must be RoadBridge, FTB, RoadBridgeFatigue or RailTraffic"
            Exit Function
    End Select
    Set oEuro = CvEuroDefaults(standardCode, vehicleType, vSel)
    If oEuro Is Nothing Then
        MovingVehicleEurocode = "MVHL: vehicleType """ & vehicleType & """ is not a " & standardCode & " vehicle"
        Exit Function
    End If
    oEuro.Add "SUB_TYPE", nSub
    If Not IsEmpty(vSel) Then
        If vehicleNo < 1 Or vehicleNo > UBound(vSel) + 1 Then
            MovingVehicleEurocode = "MVHL: vehicleNo must be 1 to " & (UBound(vSel) + 1)
            Exit Function
        End If
        oEuro.Add "SEL_VEHICLE", vSel(vehicleNo - 1)
    End If
    MovingVehicleEurocode = CvTdPut("/db/MVHL", vehId, _
                                    CvTdRec("MVLD_CODE", 11, "VEHICLE_LOAD_NAME", NAME, "VEHICLE_LOAD_NUM", 1, _
                                            "VEHICLE_TYPE_NAME", vehicleType, "VEH_EUROCODE", oEuro))
End Function

'   MovingCase 1, "KOREA", "MV1", 0, Array(Array("VL", "DB-24", 1, 1, 2, Array("L1", "L2")))
' Moving load case of every code but INDIA / EUROCODE (MovingCaseIndia /
' MovingCaseEurocode). caseType 0 general, 1 permit, 2 auto optimization.
'   subLoads      rows: vehicle type ("VL" load / "VC" class), vehicle name,
'                 scale factor, min loaded lanes, max loaded lanes, lane names
'   optimizeItems rows: vehicle type, vehicle name, scale factor
'   aslData       multiple factor, vehicle 1, vehicle 2, min lanes, max lanes,
'                 [lane names], [straddling lane names (one list or two)]
'   koreaLaneFactors  a Dictionary of the KOREA lane factor fields
' scaleFactors, laneFactorType, loadCombType, loadModel, fatigue take the
' default of the code when left out.
Public Function MovingCase(ByVal caseId As Long, ByVal code As String, ByVal NAME As String, _
                           Optional ByVal caseType As Long = 0, Optional ByVal subLoads As Variant, _
                           Optional ByVal combOption As String = "COMB", Optional ByVal laneFactorType As Variant, _
                           Optional ByVal permitVehicle As Variant, Optional ByVal refLane As Variant, _
                           Optional ByVal permitScale As Double = 1, Optional ByVal optimizeLane As Variant, _
                           Optional ByVal minVehDist As Variant, Optional ByVal minNumVehicle As Variant, _
                           Optional ByVal maxNumVehicle As Variant, Optional ByVal optimizeItems As Variant, _
                           Optional ByVal aslData As Variant, Optional ByVal fatigue As Variant, _
                           Optional ByVal loadCombType As Variant, Optional ByVal loadModel As Variant, _
                           Optional ByVal koreaLaneFactors As Variant, Optional ByVal scaleFactors As Variant, _
                           Optional ByVal DESC As String = "") As String
    Dim oRec As Object
    Dim oData As Object
    Dim col As Collection
    Dim vRow As Variant
    Dim i As Long
    Dim b As Long
    Dim sComb As String
    Dim oAsl As Object
    Dim oLines As Object
    Dim vStrad As Variant
    Dim k As Variant

    code = UCase$(code)
    If code = "KOREA" Or code = "TAIWAN" Then caseType = 0
    If IsMissing(scaleFactors) Then
        Select Case code
            Case "KSCE-LSD15": scaleFactors = Array(1, 0.9, 0.8, 0.7, 0.65, 0.65)
            Case "AASHTO LRFD", "AASHTO LRFD(PENDOT)": scaleFactors = Array(1.2, 1, 0.85, 0.65, 0.65, 0.65)
            Case "CANADA": scaleFactors = Array(1, 0.9, 0.8, 0.7, 0.6, 0.55)
            Case "AUSTRALIA": scaleFactors = Array(1, 0.8, 0.4, 0.4, 0.4, 0.4)
            Case "RUSSIA": scaleFactors = Array(0, 0, 0, 0, 0, 0)
            Case Else: scaleFactors = Array(1, 1, 0.9, 0.75, 0.75, 0.75)
        End Select
    End If
    If IsMissing(laneFactorType) And code <> "RUSSIA" Then laneFactorType = 1
    If code = "AUSTRALIA" Then
        If IsMissing(loadCombType) Then loadCombType = 1
        If IsMissing(loadModel) Then loadModel = 0
        If IsMissing(fatigue) Then fatigue = False
    ElseIf code = "RUSSIA" Then
        If IsMissing(loadCombType) Then loadCombType = 0
    End If
    Select Case UCase$(combOption)
        Case "INDE", "INDEPENDENT": sComb = "INDEPENDENT"
        Case Else: sComb = "COMBINED"
    End Select

    Set oRec = CvTdRec("LCNAME", NAME, "DESC", DESC, "TYPE", caseType)

    If Not IsMissing(aslData) Then
        b = LBound(aslData)
        Set oAsl = CvTdRec("MULTIPLE_FACTOR", aslData(b), "VEHICLE_LOAD_NAME", aslData(b + 1), _
                           "VEHICLE_LOAD_NAME2", aslData(b + 2), "MIN_LOADED_LANE", aslData(b + 3), _
                           "MAX_LOADED_LANE", aslData(b + 4))
        If UBound(aslData) >= b + 5 Then
            Set oLines = CvTdRec("NA_LLAN_NAMES", CvNumList(aslData(b + 5)), _
                                 "STRAD_LLAN1_NAMES", New Collection, "STRAD_LLAN2_NAMES", New Collection)
            If UBound(aslData) >= b + 6 Then
                vStrad = aslData(b + 6)
                If IsArray(vStrad) Then
                    If UBound(vStrad) >= LBound(vStrad) Then
                        If IsArray(vStrad(LBound(vStrad))) Then
                            Set oLines.Item("STRAD_LLAN1_NAMES") = CvNumList(vStrad(LBound(vStrad)))
                            If UBound(vStrad) > LBound(vStrad) Then
                                If IsArray(vStrad(LBound(vStrad) + 1)) Then
                                    Set oLines.Item("STRAD_LLAN2_NAMES") = CvNumList(vStrad(LBound(vStrad) + 1))
                                End If
                            End If
                        Else
                            Set oLines.Item("STRAD_LLAN1_NAMES") = CvNumList(vStrad)
                        End If
                    End If
                End If
            End If
            oAsl.Add "LINE_ITEMS", oLines
        End If
    End If

    Select Case caseType
        Case 0
            Set oData = New Dictionary
            oData.Add "SCALE_FACTORS", CvNumList(scaleFactors)
            oData.Add "COMB_OPTION", sComb
            If Not IsMissing(laneFactorType) Then oData.Add "LANE_FACTOR_TYPE", laneFactorType
            If Not IsMissing(subLoads) Then
                Set col = New Collection
                For i = LBound(subLoads) To UBound(subLoads)
                    vRow = subLoads(i)
                    b = LBound(vRow)
                    col.Add CvTdRec("VEHICLE_TYPE", CvVehType(vRow(b)), "VEHICLE_NAME", vRow(b + 1), _
                                    "SCALE_FACTOR", vRow(b + 2), "MIN_LOADED_LANE", vRow(b + 3), _
                                    "MAX_LOADED_LANE", vRow(b + 4), "LANE_NAMES", CvNumList(vRow(b + 5)))
                Next i
                oData.Add "SUB_LOAD_DATAS", col
            End If
            If Not IsMissing(loadCombType) Then oData.Add "LOAD_COMB_TYPE", loadCombType
            If Not IsMissing(fatigue) Then oData.Add "FATIGUE", fatigue
            If Not IsMissing(loadModel) Then oData.Add "LOAD_MODEL", loadModel
            If Not IsMissing(koreaLaneFactors) Then
                For Each k In koreaLaneFactors.Keys
                    oData.Item(k) = koreaLaneFactors.Item(k)
                Next k
            End If
            oRec.Add "DEFAULT", oData
            If Not oAsl Is Nothing Then oRec.Add "ASL", oAsl
        Case 1
            If IsMissing(permitVehicle) Then permitVehicle = Null
            If IsMissing(refLane) Then refLane = Null
            oRec.Add "PERMIT_LOAD", CvTdRec("VEHICLE_LOAD_NAME", permitVehicle, "REF_LANE", refLane, _
                                            "SCALE_FACTOR", permitScale)
        Case 2
            Set oData = New Dictionary
            If Not IsMissing(optimizeLane) Then oData.Add "LANE_NAME", optimizeLane
            oData.Add "SCALE_FACTORS", CvNumList(scaleFactors)
            If Not IsMissing(minVehDist) Then oData.Add "MIN_VEHL_DIST", minVehDist
            If Not IsMissing(minNumVehicle) Then oData.Add "MIN_NUM_VEHICLE", minNumVehicle
            If Not IsMissing(maxNumVehicle) Then oData.Add "MAX_NUM_VEHICLE", maxNumVehicle
            If Not IsMissing(optimizeItems) Then
                Set col = New Collection
                For i = LBound(optimizeItems) To UBound(optimizeItems)
                    vRow = optimizeItems(i)
                    b = LBound(vRow)
                    col.Add CvTdRec("VEHICLE_TYPE", CvVehType(vRow(b)), "VEHICLE_NAME", vRow(b + 1), _
                                    "SCALE_FACTOR", vRow(b + 2))
                Next i
                oData.Add "OPTIMIZE_ITEMS", col
            End If
            If Not IsMissing(loadModel) Then oData.Add "LOAD_MODEL", loadModel
            If Not IsMissing(fatigue) Then oData.Add "FATIGUE", fatigue
            oRec.Add "AUTO_OPTIMIZE", oData
            If Not oAsl Is Nothing Then oRec.Add "ASL", oAsl
    End Select
    MovingCase = CvTdPut("/db/MVLD", caseId, oRec)
End Function

'   MovingCaseIndia 1, "IRC", 2, subLoadItems:=Array(Array(1, 1, 2, "Class A", Array("L1", "L2")))
' General (default): subLoadItems rows scale factor, min lanes, max lanes,
'   vehicle class, lane names.
' optAutoLL: rows scale factor, vehicle class 1, vehicle class 2, footway,
'   lane names, [footway lane names].
' optPermit: permitVehicle, refLane, ecc, permitScale.
Public Function MovingCaseIndia(ByVal caseId As Long, ByVal NAME As String, ByVal numLoadedLanes As Long, _
                                Optional ByVal optAutoLL As Boolean = False, Optional ByVal optPermit As Boolean = False, _
                                Optional ByVal subLoadItems As Variant, Optional ByVal scaleFactor As Variant, _
                                Optional ByVal permitVehicle As Variant, Optional ByVal refLane As Variant, _
                                Optional ByVal ecc As Variant, Optional ByVal permitScale As Variant) As String
    Dim oRec As Object
    Dim oSub As Object
    Dim col As Collection
    Dim vRow As Variant
    Dim i As Long
    Dim b As Long

    If IsMissing(scaleFactor) Then scaleFactor = Array(1, 0.9, 0.8, 0.8)
    Set oRec = CvTdRec("LCNAME", NAME, "DESC", "", "SCALE_FACTOR", CvNumList(scaleFactor), _
                       "NUM_LOADED_LANES", numLoadedLanes)
    Set col = New Collection
    If optPermit Then
        If IsMissing(permitVehicle) Or IsMissing(refLane) Or IsMissing(ecc) Or IsMissing(permitScale) Then
            MovingCaseIndia = "MVLDID: a permit case needs permitVehicle, refLane, ecc and permitScale"
            Exit Function
        End If
        oRec.Add "OPT_AUTO_LL", True
        oRec.Add "OPT_LC_FOR_PERMIT_LOAD", True
        oRec.Add "PERMIT_VEHICLE", permitVehicle
        oRec.Add "REF_LANE", refLane
        oRec.Add "ECCEN", ecc
        oRec.Add "PERMIT_SCALE_FACTOR", permitScale
    Else
        If IsMissing(subLoadItems) Then
            MovingCaseIndia = "MVLDID: subLoadItems are needed"
            Exit Function
        End If
        For i = LBound(subLoadItems) To UBound(subLoadItems)
            vRow = subLoadItems(i)
            b = LBound(vRow)
            If optAutoLL Then
                Set oSub = CvTdRec("SCALE_FACTOR", vRow(b), "VEHICLE_CLASS_1", vRow(b + 1), _
                                   "VEHICLE_CLASS_2", vRow(b + 2), "FOOTWAY", vRow(b + 3), _
                                   "CARRIAGE_WAY_WIDTH", IIf(numLoadedLanes = 1, 2.3, 0), _
                                   "CARRIAGE_WAY_LOADING", IIf(numLoadedLanes = 1, 4.903325, 0), _
                                   "SELECTED_LANES", CvNumList(vRow(b + 4)))
                If UBound(vRow) >= b + 5 Then
                    If Not IsNull(vRow(b + 5)) And Not IsEmpty(vRow(b + 5)) Then
                        oSub.Add "SELECTED_FOOTWAY_LANES", CvNumList(vRow(b + 5))
                    End If
                End If
            Else
                Set oSub = CvTdRec("SCALE_FACTOR", vRow(b), "MIN_NUM_LOADED_LANES", vRow(b + 1), _
                                   "MAX_NUM_LOADED_LANES", vRow(b + 2), "VEHICLE_CLASS_1", vRow(b + 3), _
                                   "SELECTED_LANES", CvNumList(vRow(b + 4)))
            End If
            col.Add oSub
        Next i
        oRec.Add "OPT_AUTO_LL", optAutoLL
        oRec.Add "OPT_LC_FOR_PERMIT_LOAD", False
        oRec.Add "SUB_LOAD_ITEMS", col
    End If
    MovingCaseIndia = CvTdPut("/db/MVLDID", caseId, oRec)
End Function

'   MovingCaseEurocode 1, "LM1", 1, False, Array(True, "LM1", "", Array("L1"), Array(), Array())
' subLoadItems: the list by load model (1..5) and
' useOptimization. Without it only the
' name, load model and optimization flag are written.
Public Function MovingCaseEurocode(ByVal caseId As Long, ByVal NAME As String, ByVal loadModel As Long, _
                                   Optional ByVal useOptimization As Boolean = False, _
                                   Optional ByVal subLoadItems As Variant, Optional ByVal DESC As String = "") As String
    Dim oRec As Object
    Dim s As Variant
    Dim b As Long

    Set oRec = CvTdRec("LCNAME", NAME, "DESC", DESC, "TYPE_LOADMODEL", loadModel, "OPT_AUTO_OPTIMIZE", useOptimization)
    If Not IsMissing(subLoadItems) Then
        s = subLoadItems
        b = LBound(s)
        If Not useOptimization Then
            Select Case loadModel
                Case 1
                    CvAddPairs oRec, "OPT_LEADING", s(b), "VHLNAME1", s(b + 1), "VHLNAME2", s(b + 2), _
                               "SLN_LIST", CvNumList(s(b + 3)), "SRA_LIST", CvNumList(s(b + 4)), "FLN_LIST", CvNumList(s(b + 5))
                Case 2
                    CvAddPairs oRec, "OPT_LEADING", s(b), "OPT_COMB", s(b + 1), "SUB_LOAD_LIST", CvEuroSubLoads(s(b + 2))
                Case 3
                    CvAddPairs oRec, "OPT_LEADING", s(b), "VHLNAME1", s(b + 1), "VHLNAME2", s(b + 2), _
                               "SLN_LIST", CvNumList(s(b + 3)), "SRA_LIST", CvNumList(s(b + 4))
                Case 4
                    CvAddPairs oRec, "OPT_LEADING", s(b), "VHLNAME1", s(b + 1), "VHLNAME2", s(b + 2), _
                               "SLN_LIST", CvNumList(s(b + 3)), "SRA_LIST", CvNumList(s(b + 4)), "STL_LIST", CvNumList(s(b + 5))
                Case 5
                    CvAddPairs oRec, "OPT_PSI_FACTOR", s(b), "OPT_COMB", s(b + 1)
                    CvAddFactors oRec, s(b + 2), s(b + 3)
                    oRec.Add "SUB_LOAD_LIST", CvEuroSubLoads(s(b + 4))
            End Select
        Else
            Select Case loadModel
                Case 1, 3
                    CvAddPairs oRec, "OPT_LEADING", s(b), "VHLNAME1", s(b + 1), "VHLNAME2", s(b + 2), _
                               "MINVHLDIST", s(b + 3), "OPTIMIZE_LANE_NAME", s(b + 4), "LOADEDLANE", s(b + 5), _
                               "SLN_LIST", CvNumList(s(b + 6))
                Case 2
                    CvAddPairs oRec, "OPT_LEADING", s(b), "OPT_COMB", s(b + 1), "MINVHLDIST", s(b + 2), _
                               "OPTIMIZE_LANE_NAME", s(b + 3), "MIN_NUM_VHL", s(b + 4), "MAX_NUM_VHL", s(b + 5), _
                               "OPTIMIZE_LIST", CvEuroOptList(s(b + 6))
                Case 4
                    CvAddPairs oRec, "OPT_LEADING", s(b), "VHLNAME1", s(b + 1), "VHLNAME2", s(b + 2), _
                               "MINVHLDIST", s(b + 3), "OPTIMIZE_LANE_NAME", s(b + 4), "LOADEDLANE", s(b + 5), _
                               "SLN_LIST", CvNumList(s(b + 6)), "STL_LIST", CvNumList(s(b + 7))
                Case 5
                    CvAddPairs oRec, "OPT_PSI_FACTOR", s(b), "OPT_COMB", s(b + 1)
                    CvAddFactors oRec, s(b + 2), s(b + 3)
                    CvAddPairs oRec, "MINVHLDIST", s(b + 4), "OPTIMIZE_LANE_NAME", s(b + 5), _
                               "MIN_NUM_VHL", s(b + 6), "MAX_NUM_VHL", s(b + 7), "OPTIMIZE_LIST", CvEuroOptList(s(b + 8))
            End Select
        End If
    End If
    MovingCaseEurocode = CvTdPut("/db/MVLDEU", caseId, oRec)
End Function

' ---- helpers ----

Private Function CvVehType(ByVal p As Variant) As Variant
    Select Case UCase$(CStr(p))
        Case "VL", "LOAD", "VEHICLE LOAD": CvVehType = "VL"
        Case "VC", "CLASS", "VEHICLE CLASS": CvVehType = "VC"
        Case Else: CvVehType = p
    End Select
End Function

Private Sub CvAddPairs(ByVal pRec As Object, ParamArray pKeyValues() As Variant)
    Dim i As Long
    For i = LBound(pKeyValues) To UBound(pKeyValues) - 1 Step 2
        pRec.Add CStr(pKeyValues(i)), pKeyValues(i + 1)
    Next i
End Sub

' SCALE_FACTOR1..3 and MULTI_FACTOR1..3 of load model 5.
Private Sub CvAddFactors(ByVal pRec As Object, ByVal pScale As Variant, ByVal pMulti As Variant)
    Dim i As Long
    For i = 0 To 2
        pRec.Add "SCALE_FACTOR" & (i + 1), pScale(LBound(pScale) + i)
    Next i
    For i = 0 To 2
        pRec.Add "MULTI_FACTOR" & (i + 1), pMulti(LBound(pMulti) + i)
    Next i
End Sub

' Rows name, scale factor, min lane type, max lane type, lane names.
Private Function CvEuroSubLoads(ByVal pRows As Variant) As Collection
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    Set col = New Collection
    For i = LBound(pRows) To UBound(pRows)
        b = LBound(pRows(i))
        col.Add CvTdRec("TYPE", 2, "NAME", pRows(i)(b), "SCALE_FACTOR", pRows(i)(b + 1), _
                        "MIN_LOAD_LANE_TYPE", pRows(i)(b + 2), "MAX_LOAD_LANE_TYPE", pRows(i)(b + 3), _
                        "SLN_LIST", CvNumList(pRows(i)(b + 4)))
    Next i
    Set CvEuroSubLoads = col
End Function

' Rows name, scale factor.
Private Function CvEuroOptList(ByVal pRows As Variant) As Collection
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    Set col = New Collection
    For i = LBound(pRows) To UBound(pRows)
        b = LBound(pRows(i))
        col.Add CvTdRec("TYPE", 2, "NAME", pRows(i)(b), "SCALE_FACTOR", pRows(i)(b + 1))
    Next i
    Set CvEuroOptList = col
End Function

' IRS vehicles of a vehicle type (and its full name), Empty when unknown.
Private Function CvIrsVehicles(ByVal pType As String, ByRef pFullName As String) As Variant
    Dim v() As String
    Dim i As Long
    Select Case pType
        Case "BG-1676"
            pFullName = "Broad Gauge-1676mm"
            CvIrsVehicles = Array("Modified B.G. Loading 1987-1", "Modified B.G. Loading 1987-2", _
                "B.G. Standard Loading 1926-M.L.", "B.G. Standard Loading 1926-B.L.", _
                "Revised B.G. Loading 1975-WG1+WG1", "Revised B.G. Loading 1975-WAM4A+WAM4A", _
                "Revised B.G. Loading 1975-Bo-Bo+Bo-Bo", "Revised B.G. Loading 1975-WAM4A", _
                "Revised B.G. Loading 1975-WAM4A+WAM4", "Revised B.G. Loading 1975-WAM4A+WDM2", _
                "25t Loading-2008 Combination 1", "25t Loading-2008 Combination 2", _
                "25t Loading-2008 Combination 3", "25t Loading-2008 Combination 4", _
                "25t Loading-2008 Combination 5", "DFC Loading Combination 1", "DFC Loading Combination 2", _
                "DFC Loading Combination 3", "DFC Loading Combination 4", "DFC Loading Combination 5")
        Case "MG-1000"
            pFullName = "Metre Gauge-1000mm"
            CvIrsVehicles = Array("2 Co-Co Locomotives", "2 Bo-Bo Locomotives", "MGML Loading of 1929", _
                                  "M.L.", "B.L.", "C.")
        Case "NG-762"
            pFullName = "Narrow Gauge-762mm"
            CvIrsVehicles = Array("Class H: B-B or Bo-Bo Type", "Class H: C-C or Co-Co Type", _
                "Class H: Steam (Zf/1)", "Class H: Diesel Electric", "Class A: B-B or Bo-Bo Type", _
                "Class A: C-C or Co-Co Type", "Class A: Diesel Mech./Elec.", _
                "Class A: Diesel Mech./Elec.(Articulated)", "Class A: DRG No. CSO/C-873", _
                "Class B: B-B or Bo-Bo Type", "Class B: Steam Engine (Tank)", _
                "Class B: Steam Engine (Tender)", "Class B: Diesel Electric")
        Case "HML"
            pFullName = "Heavy Mineral Loadings"
            ReDim v(0 To 16)
            For i = 0 To 16
                v(i) = "Train Formation No." & (i + 1)
            Next i
            CvIrsVehicles = v
        Case "FTB"
            pFullName = "Footbridge & Footpath"
            CvIrsVehicles = Array("Footbridge & Footpath")
        Case Else
            CvIrsVehicles = Empty
    End Select
End Function

' VEH_IN defaults for an IRS vehicle.
Private Function CvIrsDefaults(ByVal pType As String, ByVal pVehicle As String) As Object
    Dim dTractive As Double
    Dim n As Long
    Select Case pType
        Case "BG-1676"
            Select Case pVehicle
                Case "Modified B.G. Loading 1987-1", "Modified B.G. Loading 1987-2", _
                     "25t Loading-2008 Combination 4", "25t Loading-2008 Combination 5", _
                     "DFC Loading Combination 4", "DFC Loading Combination 5"
                    dTractive = 490.3325
                Case "25t Loading-2008 Combination 1", "DFC Loading Combination 1": dTractive = 617.81895
                Case "25t Loading-2008 Combination 2", "DFC Loading Combination 2": dTractive = 509.9458
                Case "25t Loading-2008 Combination 3", "DFC Loading Combination 3": dTractive = 823.7586
            End Select
            Set CvIrsDefaults = CvTdRec("TRACTIVE", dTractive, "BRAKE_LOCO_RATIO", 25, "BRAKE_TRAIN_RATIO", 13.4)
        Case "MG-1000"
            If pVehicle = "2 Co-Co Locomotives" Then dTractive = 313.8128
            If pVehicle = "2 Bo-Bo Locomotives" Then dTractive = 235.3596
            Set CvIrsDefaults = CvTdRec("TRACTIVE", dTractive, "BRAKE_LOCO_RATIO", 25, "BRAKE_TRAIN_RATIO", 13.4)
        Case "NG-762"
            Set CvIrsDefaults = CvTdRec("TRACTIVE", 0, "BRAKE_LOCO_RATIO", 25, "BRAKE_TRAIN_RATIO", 13.4)
        Case "HML"
            n = CLng(Mid$(pVehicle, Len("Train Formation No.") + 1))
            Select Case n
                Case 1 To 4: Set CvIrsDefaults = CvTdRec("TRACTIVE", 588.399, "BRAKE_LOCO", 245.16625)
                Case 8: Set CvIrsDefaults = CvTdRec("TRACTIVE", 298.61249, "BRAKE_LOCO", 215.7463)
                Case 9 To 11: Set CvIrsDefaults = CvTdRec("TRACTIVE", 397.169325, "BRAKE_LOCO", 114.737805)
                Case Else: Set CvIrsDefaults = CvTdRec("TRACTIVE", 441.29925, "BRAKE_LOCO", 245.16625)
            End Select
        Case Else
            Set CvIrsDefaults = CvTdRec("FOOTWAY_WIDTH", 3, "SPAN_LENGTH", 7.5)
    End Select
End Function

' VEH_EUROCODE defaults of a Eurocode vehicle type, Nothing when unknown.
' pSelect gets the selectable vehicle list (Empty when the type has none).
Private Function CvEuroDefaults(ByVal pStd As String, ByVal pType As String, ByRef pSelect As Variant) As Object
    Dim o As Object
    Dim v() As String
    Dim i As Long

    pSelect = Empty
    Select Case pStd & "|" & pType
        Case "RoadBridge|Load Model 1"
            Set o = CvTdRec("AMP_VALUES", Array(0.75, 0.4), "TANDEM_ADJUST_VALUES", Array(1, 1, 1), _
                            "UDL_ADJUST_VALUES", Array(1, 1, 1, 1))
        Case "RoadBridge|Load Model 2": Set o = CvTdRec("ADJUSTMENT", 0.75, "ADJUSTMENT2", 1)
        Case "RoadBridge|Load Model 4": Set o = CvTdRec("ADJUSTMENT", 0.75)
        Case "RoadBridge|Load Model 3"
            Set o = CvTdRec("LM3_LOADCASE1", True, "LM3_LOADCASE2", False, "DYNAMIC_FACTOR", True, "USER_INPUT", False)
            pSelect = Array("600/150", "900/150", "1200/150/200", "1500/150/200", "1800/150/200", _
                            "2400/200", "3000/200", "3600/200")
        Case "RoadBridge|Load Model 3 (UK NA)"
            Set o = CvTdRec("DYNAMIC_FACTOR", True, "USER_INPUT", False)
            pSelect = Array("SV 80", "SV 100", "SV 196", "SOV 250", "SOV 350", "SOV 450", "SOV 600")
        Case "FTB|Uniform load (Road bridge footway)": Set o = CvTdRec("ADJUSTMENT", 0.4, "FOOTWAY", 5)
        Case "FTB|Uniform load (Footbridge)", "FTB|Uniform load (Road bridge footway) UK NA"
            Set o = CvTdRec("ADJUSTMENT", 0.4)
        Case "FTB|Concentrated Load": Set o = New Dictionary
        Case "RoadBridgeFatigue|Fatigue Load Model 1"
            Set o = CvTdRec("AMP", 1, "TANDEM_ADJUST_VALUES", Array(1, 1, 1), "UDL_ADJUST_VALUES", Array(1, 1, 1, 1))
        Case "RoadBridgeFatigue|Fatigue Load Model 3 (Two Vehicle)": Set o = CvTdRec("AMP", 1, "INTERVAL", 31.6)
        Case "RoadBridgeFatigue|Fatigue Load Model 2 (280)", "RoadBridgeFatigue|Fatigue Load Model 2 (360)", _
             "RoadBridgeFatigue|Fatigue Load Model 2 (630)", "RoadBridgeFatigue|Fatigue Load Model 2 (560)", _
             "RoadBridgeFatigue|Fatigue Load Model 2 (610)", "RoadBridgeFatigue|Fatigue Load Model 3 (One Vehicle)", _
             "RoadBridgeFatigue|Fatigue Load Model 4 (200)", "RoadBridgeFatigue|Fatigue Load Model 4 (310)", _
             "RoadBridgeFatigue|Fatigue Load Model 4 (490)", "RoadBridgeFatigue|Fatigue Load Model 4 (390)", _
             "RoadBridgeFatigue|Fatigue Load Model 4 (450)"
            Set o = CvTdRec("AMP", 1)
        Case "RailTraffic|Load Model 71": Set o = CvRail(80, 0, 0.8, 80, 0, 0.8)
        Case "RailTraffic|Load Model SW/0": Set o = CvRail(133, 15, 5.3, 133, 15, 0)
        Case "RailTraffic|Load Model SW/2": Set o = CvRail(150, 25, 7, 150, 25, 0)
        Case "RailTraffic|Unloaded Train": Set o = CvRail(10, 0, 0, 0, 0, 0)
        Case "RailTraffic|HSLM B"
            Set o = CvTdRec("V_LOAD_FACTOR", 1, "LONGI_DIST", False, "ECCEN_VERT_LOAD", False, _
                            "HSLMB_NUM", 10, "HSLMB_FORCE", 170, "HSLMB_DIST", 3.5, _
                            "PHI_DYN_EFF1", 0, "PHI_DYN_EFF2", 0)
        Case "RailTraffic|HSLM A1 ~ HSLM A10"
            Set o = CvTdRec("V_LOAD_FACTOR", 1, "LONGI_DIST", False, "ECCEN_VERT_LOAD", False, _
                            "PHI_DYN_EFF1", 0, "PHI_DYN_EFF2", 0)
            ReDim v(0 To 9)
            For i = 0 To 9
                v(i) = "A" & (i + 1)
            Next i
            pSelect = v
    End Select
    Set CvEuroDefaults = o
End Function

Private Function CvRail(ByVal w1 As Double, ByVal dd1 As Double, ByVal d1 As Double, _
                        ByVal w2 As Double, ByVal dd2 As Double, ByVal d2 As Double) As Object
    Set CvRail = CvTdRec("W1", w1, "DD1", dd1, "D1", d1, "W2", w2, "DD2", dd2, "D2", d2, _
                         "V_LOAD_FACTOR", 1, "LONGI_DIST", False, "ECCEN_VERT_LOAD", False)
End Function


'==========================================================
' [19] Analysis - controls, settlement, story, force-deformation function, boundary change
'==========================================================
'  Analysis controls, settlement, story, MLFC, boundary change.
'  A control waits in the store and goes after the model (the eigenvalue control is sent once).
'    AnalysisEigen "LANCZOS", 10
'    SettlementGroup 1, "SG1", -0.01, Array(1, 2)
'    SettlementCase 1, "SET", Array("SG1")
'==========================================================

' Main control data (ACTL).
Public Function AnalysisMain(Optional ByVal ardc As Boolean = True, Optional ByVal anrc As Boolean = True, _
                             Optional ByVal iter As Long = 20, Optional ByVal tol As Double = 0.001, _
                             Optional ByVal csecf As Boolean = False, Optional ByVal trs As Boolean = True, _
                             Optional ByVal crbar As Boolean = False, Optional ByVal bmstress As Boolean = False, _
                             Optional ByVal clats As Boolean = False) As String
    AnalysisMain = CvTdPut("/db/ACTL", 1, CvTdRec("ARDC", ardc, "ANRC", anrc, "ITER", iter, "TOL", tol, _
                                                  "CSECF", csecf, "TRS", trs, "CRBAR", crbar, _
                                                  "BMSTRESS", bmstress, "CLATS", clats))
End Function

'   AnalysisPDelta Array(Array("DL", 1), Array("LL", 0.5))
' P-Delta: load case rows name, factor.
Public Function AnalysisPDelta(ByVal loadCases As Variant, Optional ByVal iter As Long = 5, _
                               Optional ByVal tol As Double = 0.00001) As String
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    Set col = New Collection
    For i = LBound(loadCases) To UBound(loadCases)
        b = LBound(loadCases(i))
        col.Add CvTdRec("LCNAME", loadCases(i)(b), "FACTOR", loadCases(i)(b + 1))
    Next i
    AnalysisPDelta = CvTdPut("/db/PDEL", 1, CvTdRec("ITER", iter, "TOL", tol, "PDEL_CASES", col))
End Function

'   AnalysisBuckling 5, Array(Array("DL", 1, 1), Array("LL", 1, 0))
' Buckling: load case rows name, factor, type (0 variable / 1 constant).
Public Function AnalysisBuckling(ByVal modeNum As Long, ByVal loadCases As Variant, _
                                 Optional ByVal optPositive As Boolean = True, _
                                 Optional ByVal factorFrom As Double = 0, Optional ByVal factorTo As Double = 0, _
                                 Optional ByVal sturmSeq As Boolean = False, _
                                 Optional ByVal axialOnly As Boolean = False) As String
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    Set col = New Collection
    For i = LBound(loadCases) To UBound(loadCases)
        b = LBound(loadCases(i))
        col.Add CvTdRec("LCNAME", loadCases(i)(b), "FACTOR", loadCases(i)(b + 1), "LOAD_TYPE", loadCases(i)(b + 2))
    Next i
    CvStoreSet "/db/BUCK", 1, CvTdRec("MODE_NUM", modeNum, "OPT_POSITIVE", optPositive, _
                                       "OPT_CONSIDER_AXIAL_ONLY", axialOnly, "LOAD_FACTOR_FROM", factorFrom, _
                                       "LOAD_FACTOR_TO", factorTo, "OPT_STURM_SEQ", sturmSeq, "ITEMS", col)
    AnalysisBuckling = ""
End Function

' Eigenvalue control. TYPE_ "EIGEN" (nFreq, nIter, nDim, tol),
' "LANCZOS" (nFreq, freqMin + freqMax for a range, bSturm) or
' "RITZ" (loadVectors rows: load case or ACCX / ACCY / ACCZ, number of
' generations; nGLLink: number of Gel-link vectors).
Public Function AnalysisEigen(Optional ByVal TYPE_ As String = "EIGEN", Optional ByVal nFreq As Long = 1, _
                              Optional ByVal nIter As Long = 20, Optional ByVal nDim As Long = 1, _
                              Optional ByVal tol As Double = 0.0000000001, Optional ByVal freqMin As Variant, _
                              Optional ByVal freqMax As Variant, Optional ByVal bSturm As Boolean = False, _
                              Optional ByVal loadVectors As Variant, Optional ByVal nGLLink As Long = 0) As String
    Dim oRec As Object
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    Dim sName As String

    TYPE_ = UCase$(TYPE_)
    Set oRec = CvTdRec("TYPE", TYPE_)
    Select Case TYPE_
        Case "EIGEN"
            CvAddPairs oRec, "iFREQ", nFreq, "iITER", nIter, "iDIM", nDim, "TOL", tol
        Case "LANCZOS"
            If IsMissing(freqMin) Or IsMissing(freqMax) Then
                CvAddPairs oRec, "iFREQ", nFreq, "bMINMAX", False, "FRMIN", 0, "FRMAX", 1600, "bSTRUM", bSturm
            Else
                CvAddPairs oRec, "iFREQ", nFreq, "bMINMAX", True, "FRMIN", freqMin, "FRMAX", freqMax, "bSTRUM", bSturm
            End If
        Case "RITZ"
            Set col = New Collection
            If Not IsMissing(loadVectors) Then
                For i = LBound(loadVectors) To UBound(loadVectors)
                    b = LBound(loadVectors(i))
                    sName = CStr(loadVectors(i)(b))
                    If sName = "ACCX" Or sName = "ACCY" Or sName = "ACCZ" Then
                        col.Add CvTdRec("KIND", "GROUND", "GROUND", sName, "iNOG", loadVectors(i)(b + 1))
                    Else
                        col.Add CvTdRec("KIND", "CASE", "CASE", sName, "iNOG", loadVectors(i)(b + 1))
                    End If
                Next i
            End If
            CvAddPairs oRec, "bINCNL", (nGLLink <> 0), "iGNUM", nGLLink, "vRITZ", col
        Case Else
            AnalysisEigen = "EIGV: TYPE_ must be EIGEN, LANCZOS or RITZ"
            Exit Function
    End Select
    AnalysisEigen = CvTdPut("/db/EIGV", 1, oRec)
End Function

' Settlement analysis control (SMCT).
Public Function AnalysisSettlement(Optional ByVal concurrentCalc As Boolean = True, _
                                   Optional ByVal concurrentLink As Boolean = True) As String
    AnalysisSettlement = CvTdPut("/db/SMCT", 1, CvTdRec("CONCURRENT_CALC", concurrentCalc, _
                                                        "CONCURRENT_LINK", concurrentLink))
End Function

' Heat of hydration analysis control (HHCT).
' finalStage False: otherStage is the stage to stop at.
' creepShrink: TYPE_ "CREEP" / "SHRINK" / "BOTH", creepMethod "GENERAL"
' (nIter, tol) or "EFFECTIVE MODULUS" (phi1, day1, phi2, day2).
' eval "CENTER" / "GAUSS" / "NODAL". selfWeightFactor (<= 0) with selfLoad.
Public Function AnalysisHeat(Optional ByVal finalStage As Boolean = True, Optional ByVal otherStage As String = "", _
                             Optional ByVal theta As Double = 0.5, Optional ByVal initTemp As Double = 20, _
                             Optional ByVal eval As String = "GAUSS", Optional ByVal creepShrink As Boolean = True, _
                             Optional ByVal TYPE_ As String = "BOTH", Optional ByVal creepMethod As String = "GENERAL", _
                             Optional ByVal nIter As Variant, Optional ByVal tol As Variant, _
                             Optional ByVal phi1 As Variant, Optional ByVal day1 As Variant, _
                             Optional ByVal phi2 As Variant, Optional ByVal day2 As Variant, _
                             Optional ByVal equivAge As Boolean = False, Optional ByVal selfLoad As Boolean = False, _
                             Optional ByVal selfWeightFactor As Double = 0) As String
    Dim oRec As Object
    Dim oItem As Object

    If Not finalStage And Len(otherStage) = 0 Then
        AnalysisHeat = "HHCT: otherStage is needed when finalStage is False"
        Exit Function
    End If
    If selfWeightFactor > 0 Then
        AnalysisHeat = "HHCT: selfWeightFactor must be 0 or less"
        Exit Function
    End If
    Set oRec = CvTdRec("FINAL_STAGE", finalStage, "STAGE_NAME", IIf(finalStage, "", otherStage), _
                       "THETA", theta, "INIT_TEMP", initTemp, "EVAL", eval, "OPT_USE_EQUI_AGE", equivAge, _
                       "OPT_INCL_SELF_WEIGHT", selfLoad, "OPT_IS_CREEP_SHRINKAGE", creepShrink)
    If selfLoad Then oRec.Add "SELF_WEIGHT_FACTOR", selfWeightFactor
    If creepShrink Then
        If creepMethod = "GENERAL" Then
            If IsMissing(nIter) Or IsMissing(tol) Then
                AnalysisHeat = "HHCT: nIter and tol are needed for the GENERAL creep method"
                Exit Function
            End If
            Set oItem = CvTdRec("TYPE", TYPE_, "CREEP_CALC_METHOD", 0, _
                                "M_GENERAL", CvTdRec("ITER", nIter, "TOL", tol))
        ElseIf creepMethod = "EFFECTIVE MODULUS" Then
            If IsMissing(phi1) Or IsMissing(day1) Or IsMissing(phi2) Or IsMissing(day2) Then
                AnalysisHeat = "HHCT: phi1, day1, phi2 and day2 are needed for EFFECTIVE MODULUS"
                Exit Function
            End If
            Set oItem = CvTdRec("TYPE", TYPE_, "CREEP_CALC_METHOD", 1, _
                                "M_EFF_MOD", CvTdRec("PHI1", phi1, "DAY1", day1, "PHI2", phi2, "DAY2", day2))
        Else
            AnalysisHeat = "HHCT: creepMethod must be GENERAL or EFFECTIVE MODULUS"
            Exit Function
        End If
        oRec.Add "ITEM", oItem
    End If
    AnalysisHeat = CvTdPut("/db/HHCT", 1, oRec)
End Function

'   SettlementGroup 1, "SG1", -0.01, Array(1, 2, 3)
Public Function SettlementGroup(ByVal groupId As Long, ByVal NAME As String, ByVal displacement As Double, _
                                ByVal nodeIds As Variant) As String
    CvStoreSet "/db/SMPT", groupId, CvTdRec("NAME", NAME, "SETTLE", displacement, "ITEMS", CvNumList(nodeIds))
    SettlementGroup = ""
End Function

'   SettlementCase 1, "SET", Array("SG1", "SG2"), 1, 1, 2
' Settlement load case: the groups, factor, min / max groups loaded at once.
Public Function SettlementCase(ByVal caseId As Long, ByVal NAME As String, ByVal groups As Variant, _
                               Optional ByVal factor As Double = 1, Optional ByVal minGroups As Long = 1, _
                               Optional ByVal maxGroups As Long = 1, Optional ByVal DESC As String = "") As String
    SettlementCase = CvTdPut("/db/SMLC", caseId, _
                             CvTdRec("NAME", NAME, "DESC", DESC, "FACTOR", factor, "MIN", minGroups, _
                                     "MAX", maxGroups, "ST_GROUPS", CvNumList(groups)))
End Function

'   Story 1, "1F", 0
'   Story 2, "2F", 3.5, floorWidthX:=20, floorWidthY:=12
Public Function Story(ByVal storyId As Long, ByVal NAME As String, ByVal level As Double, _
                      Optional ByVal floorWidthX As Double = 0, Optional ByVal floorWidthY As Double = 0, _
                      Optional ByVal floorCenterX As Double = 0, Optional ByVal floorCenterY As Double = 0, _
                      Optional ByVal windEccX As Double = 0, Optional ByVal windEccY As Double = 0, _
                      Optional ByVal seisAccEccX As Double = 0, Optional ByVal seisAccEccY As Double = 0, _
                      Optional ByVal seisInhEccX As Double = 0, Optional ByVal seisInhEccY As Double = 0, _
                      Optional ByVal seisTorAmpX As Double = 1, Optional ByVal seisTorAmpY As Double = 1, _
                      Optional ByVal bFloorDiaph As Boolean = True) As String
    Story = CvTdPut("/db/STOR", storyId, _
                    CvTdRec("STORY_NAME", NAME, "STORY_LEVEL", level, "bFLOOR_DIAPHRAGM", bFloorDiaph, _
                            "WIND_FLOOR_WIDTH_X", floorWidthX, "WIND_FLOOR_WIDTH_Y", floorWidthY, _
                            "WIND_CENTER_X", floorCenterX, "WIND_CENTER_Y", floorCenterY, _
                            "WIND_ECCENT_X", windEccX, "WIND_ECCENT_Y", windEccY, _
                            "SEIS_ACC_ECCENT_X", seisAccEccX, "SEIS_ACC_ECCENT_Y", seisAccEccY, _
                            "SEIS_INHERENT_ECCENT_X", seisInhEccX, "SEIS_INHERENT_ECCENT_Y", seisInhEccY, _
                            "SEIS_TORSIONAL_AMP_FACTOR_X", seisTorAmpX, "SEIS_TORSIONAL_AMP_FACTOR_Y", seisTorAmpY))
End Function

'   ForceDeformFunction 1, "F1", "FORCE", True, Array(Array(0, 0), Array(0.01, 100), Array(0.02, 120))
' Multi-linear force-deformation function (MLFC) for links. TYPE_ "FORCE" / ...
' points rows x, y.
Public Function ForceDeformFunction(ByVal funcId As Long, ByVal NAME As String, _
                                    Optional ByVal TYPE_ As String = "FORCE", Optional ByVal symm As Boolean = True, _
                                    Optional ByVal points As Variant) As String
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    If IsMissing(points) Then points = Array(Array(0, 0), Array(1, 1))
    Set col = New Collection
    For i = LBound(points) To UBound(points)
        b = LBound(points(i))
        col.Add CvTdRec("X", points(i)(b), "Y", points(i)(b + 1))
    Next i
    CvStoreSet "/db/MLFC", funcId, CvTdRec("NAME", NAME, "TYPE", TYPE_, "SYMM", symm, "FUNC_ID", 0, "ITEMS", col)
    ForceDeformFunction = ""
End Function

'   BoundaryChange Array(Array("BC1", "BG1")), stAssign:=Array(Array("DL", "BC1")), MV:="BC1"
' Boundary change assignment to load cases (BCCT).
' bSPT .. bCDOF: which boundary types are changed. boundaries rows:
' boundary change name, boundary group. stAssign rows: static load case,
' boundary change name - every static load case in NX or defined here is
' listed, "UNCHANGED" when not given (a name given here is always listed,
' e.g. a combination converted by UseLoadCombination). MV / SM / THRSEV / PO /
' THNS / ULAT: the boundary change name for those analyses.
Public Function BoundaryChange(Optional ByVal boundaries As Variant, Optional ByVal stAssign As Variant, _
                               Optional ByVal MV As String = "UNCHANGED", Optional ByVal SM As String = "UNCHANGED", _
                               Optional ByVal THRSEV As String = "UNCHANGED", Optional ByVal PO As String = "UNCHANGED", _
                               Optional ByVal THNS As String = "UNCHANGED", Optional ByVal ULAT As String = "UNCHANGED", _
                               Optional ByVal bSPT As Boolean = False, Optional ByVal bSPR As Boolean = False, _
                               Optional ByVal bGSPR As Boolean = False, Optional ByVal bCGLINK As Boolean = False, _
                               Optional ByVal bSSSF As Boolean = False, Optional ByVal bPSSF As Boolean = False, _
                               Optional ByVal bRLS As Boolean = False, Optional ByVal bWSSF As Boolean = False, _
                               Optional ByVal bESSF As Boolean = False, Optional ByVal bCDOF As Boolean = False) As String
    Dim colB As Collection
    Dim colL As Collection
    Dim oAssign As Object
    Dim oCases As Object
    Dim oNames As Object
    Dim vKey As Variant
    Dim i As Long
    Dim b As Long
    Dim sName As String
    Dim vType As Variant
    Dim vName As Variant

    Set colB = New Collection
    If Not IsMissing(boundaries) Then
        For i = LBound(boundaries) To UBound(boundaries)
            b = LBound(boundaries(i))
            Set oAssign = New Collection
            oAssign.Add boundaries(i)(b + 1)
            colB.Add CvTdRec("BGCNAME", boundaries(i)(b), "vBG", oAssign)
        Next i
    End If

    Set oAssign = New Dictionary
    If Not IsMissing(stAssign) Then
        For i = LBound(stAssign) To UBound(stAssign)
            b = LBound(stAssign(i))
            oAssign.Item(CStr(stAssign(i)(b))) = stAssign(i)(b + 1)
        Next i
    End If
    ' static load cases: those in NX, those defined
    ' here, then any name given in stAssign that is in neither
    Set oNames = New Dictionary
    Set oCases = CvReadItems("/db/STLD")
    If Not oCases Is Nothing Then
        For Each vKey In oCases.Keys
            oNames.Item(CStr(oCases.Item(vKey).Item("NAME"))) = True
        Next vKey
    End If
    Set oCases = CvStoreTable("STLD")
    For Each vKey In oCases.Keys
        oNames.Item(CStr(oCases.Item(vKey).Item("NAME"))) = True
    Next vKey
    For Each vKey In oAssign.Keys
        oNames.Item(CStr(vKey)) = True
    Next vKey
    Set colL = New Collection
    For Each vKey In oNames.Keys
        sName = CStr(vKey)
        If oAssign.Exists(sName) Then
            colL.Add CvTdRec("TYPE", "ST", "BGCNAME", oAssign.Item(sName), "LCNAME", sName)
        Else
            colL.Add CvTdRec("TYPE", "ST", "BGCNAME", "UNCHANGED", "LCNAME", sName)
        End If
    Next vKey
    vType = Array("MV", "SM", "THRSEV", "PO", "THNS", "ULAT")
    vName = Array(MV, SM, THRSEV, PO, THNS, ULAT)
    For i = 0 To 5
        colL.Add CvTdRec("TYPE", vType(i), "BGCNAME", vName(i))
    Next i

    BoundaryChange = CvTdPut("/db/BCCT", 1, _
                             CvTdRec("bSPT", bSPT, "bSPR", bSPR, "bGSPR", bGSPR, "bCGLINK", bCGLINK, _
                                     "bSSSF", bSSSF, "bPSSF", bPSSF, "bRLS", bRLS, "bWSSF", bWSSF, _
                                     "bESSF", bESSF, "bCDOF", bCDOF, "vBOUNDARY", colB, "vLOADANAL", colL))
End Function

' A record whose ITEMS are not numbered items (node lists, points):
' replaces what the key holds, no appending.
Private Sub CvStoreSet(ByVal pEndpoint As String, ByVal pKey As Variant, ByVal pRec As Object)
    Dim oTable As Object
    Set oTable = CvStoreTable(CvStoreName(pEndpoint))
    If oTable.Exists(CStr(pKey)) Then oTable.Remove CStr(pKey)
    oTable.Add CStr(pKey), pRec
End Sub


'==========================================================
' [20] Dynamic - response spectrum, time history
'==========================================================
'  The options of a case
'  (damping, modal combination, subsequent / initial load, nonlinear
'  iteration) are made by the RS... / TH... functions that return an
'  Object, and handed to the case:
'    RSFunctionUser 1, "KDS", Array(Array(0, 0.4), Array(0.5, 1), Array(3, 0.2))
'    RSCase 1, "RX", "XY", 0, functions:=Array("KDS"), damping:=RSDampModal(0.05)
'  Gravity left out: from the model length unit (9.806 m/s2).
'==========================================================

' ---- response spectrum ----

' spectralType "Normalized Accel" / "Acceleration" / "Velocity" / "Displacement".
' maxValue given: the function is scaled to that maximum instead of by scaling.
'   data rows: period, value
Public Function RSFunctionUser(ByVal funcId As Long, ByVal NAME As String, ByVal data As Variant, _
                               Optional ByVal spectralType As String = "Normalized Accel", _
                               Optional ByVal scaling As Double = 1, Optional ByVal maxValue As Variant, _
                               Optional ByVal gravity As Variant, Optional ByVal damping As Double = 0.05, _
                               Optional ByVal DESC As String = "") As String
    Dim oRec As Object
    Dim col As Collection
    Dim i As Long
    Dim b As Long

    Set oRec = CvRSHead(NAME, spectralType, damping, DESC, scaling, maxValue, gravity)
    Set col = New Collection
    For i = LBound(data) To UBound(data)
        b = LBound(data(i))
        col.Add CvTdRec("PERIOD", data(i)(b), "VALUE", data(i)(b + 1))
    Next i
    oRec.Add "aFUNC", col
    RSFunctionUser = CvTdPut("/db/SPFC", funcId, oRec)
End Function

' code "IS1893(2016)" / "IS1893(2002)" / "IRC:SP114(2018)", soil Hard / Medium / Soft,
' zone II / III / IV / V. DESC left out: default description.
Public Function RSFunctionIndia(ByVal funcId As Long, ByVal NAME As String, _
                                Optional ByVal code As String = "IS1893(2002)", Optional ByVal soil As String = "Hard", _
                                Optional ByVal zone As String = "IV", Optional ByVal impFactor As Variant = 1#, _
                                Optional ByVal RRF As Variant = 1.5, Optional ByVal maxPeriod As Double = 6, _
                                Optional ByVal spectralType As String = "Normalized Accel", _
                                Optional ByVal scaling As Double = 1, Optional ByVal maxValue As Variant, _
                                Optional ByVal gravity As Variant, Optional ByVal damping As Double = 0.05, _
                                Optional ByVal DESC As String = "") As String
    Dim oRec As Object
    Dim sCode As String
    Dim nSoil As Variant
    Dim nZone As Variant

    If Len(DESC) = 0 Then
        DESC = code & ": Zone = " & zone & "  |  Soil = " & soil & " | Damping = " & CvNumText(Round(damping * 100, 2)) & _
               " % | I = " & CvNumText(impFactor) & " | R = " & CvNumText(RRF) & " "
    End If
    Select Case code
        Case "IS1893(2016)": sCode = "IS1893(2016)"
        Case "IRC:SP114(2018)": sCode = "IRC:SP:114-2018"
        Case Else: sCode = "IS2002"
    End Select
    Select Case soil
        Case "Hard": nSoil = 0
        Case "Medium": nSoil = 1
        Case "Soft": nSoil = 2
        Case Else: nSoil = "IS2002"
    End Select
    Select Case zone
        Case "II": nZone = 0
        Case "III": nZone = 1
        Case "IV": nZone = 2
        Case "V": nZone = 3
        Case Else: nZone = "IS2002"
    End Select
    Set oRec = CvRSHead(NAME, spectralType, damping, DESC, scaling, maxValue, gravity)
    oRec.Add "STR", CvTdRec("SPEC_CODE", sCode)
    oRec.Add "OPT", CvTdRec("SOILCLASS", nSoil, "iSEISZONE", nZone)
    oRec.Add "VAL", CvTdRec("DP", damping * 100, "PERIOD", maxPeriod, "IE", impFactor, "R_", RRF, "DPFAC", 1)
    oRec.Add "CALC_OPT", True
    RSFunctionIndia = CvTdPut("/db/SPFC", funcId, oRec)
End Function

' Peru E.030: zone 1..4, soil S0..S3, usage A1 / A2 / B / C. The spectrum is
' calculated every 0.05 s up to maxPeriod.
Public Function RSFunctionPeru(ByVal funcId As Long, ByVal NAME As String, Optional ByVal zone As Long = 1, _
                               Optional ByVal soil As String = "S0", Optional ByVal usage As String = "A1", _
                               Optional ByVal RRF As Variant = 1.5, Optional ByVal maxPeriod As Double = 6, _
                               Optional ByVal spectralType As String = "Normalized Accel", _
                               Optional ByVal scaling As Double = 1, Optional ByVal maxValue As Variant, _
                               Optional ByVal gravity As Variant, Optional ByVal damping As Double = 0.05, _
                               Optional ByVal DESC As String = "") As String
    Dim oRec As Object
    Dim col As Collection
    Dim z As Double
    Dim s As Double
    Dim tp As Double
    Dim tl As Double
    Dim u As Double
    Dim c As Double
    Dim t As Double
    Dim n As Long
    Dim i As Long

    If Len(DESC) = 0 Then
        DESC = "Espectro Peru E.030 : Zone = " & zone & "  |  Soil = " & soil & " |  Usage Cat = " & usage & _
               "  |  R = " & CvNumText(RRF) & " | Damping = " & CvNumText(Round(damping * 100, 2)) & " % "
    End If
    z = Array(0, 0.1, 0.25, 0.35, 0.45)(zone)
    Select Case soil
        Case "S0": s = 0.8: tp = 0.3: tl = 3
        Case "S1": s = 1: tp = 0.4: tl = 2.5
        Case "S2": s = Array(0, 1.6, 1.2, 1.15, 1.05)(zone): tp = 0.6: tl = 2
        Case "S3": s = Array(0, 2, 1.4, 1.2, 1.1)(zone): tp = 1: tl = 1.6
    End Select
    Select Case usage
        Case "A1", "A2": u = 1.5
        Case "B": u = 1.3
        Case Else: u = 1
    End Select

    Set oRec = CvRSHead(NAME, spectralType, damping, DESC, scaling, maxValue, gravity)
    Set col = New Collection
    n = -Int(-(maxPeriod / 0.05))
    For i = 0 To n - 1
        t = i * 0.05
        If t < tp Then
            c = 2.5
        ElseIf t <= tl Then
            c = 2.5 * (tp / t)
        Else
            c = 2.5 * ((tp * tl) / (t ^ 2))
        End If
        col.Add CvTdRec("PERIOD", Round(t, 3), "VALUE", Round((z * u * c * s) / RRF, 4))
    Next i
    oRec.Add "aFUNC", col
    RSFunctionPeru = CvTdPut("/db/SPFC", funcId, oRec)
End Function

'   RSCase 1, "RX", "XY", 0, functions:=Array("KDS"), damping:=RSDampModal(0.05), comb:=RSModalComb("SRSS")
' direction "XY" / "Z", interp "LINEAR" / "LOG". comb / damping: objects
' from RSModalComb / RSDamp... (left out: not written).
Public Function RSCase(ByVal caseId As Long, ByVal NAME As String, Optional ByVal direction As String = "XY", _
                       Optional ByVal angle As Double = 0, Optional ByVal scaleValue As Double = 1, _
                       Optional ByVal periodFactor As Double = 1, Optional ByVal functions As Variant, _
                       Optional ByVal interp As String = "LINEAR", Optional ByVal comb As Object, _
                       Optional ByVal damping As Object, Optional ByVal DESC As String = "", _
                       Optional ByVal bDampCorrection As Boolean = False) As String
    Dim oRec As Object
    If IsMissing(functions) Then functions = Array("RS_FUNC")
    Set oRec = CvTdRec("NAME", NAME, "DIR", direction, "ANGLE", angle, "SCALE", scaleValue, "PMFT", periodFactor, _
                       "INTERP", interp, "aFUNCNAME", CvNumList(functions), "DESC", DESC, "bDAMP", False, _
                       "bCDAMP", bDampCorrection)
    If Not damping Is Nothing Then CvMerge oRec, damping
    If Not comb Is Nothing Then CvMerge oRec, comb
    RSCase = CvTdPut("/db/SPLC", caseId, oRec)
End Function

' Modal combination: combType "CQC" / "SRSS" / "ABS" / "Linear".
' modeFactors: one factor per mode (0 = mode not used).
Public Function RSModalComb(Optional ByVal combType As String = "CQC", Optional ByVal bAddSign As Boolean = False, _
                            Optional ByVal signType As Long = 0, Optional ByVal modeFactors As Variant) As Object
    Dim o As Object
    Dim col As Collection
    Dim i As Long
    Set o = CvTdRec("COMTYPE", combType, "bADDSIGN", bAddSign, "iSIGNTYPE", signType, "bMODE", Not IsMissing(modeFactors))
    If Not IsMissing(modeFactors) Then
        Set col = New Collection
        For i = LBound(modeFactors) To UBound(modeFactors)
            If modeFactors(i) <> 0 Then
                col.Add CvTdRec("bUSE", True, "MSFACTOR", modeFactors(i))
            Else
                col.Add CvTdRec("bUSE", False, "MSFACTOR", 0)
            End If
        Next i
        o.Add "aUSEMODE", col
    End If
    Set RSModalComb = o
End Function

' Damping for a response spectrum case: one ratio for all modes.
Public Function RSDampModal(Optional ByVal ratio As Double = 0.05) As Object
    Set RSDampModal = CvTdRec("bDAMP", True, "iMDTYPE", 1, "DALL", ratio)
End Function

' Mass / stiffness proportional damping. inpType 1: massCoef / stiffCoef
' (left out: not used); inpType 2: from freq1 / damp1, freq2 / damp2.
Public Function RSDampMassStiffness(Optional ByVal inpType As Long = 1, Optional ByVal massCoef As Variant, _
                                    Optional ByVal stiffCoef As Variant, Optional ByVal freq1 As Double = 0, _
                                    Optional ByVal damp1 As Double = 0, Optional ByVal freq2 As Double = 0, _
                                    Optional ByVal damp2 As Double = 0) As Object
    Set RSDampMassStiffness = CvMassStiff(inpType, massCoef, stiffCoef, 1, freq1, damp1, freq2, damp2)
End Function

Public Function RSDampStrainEnergy() As Object
    Set RSDampStrainEnergy = CvTdRec("bDAMP", True, "iMDTYPE", 3)
End Function

' ---- time history ----

'   THFunction 1, "EQ1", Array(Array(0, 0), Array(0.01, 0.12)), "Normalized Accel"
' dataType "Normalized Accel" / "Acceleration" / "Force" / "Moment" / "Normal".
Public Function THFunction(ByVal funcId As Long, ByVal NAME As String, ByVal data As Variant, _
                           Optional ByVal dataType As String = "Normal", Optional ByVal scaling As Double = 1, _
                           Optional ByVal maxValue As Variant, Optional ByVal gravity As Variant, _
                           Optional ByVal DESC As String = "") As String
    Dim oRec As Object
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    Dim nType As Long

    Select Case dataType
        Case "Acceleration": nType = 2
        Case "Force": nType = 3
        Case "Moment": nType = 4
        Case "Normal": nType = 5
        Case Else: nType = 1
    End Select
    Set oRec = CvTdRec("NAME", NAME, "FUNCTYPE", 1, "iTYPE", nType, "DESC", DESC)
    If IsMissing(maxValue) Then
        CvAddPairs oRec, "iMETHOD", 0, "SCALE", scaling
    Else
        CvAddPairs oRec, "iMETHOD", 1, "SCALE", maxValue
    End If
    If IsMissing(gravity) Then gravity = CvGravity()
    oRec.Add "GRAV", gravity
    Set col = New Collection
    For i = LBound(data) To UBound(data)
        b = LBound(data(i))
        col.Add CvTdRec("TIME", data(i)(b), "VALUE", data(i)(b + 1))
    Next i
    oRec.Add "aFUNCDATA", col
    THFunction = CvTdPut("/db/THFC", funcId, oRec)
End Function

' Time history load cases (THIS). subsequent: THInitialLoad(...) or
' THSubsequent(...); damping: THDamp...(...); nlIter: THNonlinearIter(...).
' thType (linear modal) "Transient" / "Periodic".
Public Function THCaseLinearModal(ByVal caseId As Long, ByVal NAME As String, _
                                  Optional ByVal thType As String = "Transient", Optional ByVal endTime As Double = 1, _
                                  Optional ByVal timeInc As Double = 0.01, Optional ByVal stepOut As Long = 1, _
                                  Optional ByVal subsequent As Object, _
                                  Optional ByVal damping As Object) As String
    Dim oRec As Object
    Set oRec = CvTHCase(NAME, 1, 1, IIf(UCase$(thType) = "TRANSIENT", 1, 2), endTime, timeInc, stepOut, subsequent, damping)
    THCaseLinearModal = CvTdPut("/db/THIS", caseId, oRec)
End Function

Public Function THCaseNonlinearModal(ByVal caseId As Long, ByVal NAME As String, _
                                     Optional ByVal endTime As Double = 1, Optional ByVal timeInc As Double = 0.01, _
                                     Optional ByVal stepOut As Long = 1, Optional ByVal subsequent As Object, _
                                     Optional ByVal damping As Object, _
                                     Optional ByVal nlIter As Object) As String
    Dim oRec As Object
    Set oRec = CvTHCase(NAME, 2, 1, 1, endTime, timeInc, stepOut, subsequent, damping)
    oRec.Add "DMUPDATE", False
    If Not nlIter Is Nothing Then CvMerge oRec, nlIter
    THCaseNonlinearModal = CvTdPut("/db/THIS", caseId, oRec)
End Function

' integration "LINEAR" / "CONSTANT" or Array(gamma, beta) (Newmark).
Public Function THCaseLinearDirect(ByVal caseId As Long, ByVal NAME As String, _
                                   Optional ByVal endTime As Double = 1, Optional ByVal timeInc As Double = 0.01, _
                                   Optional ByVal stepOut As Long = 1, Optional ByVal subsequent As Object, _
                                   Optional ByVal damping As Object, _
                                   Optional ByVal integration As Variant = "LINEAR") As String
    Dim oRec As Object
    Set oRec = CvTHCase(NAME, 1, 2, 1, endTime, timeInc, stepOut, subsequent, damping)
    CvTHIntegration oRec, integration
    THCaseLinearDirect = CvTdPut("/db/THIS", caseId, oRec)
End Function

' dmUpdate "YES" / "NO": update the damping matrix.
Public Function THCaseNonlinearDirect(ByVal caseId As Long, ByVal NAME As String, _
                                      Optional ByVal endTime As Double = 1, Optional ByVal timeInc As Double = 0.01, _
                                      Optional ByVal stepOut As Long = 1, Optional ByVal subsequent As Object, _
                                      Optional ByVal damping As Object, _
                                      Optional ByVal integration As Variant = "LINEAR", _
                                      Optional ByVal nlIter As Object, _
                                      Optional ByVal dmUpdate As String = "NO") As String
    Dim oRec As Object
    Set oRec = CvTHCase(NAME, 2, 2, 1, endTime, timeInc, stepOut, subsequent, damping)
    oRec.Add "DMUPDATE", (UCase$(dmUpdate) = "YES")
    CvTHIntegration oRec, integration
    If Not nlIter Is Nothing Then CvMerge oRec, nlIter
    THCaseNonlinearDirect = CvTdPut("/db/THIS", caseId, oRec)
End Function

' Nonlinear static. Increment control: masterNode = Array(node, dir, inc),
' else globalDisp <> 0 (max translation), else the load scale factor.
Public Function THCaseNonlinearStatic(ByVal caseId As Long, ByVal NAME As String, _
                                      Optional ByVal endTime As Double = 1, Optional ByVal incSteps As Long = 1, _
                                      Optional ByVal subsequent As Object, _
                                      Optional ByVal loadScale As Double = 1, Optional ByVal globalDisp As Double = 0, _
                                      Optional ByVal masterNode As Variant, Optional ByVal loadOutput As Boolean = False, _
                                      Optional ByVal nlIter As Object) As String
    Dim oRec As Object
    Dim oCommon As Object
    Dim b As Long

    Set oCommon = CvTdRec("NAME", NAME, "DESC", "", "iATYPE", 2, "iAMETHOD", 3, "iTHTYPE", 1, _
                          "ENDTIME", endTime, "iISTEP", incSteps, "iOUT", 1, "INITLOAD", 0, "INITMETHOD", "ORDER")
    CvTHSubseq oCommon, subsequent
    Set oRec = CvTdRec("COMMON", oCommon, "bCUMULATE", loadOutput, "DMUPDATE", False)
    Dim bMaster As Boolean
    If Not IsMissing(masterNode) Then bMaster = (UBound(masterNode) - LBound(masterNode) = 2)
    If bMaster Then
        b = LBound(masterNode)
        CvAddPairs oRec, "iINCCTRL", 1, "iCTRL", 1, "MNODE", masterNode(b), "MDIR", masterNode(b + 1), _
                   "TINC", masterNode(b + 2)
    ElseIf globalDisp <> 0 Then
        CvAddPairs oRec, "iINCCTRL", 1, "iCTRL", 0, "TINC", globalDisp
    Else
        CvAddPairs oRec, "iINCCTRL", 0, "SCALE", loadScale
    End If
    If Not nlIter Is Nothing Then
        CvMerge oRec, nlIter
        If oRec.Exists("MINSSS") Then oRec.Remove "MINSSS"
    End If
    THCaseNonlinearStatic = CvTdPut("/db/THIS", caseId, oRec)
End Function

' Initial load control for a case (instead of THSubsequent).
Public Function THInitialLoad(Optional ByVal useInitialLoad As Boolean = True, Optional ByVal cumulateDVA As Boolean = False, _
                              Optional ByVal keepFinalLoads As Boolean = False, _
                              Optional ByVal geomNonlinear As Boolean = False) As Object
    Dim o As Object
    Set o = CvTdRec("iGEOM", IIf(geomNonlinear, 1, 0), "bSUBSEQ", False, "INITMETHOD", "INIT", _
                    "INITLOAD", IIf(useInitialLoad, 0, 1))
    If useInitialLoad Then
        o.Add "bDVA", cumulateDVA
        o.Add "bKEEP", keepFinalLoads
    End If
    Set THInitialLoad = o
End Function

' Subsequent to: loadCase = Array(type, name[, cumulate DVA, keep loads])
' (type "ST", "TH", ...), or the initial element forces table, or the
' initial forces for geometric stiffness.
Public Function THSubsequent(Optional ByVal loadCase As Variant, Optional ByVal initElemForces As Boolean = False, _
                             Optional ByVal initGeomStiffness As Boolean = False, _
                             Optional ByVal geomNonlinear As Boolean = False) As Object
    Dim o As Object
    Dim b As Long
    Dim n As Long
    Set o = CvTdRec("iGEOM", IIf(geomNonlinear, 1, 0))
    If IsMissing(loadCase) And Not initElemForces And Not initGeomStiffness Then
        o.Add "bSUBSEQ", False
        Set THSubsequent = o
        Exit Function
    End If
    o.Add "bSUBSEQ", True
    If Not IsMissing(loadCase) Then
        b = LBound(loadCase)
        n = UBound(loadCase) - b + 1
        CvAddPairs o, "SUBSEQ", 0, "LCTYPE", loadCase(b), "CASE", loadCase(b + 1)
        If loadCase(b) = "TH" Then
            If n > 2 Then o.Add "bDVA", loadCase(b + 2) Else o.Add "bDVA", False
            If n > 3 Then o.Add "bKEEP", loadCase(b + 3) Else o.Add "bKEEP", False
        End If
    ElseIf initElemForces Then
        o.Add "SUBSEQ", 1
    Else
        o.Add "SUBSEQ", 2
    End If
    Set THSubsequent = o
End Function

' Nonlinear iteration control. Norms left out are not checked (0.001 written).
' rungeKutta "FEHLBERG" / "CASHKARP".
Public Function THNonlinearIter(Optional ByVal maxIter As Long = 10, Optional ByVal minStepSize As Double = 0.00001, _
                                Optional ByVal maxSubSteps As Long = 10, Optional ByVal dispNorm As Variant, _
                                Optional ByVal forceNorm As Variant, Optional ByVal energyNorm As Variant, _
                                Optional ByVal lineSearchIter As Variant, Optional ByVal rungeKutta As String = "FEHLBERG", _
                                Optional ByVal tol As Double = 0.00000001, _
                                Optional ByVal checkConvergence As Boolean = True) As Object
    Set THNonlinearIter = CvTdRec("bITER", True, "bCONV", checkConvergence, "iMAXITER", maxIter, _
                                  "iMSTEP", maxSubSteps, "bDN", Not IsMissing(dispNorm), _
                                  "DN", CvOr(dispNorm, 0.001), "bFN", Not IsMissing(forceNorm), _
                                  "FN", CvOr(forceNorm, 0.001), "bEN", Not IsMissing(energyNorm), _
                                  "EN", CvOr(energyNorm, 0.001), _
                                  "iRKM", IIf(UCase$(Trim$(rungeKutta)) = "CASHKARP", 1, 0), "dTOL", tol, _
                                  "bULSM", Not IsMissing(lineSearchIter), "ULSM", CvOr(lineSearchIter, 5), _
                                  "MINSSS", minStepSize)
End Function

' Modal damping for a time history case; modeOverrides rows mode, ratio.
Public Function THDampModal(Optional ByVal ratio As Double = 0.05, Optional ByVal modeOverrides As Variant) As Object
    Dim o As Object
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    Set o = CvTdRec("bDAMP", True, "iMDTYPE", 1, "DALL", ratio)
    If Not IsMissing(modeOverrides) Then
        Set col = New Collection
        For i = LBound(modeOverrides) To UBound(modeOverrides)
            b = LBound(modeOverrides(i))
            col.Add CvTdRec("iMODE", modeOverrides(i)(b), "DAMPING", modeOverrides(i)(b + 1))
        Next i
        If col.Count > 0 Then o.Add "aDAMP", col
    End If
    Set THDampModal = o
End Function

' Mass / stiffness proportional damping for a time history case.
' inpType 2: from freq1 / freq2 (or period1 / period2) with damp1 / damp2.
Public Function THDampMassStiffness(Optional ByVal inpType As Long = 1, Optional ByVal massCoef As Variant, _
                                    Optional ByVal stiffCoef As Variant, Optional ByVal freq1 As Variant, _
                                    Optional ByVal damp1 As Double = 0, Optional ByVal freq2 As Variant, _
                                    Optional ByVal damp2 As Double = 0, Optional ByVal period1 As Variant, _
                                    Optional ByVal period2 As Variant) As Object
    If Not IsMissing(period1) Or Not IsMissing(period2) Then
        Set THDampMassStiffness = CvMassStiff(inpType, massCoef, stiffCoef, 2, CvOr(period1, 0), damp1, _
                                              CvOr(period2, 0), damp2)
    Else
        Set THDampMassStiffness = CvMassStiff(inpType, massCoef, stiffCoef, 1, CvOr(freq1, 0), damp1, _
                                              CvOr(freq2, 0), damp2)
    End If
End Function

Public Function THDampStrainEnergy() As Object
    Set THDampStrainEnergy = CvTdRec("bDAMP", True, "iMDTYPE", 3)
End Function

Public Function THDampElementMassStiffness() As Object
    Set THDampElementMassStiffness = CvTdRec("bDAMP", True, "iMDTYPE", 4)
End Function

'   THGroundAccel 1, "TH1", 0, "EQ1", 1
' Ground acceleration (THGA): function / scale / arrival time per direction.
Public Function THGroundAccel(ByVal gaId As Long, ByVal thCase As String, ByVal angle As Double, _
                              Optional ByVal funcX As Variant, Optional ByVal scaleX As Double = 1, _
                              Optional ByVal timeX As Double = 0, Optional ByVal funcY As Variant, _
                              Optional ByVal scaleY As Double = 1, Optional ByVal timeY As Double = 0, _
                              Optional ByVal funcZ As Variant, Optional ByVal scaleZ As Double = 1, _
                              Optional ByVal timeZ As Double = 0) As String
    Dim oRec As Object
    Set oRec = CvTdRec("NAME", thCase)
    CvTHDir oRec, "X", funcX, scaleX, timeX
    CvTHDir oRec, "Y", funcY, scaleY, timeY
    CvTHDir oRec, "Z", funcZ, scaleZ, timeZ
    oRec.Add "ANGLE", angle
    THGroundAccel = CvTdPut("/db/THGA", gaId, oRec)
End Function

' Time varying static load (THSL).
Public Function THStaticLoad(ByVal slId As Long, ByVal thCase As String, ByVal staticCase As String, _
                             ByVal funcName As String, ByVal scaleValue As Double, ByVal arrivalTime As Double) As String
    THStaticLoad = CvTdPut("/db/THSL", slId, CvTdRec("THIS_LCNAME", thCase, "SLOAD", staticCase, _
                                                     "THIS_FUNCNAME", funcName, "ATIME", arrivalTime, "SCALE", scaleValue))
End Function

'   THNodalLoad Array(5, 6), "TH1", "F1", "FZ", 0, 1
' Dynamic nodal load (THNL).
Public Function THNodalLoad(ByVal nodeIds As Variant, ByVal thCase As String, ByVal funcName As String, _
                            ByVal direction As String, ByVal arrivalTime As Double, ByVal scaleValue As Double) As String
    THNodalLoad = CvItemsPut("/db/THNL", nodeIds, _
                             CvTdRec("ID", 1, "THLCNAME", thCase, "FUNC_NAME", funcName, "DIR", direction, _
                                     "ARRIVAL_TIME", arrivalTime, "SCALE_FACTOR", scaleValue), True)
End Function

' Multiple support excitation (THMS) at nodes.
Public Function THSupportExcitation(ByVal nodeIds As Variant, ByVal thCase As String, Optional ByVal angle As Double = 0, _
                                    Optional ByVal funcX As Variant, Optional ByVal scaleX As Variant, _
                                    Optional ByVal timeX As Double = 0, Optional ByVal funcY As Variant, _
                                    Optional ByVal scaleY As Variant, Optional ByVal timeY As Double = 0, _
                                    Optional ByVal funcZ As Variant, Optional ByVal scaleZ As Variant, _
                                    Optional ByVal timeZ As Double = 0) As String
    Dim oRec As Object
    Set oRec = CvTdRec("ID", 1, "LCNAME", thCase, "ANGLE", angle)
    CvTHDir oRec, "X", funcX, CvOr(scaleX, Null), timeX
    CvTHDir oRec, "Y", funcY, CvOr(scaleY, Null), timeY
    CvTHDir oRec, "Z", funcZ, CvOr(scaleZ, Null), timeZ
    THSupportExcitation = CvItemsPut("/db/THMS", nodeIds, oRec, True)
End Function

' ---- helpers ----

' Common head of a spectrum function (SPFC).
Private Function CvRSHead(ByVal pName As String, ByVal pType As String, ByVal pDamp As Double, ByVal pDesc As String, _
                          ByVal pScale As Double, ByVal pMax As Variant, ByVal pGrav As Variant) As Object
    Dim o As Object
    Dim nType As Long
    Select Case pType
        Case "Acceleration": nType = 2
        Case "Velocity": nType = 3
        Case "Displacement": nType = 4
        Case Else: nType = 1
    End Select
    Set o = CvTdRec("NAME", pName, "iTYPE", nType, "DRATIO", pDamp, "DESC", pDesc)
    If IsMissing(pMax) Then
        CvAddPairs o, "iMETHOD", 0, "SCALE", pScale
    Else
        CvAddPairs o, "iMETHOD", 1, "SCALE", pMax
    End If
    If nType = 1 Then
        If IsMissing(pGrav) Then pGrav = CvGravity()
        o.Add "GRAV", pGrav
    End If
    Set CvRSHead = o
End Function

Private Function CvMassStiff(ByVal pType As Long, ByVal pMass As Variant, ByVal pStiff As Variant, ByVal pCalc As Long, _
                             ByVal f1 As Variant, ByVal d1 As Double, ByVal f2 As Variant, ByVal d2 As Double) As Object
    Dim o As Object
    Set o = CvTdRec("bDAMP", True, "iMDTYPE", 2, "iCOEF", pType, "bMASSP", Not IsMissing(pMass), _
                    "bSTIFFP", Not IsMissing(pStiff))
    If pType = 2 Then
        CvAddPairs o, "iCALC", pCalc, "FP1", f1, "DR1", d1, "FP2", f2, "DR2", d2
        o.Item("bMASSP") = True
        o.Item("bSTIFFP") = True
    Else
        CvAddPairs o, "MASSC", CvOr(pMass, 0), "STIFFC", CvOr(pStiff, 0)
    End If
    Set CvMassStiff = o
End Function

' COMMON block of a time history case, and the damping keys at the top.
Private Function CvTHCase(ByVal pName As String, ByVal pAType As Long, ByVal pAMethod As Long, ByVal pTHType As Long, _
                          ByVal pEnd As Double, ByVal pInc As Double, ByVal pOut As Long, _
                          ByVal pSub As Object, ByVal pDamp As Object) As Object
    Dim oCommon As Object
    Dim oRec As Object
    Dim k As Variant

    Set oCommon = CvTdRec("NAME", pName, "DESC", "", "iATYPE", pAType, "iAMETHOD", pAMethod, "iTHTYPE", pTHType, _
                          "ENDTIME", pEnd, "INC", pInc, "iOUT", pOut, "INITLOAD", 0, "INITMETHOD", "ORDER")
    CvTHSubseq oCommon, pSub
    Set oRec = CvTdRec("COMMON", oCommon)
    If Not pDamp Is Nothing Then
        For Each k In pDamp.Keys
            If k = "iMDTYPE" Then
                oCommon.Item("iMDTYPE") = pDamp.Item(k)
            ElseIf k <> "bDAMP" Then
                If IsObject(pDamp.Item(k)) Then Set oRec.Item(k) = pDamp.Item(k) Else oRec.Item(k) = pDamp.Item(k)
            End If
        Next k
    End If
    Set CvTHCase = oRec
End Function

Private Sub CvTHSubseq(ByVal pCommon As Object, ByVal pSub As Object)
    If pSub Is Nothing Then
        pCommon.Item("bSUBSEQ") = False
        pCommon.Item("iGEOM") = 0
    Else
        CvMerge pCommon, pSub
    End If
End Sub

Private Sub CvTHIntegration(ByVal pRec As Object, ByVal pInt As Variant)
    If IsArray(pInt) Then
        pRec.Add "iNMM", 3
        pRec.Add "GAMMA", pInt(LBound(pInt))
        pRec.Add "BETA", pInt(LBound(pInt) + 1)
    ElseIf UCase$(CStr(pInt)) = "CONSTANT" Then
        pRec.Add "iNMM", 1
    ElseIf UCase$(CStr(pInt)) = "LINEAR" Then
        pRec.Add "iNMM", 2
    End If
End Sub

' FUNCx / SCALEx / ATIMEx - an empty function when none is given.
Private Sub CvTHDir(ByVal pRec As Object, ByVal pDir As String, ByVal pFunc As Variant, _
                    ByVal pScale As Variant, ByVal pTime As Double)
    If IsMissing(pFunc) Then
        pRec.Add "FUNC" & pDir, ""
        pRec.Add "SCALE" & pDir, 1
        pRec.Add "ATIME" & pDir, 1
    Else
        pRec.Add "FUNC" & pDir, pFunc
        pRec.Add "SCALE" & pDir, pScale
        pRec.Add "ATIME" & pDir, pTime
    End If
End Sub

' Copy every key of pFrom into pTo (later wins).
Private Sub CvMerge(ByVal pTo As Object, ByVal pFrom As Object)
    Dim k As Variant
    For Each k In pFrom.Keys
        If IsObject(pFrom.Item(k)) Then
            Set pTo.Item(k) = pFrom.Item(k)
        Else
            pTo.Item(k) = pFrom.Item(k)
        End If
    Next k
End Sub

Private Function CvOr(ByVal p As Variant, ByVal pDefault As Variant) As Variant
    If IsMissing(p) Then CvOr = pDefault Else CvOr = p
End Function

' Gravity in the model length unit: the UNIT defined here, else the one in NX, else m.
Private Function CvGravity() As Double
    Dim sDist As String
    Dim oUnit As Object
    Dim oJson As Object

    Set oUnit = StoreGet("/db/UNIT", 1)
    If Not oUnit Is Nothing Then
        If oUnit.Exists("DIST") Then sDist = UCase$(CStr(oUnit.Item("DIST")))
    End If
    If Len(sDist) = 0 Then
        On Error Resume Next
        Set oJson = JsonConverter.ParseJson(CallGet("/db/UNIT"))
        sDist = UCase$(CStr(oJson("UNIT")("1")("DIST")))
        On Error GoTo 0
    End If
    Select Case sDist
        Case "CM": CvGravity = 9.806 * 100
        Case "MM": CvGravity = 9.806 * 1000
        Case "IN": CvGravity = 9.806 * 39.3701
        Case "FT": CvGravity = 9.806 * 3.28084
        Case Else: CvGravity = 9.806
    End Select
End Function

' A number in the description text: 8 -> "8",
' 8# / 1.5 -> "8.0" / "1.5" (pass 8# to get "8.0").
Private Function CvNumText(ByVal d As Variant) As String
    Select Case VarType(d)
        Case vbInteger, vbLong, vbByte
            CvNumText = CStr(d)
        Case Else
            If CDbl(d) = Int(CDbl(d)) Then CvNumText = CStr(CDbl(d)) & ".0" Else CvNumText = Replace(CStr(CDbl(d)), ",", ".")
    End Select
End Function


'==========================================================
' [21] Heat of hydration - pipe cooling, temperatures, convection, heat source, stages, result graph
'==========================================================
'  Analysis control: AnalysisHeat ([19]).
'    HeatAmbientConst 1, "AIR", 20
'    HeatConvectionConst 1, "C1", 12
'    HeatConvection 101, 1, "C1", "AIR"
'    HeatSourceCode 1, "HS", useConcreteData:=True, cementType:="Normal", temperature:=20, cementContent:=350
'    HeatSource Array(101, 102), "HS"
'    HeatStage 1, "H1", 20, actElem:="Pour1"
'==========================================================

'   HeatPipeCooling 1, "P1", 0.025, 1300, 4.2, 1000, 15, 1.2, "H1", "H2", 0, 48, Array(1, 2, 3)
Public Function HeatPipeCooling(ByVal pipeId As Long, ByVal NAME As String, ByVal diameter As Double, _
                                ByVal convection As Double, ByVal specificHeat As Double, ByVal density As Double, _
                                ByVal inletTemp As Double, ByVal flowRate As Double, ByVal startStage As String, _
                                ByVal endStage As String, ByVal startTime As Double, ByVal endTime As Double, _
                                ByVal nodeIds As Variant) As String
    CvStoreSet "/db/HPCE", pipeId, _
               CvTdRec("NAME", NAME, "DIAMETER", diameter, "COEF", convection, "HEAT", specificHeat, _
                       "DENSITY", density, "TEMPER", inletTemp, "FLOW_RATE", flowRate, _
                       "START_STAGE", startStage, "END_STAGE", endStage, "START_TIME", startTime, _
                       "END_TIME", endTime, "ITEMS", CvNumList(nodeIds))
    HeatPipeCooling = ""
End Function

'   HeatPrescribedTemp Array(1, 2), 25
Public Function HeatPrescribedTemp(ByVal nodeIds As Variant, ByVal temperature As Double, _
                                   Optional ByVal GROUP_NAME As String = "") As String
    HeatPrescribedTemp = CvItemsPut("/db/HSPT", nodeIds, _
                                    CvTdRec("ID", 1, "GROUP_NAME", GROUP_NAME, "TEMPER", temperature), True)
End Function

' Ambient temperature functions (ETFC).
Public Function HeatAmbientConst(ByVal funcId As Long, ByVal NAME As String, ByVal temperature As Double) As String
    HeatAmbientConst = CvTdPut("/db/ETFC", funcId, CvTdRec("NAME", NAME, "TYPE", "CONST", "TEMP", temperature))
End Function

Public Function HeatAmbientSine(ByVal funcId As Long, ByVal NAME As String, ByVal maxTemp As Double, _
                                ByVal meanTemp As Double, ByVal delayTime As Double) As String
    HeatAmbientSine = CvTdPut("/db/ETFC", funcId, CvTdRec("NAME", NAME, "TYPE", "SINE", "MAX_TEMP", maxTemp, _
                                                          "MEAN_TEMP", meanTemp, "DELAY_TIME", delayTime))
End Function

'   data rows: time, temperature
Public Function HeatAmbientUser(ByVal funcId As Long, ByVal NAME As String, ByVal scaleFactor As Double, _
                                ByVal data As Variant) As String
    HeatAmbientUser = CvTdPut("/db/ETFC", funcId, CvTdRec("NAME", NAME, "TYPE", "USER", _
                                                          "SCALE_FACTOR", scaleFactor, "ITEM", CvTimeValues(data)))
End Function

' Convection coefficient functions (CCFC).
Public Function HeatConvectionConst(ByVal funcId As Long, ByVal NAME As String, ByVal coefficient As Double) As String
    HeatConvectionConst = CvTdPut("/db/CCFC", funcId, CvTdRec("NAME", NAME, "TYPE", "CONST", "COEF", coefficient))
End Function

Public Function HeatConvectionUser(ByVal funcId As Long, ByVal NAME As String, ByVal scaleFactor As Double, _
                                   ByVal data As Variant) As String
    HeatConvectionUser = CvTdPut("/db/CCFC", funcId, CvTdRec("NAME", NAME, "TYPE", "USER", _
                                                             "SCALE_FACTOR", scaleFactor, "ITEM", CvTimeValues(data)))
End Function

'   HeatConvection 101, 1, "C1", "AIR", "HB"
' Convection boundary on an element face (HECB). A group given becomes a
' boundary group.
Public Function HeatConvection(ByVal elemIds As Variant, ByVal faceNo As Long, Optional ByVal convFunc As String = "", _
                               Optional ByVal ambientFunc As String = "", Optional ByVal GROUP_NAME As String = "") As String
    HeatConvection = CvItemsPut("/db/HECB", elemIds, _
                                CvTdRec("ID", 1, "GROUP_NAME", GROUP_NAME, "FACE_NO", faceNo, _
                                        "CCFC_NAME", convFunc, "ETFC_NAME", ambientFunc), True)
End Function

' Heat source functions (HSFC).
Public Function HeatSourceConst(ByVal funcId As Long, ByVal NAME As String, ByVal heatSource As Double) As String
    HeatSourceConst = CvTdPut("/db/HSFC", funcId, CvTdRec("NAME", NAME, "TYPE", "CONST", "TEMP_CONST", heatSource))
End Function

' useConcreteData: cementType (Normal / Moderate Heat / High-early-strength /
' Blast-furnace Slag / Fly Ash), temperature 10 / 20 / 30, cementContent;
' otherwise k and alpha.
Public Function HeatSourceCode(ByVal funcId As Long, ByVal NAME As String, Optional ByVal useConcreteData As Boolean = False, _
                               Optional ByVal k As Variant, Optional ByVal alpha As Variant, _
                               Optional ByVal cementType As String = "", Optional ByVal temperature As Variant, _
                               Optional ByVal cementContent As Variant) As String
    Dim oRec As Object
    Dim nCement As Long
    Dim nTemp As Long

    Set oRec = CvTdRec("NAME", NAME, "TYPE", "FUNC", "OPT_USE_CONC_DATA", useConcreteData)
    If useConcreteData Then
        Select Case cementType
            Case "Moderate Heat": nCement = 1
            Case "High-early-strength": nCement = 2
            Case "Blast-furnace Slag": nCement = 3
            Case "Fly Ash": nCement = 4
        End Select
        If Not IsMissing(temperature) Then
            Select Case CStr(temperature)
                Case "20": nTemp = 1
                Case "30": nTemp = 2
            End Select
        End If
        CvAddPairs oRec, "CEMENT_TYPE", nCement, "TEMP_FUNC", nTemp, "CEMENT_CONT", CvOr(cementContent, Null)
    Else
        CvAddPairs oRec, "K", CvOr(k, Null), "ALPHA", CvOr(alpha, Null)
    End If
    HeatSourceCode = CvTdPut("/db/HSFC", funcId, oRec)
End Function

Public Function HeatSourceUser(ByVal funcId As Long, ByVal NAME As String, ByVal scaleFactor As Double, _
                               ByVal data As Variant, Optional ByVal isAdiabaticTemp As Boolean = True) As String
    HeatSourceUser = CvTdPut("/db/HSFC", funcId, CvTdRec("NAME", NAME, "TYPE", "USER", "IS_ADIABATIC_TEMP", isAdiabaticTemp, _
                                                         "SCALE_FACTOR", scaleFactor, "ITEM", CvTimeValues(data)))
End Function

'   HeatSource Array(101, 102), "HS"
' Assign a heat source function to elements (HAHS).
Public Function HeatSource(ByVal elemIds As Variant, ByVal funcName As String) As String
    Dim v As Variant
    Dim i As Long
    v = CvIds(elemIds)
    For i = LBound(v) To UBound(v)
        CvTdPut "/db/HAHS", CLng(v(i)), CvTdRec("FUNC_NAME", funcName)
    Next i
    HeatSource = ""
End Function

'   HeatStage 1, "H1", 20, actElem:=Array("Pour1"), actBndr:="Conv1", actLoad:="LG", actDay:=0
' Heat of hydration construction stage (HSTG). initialTemp left out: none.
' Days (one for all or one per load group) default "0.000000".
Public Function HeatStage(ByVal stageId As Long, ByVal NAME As String, Optional ByVal initialTemp As Variant, _
                          Optional ByVal addStep As Variant, Optional ByVal actElem As Variant, _
                          Optional ByVal actBndr As Variant, Optional ByVal deactBndr As Variant, _
                          Optional ByVal actLoad As Variant, Optional ByVal actDay As Variant, _
                          Optional ByVal deactLoad As Variant, Optional ByVal deactDay As Variant) As String
    Dim oRec As Object
    Set oRec = CvTdRec("NAME", NAME, "bINITAL_TEMP", Not IsMissing(initialTemp), "ADD_STEP", CvNumList(addStep))
    If Not IsMissing(initialTemp) Then oRec.Add "INITIAL_TEMP", initialTemp
    oRec.Add "ACT_ELEM", CvNameList(actElem)
    oRec.Add "ACT_BNGR", CvNameList(actBndr)
    oRec.Add "DACT_BNGR", CvNameList(deactBndr)
    oRec.Add "ACT_LOAD", CvLoadDays(actLoad, actDay)
    oRec.Add "DACT_LOAD", CvLoadDays(deactLoad, deactDay)
    HeatStage = CvTdPut("/db/HSTG", stageId, oRec)
End Function

'   HeatResultGraph Array(5, 6), "Max"
' Stress history graph at nodes (HHND). component Sig_xx / Sig_yy / Sig_zz /
' Max / Sig_P1 / Sig_P2 / Sig_P3. graphId: the first id (the next nodes get
' the next ids).
Public Function HeatResultGraph(ByVal graphId As Long, ByVal nodeIds As Variant, ByVal component As String) As String
    Dim v As Variant
    Dim i As Long
    Dim nComp As Long
    Dim sSuffix As String

    Select Case component
        Case "Sig_yy": nComp = 1: sSuffix = "Y"
        Case "Sig_zz": nComp = 2: sSuffix = "Z"
        Case "Max": nComp = 3: sSuffix = "Max"
        Case "Sig_P1": nComp = 4: sSuffix = "P1"
        Case "Sig_P2": nComp = 5: sSuffix = "P2"
        Case "Sig_P3": nComp = 6: sSuffix = "P3"
        Case Else: nComp = 0: sSuffix = "X"
    End Select
    v = CvIds(nodeIds)
    For i = LBound(v) To UBound(v)
        CvTdPut "/db/HHND", graphId + i - LBound(v), _
                CvTdRec("NAME", "N" & v(i) & " - " & sSuffix, "TYPE", 0, "NODE_KEY", v(i), "ELEM_KEY", 0, "COMP", nComp)
    Next i
    HeatResultGraph = ""
End Function

' ---- helpers ----

Private Function CvTimeValues(ByVal pData As Variant) As Collection
    Dim col As Collection
    Dim i As Long
    Dim b As Long
    Set col = New Collection
    For i = LBound(pData) To UBound(pData)
        b = LBound(pData(i))
        col.Add CvTdRec("TIME", pData(i)(b), "VALUE", pData(i)(b + 1))
    Next i
    Set CvTimeValues = col
End Function

Private Function CvNameList(ByVal p As Variant) As Collection
    If CvHasNames(p) Then Set CvNameList = CvNumList(p) Else Set CvNameList = New Collection
End Function

Private Function CvLoadDays(ByVal pNames As Variant, ByVal pDays As Variant) As Collection
    Dim col As Collection
    Dim v As Variant
    Dim i As Long
    Set col = New Collection
    If CvHasNames(pNames) Then
        v = CvIds(pNames)
        For i = LBound(v) To UBound(v)
            col.Add CvTdRec("LOAD_NAME", v(i), "DAY", CStr(CvPick(pDays, i - LBound(v), "0.000000")))
        Next i
    End If
    Set CvLoadDays = col
End Function


'==========================================================
' [22] Result view - cutting line, cutting plane, selection in NX
'==========================================================
'  Cutting lines / planes and the selection in NX. Lines and planes are kept with the model and sent
'  by ModelCreate; the selection is read from NX straight away.
'==========================================================

'   CuttingLine 1, "CL1", Array(0, 0, 0), Array(10, 0, 0)
Public Function CuttingLine(ByVal lineId As Long, ByVal NAME As String, ByVal pt1 As Variant, ByVal pt2 As Variant) As String
    Dim a As Variant
    Dim b As Variant
    a = PscNums(pt1, 3)
    b = PscNums(pt2, 3)
    CuttingLine = CvTdPut("/db/CUTL", lineId, _
                          CvTdRec("NAME", NAME, "DIR", "NORMAL", "PT1X", a(0), "PT1Y", a(1), "PT1Z", a(2), _
                                  "PT2X", b(0), "PT2Y", b(1), "PT2Z", b(2), "R", 255, "G", 0, "B", 0, "TYPE", 0))
End Function

'   CuttingPlane 1, "CP1", "NORMAL", Array(0, 0, 0), Array(1, 0, 0), Array(0, 1, 0)
' direction "NORMAL" / "PLANE".
Public Function CuttingPlane(ByVal planeId As Long, ByVal NAME As String, ByVal direction As String, _
                             ByVal pt1 As Variant, ByVal pt2 As Variant, ByVal pt3 As Variant) As String
    Dim a As Variant
    Dim b As Variant
    Dim c As Variant
    a = PscNums(pt1, 3)
    b = PscNums(pt2, 3)
    c = PscNums(pt3, 3)
    CuttingPlane = CvTdPut("/db/CLWP", planeId, _
                           CvTdRec("NAME", NAME, "DIR", direction, "PT1X", a(0), "PT1Y", a(1), "PT1Z", a(2), _
                                   "PT2X", b(0), "PT2Y", b(1), "PT2Z", b(2), "PT3X", c(0), "PT3Y", c(1), "PT3Z", c(2), _
                                   "R", 255, "G", 0, "B", 0))
End Function

' Ids of the nodes / elements selected in NX now (an empty array when none).
'   v = SelectedElements()
Public Function SelectedNodes() As Variant
    SelectedNodes = CvSelected("NODE_LIST")
End Function

Public Function SelectedElements() As Variant
    SelectedElements = CvSelected("ELEM_LIST")
End Function

Private Function CvSelected(ByVal pKey As String) As Variant
    Dim oJson As Object
    Dim oList As Object
    Dim out() As Long
    Dim v As Variant
    Dim i As Long

    CvSelected = Array()
    On Error GoTo Done
    Set oJson = JsonConverter.ParseJson(CallGet("/view/SELECT"))
    Set oList = oJson("SELECT")(pKey)
    If oList.Count = 0 Then Exit Function
    ReDim out(0 To oList.Count - 1)
    For Each v In oList
        out(i) = CLng(v)
        i = i + 1
    Next v
    CvSelected = out
Done:
End Function


'==========================================================
' [23] Result graphics - result display, load display, view angle, capture
'==========================================================
'  Draws in the CIVIL NX window (view/RESULTGRAPHIC, view/DISPLAY,
'  view/ANGLE, view/CAPTURE). Kept for the
'  calculation sheets that paste result pictures.
'  The model still waiting in the store is sent first.
'    ShowBeamForce "CB", "COMB1", COMP:="My"
'    SetView "Front"
'    CaptureToCell Range("H2")
'==========================================================


Function ShowDeformedShape(ByVal loadType As String, ByVal loadName As String, _
                           Optional ByVal stepIndex As Long = 2, _
                           Optional ByVal MINMAX As String = "Max", _
                           Optional ByVal thOption As String = "Displacement", _
                           Optional ByVal COMP As String = "DZ", _
                           Optional ByVal useLocal As Boolean = False, _
                           Optional ByVal showDeform As Boolean = False, _
                           Optional ByVal deformScale As Double = 1, _
                           Optional ByVal realDeform As Boolean = False, _
                           Optional ByVal realDisp As Boolean = False, _
                           Optional ByVal relativeDisp As Boolean = False, _
                           Optional ByVal showValues As Boolean = False, _
                           Optional ByVal valuesDecimal As Long = 2, _
                           Optional ByVal showLegend As Boolean = False, _
                           Optional ByVal showMirrored As Boolean = False, _
                           Optional ByVal showUndeformed As Boolean = True) As String
    CvFlushPending

    Dim loadCaseComb As Object
    Set loadCaseComb = New Dictionary
    loadCaseComb.Add "TYPE", loadType
    loadCaseComb.Add "NAME", loadName
    loadCaseComb.Add "MINMAX", MINMAX
    loadCaseComb.Add "STEP_INDEX", stepIndex
    loadCaseComb.Add "TH_OPTION", thOption

    Dim components As Object
    Set components = New Dictionary
    components.Add "COMP", COMP
    components.Add "OPT_LOCAL_CHECK", useLocal

    Dim typeOfDisplay As Object
    Set typeOfDisplay = New Dictionary
    typeOfDisplay.Add "DEFORM", BuildDeform(showDeform, deformScale, realDeform, realDisp, relativeDisp)
    typeOfDisplay.Add "VALUES", BuildValues(showValues, decimalPt:=valuesDecimal)
    typeOfDisplay.Add "LEGEND", BuildLegend(showLegend)
    typeOfDisplay.Add "MIRRORED", BuildCheckOnly(showMirrored)
    typeOfDisplay.Add "UNDEFORMED", BuildCheckOnly(showUndeformed)
    typeOfDisplay.Add "OPT_CUR_STEP_DISPLACEMENT", True
    typeOfDisplay.Add "OPT_STAGE_STEP_REAL_DISPLACEMENT", True
    typeOfDisplay.Add "OPT_INCLUDING_CAMBER_DISPLACEMENT", True

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "CURRENT_MODE", "DeformedShape"
    arg.Add "LOAD_CASE_COMB", loadCaseComb
    arg.Add "COMPONENTS", components
    arg.Add "TYPE_OF_DISPLAY", typeOfDisplay

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    ShowDeformedShape = CallPost("/view/RESULTGRAPHIC", JsonConverter.ConvertToJson(body))

End Function

Function ShowReactionResult(ByVal loadType As String, ByVal loadName As String, _
                            Optional ByVal stepIndex As Long = 2, _
                            Optional ByVal MINMAX As String = "Max", _
                            Optional ByVal COMP As String = "FXYZ", _
                            Optional ByVal useLocal As Boolean = False, _
                            Optional ByVal showValues As Boolean = False, _
                            Optional ByVal valuesDecimal As Long = 2, _
                            Optional ByVal showLegend As Boolean = False, _
                            Optional ByVal legendPosition As String = "right", _
                            Optional ByVal legendDecimal As Long = 2, _
                            Optional ByVal arrowScale As Double = 1) As String
    CvFlushPending

    Dim loadCaseComb As Object
    Set loadCaseComb = New Dictionary
    loadCaseComb.Add "TYPE", loadType
    loadCaseComb.Add "NAME", loadName
    loadCaseComb.Add "MINMAX", MINMAX
    loadCaseComb.Add "STEP_INDEX", stepIndex

    Dim components As Object
    Set components = New Dictionary
    components.Add "COMP", COMP
    components.Add "OPT_LOCAL_CHECK", useLocal

    Dim typeOfDisplay As Object
    Set typeOfDisplay = New Dictionary
    typeOfDisplay.Add "LEGEND", BuildLegend(showLegend, legendPosition, decimalPt:=legendDecimal)
    typeOfDisplay.Add "VALUES", BuildValues(showValues, decimalPt:=valuesDecimal)
    typeOfDisplay.Add "ARROW_SCALE_FACTOR", arrowScale

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "CURRENT_MODE", "ReactionForces/Moments"
    arg.Add "LOAD_CASE_COMB", loadCaseComb
    arg.Add "COMPONENTS", components
    arg.Add "TYPE_OF_DISPLAY", typeOfDisplay

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    ShowReactionResult = CallPost("/view/RESULTGRAPHIC", JsonConverter.ConvertToJson(body))

End Function

Function ShowBeamForce(ByVal loadType As String, ByVal loadName As String, _
                       Optional ByVal stepIndex As Long = 2, _
                       Optional ByVal MINMAX As String = "Max", _
                       Optional ByVal part As String = "total", _
                       Optional ByVal COMP As String = "Fx", _
                       Optional ByVal showValues As Boolean = False, _
                       Optional ByVal valuesDecimal As Long = 2, _
                       Optional ByVal showLegend As Boolean = False, _
                       Optional ByVal legendPosition As String = "right", _
                       Optional ByVal legendDecimal As Long = 2) As String
    CvFlushPending

    Dim loadCaseComb As Object
    Set loadCaseComb = New Dictionary
    loadCaseComb.Add "TYPE", loadType
    loadCaseComb.Add "MINMAX", MINMAX
    loadCaseComb.Add "NAME", loadName
    loadCaseComb.Add "STEP_INDEX", stepIndex

    Dim components As Object
    Set components = New Dictionary
    components.Add "PART", part
    components.Add "COMP", COMP
    components.Add "OPT_SHOW_TRUSS_FORCES", True

    Dim typeOfDisplay As Object
    Set typeOfDisplay = New Dictionary
    typeOfDisplay.Add "CONTOUR", BuildContour()
    typeOfDisplay.Add "DEFORM", BuildDeform()
    typeOfDisplay.Add "VALUES", BuildValues(showValues, decimalPt:=valuesDecimal)
    typeOfDisplay.Add "LEGEND", BuildLegend(showLegend, legendPosition, decimalPt:=legendDecimal)
    typeOfDisplay.Add "MIRRORED", BuildCheckOnly(False)
    typeOfDisplay.Add "UNDEFORMED", BuildCheckOnly(True)
    typeOfDisplay.Add "OPT_CUR_STEP_FORCE", False
    typeOfDisplay.Add "YIELD_POINT", BuildCheckOnly(False)

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "CURRENT_MODE", "BeamForces/Moments"
    arg.Add "LOAD_CASE_COMB", loadCaseComb
    arg.Add "COMPONENTS", components
    arg.Add "TYPE_OF_DISPLAY", typeOfDisplay

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    ShowBeamForce = CallPost("/view/RESULTGRAPHIC", JsonConverter.ConvertToJson(body))

End Function

Function ShowBeamDiagram(ByVal loadType As String, ByVal loadName As String, _
                         Optional ByVal stepIndex As Long = 1, _
                         Optional ByVal MINMAX As String = "Max", _
                         Optional ByVal part As String = "Total", _
                         Optional ByVal COMP As String = "Fx", _
                         Optional ByVal fidelity As String = "5 Points", _
                         Optional ByVal fill As String = "Solid", _
                         Optional ByVal diagramScale As Double = 1, _
                         Optional ByVal showContour As Boolean = True, _
                         Optional ByVal showValues As Boolean = False, _
                         Optional ByVal valuesDecimal As Long = 2, _
                         Optional ByVal showLegend As Boolean = True, _
                         Optional ByVal legendPosition As String = "right", _
                         Optional ByVal legendDecimal As Long = 2) As String
    CvFlushPending

    Dim loadCaseComb As Object
    Set loadCaseComb = New Dictionary
    loadCaseComb.Add "TYPE", loadType
    loadCaseComb.Add "NAME", loadName
    loadCaseComb.Add "STEP_INDEX", stepIndex
    loadCaseComb.Add "MINMAX", MINMAX

    Dim components As Object
    Set components = New Dictionary
    components.Add "PART", part
    components.Add "COMP", COMP

    Dim displayOptions As Object
    Set displayOptions = New Dictionary
    displayOptions.Add "FIDELITY", fidelity
    displayOptions.Add "FILL", fill
    displayOptions.Add "SCALE", diagramScale

    Dim typeOfDisplay As Object
    Set typeOfDisplay = New Dictionary
    typeOfDisplay.Add "CONTOUR", BuildContour(showContour)
    typeOfDisplay.Add "VALUES", BuildValues(showValues, decimalPt:=valuesDecimal)
    typeOfDisplay.Add "LEGEND", BuildLegend(showLegend, legendPosition, decimalPt:=legendDecimal)

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "CURRENT_MODE", "BeamDiagrams"
    arg.Add "LOAD_CASE_COMB", loadCaseComb
    arg.Add "COMPONENTS", components
    arg.Add "DISPLAY_OPTIONS", displayOptions
    arg.Add "TYPE_OF_DISPLAY", typeOfDisplay

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    ShowBeamDiagram = CallPost("/view/RESULTGRAPHIC", JsonConverter.ConvertToJson(body))

End Function

Function ShowLoad(ByVal LoadCase As String, Optional ByVal loadType As String = "ST", _
                  Optional ByVal showNodal As Boolean = False, _
                  Optional ByVal showBeamLoad As Boolean = True) As String
    CvFlushPending

    Dim loadValue As Object
    Set loadValue = New Dictionary
    loadValue.Add "FORMAT", "Fixed"
    loadValue.Add "PLACE", 1

    Dim loadDisplay As Object
    Set loadDisplay = New Dictionary
    loadDisplay.Add "LOAD_VALUE", loadValue
    loadDisplay.Add "NODAL_BODY_FORCE", False
    loadDisplay.Add "NODAL_LOAD", showNodal
    loadDisplay.Add "SPECIFIED_DISPLACEMENT", False
    loadDisplay.Add "BEAM_LOAD", showBeamLoad
    loadDisplay.Add "PRESTRESS_LOAD", False
    loadDisplay.Add "PRETENSION_LOAD", False
    loadDisplay.Add "FLOOR_LOAD", False
    loadDisplay.Add "FINISHING_MATERIAL_LOAD", False
    loadDisplay.Add "PRESSURE_LOAD", False
    loadDisplay.Add "PLANE_LOAD", False
    loadDisplay.Add "NODAL_TEMPERATURE", False
    loadDisplay.Add "ELEMENT_TEMPERATURE", False
    loadDisplay.Add "TEMPERATURE_GRADIENT", False
    loadDisplay.Add "BEAM_SECTION_TEMPERATURE", False

    Dim caseSelection As Object
    Set caseSelection = New Dictionary
    caseSelection.Add "TYPE", loadType
    caseSelection.Add "NAME", LoadCase
    loadDisplay.Add "CASE_SELECTION", caseSelection

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "LOAD", loadDisplay

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    ShowLoad = CallPost("/view/DISPLAY", JsonConverter.ConvertToJson(body))

End Function

Function HideLoad() As String

    Dim loadValue As Object
    Set loadValue = New Dictionary
    loadValue.Add "FORMAT", "Fixed"
    loadValue.Add "PLACE", 1

    Dim loadDisplay As Object
    Set loadDisplay = New Dictionary
    loadDisplay.Add "LOAD_VALUE", loadValue
    loadDisplay.Add "NODAL_BODY_FORCE", False
    loadDisplay.Add "NODAL_LOAD", False
    loadDisplay.Add "SPECIFIED_DISPLACEMENT", False
    loadDisplay.Add "BEAM_LOAD", False
    loadDisplay.Add "PRESTRESS_LOAD", False
    loadDisplay.Add "PRETENSION_LOAD", False
    loadDisplay.Add "FLOOR_LOAD", False
    loadDisplay.Add "FINISHING_MATERIAL_LOAD", False
    loadDisplay.Add "PRESSURE_LOAD", False
    loadDisplay.Add "PLANE_LOAD", False
    loadDisplay.Add "NODAL_TEMPERATURE", False
    loadDisplay.Add "ELEMENT_TEMPERATURE", False
    loadDisplay.Add "TEMPERATURE_GRADIENT", False
    loadDisplay.Add "BEAM_SECTION_TEMPERATURE", False

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "LOAD", loadDisplay

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    HideLoad = CallPost("/view/DISPLAY", JsonConverter.ConvertToJson(body))

End Function

Function SetView(ByVal viewName As String) As String

    Dim h As Double, v As Double

    Select Case UCase(viewName)
        Case "FRONT": h = 0: v = 0
        Case "BACK": h = 180: v = 0
        Case "LEFT": h = 90: v = 0
        Case "RIGHT": h = 270: v = 0
        Case "TOP": h = 0: v = 90
        Case "BOTTOM": h = 0: v = -90
        Case "ISOMETRIC", "ISO": h = 35: v = 10
        Case Else
            SetView = "{""error"":""Unknown view name (Front/Back/Left/Right/Top/Bottom/Isometric)""}"
            Exit Function
    End Select

    SetView = SetAngle(h, v)

End Function

Function SetAngle(Optional ByVal horizontal As Double = 0, Optional ByVal vertical As Double = 0) As String

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "HORIZONTAL", horizontal
    arg.Add "VERTICAL", vertical

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    SetAngle = CallPost("/view/ANGLE", JsonConverter.ConvertToJson(body))

End Function

Function CaptureToCell(ByVal cell As Range, Optional ByVal mode As String = "post", _
                       Optional ByVal captureWidth As Long = 800, Optional ByVal captureHeight As Long = 600, _
                       Optional ByVal hidden As Boolean = False) As String
    CvFlushPending

    Dim folderPath As String
    folderPath = ThisWorkbook.Path & "\"

    Dim exportPath As String
    Dim i As Long
    i = 1
    Do
        exportPath = folderPath & "capture_" & i & ".jpg"
        i = i + 1
    Loop While DIR(exportPath) <> ""

    Dim arg As Object
    Set arg = New Dictionary
    arg.Add "EXPORT_PATH", exportPath
    arg.Add "WIDTH", captureWidth
    arg.Add "HEIGHT", captureHeight
    arg.Add "SET_MODE", mode
    arg.Add "SET_HIDDEN", hidden

    Dim body As Object
    Set body = New Dictionary
    body.Add "Argument", arg

    CaptureToCell = CallPost("/view/CAPTURE", JsonConverter.ConvertToJson(body))

    Dim pic As Object
    Set pic = cell.Parent.Pictures.Insert(exportPath)
    pic.Left = cell.Left
    pic.Top = cell.Top
    pic.width = cell.MergeArea.width
    pic.height = cell.MergeArea.height

End Function

Private Function BuildContour(Optional ByVal use As Boolean = True, Optional ByVal numColor As Long = 12, Optional ByVal colorType As String = "rgb") As Object
    Dim d As Object
    Set d = New Dictionary
    d.Add "OPT_CHECK", use
    d.Add "NUM_OF_COLOR", numColor
    d.Add "COLOR_TYPE", colorType
    Set BuildContour = d
End Function

Private Function BuildDeform(Optional ByVal use As Boolean = False, Optional ByVal deformScale As Double = 1, _
                             Optional ByVal realDeform As Boolean = False, Optional ByVal realDisp As Boolean = False, _
                             Optional ByVal relDisp As Boolean = False) As Object
    Dim d As Object
    Set d = New Dictionary
    d.Add "OPT_CHECK", use
    d.Add "SCALE_FACTOR", deformScale
    d.Add "REL_DISP", relDisp
    d.Add "REAL_DISP", realDisp
    d.Add "REAL_DEFORM", realDeform
    Set BuildDeform = d
End Function

Private Function BuildValues(Optional ByVal use As Boolean = False, Optional ByVal expo As Boolean = False, _
                             Optional ByVal decimalPt As Long = 2, Optional ByVal orient As Long = 0) As Object
    Dim d As Object
    Set d = New Dictionary
    d.Add "OPT_CHECK", use
    d.Add "VALUE_EXP", expo
    d.Add "DECIMAL_PT", decimalPt
    d.Add "SET_ORIENT", orient
    Set BuildValues = d
End Function

Private Function BuildLegend(Optional ByVal use As Boolean = True, Optional ByVal position As String = "right", _
                             Optional ByVal expo As Boolean = False, Optional ByVal decimalPt As Long = 2) As Object
    Dim d As Object
    Set d = New Dictionary
    d.Add "OPT_CHECK", use
    d.Add "POSITION", position
    d.Add "VALUE_EXP", expo
    d.Add "DECIMAL_PT", decimalPt
    Set BuildLegend = d
End Function

Private Function BuildCheckOnly(ByVal use As Boolean) As Object
    Dim d As Object
    Set d = New Dictionary
    d.Add "OPT_CHECK", use
    Set BuildCheckOnly = d
End Function


'==========================================================
' [24] Geometry by location - nodes / elements from points, selection, groups
'==========================================================
'  Nodes and elements without ids, made along lines, lofted or extruded;
'  selection by box / line / polygon; ids in a group; convection by nodes.
'    - a node at a place that already has one (within 0.00001) is not made
'      again; its id is used instead
'    - ids follow on from the largest node / element id written so far
'    - points are Array(x, y, z); the functions that make several items
'      return their ids as an array (0 based)
'      Dim e As Variant
'      e = BeamSE(Array(0, 0, 0), Array(10, 0, 0), 5, 1, 1)   ' 5 beams
'      Support SelectBox(Array(0, 0, 0), Array(10, 0, 0)), "fix"
'==========================================================

' Node at a place. GROUP: the structure group(s) it joins ("A,B").
' merge False always makes a new node.
Public Function NodeAt(ByVal x As Double, ByVal y As Double, ByVal z As Double, _
                       Optional ByVal GROUP As String = "", Optional ByVal merge As Boolean = True) As Long
    Dim sCell As String
    Dim v As Variant
    Dim p As Variant
    Dim id As Long

    CvGeoSync
    x = CvR6(x): y = CvR6(y): z = CvR6(z)
    sCell = CvCell(x, y, z)
    If merge And mGeoGrid.Exists(sCell) Then
        For Each v In mGeoGrid.Item(sCell)
            p = mGeoXYZ.Item(CStr(v))
            If Sqr((x - p(0)) ^ 2 + (y - p(1)) ^ 2 + (z - p(2)) ^ 2) < 0.00001 Then
                If Len(GROUP) > 0 Then CvGroupJoin GROUP, CLng(v), Null
                NodeAt = CLng(v)
                Exit Function
            End If
        Next v
    End If
    id = mGeoNodeMax + 1
    Node id, x, y, z, GROUP
    CvGeoAddNode id, x, y, z
    mGeoNodeSeen = mGeoNodeSeen + 1
    NodeAt = id
End Function

' Nodes on the line from sLoc to eLoc, n segments (n + 1 nodes).
Public Function NodeSE(ByVal sLoc As Variant, ByVal eLoc As Variant, Optional ByVal n As Long = 1, _
                       Optional ByVal GROUP As String = "", Optional ByVal merge As Boolean = True) As Variant
    Dim pts As Variant
    Dim out() As Long
    Dim i As Long
    pts = CvLinspace(sLoc, eLoc, n)
    ReDim out(0 To n)
    For i = 0 To n
        out(i) = NodeAt(pts(i)(0), pts(i)(1), pts(i)(2), GROUP, merge)
    Next i
    NodeSE = out
End Function

' Nodes from sLoc along a direction for a length, n segments.
Public Function NodeSDL(ByVal sLoc As Variant, ByVal dirVec As Variant, ByVal l As Double, Optional ByVal n As Long = 1, _
                        Optional ByVal GROUP As String = "", Optional ByVal merge As Boolean = True) As Variant
    Dim pts As Variant
    Dim out() As Long
    Dim i As Long
    pts = CvSdlPoints(sLoc, dirVec, l, n)
    ReDim out(0 To n)
    For i = 0 To n
        out(i) = NodeAt(pts(i)(0), pts(i)(1), pts(i)(2), GROUP, merge)
    Next i
    NodeSDL = out
End Function

' The next free node / element id.
Public Function NextNodeId() As Long
    CvGeoSync
    NextNodeId = mGeoNodeMax + 1
End Function

Public Function NextElemId() As Long
    CvGeoSync
    NextElemId = mGeoElemMax + 1
End Function

' Coordinates of a node made or read here: Array(x, y, z).
Public Function NodeXYZ(ByVal nodeId As Long) As Variant
    CvGeoSync
    If mGeoXYZ.Exists(CStr(nodeId)) Then NodeXYZ = mGeoXYZ.Item(CStr(nodeId)) Else NodeXYZ = Empty
End Function

' End of the last beam / truss made (Element.lastLoc).
Public Function LastLoc() As Variant
    If IsEmpty(mGeoLast) Then LastLoc = Array(0, 0, 0) Else LastLoc = mGeoLast
End Function

' Elements with the next id (Element.Beam(i, j) / Truss / Plate / Solid).
Public Function AddBeam(ByVal ni As Long, ByVal nj As Long, Optional ByVal mat As Long = 1, Optional ByVal sect As Long = 1, _
                        Optional ByVal angle As Double = 0, Optional ByVal GROUP As String = "") As Long
    Dim id As Long
    id = NextElemId()
    Beam id, mat, sect, ni, nj, angle, GROUP
    CvGeoAddElem id
    mGeoLast = NodeXYZ(nj)
    AddBeam = id
End Function

Public Function AddTruss(ByVal ni As Long, ByVal nj As Long, Optional ByVal mat As Long = 1, Optional ByVal sect As Long = 1, _
                         Optional ByVal angle As Double = 0, Optional ByVal GROUP As String = "") As Long
    Dim id As Long
    id = NextElemId()
    Truss id, mat, sect, ni, nj, angle, GROUP
    CvGeoAddElem id
    mGeoLast = NodeXYZ(nj)
    AddTruss = id
End Function

' A node repeated in the list makes a triangle.
Public Function AddPlate(ByVal nodeIds As Variant, Optional ByVal STYPE As Long = 1, Optional ByVal mat As Long = 1, _
                         Optional ByVal sect As Long = 1, Optional ByVal angle As Double = 0, _
                         Optional ByVal GROUP As String = "") As Long
    Dim id As Long
    Dim oUniq As Object
    Dim v As Variant
    Dim nodes As Variant

    Set oUniq = New Dictionary
    For Each v In nodeIds
        oUniq.Item(CStr(v)) = CLng(v)
    Next v
    If oUniq.Count = 3 Then nodes = CvLongs(oUniq.Items) Else nodes = CvLongs(nodeIds)
    id = NextElemId()
    Plate id, mat, sect, nodes, angle, STYPE, GROUP
    CvGeoAddElem id
    AddPlate = id
End Function

' 4, 6 or 8 nodes. The two faces are swapped when the first face points away
' from the rest.
Public Function AddSolid(ByVal nodeIds As Variant, Optional ByVal mat As Long = 1, Optional ByVal GROUP As String = "") As Long
    Dim id As Long
    Dim n() As Long
    Dim nn As Long
    Dim i As Long
    Dim c(2) As Double
    Dim p0 As Variant, p1 As Variant, p2 As Variant, pt As Variant
    Dim zb(2) As Double
    Dim a(2) As Double, b(2) As Double
    Dim d As Double
    Dim m As Long

    n = CvLongs(nodeIds)
    nn = UBound(n) + 1
    For i = 0 To nn - 1
        pt = NodeXYZ(n(i))
        c(0) = c(0) + pt(0): c(1) = c(1) + pt(1): c(2) = c(2) + pt(2)
    Next i
    c(0) = c(0) / nn: c(1) = c(1) / nn: c(2) = c(2) / nn
    p0 = NodeXYZ(n(0)): p1 = NodeXYZ(n(1)): p2 = NodeXYZ(n(2))
    For i = 0 To 2
        a(i) = p1(i) - p0(i)
        b(i) = p2(i) - p1(i)
    Next i
    zb(0) = a(1) * b(2) - a(2) * b(1)
    zb(1) = a(2) * b(0) - a(0) * b(2)
    zb(2) = a(0) * b(1) - a(1) * b(0)
    If nn = 8 Then m = 4 Else m = 3
    pt = NodeXYZ(n(m))
    d = (pt(0) - c(0)) * zb(0) + (pt(1) - c(1)) * zb(1) + (pt(2) - c(2)) * zb(2)
    If d < 0 Then
        Dim r() As Long
        ReDim r(0 To nn - 1)
        If nn = 4 Then
            r(0) = n(2): r(1) = n(1): r(2) = n(0): r(3) = n(3)
        Else
            For i = 0 To nn - 1
                r(i) = n((i + m) Mod nn)
            Next i
        End If
        n = r
    End If
    id = NextElemId()
    Solid id, mat, n, GROUP
    CvGeoAddElem id
    AddSolid = id
End Function

' Beams on the line sLoc - eLoc in n pieces (Element.Beam.SE).
Public Function BeamSE(ByVal sLoc As Variant, ByVal eLoc As Variant, Optional ByVal n As Long = 1, _
                       Optional ByVal mat As Long = 1, Optional ByVal sect As Long = 1, _
                       Optional ByVal angle As Double = 0, Optional ByVal GROUP As String = "") As Variant
    BeamSE = CvLineElems("BEAM", CvLinspace(sLoc, eLoc, n), n, mat, sect, angle, GROUP)
End Function

' Beams from sLoc along a direction for a length, n pieces (Element.Beam.SDL).
Public Function BeamSDL(ByVal sLoc As Variant, ByVal dirVec As Variant, ByVal l As Double, Optional ByVal n As Long = 1, _
                        Optional ByVal mat As Long = 1, Optional ByVal sect As Long = 1, _
                        Optional ByVal angle As Double = 0, Optional ByVal GROUP As String = "") As Variant
    BeamSDL = CvLineElems("BEAM", CvSdlPoints(sLoc, dirVec, l, n), n, mat, sect, angle, GROUP)
End Function

Public Function TrussSE(ByVal sLoc As Variant, ByVal eLoc As Variant, Optional ByVal n As Long = 1, _
                        Optional ByVal mat As Long = 1, Optional ByVal sect As Long = 1, _
                        Optional ByVal angle As Double = 0, Optional ByVal GROUP As String = "") As Variant
    TrussSE = CvLineElems("TRUSS", CvLinspace(sLoc, eLoc, n), n, mat, sect, angle, GROUP)
End Function

Public Function TrussSDL(ByVal sLoc As Variant, ByVal dirVec As Variant, ByVal l As Double, Optional ByVal n As Long = 1, _
                         Optional ByVal mat As Long = 1, Optional ByVal sect As Long = 1, _
                         Optional ByVal angle As Double = 0, Optional ByVal GROUP As String = "") As Variant
    TrussSDL = CvLineElems("TRUSS", CvSdlPoints(sLoc, dirVec, l, n), n, mat, sect, angle, GROUP)
End Function

'   BeamPLine Array(Array(0, 0, 0), Array(25, 0, 10), Array(50, 0, 0)), 10, 2, divAxis:="X"
' Beams along a curve through the points (Element.Beam.PLine).
' nDiv 0: straight between the given points. deg 1: polyline; a higher
' degree is supported when it is the number of points - 1 (one curve through
' all points, e.g. an arch through 3 points with deg 2).
' divAxis "L" equal length, "X" / "Y" / "Z" equal steps along that axis.
Public Function BeamPLine(ByVal points As Variant, Optional ByVal nDiv As Long = 0, Optional ByVal deg As Long = 1, _
                          Optional ByVal mat As Long = 1, Optional ByVal sect As Long = 1, _
                          Optional ByVal angle As Double = 0, Optional ByVal GROUP As String = "", _
                          Optional ByVal divAxis As String = "L") As Variant
    Dim pts As Variant
    Dim m As Long

    m = UBound(points) - LBound(points) + 1
    If nDiv = 0 Then
        pts = CvPointList(points)
        BeamPLine = CvLineElems("BEAM", pts, m - 1, mat, sect, angle, GROUP)
        Exit Function
    End If
    If deg < 1 Then deg = 1
    If deg > m - 1 Then deg = m - 1
    If deg <> 1 And deg <> m - 1 Then
        BeamPLine = Array()
        mLastError = "BeamPLine: deg must be 1 or the number of points - 1"
        Exit Function
    End If
    pts = CvCurvePoints(CvPointList(points), nDiv, deg, UCase$(divAxis))
    BeamPLine = CvLineElems("BEAM", pts, nDiv, mat, sect, angle, GROUP)
End Function

'   PlateLoft Array("#Edge_A", "#Edge_B"), nDiv:=3, sect:=1
' Plates between the nodes of structure groups taken in turn
' (Element.Plate.loftGroups). bClose joins the last node back to the first.
Public Function PlateLoft(ByVal groups As Variant, Optional ByVal STYPE As Long = 1, Optional ByVal mat As Long = 1, _
                          Optional ByVal sect As Long = 1, Optional ByVal angle As Double = 0, _
                          Optional ByVal GROUP As String = "", Optional ByVal nDiv As Long = 1, _
                          Optional ByVal bClose As Boolean = False) As Variant
    Dim col As Collection
    Dim g As Long
    Dim vG As Variant
    Dim a As Variant, b As Variant

    vG = CvIds(groups)
    Set col = New Collection
    For g = LBound(vG) To UBound(vG) - 1
        a = NodesInGroup(vG(g))
        b = NodesInGroup(vG(g + 1))
        If bClose Then
            a = CvAppend(a, a(0))
            b = CvAppend(b, b(0))
        End If
        CvLongest a, b
        CvPlatesBetween col, a, b, nDiv, STYPE, mat, sect, angle, GROUP
    Next g
    PlateLoft = CvColToArray(col)
End Function

'   PlateExtrude SelectLine(Array(0, 0, 1), Array(2, 0, 1)), Array(0, 0, 0.5), 5, sect:=3
' Plates made by moving a line of nodes (node ids, or points Array(x, y, z))
' along a vector in nDiv steps (Element.Plate.extrude). nDiv is cut to a whole number.
Public Function PlateExtrude(ByVal nodesOrPoints As Variant, ByVal dirVec As Variant, Optional ByVal nDiv As Double = 1, _
                             Optional ByVal bClose As Boolean = False, Optional ByVal STYPE As Long = 1, _
                             Optional ByVal mat As Long = 1, Optional ByVal sect As Long = 1, _
                             Optional ByVal angle As Double = 0, Optional ByVal GROUP As String = "") As Variant
    Dim col As Collection
    Dim a() As Long, b() As Long
    Dim v As Variant
    Dim p As Variant
    Dim i As Long
    Dim n As Long

    nDiv = Fix(nDiv)          ' a count like 10 * dy is cut down to a whole number
    n = UBound(nodesOrPoints) - LBound(nodesOrPoints) + 1
    Set col = New Collection
    If n = 0 Then
        PlateExtrude = Array()
        Exit Function
    End If
    ReDim a(0 To n - 1): ReDim b(0 To n - 1)
    i = 0
    For Each v In nodesOrPoints
        If IsArray(v) Then
            p = v
            a(i) = NodeAt(p(0), p(1), p(2))
        Else
            a(i) = CLng(v)
        End If
        i = i + 1
    Next v
    For i = 0 To n - 1
        p = NodeXYZ(a(i))
        b(i) = NodeAt(p(0) + dirVec(0), p(1) + dirVec(1), p(2) + dirVec(2))
    Next i
    Dim va As Variant, vb As Variant
    va = a: vb = b
    If bClose Then
        va = CvAppend(va, a(0))
        vb = CvAppend(vb, b(0))
    End If
    CvPlatesBetween col, va, vb, nDiv, STYPE, mat, sect, angle, GROUP
    PlateExtrude = CvColToArray(col)
End Function

'   SolidExtrude ElemsInGroup("#Base"), Array(0, 0, 3), 5, 1, "Soil", True
' Solids made by moving plate elements along a vector in nDiv layers
' (Element.Solid.extrudeFromPlates). bDeletePlate removes the plates.
Public Function SolidExtrude(ByVal elemIds As Variant, Optional ByVal dirVec As Variant, Optional ByVal nDiv As Long = 1, _
                             Optional ByVal mat As Long = 1, Optional ByVal GROUP As String = "", _
                             Optional ByVal bDeletePlate As Boolean = False) As Variant
    Dim col As Collection
    Dim oElem As Object
    Dim plates As Collection
    Dim v As Variant
    Dim rec As Object
    Dim nA As Variant
    Dim nB() As Long
    Dim all() As Long
    Dim ex(2) As Double
    Dim q As Long
    Dim i As Long
    Dim k As Long
    Dim p As Variant

    If IsMissing(dirVec) Then dirVec = Array(0, 0, 1)
    ex(0) = dirVec(0) / nDiv: ex(1) = dirVec(1) / nDiv: ex(2) = dirVec(2) / nDiv
    Set oElem = CvStoreTable("ELEM")
    Set plates = New Collection
    For Each v In CvIds(elemIds)
        If oElem.Exists(CStr(v)) Then
            If oElem.Item(CStr(v)).Item("TYPE") = "PLATE" Or oElem.Item(CStr(v)).Item("TYPE") = "WALL" Then plates.Add CLng(v)
        End If
    Next v
    Set col = New Collection
    For Each v In plates
        Set rec = oElem.Item(CStr(v))
        nA = CvLongs(CvNonZero(rec.Item("NODE")))
        For q = 1 To nDiv
            ReDim nB(0 To UBound(nA))
            For i = 0 To UBound(nA)
                p = NodeXYZ(nA(i))
                nB(i) = NodeAt(p(0) + ex(0), p(1) + ex(1), p(2) + ex(2))
            Next i
            ReDim all(0 To 2 * (UBound(nA) + 1) - 1)
            For i = 0 To UBound(nA)
                all(i) = nA(i)
                all(i + UBound(nA) + 1) = nB(i)
            Next i
            col.Add AddSolid(all, mat, GROUP)
            nA = nB
        Next q
    Next v
    If bDeletePlate Then
        For Each v In plates
            oElem.Remove CStr(v)
            mGeoElemSeen = mGeoElemSeen - 1
        Next v
    End If
    SolidExtrude = CvColToArray(col)
End Function

' Ids in structure groups (one name or Array of names), in the order they
' joined, each once (nodesInGroup / elemsInGroup). bReverse reverses.
Public Function NodesInGroup(ByVal names As Variant, Optional ByVal bReverse As Boolean = False) As Variant
    NodesInGroup = CvGroupIds(names, "N_LIST", bReverse)
End Function

Public Function ElemsInGroup(ByVal names As Variant, Optional ByVal bReverse As Boolean = False) As Variant
    ElemsInGroup = CvGroupIds(names, "E_LIST", bReverse)
End Function

' The ids without the ones given (one id or an array), order kept.
'   RigidLink master, Without(SelectBox(p1, p2), master)
Public Function Without(ByVal ids As Variant, ByVal removeIds As Variant) As Variant
    Dim oOut As Object
    Dim col As Collection
    Dim v As Variant
    Set oOut = New Dictionary
    For Each v In CvIds(removeIds)
        oOut.Item(CStr(v)) = True
    Next v
    Set col = New Collection
    For Each v In CvIds(ids)
        If Not oOut.Exists(CStr(v)) Then col.Add CLng(v)
    Next v
    Without = CvColToArray(col)
End Function

' Ids from several lists, each once, ascending.
'   pins = IdsUnion(SelectPolygon(face1), SelectPolygon(face2))
Public Function IdsUnion(ParamArray lists() As Variant) As Variant
    Dim oSeen As Object
    Dim col As Collection
    Dim i As Long
    Dim v As Variant
    Set oSeen = New Dictionary
    Set col = New Collection
    For i = LBound(lists) To UBound(lists)
        For Each v In CvIds(lists(i))
            If Not oSeen.Exists(CStr(v)) Then
                oSeen.Add CStr(v), True
                col.Add CLng(v)
            End If
        Next v
    Next i
    IdsUnion = CvSortLongs(CvColToArray(col))
End Function

' Element ids in the order they were made, optionally only of a type
' ("BEAM", "TRUSS", "PLATE", "SOLID" ...), a section or a material.
Public Function ElemList(Optional ByVal elemType As String = "", Optional ByVal sectId As Long = 0, _
                         Optional ByVal matlId As Long = 0) As Variant
    Dim col As Collection
    Dim oElem As Object
    Dim vKey As Variant
    Dim rec As Object
    Set col = New Collection
    Set oElem = CvStoreTable("ELEM")
    For Each vKey In oElem.Keys
        Set rec = oElem.Item(vKey)
        If (Len(elemType) = 0 Or UCase$(rec.Item("TYPE")) = UCase$(elemType)) And _
           (sectId = 0 Or CLng(rec.Item("SECT")) = sectId) And _
           (matlId = 0 Or CLng(rec.Item("MATL")) = matlId) Then col.Add CLng(vKey)
    Next vKey
    ElemList = CvColToArray(col)
End Function

'   Support SelectBox(Array(0, 0, 0), Array(20, 10, 0)), "fix"
' Nodes (output "NODE_ID") or elements ("ELEM_ID", by their centre) in a box,
' ascending. elemType keeps only that element type.
Public Function SelectBox(ByVal pt1 As Variant, ByVal pt2 As Variant, Optional ByVal output As String = "NODE_ID", _
                          Optional ByVal elemType As String = "") As Variant
    Const TOL As Double = 0.001
    Dim lo(2) As Double, hi(2) As Double
    Dim i As Long
    Dim col As Collection
    Dim it As Variant
    For i = 0 To 2
        lo(i) = CvMin(pt1(i), pt2(i)) - TOL
        hi(i) = CvMax(pt1(i), pt2(i)) + TOL
    Next i
    Set col = New Collection
    For Each it In CvSelectables(output, elemType)
        If it(1) >= lo(0) And it(1) <= hi(0) And it(2) >= lo(1) And it(2) <= hi(1) And _
           it(3) >= lo(2) And it(3) <= hi(2) Then col.Add it(0)
    Next it
    SelectBox = CvSortLongs(CvColToArray(col))
End Function

' Nodes / elements within radius of the line pt1 - pt2, nearest pt1 first.
Public Function SelectLine(ByVal pt1 As Variant, ByVal pt2 As Variant, Optional ByVal output As String = "NODE_ID", _
                           Optional ByVal radius As Double = 0.001, Optional ByVal elemType As String = "") As Variant
    Dim lo(2) As Double, hi(2) As Double
    Dim dv(2) As Double
    Dim i As Long
    Dim it As Variant
    Dim df(2) As Double
    Dim cr(2) As Double
    Dim along As Double
    Dim perp As Double
    Dim dlen As Double
    Dim keys As Collection
    Dim ids As Collection

    For i = 0 To 2
        lo(i) = CvMin(pt1(i) - radius, pt2(i) - radius)
        hi(i) = CvMax(pt1(i) + radius, pt2(i) + radius)
        dv(i) = pt2(i) - pt1(i)
    Next i
    dlen = Sqr(dv(0) ^ 2 + dv(1) ^ 2 + dv(2) ^ 2)
    Set keys = New Collection
    Set ids = New Collection
    For Each it In CvSelectables(output, elemType)
        If it(1) >= lo(0) And it(1) <= hi(0) And it(2) >= lo(1) And it(2) <= hi(1) And _
           it(3) >= lo(2) And it(3) <= hi(2) Then
            For i = 0 To 2
                df(i) = it(i + 1) - pt1(i)
            Next i
            cr(0) = df(1) * dv(2) - df(2) * dv(1)
            cr(1) = df(2) * dv(0) - df(0) * dv(2)
            cr(2) = df(0) * dv(1) - df(1) * dv(0)
            along = Sqr(df(0) ^ 2 + df(1) ^ 2 + df(2) ^ 2)
            perp = Sqr(cr(0) ^ 2 + cr(1) ^ 2 + cr(2) ^ 2) / dlen
            If perp < radius Then
                keys.Add along
                ids.Add it(0)
            End If
        End If
    Next it
    SelectLine = CvSortByKey(keys, ids)
End Function

'   SelectPolygon Array(Array(0, 0, 0), Array(10, 0, 0), Array(10, 0, 5), Array(0, 0, 5))
' Nodes / elements in a flat polygon (its edges included), ascending.
Public Function SelectPolygon(ByVal points As Variant, Optional ByVal output As String = "NODE_ID", _
                              Optional ByVal elemType As String = "") As Variant
    Const TOL As Double = 0.001
    Dim pts As Variant
    Dim n As Long
    Dim i As Long
    Dim lo(2) As Double, hi(2) As Double
    Dim a As Long, b As Long, c As Long
    Dim xd As Double, yd As Double, zd As Double
    Dim pu() As Double, pv() As Double
    Dim c1 As Double, c2 As Double
    Dim col As Collection
    Dim it As Variant

    pts = CvPointList(points)
    n = UBound(pts) + 1
    For i = 0 To 2
        lo(i) = 1E+300: hi(i) = -1E+300
    Next i
    For i = 0 To n - 1
        For a = 0 To 2
            If pts(i)(a) < lo(a) Then lo(a) = pts(i)(a)
            If pts(i)(a) > hi(a) Then hi(a) = pts(i)(a)
        Next a
    Next i
    For i = 0 To 2
        lo(i) = lo(i) - TOL: hi(i) = hi(i) + TOL
    Next i
    xd = hi(0) - lo(0): yd = hi(1) - lo(1): zd = hi(2) - lo(2)
    ' plane choice (x is compared with y only)
    If xd < yd Then
        a = 1: b = 2: c = 0
    ElseIf yd < CvMin(xd, zd) Then
        a = 0: b = 2: c = 1
    Else
        a = 0: b = 1: c = 2
    End If
    ReDim pu(0 To n - 1): ReDim pv(0 To n - 1)
    c1 = 1E+300: c2 = -1E+300
    For i = 0 To n - 1
        pu(i) = pts(i)(a): pv(i) = pts(i)(b)
        If pts(i)(c) < c1 Then c1 = pts(i)(c)
        If pts(i)(c) > c2 Then c2 = pts(i)(c)
    Next i
    c1 = c1 - TOL: c2 = c2 + TOL
    Set col = New Collection
    For Each it In CvSelectables(output, elemType)
        If it(1) >= lo(0) And it(1) <= hi(0) And it(2) >= lo(1) And it(2) <= hi(1) And _
           it(3) >= lo(2) And it(3) <= hi(2) Then
            If it(c + 1) >= c1 And it(c + 1) <= c2 Then
                If CvInPolygon(it(a + 1), it(b + 1), pu, pv, TOL) Then col.Add it(0)
            End If
        End If
    Next it
    SelectPolygon = CvSortLongs(CvColToArray(col))
End Function

' Convection boundaries on every solid face whose nodes are all in nodeIds
' (HoH.Convection.Boundary.bySelectedNodes).
Public Function HeatConvectionByNodes(ByVal nodeIds As Variant, Optional ByVal convFunc As String = "", _
                                      Optional ByVal ambientFunc As String = "", Optional ByVal GROUP_NAME As String = "") As String
    Dim oSel As Object
    Dim oElem As Object
    Dim vKey As Variant
    Dim rec As Object
    Dim nodes As Variant
    Dim v As Variant
    Dim s As String
    Dim face As Long

    Set oSel = New Dictionary
    For Each v In CvIds(nodeIds)
        oSel.Item(CStr(v)) = True
    Next v
    Set oElem = CvStoreTable("ELEM")
    For Each vKey In oElem.Keys
        Set rec = oElem.Item(vKey)
        If rec.Item("TYPE") = "SOLID" Then
            nodes = CvNonZero(rec.Item("NODE"))
            s = ""
            For Each v In nodes
                If oSel.Exists(CStr(v)) Then s = s & "T" Else s = s & "F"
            Next v
            face = CvSolidFace(s)
            If face > 0 Then HeatConvection CLng(vKey), face, convFunc, ambientFunc, GROUP_NAME
        End If
    Next vKey
    HeatConvectionByNodes = ""
End Function

' ---- helpers ----

' "fix" / "pin" / "roller", or a short string of 0 / 1 -> 7 digits.
Private Function CvConstraint(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Select Case LCase$(s)
        Case "pin": s = "111"
        Case "fix": s = "1111111"
        Case "roller": s = "001"
    End Select
    s = Left$(s & String(7, "0"), 7)
    For i = 1 To 7
        If Mid$(s, i, 1) <> "0" Then out = out & "1" Else out = out & "0"
    Next i
    CvConstraint = out
End Function

Private Sub CvGeoReset()
    Set mGeoXYZ = Nothing
    Set mGeoGrid = Nothing
    mGeoNodeMax = 0
    mGeoElemMax = 0
    mGeoNodeSeen = 0
    mGeoElemSeen = 0
    mGeoLast = Empty
End Sub

' Bring the registry up to date with nodes / elements written by the other
' helpers (Node, Beam ...). The registry outlives ModelCreate.
Private Sub CvGeoSync()
    Dim oT As Object
    Dim vKey As Variant
    Dim rec As Object

    If mGeoXYZ Is Nothing Then
        Set mGeoXYZ = New Dictionary
        Set mGeoGrid = New Dictionary
    End If
    Set oT = CvStoreTable("NODE")
    If oT.Count <> mGeoNodeSeen Then
        For Each vKey In oT.Keys
            If Not mGeoXYZ.Exists(CStr(vKey)) Then
                Set rec = oT.Item(vKey)
                CvGeoAddNode CLng(vKey), CDbl(rec.Item("X")), CDbl(rec.Item("Y")), CDbl(rec.Item("Z"))
            End If
        Next vKey
        mGeoNodeSeen = oT.Count
    End If
    Set oT = CvStoreTable("ELEM")
    If oT.Count <> mGeoElemSeen Then
        For Each vKey In oT.Keys
            If CLng(vKey) > mGeoElemMax Then mGeoElemMax = CLng(vKey)
        Next vKey
        mGeoElemSeen = oT.Count
    End If
End Sub

Private Sub CvGeoAddNode(ByVal pId As Long, ByVal x As Double, ByVal y As Double, ByVal z As Double)
    Dim sCell As String
    mGeoXYZ.Item(CStr(pId)) = Array(x, y, z)
    sCell = CvCell(x, y, z)
    If Not mGeoGrid.Exists(sCell) Then mGeoGrid.Add sCell, New Collection
    mGeoGrid.Item(sCell).Add pId
    If pId > mGeoNodeMax Then mGeoNodeMax = pId
End Sub

Private Sub CvGeoAddElem(ByVal pId As Long)
    If pId > mGeoElemMax Then mGeoElemMax = pId
    mGeoElemSeen = mGeoElemSeen + 1
End Sub

Private Function CvCell(ByVal x As Double, ByVal y As Double, ByVal z As Double) As String
    CvCell = Fix(x) & "," & Fix(y) & "," & Fix(z)
End Function

' round(v, 6) (half to even on the decimal value).
Private Function CvR6(ByVal v As Double) As Double
    CvR6 = CDbl(Round(CDec(v), 6))
End Function

' n + 1 points from s to e (numpy.linspace).
Private Function CvLinspace(ByVal s As Variant, ByVal e As Variant, ByVal n As Long) As Variant
    Dim out() As Variant
    Dim i As Long
    Dim st(2) As Double
    Dim k As Long
    ReDim out(0 To n)
    For k = 0 To 2
        st(k) = (e(k) - s(k)) / n
    Next k
    For i = 0 To n
        If i = n Then
            out(i) = Array(CDbl(e(0)), CDbl(e(1)), CDbl(e(2)))
        Else
            out(i) = Array(i * st(0) + s(0), i * st(1) + s(1), i * st(2) + s(2))
        End If
    Next i
    CvLinspace = out
End Function

' n + 1 points from s along dirVec for length l.
Private Function CvSdlPoints(ByVal s As Variant, ByVal dirVec As Variant, ByVal l As Double, ByVal n As Long) As Variant
    Dim out() As Variant
    Dim u(2) As Double
    Dim nrm As Double
    Dim i As Long
    nrm = Sqr(dirVec(0) ^ 2 + dirVec(1) ^ 2 + dirVec(2) ^ 2)
    u(0) = dirVec(0) / nrm: u(1) = dirVec(1) / nrm: u(2) = dirVec(2) / nrm
    ReDim out(0 To n)
    For i = 0 To n
        out(i) = Array(s(0) + i * l * u(0) / n, s(1) + i * l * u(1) / n, s(2) + i * l * u(2) / n)
    Next i
    CvSdlPoints = out
End Function

' Nodes at the points (none made with a group), then n beams / trusses joining them.
Private Function CvLineElems(ByVal pType As String, ByVal pts As Variant, ByVal n As Long, ByVal mat As Long, _
                             ByVal sect As Long, ByVal angle As Double, ByVal GROUP As String) As Variant
    Dim nid() As Long
    Dim out() As Long
    Dim i As Long
    ReDim nid(0 To n)
    For i = 0 To n
        nid(i) = NodeAt(pts(i)(0), pts(i)(1), pts(i)(2))
    Next i
    mGeoLast = Array(CvR6(pts(n)(0)), CvR6(pts(n)(1)), CvR6(pts(n)(2)))
    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        If pType = "TRUSS" Then
            out(i) = AddTruss(nid(i), nid(i + 1), mat, sect, angle, GROUP)
        Else
            out(i) = AddBeam(nid(i), nid(i + 1), mat, sect, angle, GROUP)
        End If
    Next i
    CvLineElems = out
End Function

' Plates between two equal node rows, nDiv rows of plates (new nodes between).
Private Sub CvPlatesBetween(ByVal col As Collection, ByVal a As Variant, ByVal b As Variant, ByVal nDiv As Long, _
                            ByVal STYPE As Long, ByVal mat As Long, ByVal sect As Long, ByVal angle As Double, _
                            ByVal GROUP As String)
    Dim rows() As Variant
    Dim m As Long
    Dim i As Long
    Dim j As Long
    Dim q As Long
    Dim r() As Long
    Dim pts As Variant

    m = UBound(a) - LBound(a) + 1
    If m < 2 Then Exit Sub
    ReDim rows(0 To nDiv)
    rows(0) = a
    rows(nDiv) = b
    If nDiv > 1 Then
        For j = 1 To nDiv - 1
            ReDim r(0 To m - 1)
            rows(j) = r
        Next j
        For i = 0 To m - 1
            pts = CvLinspace(NodeXYZ(a(i)), NodeXYZ(b(i)), nDiv)
            For j = 1 To nDiv - 1
                r = rows(j)
                r(i) = NodeAt(pts(j)(0), pts(j)(1), pts(j)(2))
                rows(j) = r
            Next j
        Next i
    End If
    For q = 0 To nDiv - 1
        For i = 0 To m - 2
            col.Add AddPlate(Array(rows(q)(i), rows(q + 1)(i), rows(q + 1)(i + 1), rows(q)(i + 1)), STYPE, mat, sect, angle, GROUP)
        Next i
    Next q
End Sub

' Make two lists the same length by repeating the last item of the shorter.
Private Sub CvLongest(ByRef a As Variant, ByRef b As Variant)
    Do While UBound(a) < UBound(b)
        a = CvAppend(a, a(UBound(a)))
    Loop
    Do While UBound(b) < UBound(a)
        b = CvAppend(b, b(UBound(b)))
    Loop
End Sub

Private Function CvAppend(ByVal arr As Variant, ByVal v As Variant) As Variant
    Dim out() As Variant
    Dim i As Long
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    ReDim out(0 To n)
    For i = 0 To n - 1
        out(i) = arr(LBound(arr) + i)
    Next i
    out(n) = v
    CvAppend = out
End Function

Private Function CvColToArray(ByVal col As Collection) As Variant
    Dim out() As Long
    Dim i As Long
    If col.Count = 0 Then
        CvColToArray = Array()
        Exit Function
    End If
    ReDim out(0 To col.Count - 1)
    For i = 1 To col.Count
        out(i - 1) = col.Item(i)
    Next i
    CvColToArray = out
End Function

Private Function CvLongs(ByVal p As Variant) As Variant
    Dim out() As Long
    Dim v As Variant
    Dim i As Long
    Dim n As Long
    For Each v In p
        n = n + 1
    Next v
    If n = 0 Then
        CvLongs = Array()
        Exit Function
    End If
    ReDim out(0 To n - 1)
    For Each v In p
        out(i) = CLng(v)
        i = i + 1
    Next v
    CvLongs = out
End Function

Private Function CvNonZero(ByVal p As Variant) As Variant
    Dim col As Collection
    Dim v As Variant
    Set col = New Collection
    For Each v In p
        If CLng(v) <> 0 Then col.Add CLng(v)
    Next v
    CvNonZero = CvColToArray(col)
End Function

' Array of Array(x, y, z) from an array of points or a 3-column range.
Private Function CvPointList(ByVal p As Variant) As Variant
    Dim out() As Variant
    Dim v As Variant
    Dim i As Long
    Dim r As Long
    If IsObject(p) Then p = p.Value
    If CvIs2D(p) Then
        ReDim out(0 To UBound(p, 1) - LBound(p, 1))
        For r = LBound(p, 1) To UBound(p, 1)
            out(r - LBound(p, 1)) = Array(CDbl(p(r, LBound(p, 2))), CDbl(p(r, LBound(p, 2) + 1)), CDbl(p(r, LBound(p, 2) + 2)))
        Next r
    Else
        ReDim out(0 To UBound(p) - LBound(p))
        For Each v In p
            out(i) = Array(CDbl(v(LBound(v))), CDbl(v(LBound(v) + 1)), CDbl(v(LBound(v) + 2)))
            i = i + 1
        Next v
    End If
    CvPointList = out
End Function

Private Function CvMin(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then CvMin = a Else CvMin = b
End Function

Private Function CvMax(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then CvMax = a Else CvMax = b
End Function

' Each node (or element centre) as Array(id, x, y, z).
Private Function CvSelectables(ByVal output As String, ByVal elemType As String) As Collection
    Dim col As Collection
    Dim vKey As Variant
    Dim p As Variant
    Dim oElem As Object
    Dim rec As Object
    Dim v As Variant
    Dim c(2) As Double
    Dim n As Long

    CvGeoSync
    Set col = New Collection
    If UCase$(Left$(output, 4)) = "ELEM" Then
        Set oElem = CvStoreTable("ELEM")
        For Each vKey In oElem.Keys
            Set rec = oElem.Item(vKey)
            If Len(elemType) = 0 Or UCase$(rec.Item("TYPE")) = UCase$(elemType) Then
                c(0) = 0: c(1) = 0: c(2) = 0: n = 0
                For Each v In CvNonZero(rec.Item("NODE"))
                    p = mGeoXYZ.Item(CStr(v))
                    c(0) = c(0) + p(0): c(1) = c(1) + p(1): c(2) = c(2) + p(2)
                    n = n + 1
                Next v
                If n > 0 Then col.Add Array(CLng(vKey), c(0) / n, c(1) / n, c(2) / n)
            End If
        Next vKey
    Else
        For Each vKey In mGeoXYZ.Keys
            p = mGeoXYZ.Item(vKey)
            col.Add Array(CLng(vKey), p(0), p(1), p(2))
        Next vKey
    End If
    Set CvSelectables = col
End Function

Private Function CvSortLongs(ByVal arr As Variant) As Variant
    Dim i As Long, j As Long
    Dim t As Long
    If UBound(arr) < LBound(arr) Then
        CvSortLongs = arr
        Exit Function
    End If
    For i = LBound(arr) + 1 To UBound(arr)
        t = arr(i)
        j = i - 1
        Do While j >= LBound(arr)
            If arr(j) <= t Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = t
    Next i
    CvSortLongs = arr
End Function

' Ids ordered by key, then by id.
Private Function CvSortByKey(ByVal keys As Collection, ByVal ids As Collection) As Variant
    Dim n As Long
    Dim k() As Double
    Dim d() As Long
    Dim i As Long, j As Long
    Dim tk As Double, td As Long
    n = keys.Count
    If n = 0 Then
        CvSortByKey = Array()
        Exit Function
    End If
    ReDim k(0 To n - 1): ReDim d(0 To n - 1)
    For i = 1 To n
        k(i - 1) = keys.Item(i): d(i - 1) = ids.Item(i)
    Next i
    For i = 1 To n - 1
        tk = k(i): td = d(i)
        j = i - 1
        Do While j >= 0
            If k(j) < tk Or (k(j) = tk And d(j) <= td) Then Exit Do
            k(j + 1) = k(j): d(j + 1) = d(j)
            j = j - 1
        Loop
        k(j + 1) = tk: d(j + 1) = td
    Next i
    CvSortByKey = d
End Function

' Even-odd test, the edges (within tol) counting as inside.
Private Function CvInPolygon(ByVal px As Double, ByVal py As Double, pu() As Double, pv() As Double, ByVal tol As Double) As Boolean
    Dim n As Long
    Dim i As Long, j As Long
    Dim du As Double, dv As Double, seg As Double, t As Double
    Dim cu As Double, cv As Double
    Dim inside As Boolean

    n = UBound(pu) + 1
    j = n - 1
    For i = 0 To n - 1
        du = pu(j) - pu(i): dv = pv(j) - pv(i)
        seg = du * du + dv * dv
        If seg = 0 Then
            If (px - pu(i)) ^ 2 + (py - pv(i)) ^ 2 <= tol * tol Then CvInPolygon = True: Exit Function
        Else
            t = ((px - pu(i)) * du + (py - pv(i)) * dv) / seg
            If t < 0 Then t = 0
            If t > 1 Then t = 1
            cu = pu(i) + t * du: cv = pv(i) + t * dv
            If (px - cu) ^ 2 + (py - cv) ^ 2 <= tol * tol Then CvInPolygon = True: Exit Function
        End If
        j = i
    Next i
    j = n - 1
    For i = 0 To n - 1
        If (pv(i) > py) <> (pv(j) > py) Then
            If px < (pu(j) - pu(i)) * (py - pv(i)) / (pv(j) - pv(i)) + pu(i) Then inside = Not inside
        End If
        j = i
    Next i
    CvInPolygon = inside
End Function

Private Function CvSolidFace(ByVal s As String) As Long
    Select Case s
        Case "TTTTFFFF", "TTTFFF", "TTTF": CvSolidFace = 1
        Case "FFFFTTTT", "FFFTTT", "TTFT": CvSolidFace = 2
        Case "TTFFTTFF", "TTFTTF", "FTTT": CvSolidFace = 3
        Case "FTTFFTTF", "FTTFTT", "TFTT": CvSolidFace = 4
        Case "FFTTFFTT", "TFTTFT": CvSolidFace = 5
        Case "TFFTTFFT": CvSolidFace = 6
    End Select
End Function

Private Function CvGroupIds(ByVal names As Variant, ByVal pList As String, ByVal bReverse As Boolean) As Variant
    Dim oGrp As Object
    Dim vKey As Variant
    Dim vName As Variant
    Dim oSeen As Object
    Dim col As Collection
    Dim v As Variant
    Dim out As Variant
    Dim i As Long
    Dim t As Long

    Set oGrp = CvStoreTable("GRUP")
    Set oSeen = New Dictionary
    Set col = New Collection
    For Each vName In CvIds(names)
        For Each vKey In oGrp.Keys
            If CStr(oGrp.Item(vKey).Item("NAME")) = CStr(vName) Then
                For Each v In oGrp.Item(vKey).Item(pList)
                    If Not oSeen.Exists(CStr(v)) Then
                        oSeen.Add CStr(v), True
                        col.Add CLng(v)
                    End If
                Next v
            End If
        Next vKey
    Next vName
    out = CvColToArray(col)
    If bReverse And UBound(out) > 0 Then
        For i = 0 To (UBound(out) - 1) \ 2
            t = out(i): out(i) = out(UBound(out) - i): out(UBound(out) - i) = t
        Next i
    End If
    CvGroupIds = out
End Function

' Points along a curve through pts (one polynomial in the chord parameter
' when deg = number of points - 1, straight pieces when deg = 1), sampled at
' 500 parameter values and divided into n by length or along an axis.
Private Function CvCurvePoints(ByVal pts As Variant, ByVal n As Long, ByVal deg As Long, ByVal divAxis As String) As Variant
    Const NS As Long = 500
    Dim m As Long
    Dim u() As Double
    Dim tot As Double
    Dim i As Long, k As Long
    Dim uf() As Double
    Dim den() As Double
    Dim cl() As Double
    Dim eq As Double
    Dim t As Double
    Dim out() As Variant
    Dim ax As Long
    Dim xp() As Double

    m = UBound(pts) + 1
    ReDim u(0 To m - 1)
    For i = 1 To m - 1
        u(i) = u(i - 1) + Sqr((pts(i)(0) - pts(i - 1)(0)) ^ 2 + (pts(i)(1) - pts(i - 1)(1)) ^ 2 + (pts(i)(2) - pts(i - 1)(2)) ^ 2)
    Next i
    tot = u(m - 1)
    For i = 0 To m - 1
        u(i) = u(i) / tot
    Next i
    ReDim uf(0 To NS - 1): ReDim den(0 To NS - 1, 0 To 2): ReDim cl(0 To NS - 1)
    For i = 0 To NS - 1
        uf(i) = i / (NS - 1)
        For k = 0 To 2
            den(i, k) = CvCurveAt(pts, u, deg, uf(i), k)
        Next k
        If i > 0 Then cl(i) = cl(i - 1) + Sqr((den(i, 0) - den(i - 1, 0)) ^ 2 + (den(i, 1) - den(i - 1, 1)) ^ 2 + (den(i, 2) - den(i - 1, 2)) ^ 2)
    Next i
    ReDim xp(0 To NS - 1)
    Select Case divAxis
        Case "X": ax = 0
        Case "Y": ax = 1
        Case "Z": ax = 2
        Case Else: ax = -1
    End Select
    For i = 0 To NS - 1
        If ax < 0 Then xp(i) = cl(i) Else xp(i) = den(i, ax)
    Next i
    ReDim out(0 To n)
    For i = 0 To n
        If ax < 0 Then
            eq = i * (cl(NS - 1) / n)
        Else
            eq = pts(0)(ax) + i * ((pts(m - 1)(ax) - pts(0)(ax)) / n)
            If i = n Then eq = pts(m - 1)(ax)
        End If
        If ax < 0 And i = n Then eq = cl(NS - 1)
        t = CvInterp(eq, xp, uf)
        out(i) = Array(CvR6(CvCurveAt(pts, u, deg, t, 0)), CvR6(CvCurveAt(pts, u, deg, t, 1)), CvR6(CvCurveAt(pts, u, deg, t, 2)))
    Next i
    CvCurvePoints = out
End Function

Private Function CvCurveAt(ByVal pts As Variant, u() As Double, ByVal deg As Long, ByVal t As Double, ByVal k As Long) As Double
    Dim m As Long
    Dim i As Long, j As Long
    Dim s As Double
    Dim L As Double
    m = UBound(u) + 1
    If deg = 1 Then
        For i = 1 To m - 1
            If t <= u(i) Or i = m - 1 Then
                CvCurveAt = pts(i - 1)(k) + (pts(i)(k) - pts(i - 1)(k)) * (t - u(i - 1)) / (u(i) - u(i - 1))
                Exit Function
            End If
        Next i
    Else
        For i = 0 To m - 1
            L = 1
            For j = 0 To m - 1
                If j <> i Then L = L * (t - u(j)) / (u(i) - u(j))
            Next j
            s = s + L * pts(i)(k)
        Next i
        CvCurveAt = s
    End If
End Function

' numpy.interp: x in increasing xp -> value in fp (ends held).
Private Function CvInterp(ByVal x As Double, xp() As Double, fp() As Double) As Double
    Dim n As Long
    Dim lo As Long, hi As Long, md As Long
    n = UBound(xp) + 1
    If x <= xp(0) Then CvInterp = fp(0): Exit Function
    If x >= xp(n - 1) Then CvInterp = fp(n - 1): Exit Function
    lo = 0: hi = n - 1
    Do While hi - lo > 1
        md = (lo + hi) \ 2
        If xp(md) <= x Then lo = md Else hi = md
    Loop
    CvInterp = fp(lo) + (fp(hi) - fp(lo)) * (x - xp(lo)) / (xp(hi) - xp(lo))
End Function
