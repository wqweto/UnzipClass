VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' $Header: $
'
'   VB6 Unzip Support
'   Copyright (c) 2012 Unicontsoft
'
'   Based on public submissions by Marco v/d Berg and optimizations by John Korejwa.
'   [_+Compression Methods+_ V 1.04](http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37867&lngWId=1)
'
'   Original code is completely refactored to be wrapped in a class module.
'
'   Preprocessor consts:
'     IsPublic - set this if class is to be exposed as public class of an ActiveX
'         DLL or OCX. In this case class is completely self-contained. Otherwise
'         you will need ZcFlagsEnum and ZcFileInfoType in a module of your project.
'         Just search and copy these from source below.
'
'     SkipCrcCheck - controls if CRC is checked after extraction. Disabling it
'         greatly increases extraction speed both in IDE and compiled.
'
'     UnrollInflate - if set unrolls inflate routine to increase speed.
'
'     ProfileCode - if set collects profiling info.
'
'   Known limitations:
'     Has some troubles with memory mapped files if archive or extracted files
'     are too big to fit in free process memory (usually about 1GB)
'
' $Log: $
'
'=========================================================================
Option Explicit
DefObj A-Z

#Const IsPublic = False
#Const ProfileCode = False
#Const SkipCrcCheck = True
#Const UnrollInflate = True

'=========================================================================
' Public events
'=========================================================================

Event Error(ByVal Idx As Long, Error As String)
Event Progress(ByVal Idx As Long, ByVal Pos As Long, ByVal Size As Long)

'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const CREATE_ALWAYS         As Long = 2
Private Const OPEN_EXISTING         As Long = &H3
Private Const GENERIC_WRITE         As Long = &H40000000
Private Const GENERIC_READ          As Long = &H80000000
'--- for CreateFileMapping
Private Const PAGE_READONLY         As Long = 2
Private Const PAGE_READWRITE        As Long = 4
'--- for MapViewOfFile
Private Const FILE_MAP_WRITE        As Long = &H2
Private Const FILE_MAP_READ         As Long = &H4
'--- for MultiByteToWideChar
Private Const CP_OEMCP              As Long = 1

Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Long
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Boolean
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cbMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Private Type SAFEARRAY
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    cElements           As Long
    lLbound             As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LEN_LOCALHEADER       As Long = 30
Private Const LEN_EXTLOCALHEADER    As Long = 16
Private Const LEN_CENTRALHEADER     As Long = 46
Private Const LEN_ENDHEADER         As Long = 22

#If IsPublic Then
    Public Enum ZcFlagsEnum
        zcFlgEncrypted = 2 ^ 0                  'bit 0 set = file is encrypted
        zcFlgUsedMed = 2 ^ 1 + 2 ^ 2            'bit 1+2 depending on compression type
                                                'type = 6 (imploding)
                                                'bit 1 set = use 8k dictionary else 4k dictionary
                                                'bit 2 set = use 3 trees else use 2 trees
                                                'type = 8 (deflating)
                                                'bit 2 : 1
                                                '    0   0 = Normal (-en) compression option was used.
                                                '    0   1 = Maximum (-exx/-ex) compression option was used
                                                '    1   0 = Fast (-ef) compression option was used
                                                '    1   1 = Super Fast (-es) compression option was used
                                                'bits are undefined if other methods are used
        zcFlgExtLocHead = 2 ^ 3                 'bit 3 set = Extended local header is used to store CRC and size
        zcFlgRes64 = 2 ^ 4                      'bit 4 Reserved for ZIP64
        zcFlgPathed = 2 ^ 5                     'bit 5 set = file is compressed pathed data
        zcFlgEncStrong = 2 ^ 6                  'bit 6 set = file is encrypted using strong encryption
    End Enum
    
    Public Type ZcFileInfoType
        SystemMadeBy    As String
        VersionMadeBy   As String
        SystemNeeded    As String
        VersionNeeded   As String
        Flags           As ZcFlagsEnum
        Method          As Integer
        DateTime        As Date
        CRC32           As Long
        CompressedSize  As Long
        Size            As Long
        DiskStart       As Integer
        AttribI         As Integer
        AttribX         As Long
        Offset          As Long
        FileName        As String
        IsDir           As Boolean
        Extension()     As Byte
        Comment         As String
    End Type
#End If

Private Type ZcLocalHeaderType
    Signature       As Long                     'Signature
    VerExt          As Integer                  'version needed to extract
    Flags           As Integer                  'encrypt and compression flags
    Method          As Integer                  'compression method
    FTime           As Integer                  'time last modifies, dos format
    FDate           As Integer                  'date last modifies, dos format
    '--- padding
    CRC32           As Long                     'CRC32 for uncompressed file
    CSize           As Long                     'compressed size
    USize           As Long                     'uncompressed size
    LenFname        As Integer                  'Length filename
    LenExt          As Integer                  'Length for extra field
End Type

Private Type ZcCentralHeaderType
    Signature       As Long                     'Signature
    VerMade         As Integer                  'version made by
    VerExt          As Integer                  'version needed to extract
    Flags           As Integer                  'encrypt and compression flags
    Method          As Integer                  'compression method
    FTime           As Integer                  'time last modifies, dos format
    FDate           As Integer                  'date last modifies, dos format
    '--- padding
    CRC32           As Long                     'CRC32 for uncompressed file
    CSize           As Long                     'compressed size
    USize           As Long                     'uncompressed size
    LenFname        As Integer                  'Length filename
    LenExt          As Integer                  'Length for extra field
    LenCom          As Integer                  'Length for comment field
    DiskStart       As Integer                  'start disk number
    AttribI         As Integer                  'internal file attributes
    '--- padding
    AttribX         As Long                     'external file attributes
    Offset          As Long                     'relative offset of local header
End Type

Private Type ZcEndHeaderType
    Signature       As Long                     'Signature
    DiskNum         As Integer                  'this disk number
    DiskStart       As Integer                  'start disk number
    Entries         As Integer                  'Entries on this disk
    TotEntr         As Integer                  'Number of total entries
    CenSize         As Long                     'size of entire cetral directory
    CenOff          As Long                     'offset of central on starting disk
    LenCom          As Integer                  'Length of comment field
End Type

Private Type ZcMapFileType
    uSA             As SAFEARRAY
    hFile           As Long
    hFileMap        As Long
    lPtr            As Long
End Type

Private Type ZcCodesType
    Length()        As Long
    Code()          As Long
End Type

Private Type ZcUnpackContextType
    '--- buffer management
    baOutStream()   As Byte
    lOutPos         As Long
    uOutSA          As SAFEARRAY
    baInStream()    As Byte
    lInPos          As Long
    uInSA           As SAFEARRAY
    lInBuff         As Long
    lInBits         As Long
    aPow2(16)       As Long
    aMask(16)       As Long
    '--- huffman coding
    uLitLen         As ZcCodesType          'Literal/length tree
    uDist           As ZcCodesType          'Distance tree
    aLenOrder()     As Long
    uLC             As ZcCodesType
    uDC             As ZcCodesType
    lMinLitLen      As Long           'Minimum length used in literal/length codes
    lMaxLitLen      As Long           'Maximum length used in literal/length codes
    lMinDist        As Long           'Minimum length used in distance codes
    lMaxDist        As Long           'Maximum length used in distance codes
    bStaticInit     As Boolean
    uStaticLitLen   As ZcCodesType
    uStaticDist     As ZcCodesType
    #If ProfileCode Then
        lCountChar      As Long
        lCountDist      As Long
        lCountCopy      As Long
        lCountLoop      As Long
        lCountAPI       As Long
        lCountBits      As Long
    #End If
End Type

'Signatures long version
Private Const LNG_LOCAL_SIG         As Long = &H4034B50
Private Const LNG_CENTRAL_SIG       As Long = &H2014B50
Private Const LNG_END_SIG           As Long = &H6054B50
Private Const LNG_EXTLOCAL_SIG      As Long = &H8074B50

Private m_uFileData()       As ZcFileInfoType
Private m_sZipName          As String
Private m_sComments         As String
Private m_lCrc32Table(255)  As Long
Private m_lCurrentFile      As Long
Private m_sLastError        As String

'=========================================================================
' Properties
'=========================================================================

Property Get ZipName() As String
    ZipName = m_sZipName
End Property

Public Property Get Comments() As String
    Comments = m_sComments
End Property

Public Property Get Count() As Long
    On Error Resume Next
    Count = UBound(m_uFileData)
    On Error GoTo 0
End Property

Public Property Get LastError() As String
    LastError = m_sLastError
End Property

#If IsPublic Then
Public Property Get File(ByVal IdxZeroBased As Long) As ZcFileInfoType
#Else
Friend Property Get File(ByVal IdxZeroBased As Long) As ZcFileInfoType
#End If
    If IdxZeroBased >= 0 Then
        File = m_uFileData(IdxZeroBased)
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sZipName As String) As Boolean
    Dim uEndHdr         As ZcEndHeaderType
    Dim uCentralHdr     As ZcCentralHeaderType
    Dim uLocalHdr       As ZcLocalHeaderType
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim uZipMap         As ZcMapFileType
    Dim baZipMap()      As Byte
    Dim cOffset         As Currency
    
    On Error GoTo EH
    '--- init member vars
    m_sZipName = sZipName
    m_sComments = vbNullString
    m_lCurrentFile = -1
    Erase m_uFileData
    If Not pvMapFile(baZipMap, uZipMap, 0, 1, sZipName) Then
        GoTo QH
    End If
    '--- try end header direct lookup at EOF - LEN_ENDHEADER
    Call GetFileSizeEx(uZipMap.hFile, cOffset)
    If Not pvMapFile(baZipMap, uZipMap, cOffset - LEN_ENDHEADER / 10000@, LEN_ENDHEADER) Then
        GoTo QH
    End If
    Call CopyMemory(uEndHdr, baZipMap(0), LEN_ENDHEADER)
    If uEndHdr.Signature = LNG_END_SIG Then
        cOffset = pvInc(0, uEndHdr.CenOff)
        lCount = uEndHdr.Entries
    Else
        '--- fall back to sequential lookup from start of file
        lCount = 0
        cOffset = 0
        Do
            If Not pvMapFile(baZipMap, uZipMap, (cOffset), LEN_LOCALHEADER) Then
                GoTo QH
            End If
            Call CopyMemory(uLocalHdr, baZipMap(0), 4)
            Select Case uLocalHdr.Signature
            Case LNG_LOCAL_SIG
                Call CopyMemory(uLocalHdr, baZipMap(0), 14)
                Call CopyMemory(uLocalHdr.CRC32, baZipMap(14), LEN_LOCALHEADER - 14)
                cOffset = pvInc(cOffset, LEN_LOCALHEADER + uLocalHdr.CSize + uLocalHdr.LenFname + uLocalHdr.LenExt)
                lCount = lCount + 1
            Case LNG_CENTRAL_SIG
                Exit Do
            Case LNG_EXTLOCAL_SIG
                cOffset = pvInc(cOffset, LEN_EXTLOCALHEADER)
            Case Else
                pvError "Unknown header"
                GoTo QH
            End Select
        Loop
    End If
    '--- collect info from central headers
    ReDim m_uFileData(0 To lCount) As ZcFileInfoType
    For lIdx = 0 To lCount - 1
        m_lCurrentFile = lIdx
        '--- map central header
        If Not pvMapFile(baZipMap, uZipMap, cOffset, LEN_CENTRALHEADER) Then
            GoTo QH
        End If
        Call CopyMemory(uCentralHdr, baZipMap(0), 16)
        Call CopyMemory(uCentralHdr.CRC32, baZipMap(16), 22)
        Call CopyMemory(uCentralHdr.AttribX, baZipMap(16 + 22), LEN_CENTRALHEADER - 16 - 22)
        If uCentralHdr.Signature <> LNG_CENTRAL_SIG Then
            pvError "Invalid center header"
            GoTo QH
        End If
        '--- extract file info
        With m_uFileData(lIdx)
            .SystemMadeBy = ((uCentralHdr.VerMade \ 256) And 255)
            .VersionMadeBy = (uCentralHdr.VerMade And 255) \ 10 & "." & (uCentralHdr.VerMade Mod 10)
            .SystemNeeded = ((uCentralHdr.VerExt \ 256) And 255)
            .VersionNeeded = (uCentralHdr.VerExt And 255) \ 10 & "." & (uCentralHdr.VerExt Mod 10)
            .Flags = uCentralHdr.Flags
            .Method = uCentralHdr.Method
            .DateTime = pvToDate(uCentralHdr.FDate, uCentralHdr.FTime)
            .CRC32 = uCentralHdr.CRC32
            .CompressedSize = uCentralHdr.CSize
            .Size = uCentralHdr.USize
            .DiskStart = uCentralHdr.DiskStart
            .AttribI = uCentralHdr.AttribI
            .AttribX = uCentralHdr.AttribX
            .Offset = uCentralHdr.Offset
            If Not pvMapFile(baZipMap, uZipMap, cOffset, uCentralHdr.LenFname + uCentralHdr.LenExt + uCentralHdr.LenCom) Then
                GoTo QH
            End If
            .FileName = Replace(pvToString(baZipMap, 0, uCentralHdr.LenFname), "/", "\")
            .IsDir = (.CompressedSize = 0 And Right$(.FileName, 1) = "\")
            If uCentralHdr.LenExt <> 0 Then
                ReDim .Extension(0 To pvToUnsigned(uCentralHdr.LenExt) - 1) As Byte
                Call CopyMemory(.Extension(0), baZipMap(pvToUnsigned(uCentralHdr.LenFname)), UBound(.Extension) + 1)
            End If
            .Comment = pvToString(baZipMap, uCentralHdr.LenFname + uCentralHdr.LenExt, uCentralHdr.LenCom)
        End With
    Next
    '--- extract zip comment from end header
    If Not pvMapFile(baZipMap, uZipMap, cOffset, LEN_ENDHEADER) Then
        GoTo QH
    End If
    Call CopyMemory(uEndHdr, baZipMap(0), LEN_ENDHEADER)
    If uEndHdr.LenCom > 0 Then
        If Not pvMapFile(baZipMap, uZipMap, cOffset, uEndHdr.LenCom) Then
            GoTo QH
        End If
        m_sComments = pvToString(baZipMap, 0, uEndHdr.LenCom)
    End If
    '--- success
    Init = True
QH:
    pvMapFile baZipMap, uZipMap
    Exit Function
EH:
    pvMapFile baZipMap, uZipMap
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume
End Function

Public Function Unzip( _
            sTargetPath As String, _
            Optional vFilter As Variant) As Boolean
    Dim uLocalHdr       As ZcLocalHeaderType
    Dim lIdx            As Long
    Dim lOutputCrc32    As Long
    Dim uZipMap         As ZcMapFileType
    Dim baZipMap()      As Byte
    Dim uUnpackMap      As ZcMapFileType
    Dim baUnpackMap()   As Byte
    Dim cOffset         As Currency
    Dim sOutputFile       As String
    
    On Error GoTo EH
    m_lCurrentFile = -1
    If Not pvMapFile(baZipMap, uZipMap, 0, 1, m_sZipName) Then
        GoTo QH
    End If
    For lIdx = 0 To Count - 1
        m_lCurrentFile = lIdx
        '--- check if current entry is filtered
        If IsArray(vFilter) Then
            If Not At(vFilter, lIdx) Then
                GoTo SkipFile
            End If
        ElseIf Not IsMissing(vFilter) Then
            If Not m_uFileData(lIdx).FileName Like IIf(vFilter = "*.*", "*", vFilter) Then
                GoTo SkipFile
            End If
        End If
        '--- handle folders
        If m_uFileData(lIdx).IsDir Then
            If Not MkPath(sTargetPath & IIf(Right$(sTargetPath, 1) <> "\", "\", vbNullString) & m_uFileData(lIdx).FileName) Then
                pvError "Error creating folder"
            End If
            GoTo SkipFile
        End If
        '--- read local header
        cOffset = pvInc(0, m_uFileData(lIdx).Offset)
        If Not pvMapFile(baZipMap, uZipMap, cOffset, LEN_LOCALHEADER) Then
            GoTo QH
        End If
        Call CopyMemory(uLocalHdr, baZipMap(0), 14)
        Call CopyMemory(uLocalHdr.CRC32, baZipMap(14), LEN_LOCALHEADER - 14)
        cOffset = pvInc(cOffset, uLocalHdr.LenFname + uLocalHdr.LenExt)
        '--- check supported features
        If uLocalHdr.Signature <> LNG_LOCAL_SIG Then
            pvError "Invalid signature"
        End If
        If (m_uFileData(lIdx).Flags And 1) <> 0 Then
            pvError "Encryption not supported"
            GoTo SkipFile
        End If
        Select Case m_uFileData(lIdx).Method
        Case 0, 8, 9
        Case 2, 3, 4, 5
            pvError "UnReduce not supported"
            GoTo SkipFile
        Case Else '--- 1, 6, 7, 10
            pvError "Compression not supported"
            GoTo SkipFile
        End Select
        '--- map compressed portion
        If m_uFileData(lIdx).CompressedSize <> 0 Then
            If Not pvMapFile(baZipMap, uZipMap, cOffset, m_uFileData(lIdx).CompressedSize) Then
                GoTo SkipFile
            End If
        End If
        '--- map output file
        sOutputFile = sTargetPath & IIf(Right$(sTargetPath, 1) <> "\", "\", vbNullString) & m_uFileData(lIdx).FileName
        If Not MkPath(Left$(sOutputFile, InStrRev(sOutputFile, "\"))) Then
            pvError "Error creating output folder"
            GoTo SkipFile
        End If
        If Not pvMapFile(baUnpackMap, uUnpackMap, 0, -m_uFileData(lIdx).Size, sOutputFile) Then
            GoTo SkipFile
        End If
        '--- decompress/move bits and calc CRC32
        If m_uFileData(lIdx).Size = 0 Then
            lOutputCrc32 = 0
        ElseIf m_uFileData(lIdx).Method = 0 Then
            If Not pvMoveBits(VarPtr(baZipMap(0)), m_uFileData(lIdx).CompressedSize, VarPtr(baUnpackMap(0)), m_uFileData(lIdx).Size, lOutputCrc32) Then
                pvMapFile baUnpackMap, uUnpackMap
                GoTo SkipFile
            End If
        Else '--- 8, 9
            If Not pvInflate(VarPtr(baZipMap(0)), m_uFileData(lIdx).CompressedSize, VarPtr(baUnpackMap(0)), m_uFileData(lIdx).Size, lOutputCrc32, m_uFileData(lIdx).Method = 9) Then
                pvMapFile baUnpackMap, uUnpackMap
                GoTo SkipFile
            End If
        End If
        #If Not SkipCrcCheck Then
            If m_uFileData(lIdx).CRC32 <> lOutputCrc32 Then
                pvError "CRC check failed"
            End If
        #End If
        Call SetFileTime(uUnpackMap.hFile, pvToFileTime(m_uFileData(lIdx).DateTime), pvToFileTime(Now), pvToFileTime(m_uFileData(lIdx).DateTime))
        pvMapFile baUnpackMap, uUnpackMap
SkipFile:
    Next
    '--- success
    Unzip = True
QH:
    pvMapFile baZipMap, uZipMap
    Exit Function
EH:
    pvMapFile baZipMap, uZipMap
    pvMapFile baUnpackMap, uUnpackMap
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function pvToDate(ByVal nDate As Integer, ByVal nTime As Integer) As Date
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME

    If nDate <> 0 Or nTime <> 0 Then
        Call DosDateTimeToFileTime(nDate, nTime, uFileTime)
        Call FileTimeToSystemTime(uFileTime, uSysTime)
        With uSysTime
            pvToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
        End With
    End If
End Function

Private Function pvToFileTime(ByVal dCreated As Date) As FILETIME
    Dim uSysTime        As SYSTEMTIME
    Dim uFileTime       As FILETIME
    
    With uSysTime                           'Convert Dates to FILETIMEs
        .wYear = Year(dCreated)
        .wMonth = Month(dCreated)
        .wDay = Day(dCreated)
        .wHour = Hour(dCreated)
        .wMinute = Minute(dCreated)
        .wSecond = Second(dCreated)
    End With
    Call SystemTimeToFileTime(uSysTime, uFileTime)
    Call LocalFileTimeToFileTime(uFileTime, pvToFileTime)
End Function

Private Function pvToString(baMap() As Byte, ByVal nOffset As Integer, ByVal nSize As Integer) As String
    If nSize <> 0 Then
        pvToString = String(pvToUnsigned(nSize), 0)
        Call MultiByteToWideChar(CP_OEMCP, 0, baMap(pvToUnsigned(nOffset)), pvToUnsigned(nSize), ByVal StrPtr(pvToString), Len(pvToString))
    End If
End Function

Private Function pvToUnsigned(ByVal nValue As Integer) As Long
    pvToUnsigned = IIf(nValue >= 0, nValue, &HFFFF& + nValue + 1)
End Function

Private Function MkPath(sPath As String) As Boolean
    Dim lPos            As Long
    
    On Error Resume Next
    If LenB(sPath) <> 0 Then
        If GetAttr(sPath) = -1 Then
            Err.Clear
            MkDir sPath
            If Err.Number <> 0 Then
                lPos = Len(sPath) + 1
                Do
                    lPos = InStrRev(sPath, "\", lPos - 1)
                    If lPos <= 1 Then
                        Exit Do
                    End If
                    If GetAttr(Left$(sPath, lPos - 1)) = -1 Then
                    Else
                        Exit Do
                    End If
                Loop
                Do
                    lPos = InStr(lPos + 1, sPath, "\")
                    If lPos = 0 Then
                        Exit Do
                    End If
                    MkDir Left$(sPath, lPos - 1)
                Loop
                Err.Clear
                If GetAttr(sPath) = -1 Then
                    MkDir sPath
                End If
            End If
        End If
    End If
    '--- success (or failure)
    MkPath = (Err.Number = 0)
End Function

Private Sub pvError(sError As String)
    Debug.Print sError, File(m_lCurrentFile).FileName, Timer
    m_sLastError = sError
    RaiseEvent Error(m_lCurrentFile, m_sLastError)
End Sub

Private Function pvMapFile(baBits() As Byte, uMap As ZcMapFileType, Optional cOffset As Currency, Optional ByVal lSize As Long, Optional sFile As String) As Boolean
    Dim bReadOnly       As Boolean
    Dim aQuad(0 To 1)   As Long
    Dim lOffset         As Long
    
    If LenB(sFile) = 0 And cOffset = 0 And lSize = 0 Then
        Call CopyMemory(ByVal ArrPtr(baBits()), 0&, 4)
        '--- success
        pvMapFile = True
        GoTo CleanUp
    Else
        bReadOnly = lSize > 0
        lSize = Abs(lSize)
        If uMap.lPtr <> 0 Then
            Call UnmapViewOfFile(uMap.lPtr)
            uMap.lPtr = 0
        End If
        If LenB(sFile) <> 0 Then
            If uMap.hFileMap <> 0 Then
                Call CloseHandle(uMap.hFileMap)
                uMap.hFileMap = 0
            End If
            If uMap.hFile <> -1 Then
                Call CloseHandle(uMap.hFile)
                uMap.hFile = -1
            End If
            uMap.hFile = CreateFile(sFile, GENERIC_READ Or GENERIC_WRITE, 0, 0, IIf(bReadOnly, OPEN_EXISTING, CREATE_ALWAYS), 0, 0)
            If uMap.hFile = -1 Then
                GoTo QH
            End If
            uMap.hFileMap = CreateFileMapping(uMap.hFile, 0, IIf(bReadOnly, PAGE_READONLY, PAGE_READWRITE), 0, IIf(bReadOnly, 0, lSize), 0)
            If uMap.hFileMap = 0 Then
                GoTo QH
            End If
        End If
        Call CopyMemory(aQuad(0), cOffset, 8)
        lOffset = aQuad(0) And &HFFFF&
        aQuad(0) = aQuad(0) And Not &HFFFF&
        uMap.lPtr = MapViewOfFile(uMap.hFileMap, IIf(bReadOnly, FILE_MAP_READ, FILE_MAP_WRITE), aQuad(1), aQuad(0), 0) ' lOffset + lSize)
        If uMap.lPtr = 0 Then
            GoTo QH
        End If
        With uMap.uSA
            .cbElements = 1
            .cDims = 1
            .cElements = lSize
            .pvData = uMap.lPtr + lOffset
        End With
        Call CopyMemory(ByVal ArrPtr(baBits()), VarPtr(uMap.uSA), 4)
        cOffset = pvInc(cOffset, lSize)
    End If
    '--- success
    pvMapFile = True
    Exit Function
QH:
    pvError "Error mapping file " & sFile
CleanUp:
    If uMap.lPtr <> 0 Then
        Call UnmapViewOfFile(uMap.lPtr)
        uMap.lPtr = 0
    End If
    If uMap.hFileMap <> 0 Then
        Call CloseHandle(uMap.hFileMap)
        uMap.hFileMap = 0
    End If
    If uMap.hFile <> -1 Then
        Call CloseHandle(uMap.hFile)
        uMap.hFile = -1
    End If
End Function

Private Function pvInc(ByVal cOffset As Currency, ByVal lInc As Long) As Currency
    pvInc = cOffset + lInc / 10000@ + IIf(lInc < 0, 429496.7296@, 0@)
End Function

Private Function At(vArray As Variant, ByVal lIdx As Long) As Boolean
    On Error Resume Next
    At = vArray(lIdx)
    On Error GoTo 0
End Function

Private Function pvMoveBits( _
            ByVal lPackedPtr As Long, _
            ByVal lPackedSize As Long, _
            ByVal lOutputPtr As Long, _
            lOutputSize As Long, _
            lOutputCrc32 As Long) As Boolean
    Const CHUNK_SIZE    As Long = &H10000
    Dim lOffset         As Long
    Dim lSize           As Long
    Dim lIdx            As Long
    Dim uCtx            As ZcUnpackContextType
    
    On Error GoTo EH
    '--- map byte arrays
    pvMapArray uCtx.baInStream, uCtx.uInSA, lPackedPtr, lPackedSize
    pvMapArray uCtx.baOutStream, uCtx.uOutSA, lOutputPtr, lOutputSize
    lOutputCrc32 = -1
    Do While lOffset < lPackedSize
        If lOffset + CHUNK_SIZE > lPackedSize Then
            lSize = lPackedSize - lOffset
        Else
            lSize = CHUNK_SIZE
        End If
        Call CopyMemory(uCtx.baOutStream(lOffset), uCtx.baInStream(lOffset), lSize)
        #If Not SkipCrcCheck Then
            For lIdx = 0 To lSize - 1
                lOutputCrc32 = (((lOutputCrc32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (m_lCrc32Table((lOutputCrc32 And &HFF) Xor uCtx.baInStream(lOffset + lIdx)))
            Next
        #End If
        lOffset = lOffset + lSize
    Loop
    lOutputCrc32 = lOutputCrc32 Xor &HFFFFFFFF
    '--- success
    pvMoveBits = True
QH:
    pvMapArray uCtx.baInStream, uCtx.uInSA
    pvMapArray uCtx.baOutStream, uCtx.uOutSA
    Exit Function
EH:
    pvMapArray uCtx.baInStream, uCtx.uInSA
    pvMapArray uCtx.baOutStream, uCtx.uOutSA
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'= Inflate routines ======================================================

Private Function pvInflate( _
            ByVal lPackedPtr As Long, _
            ByVal lPackedSize As Long, _
            ByVal lOutputPtr As Long, _
            lOutputSize As Long, _
            lOutputCrc32 As Long, _
            ByVal bZIP64 As Boolean) As Boolean
    Dim bIsLastBlock    As Boolean
    Dim lBlockType      As Long
    Dim lChar           As Long
    Dim lBits           As Long
    Dim lSize           As Long
    Dim lComplement     As Long
    Dim lDist           As Long
    Dim lIdx            As Long
    Dim lCrc32Pos       As Long
    Dim uCtx            As ZcUnpackContextType
    Dim aPow2(16)       As Long
    Dim aMask(16)       As Long
    #If UnrollInflate Then
        Dim lInBuff         As Long
        Dim lInBits         As Long
        Dim lNumBits        As Long
        Dim lInPos          As Long
        Dim lOutPos         As Long
    #End If
    
    On Error GoTo EH
    '--- init inflate
    pvSetArray uCtx.aLenOrder, Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    pvSetArray uCtx.uLC.Code, Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    pvSetArray uCtx.uLC.Length, Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    pvSetArray uCtx.uDC.Code, Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153)
    pvSetArray uCtx.uDC.Length, Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14)
    For lIdx = 0 To 16
        uCtx.aPow2(lIdx) = 2 ^ lIdx
        uCtx.aMask(lIdx) = uCtx.aPow2(lIdx) - 1
        aPow2(lIdx) = uCtx.aPow2(lIdx)
        aMask(lIdx) = uCtx.aMask(lIdx)
    Next
    lOutputCrc32 = -1
    '--- map byte arrays
    pvMapArray uCtx.baInStream, uCtx.uInSA, lPackedPtr, lPackedSize
    pvMapArray uCtx.baOutStream, uCtx.uOutSA, lOutputPtr, lOutputSize
    Do
        bIsLastBlock = pvGetBits(uCtx, 1)
        lBlockType = pvGetBits(uCtx, 2)
        If lBlockType = 0 Then              'Block is Stored
            'this is done couse if lInBits >= then next byte is already in lInBuff
            Do While uCtx.lInBits >= 8
                uCtx.lInPos = uCtx.lInPos - 1
                uCtx.lInBits = uCtx.lInBits - 8
            Loop
            If uCtx.lInPos + 4 > UBound(uCtx.baInStream) Then
                pvError "InStream depleated"
                GoTo QH
            End If
            Call CopyMemory(lSize, uCtx.baInStream(uCtx.lInPos), 2)       'Read Count
            Call CopyMemory(lComplement, uCtx.baInStream(uCtx.lInPos + 2), 2)   'Read ones compliment of Count
            uCtx.lInPos = uCtx.lInPos + 4
            If lSize <> (Not lComplement And &HFFFF&) Then
                pvError "Invalid block size"
                GoTo QH
            End If
            If uCtx.lInPos + lSize - 1 > UBound(uCtx.baInStream) Then
                pvError "InStream depleated"
                GoTo QH
            End If
            If uCtx.lOutPos + lSize - 1 > UBound(uCtx.baOutStream) Then
                pvError "OutStream overflow"
                GoTo QH
            End If
            Call CopyMemory(uCtx.baOutStream(uCtx.lOutPos), uCtx.baInStream(uCtx.lInPos), lSize)
            uCtx.lOutPos = uCtx.lOutPos + lSize
            uCtx.lInPos = uCtx.lInPos + lSize
            uCtx.lInBuff = 0
            uCtx.lInBits = 0
        ElseIf lBlockType = 3 Then
            pvError "Error in compressed data"
            GoTo QH
        Else
            If lBlockType = 1 Then          'Static Compression
                If Not pvCreateStaticTree(uCtx) Then
                    GoTo QH
                End If
            Else 'lBlockType = 2            'Dynamic Compression
                If Not pvCreateDynamicTree(uCtx) Then
                    GoTo QH
                End If
            End If
            #If UnrollInflate Then
                lInBuff = uCtx.lInBuff
                lInBits = uCtx.lInBits
                lInPos = uCtx.lInPos
                lOutPos = uCtx.lOutPos
                Do
                    '{  pvNeedBits uCtx.lMaxLitLen
                    Do While lInBits < uCtx.lMaxLitLen
                        If lInPos <= UBound(uCtx.baInStream) Then
                            lInBuff = lInBuff + (uCtx.baInStream(lInPos) * aPow2(lInBits))
                        End If
                        lInBits = lInBits + 8
                        lInPos = lInPos + 1
                    Loop
                    '}
                    lBits = uCtx.lMinLitLen
                    Do While uCtx.uLitLen.Length(lInBuff And aMask(lBits)) <> lBits
                        lBits = lBits + 1
                    Loop
                    lChar = uCtx.uLitLen.Code(lInBuff And aMask(lBits))
                    '{  pvDropBits lBits
                    lInBuff = lInBuff \ aPow2(lBits)
                    lInBits = lInBits - lBits
                    '}
                    If lChar < 256 Then  'Character is Literal
                        uCtx.baOutStream(lOutPos) = lChar
                        lOutPos = lOutPos + 1
                    ElseIf lChar > 256 Then 'Character is Length Symbol
                        lChar = lChar - 257
                        '{  lSize = uCtx.uLC.Code(lChar) + pvGetBits(uCtx, uCtx.uLC.Length(lChar))
                        lSize = uCtx.uLC.Code(lChar)
                        lNumBits = uCtx.uLC.Length(lChar)
                        Do While lInBits < lNumBits
                            If lInPos <= UBound(uCtx.baInStream) Then
                                lInBuff = lInBuff + (uCtx.baInStream(lInPos) * aPow2(lInBits))
                            End If
                            lInBits = lInBits + 8
                            lInPos = lInPos + 1
                        Loop
                        lSize = lSize + (lInBuff And aMask(lNumBits))
                        lInBuff = lInBuff \ aPow2(lNumBits)
                        lInBits = lInBits - lNumBits
                        '}
                        If (lSize = 258) And bZIP64 Then
                            lSize = pvGetBits(uCtx, 16) + 3
                        End If
                        '{  pvNeedBits uCtx.lMaxDist
                        Do While lInBits < uCtx.lMaxDist
                            If lInPos <= UBound(uCtx.baInStream) Then
                                lInBuff = lInBuff + (uCtx.baInStream(lInPos) * aPow2(lInBits))
                            End If
                            lInBits = lInBits + 8
                            lInPos = lInPos + 1
                        Loop
                        '}
                        lBits = uCtx.lMinDist
                        Do While uCtx.uDist.Length(lInBuff And aMask(lBits)) <> lBits
                            lBits = lBits + 1
                        Loop
                        lChar = uCtx.uDist.Code(lInBuff And aMask(lBits))
                        '{  pvDropBits lBits
                        lInBuff = lInBuff \ aPow2(lBits)
                        lInBits = lInBits - lBits
                        '}
                        '{  lDist = uCtx.uDC.Code(lChar) + pvGetBits(uCtx, uCtx.uDC.Length(lChar))
                        lDist = uCtx.uDC.Code(lChar)
                        lNumBits = uCtx.uDC.Length(lChar)
                        Do While lInBits < lNumBits
                            If lInPos <= UBound(uCtx.baInStream) Then
                                lInBuff = lInBuff + (uCtx.baInStream(lInPos) * aPow2(lInBits))
                            End If
                            lInBits = lInBits + 8
                            lInPos = lInPos + 1
                        Loop
                        lDist = lDist + (lInBuff And aMask(lNumBits))
                        lInBuff = lInBuff \ aPow2(lNumBits)
                        lInBits = lInBits - lNumBits
                        '}
                        If lDist < lSize Then
                            For lIdx = 0 To lSize - 1
                                uCtx.baOutStream(lOutPos) = uCtx.baOutStream(lOutPos - lDist)
                                lOutPos = lOutPos + 1
                            Next
                        Else
                            Call CopyMemory(uCtx.baOutStream(lOutPos), uCtx.baOutStream(lOutPos - lDist), lSize)
                            lOutPos = lOutPos + lSize
                        End If
                    Else '--- 256
                        Exit Do
                    End If
                Loop
                uCtx.lInBuff = lInBuff
                uCtx.lInBits = lInBits
                uCtx.lInPos = lInPos
                uCtx.lOutPos = lOutPos
            #Else
                Do
                    pvNeedBits uCtx, uCtx.lMaxLitLen
                    lBits = uCtx.lMinLitLen
                    Do While uCtx.uLitLen.Length(uCtx.lInBuff And aMask(lBits)) <> lBits
                        #If ProfileCode Then
                            uCtx.lCountBits = uCtx.lCountBits + 1
                        #End If
                        lBits = lBits + 1
                    Loop
                    lChar = uCtx.uLitLen.Code(uCtx.lInBuff And aMask(lBits))
                    pvDropBits uCtx, lBits
                    If lChar < 256 Then  'Character is Literal
                        #If ProfileCode Then
                            uCtx.lCountChar = uCtx.lCountChar + 1
                        #End If
                        uCtx.baOutStream(uCtx.lOutPos) = lChar
                        uCtx.lOutPos = uCtx.lOutPos + 1
                    ElseIf lChar > 256 Then 'Character is Length Symbol
                        #If ProfileCode Then
                            uCtx.lCountDist = uCtx.lCountDist + 1
                        #End If
                        lChar = lChar - 257
                        lSize = uCtx.uLC.Code(lChar) + pvGetBits(uCtx, uCtx.uLC.Length(lChar))
                        If (lSize = 258) And bZIP64 Then
                            lSize = pvGetBits(uCtx, 16) + 3
                        End If
                        pvNeedBits uCtx, uCtx.lMaxDist
                        lBits = uCtx.lMinDist
                        Do While uCtx.uDist.Length(uCtx.lInBuff And aMask(lBits)) <> lBits
                            lBits = lBits + 1
                        Loop
                        lChar = uCtx.uDist.Code(uCtx.lInBuff And aMask(lBits))
                        pvDropBits uCtx, lBits
                        lDist = uCtx.uDC.Code(lChar) + pvGetBits(uCtx, uCtx.uDC.Length(lChar))
                        If lDist < lSize Then
                            #If ProfileCode Then
                                uCtx.lCountLoop = uCtx.lCountLoop + 1
                                uCtx.lCountCopy = uCtx.lCountCopy + lSize
                            #End If
                            For lIdx = 0 To lSize - 1
                                uCtx.baOutStream(uCtx.lOutPos) = uCtx.baOutStream(uCtx.lOutPos - lDist)
                                uCtx.lOutPos = uCtx.lOutPos + 1
                            Next
                        Else
                            #If ProfileCode Then
                                uCtx.lCountAPI = uCtx.lCountAPI + 1
                                uCtx.lCountCopy = uCtx.lCountCopy + lSize
                            #End If
                            Call CopyMemory(uCtx.baOutStream(uCtx.lOutPos), uCtx.baOutStream(uCtx.lOutPos - lDist), lSize)
                            uCtx.lOutPos = uCtx.lOutPos + lSize
                        End If
                    Else '--- 256
                        Exit Do
                    End If
                Loop
            #End If
        End If
        #If Not SkipCrcCheck Then
            Do While lCrc32Pos < uCtx.lOutPos
                lOutputCrc32 = (((lOutputCrc32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (m_lCrc32Table((lOutputCrc32 And &HFF) Xor uCtx.baOutStream(lCrc32Pos)))
                lCrc32Pos = lCrc32Pos + 1
            Loop
        #End If
        RaiseEvent Progress(m_lCurrentFile, uCtx.lInPos, lPackedSize)
    Loop While Not bIsLastBlock
    lOutputSize = uCtx.lOutPos
    lOutputCrc32 = lOutputCrc32 Xor &HFFFFFFFF
    #If ProfileCode Then
        Debug.Print m_uFileData(m_lCurrentFile).FileName, uCtx.lCountChar; "(" & Format(uCtx.lCountChar * 100# / (uCtx.lCountChar + uCtx.lCountCopy), "0.0"); ")", uCtx.lCountCopy; "("; Format(uCtx.lCountCopy * 100# / (uCtx.lCountChar + uCtx.lCountCopy), "0.0"); ")", uCtx.lCountDist, uCtx.lCountLoop, uCtx.lCountAPI, uCtx.lCountBits
    #End If
    '--- success
    pvInflate = True
QH:
    pvMapArray uCtx.baInStream, uCtx.uInSA
    pvMapArray uCtx.baOutStream, uCtx.uOutSA
    Exit Function
EH:
    pvMapArray uCtx.baInStream, uCtx.uInSA
    pvMapArray uCtx.baOutStream, uCtx.uOutSA
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume
End Function

'This sub is used to create a static huffmann tree for inflate
Private Function pvCreateStaticTree(uCtx As ZcUnpackContextType) As Boolean
    Dim lIdx            As Long
    Dim aLengths(287)   As Long
    
    On Error GoTo EH
    If Not uCtx.bStaticInit Then
        For lIdx = 0 To 287
            If lIdx <= 143 Then
                aLengths(lIdx) = 8
            ElseIf lIdx <= 255 Then
                aLengths(lIdx) = 9
            ElseIf lIdx <= 279 Then
                aLengths(lIdx) = 7
            Else
                aLengths(lIdx) = 8
            End If
        Next
        If Not pvCreateCodes(uCtx.uStaticLitLen, aLengths, 287, uCtx.lMaxLitLen, uCtx.lMinLitLen, "Static tree") Then
            GoTo QH
        End If
        For lIdx = 0 To 31
            aLengths(lIdx) = 5
        Next
        If Not pvCreateCodes(uCtx.uStaticDist, aLengths, 31, uCtx.lMaxDist, uCtx.lMinDist, "Static tree") Then
            GoTo QH
        End If
        uCtx.bStaticInit = True
    Else
        uCtx.lMinLitLen = 7
        uCtx.lMaxLitLen = 9
        uCtx.lMinDist = 5
        uCtx.lMaxDist = 5
    End If
    uCtx.uLitLen = uCtx.uStaticLitLen
    uCtx.uDist = uCtx.uStaticDist
    '--- success
    pvCreateStaticTree = True
QH:
    Exit Function
EH:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume
End Function

'This sub is used to create a dynamic tree for inflate
Private Function pvCreateDynamicTree(uCtx As ZcUnpackContextType) As Boolean
    Dim aLengths()      As Long
    Dim uBlTree         As ZcCodesType
    Dim lMinBL          As Long
    Dim lMaxBL          As Long
    Dim lNumLen         As Long
    Dim lNumDist        As Long
    Dim lNumCode        As Long
    Dim lChar           As Long
    Dim lBits           As Long
    Dim lLenNext        As Long
    Dim lPos            As Long
    Dim lIdx            As Long

    On Error GoTo EH
    lNumLen = pvGetBits(uCtx, 5) + 257    'Get length of the literal/length tree
    lNumDist = pvGetBits(uCtx, 5) + 1     'Get length of the distance tree
    lNumCode = pvGetBits(uCtx, 4) + 4     'Get number of codes for the tree to form the other trees
    ReDim aLengths(18) As Long
    'read the length per code
    For lIdx = 0 To lNumCode - 1
        aLengths(uCtx.aLenOrder(lIdx)) = pvGetBits(uCtx, 3)
    Next
    'codes not used get length 0
    For lIdx = lNumCode To 18
        aLengths(uCtx.aLenOrder(lIdx)) = 0
    Next
    'create the construction tree
    If Not pvCreateCodes(uBlTree, aLengths, 18, lMaxBL, lMinBL, "Dynamic tree") Then
        GoTo QH
    End If
    'Get the codes for the literal/length and distance trees
    ReDim aLengths(lNumLen + lNumDist) As Long
    lPos = 0
    Do While lPos < lNumLen + lNumDist
        pvNeedBits uCtx, lMaxBL
        lBits = lMinBL
        Do While uBlTree.Length(uCtx.lInBuff And uCtx.aMask(lBits)) <> lBits
            lBits = lBits + 1
        Loop
        lChar = uBlTree.Code(uCtx.lInBuff And uCtx.aMask(lBits))
        pvDropBits uCtx, lBits
        If lChar < 16 Then
            aLengths(lPos) = lChar
            lPos = lPos + 1
        Else
            If lChar = 16 Then
                If lPos = 0 Then
                    pvError "Dynamic tree: no last Length"
                    GoTo QH
                End If
                lLenNext = aLengths(lPos - 1)
                lChar = 3 + pvGetBits(uCtx, 2)
            ElseIf lChar = 17 Then
                lChar = 3 + pvGetBits(uCtx, 3)
                lLenNext = 0
            Else
                lChar = 11 + pvGetBits(uCtx, 7)
                lLenNext = 0
            End If
            If lPos + lChar > lNumLen + lNumDist Then
                pvError "Dynamic tree: too many Lengths"
                GoTo QH
            End If
            Do While lChar > 0
                lChar = lChar - 1
                aLengths(lPos) = lLenNext
                lPos = lPos + 1
            Loop
        End If
    Loop
    'create the literal/length tree
    If Not pvCreateCodes(uCtx.uLitLen, aLengths, lNumLen - 1, uCtx.lMaxLitLen, uCtx.lMinLitLen, "Dynamic tree") Then
        GoTo QH
    End If
    For lIdx = 0 To lNumDist
        aLengths(lIdx) = aLengths(lIdx + lNumLen)
    Next
    'create the distance tree
    If Not pvCreateCodes(uCtx.uDist, aLengths, lNumDist - 1, uCtx.lMaxDist, uCtx.lMinDist, "Dynamic tree") Then
        GoTo QH
    End If
    '--- success
    pvCreateDynamicTree = True
QH:
    Exit Function
EH:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume
End Function

'This function is used to retrieve the codes belonging to the huffman-trees
Private Function pvCreateCodes( _
            uTree As ZcCodesType, _
            aLengths() As Long, _
            ByVal lNumCodes As Long, _
            lMaxBits As Long, _
            lMinBits As Long, _
            sContext As String) As Boolean
    Dim aBits(16)       As Long
    Dim aNextCode(16)   As Long
    Dim lCode           As Long
    Dim lLenNext        As Long
    Dim lIdx            As Long

    On Error GoTo EH
    'retrieve the bitLength count and minimum and maximum bitLengths
    lMinBits = 16
    For lIdx = 0 To lNumCodes
        aBits(aLengths(lIdx)) = aBits(aLengths(lIdx)) + 1
        If aLengths(lIdx) > lMaxBits Then
            lMaxBits = aLengths(lIdx)
        End If
        If aLengths(lIdx) < lMinBits And aLengths(lIdx) > 0 Then
            lMinBits = aLengths(lIdx)
        End If
    Next
    lLenNext = 1
    For lIdx = 1 To lMaxBits
        lLenNext = lLenNext + lLenNext
        lLenNext = lLenNext - aBits(lIdx)
        If lLenNext < 0 Then
            pvError sContext & ": Huffman tree over subscribe"
            GoTo QH
        End If
    Next
    ReDim uTree.Code(2 ^ lMaxBits - 1) As Long 'set the right dimensions
    ReDim uTree.Length(2 ^ lMaxBits - 1) As Long
    lCode = 0
    aBits(0) = 0
    For lIdx = 1 To lMaxBits
        lCode = (lCode + aBits(lIdx - 1)) * 2
        aNextCode(lIdx) = lCode
    Next
    For lIdx = 0 To lNumCodes
        lLenNext = aLengths(lIdx)
        If lLenNext <> 0 Then
            lCode = pvBitReverse(aNextCode(lLenNext), lLenNext)
            uTree.Length(lCode) = lLenNext
            uTree.Code(lCode) = lIdx
            aNextCode(lLenNext) = aNextCode(lLenNext) + 1
        End If
    Next
    '--- success
    pvCreateCodes = True
QH:
    Exit Function
EH:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume
End Function

'Inflated codes are stored in reversed order so this funtion will reverse the stored order to get the original value back
Private Function pvBitReverse(ByVal lValue As Long, ByVal lNumBits As Long) As Long
    Do While lNumBits > 0
        pvBitReverse = pvBitReverse * 2 + (lValue And 1)
        lNumBits = lNumBits - 1
        lValue = lValue \ 2
    Loop
End Function

'This sub makes sure that there are at least the number of requested bits in lInBuff
Private Function pvNeedBits(uCtx As ZcUnpackContextType, ByVal lNumBits As Long) As Boolean
    Do While uCtx.lInBits < lNumBits
        If uCtx.lInPos <= UBound(uCtx.baInStream) Then
            uCtx.lInBuff = uCtx.lInBuff + (uCtx.baInStream(uCtx.lInPos) * uCtx.aPow2(uCtx.lInBits))
        End If
        uCtx.lInBits = uCtx.lInBits + 8
        uCtx.lInPos = uCtx.lInPos + 1
    Loop
End Function

'This sub will drop the amount of bits requested
Private Sub pvDropBits(uCtx As ZcUnpackContextType, ByVal lNumBits As Long)
    uCtx.lInBuff = uCtx.lInBuff \ uCtx.aPow2(lNumBits)
    uCtx.lInBits = uCtx.lInBits - lNumBits
End Sub

Private Function pvGetBits(uCtx As ZcUnpackContextType, ByVal lNumBits As Long) As Long
    Do While uCtx.lInBits < lNumBits
        If uCtx.lInPos <= UBound(uCtx.baInStream) Then
            uCtx.lInBuff = uCtx.lInBuff + (uCtx.baInStream(uCtx.lInPos) * uCtx.aPow2(uCtx.lInBits))
        End If
        uCtx.lInBits = uCtx.lInBits + 8
        uCtx.lInPos = uCtx.lInPos + 1
    Loop
    pvGetBits = uCtx.lInBuff And uCtx.aMask(lNumBits)
    uCtx.lInBuff = uCtx.lInBuff \ uCtx.aPow2(lNumBits)
    uCtx.lInBits = uCtx.lInBits - lNumBits
End Function

Private Sub pvMapArray(baBits() As Byte, uSA As SAFEARRAY, Optional ByVal lPtr As Long, Optional ByVal lSize As Long)
    If lPtr = 0 Then
        Call CopyMemory(ByVal ArrPtr(baBits()), 0&, 4)
    Else
        With uSA
            .cbElements = 1
            .cDims = 1
            .lLbound = 0
            .cElements = lSize
            .pvData = lPtr
        End With
        Call CopyMemory(ByVal ArrPtr(baBits()), VarPtr(uSA), 4)
    End If
End Sub

Private Sub pvSetArray(aDest() As Long, vTemp As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(vTemp)) As Long
    For lIdx = 0 To UBound(vTemp)
        aDest(lIdx) = vTemp(lIdx)
    Next
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lValue          As Long
    
    For lIdx = 0 To 255
        lValue = lIdx
        For lJdx = 0 To 7
            If (lValue And 1) <> 0 Then
              lValue = (((lValue And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
            Else
              lValue = ((lValue And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next
        m_lCrc32Table(lIdx) = lValue
    Next
End Sub

