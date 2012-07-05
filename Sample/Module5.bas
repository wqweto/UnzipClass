Attribute VB_Name = "Module5"
'=========================================================================
' $Header: $
'
'   VB6 Unzip Support
'   Copyright (c) 2012 Unicontsoft
'
'   Projectwise public ZcFlagsEnum and ZcFileInfoType
'
' $Log: $
'
'=========================================================================
Option Explicit

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
