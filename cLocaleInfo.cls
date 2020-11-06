VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLocaleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' CLocaleInfo
' Clase para obtener la información del idioma actual del sistema
'
Option Explicit

' LCID del idioma actual
Private LCID As Long
' Longitud del buffer para obtener la información
Private Const MAX_LOCALE_BUF As Long = 256&

' formato de la fecha
Public Enum eDateOrder
    MDY = 0 ' Month-Day-Year
    DMY = 1 ' Day-Month-Year
    YMD = 2 ' Year-Month-Day
End Enum

'------------------------------------------------------------------------------
' Funciones del API para la información local
'------------------------------------------------------------------------------
'Private Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" _
    (ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
    (ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" _
    (ByVal Locale As Long, ByVal dwFlags As Long, _
    lpTime As SYSTEMTIME, ByVal lpFormat As String, _
    ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" _
    (ByVal Locale As Long, ByVal dwFlags As Long, _
    lpDate As SYSTEMTIME, ByVal lpFormat As String, _
    ByVal lpDateStr As String, ByVal cchDate As Long) As Long

'------------------------------------------------------------------------------
' Constantes
'------------------------------------------------------------------------------
' Locale Types.
' These types are used for the GetLocaleInfo NLS API routine.
'
Public Enum eLocaleInfo
    LOCALE_ILANGUAGE = &H1                  '  language id
    LOCALE_SLANGUAGE = &H2                  '  localized name of language
    LOCALE_SENGLANGUAGE = &H1001            '  English name of language
    LOCALE_SABBREVLANGNAME = &H3            '  abbreviated language name
    LOCALE_SNATIVELANGNAME = &H4            '  native name of language
    LOCALE_ICOUNTRY = &H5                   '  country code
    LOCALE_SCOUNTRY = &H6                   '  localized name of country
    LOCALE_SENGCOUNTRY = &H1002             '  English name of country
    LOCALE_SABBREVCTRYNAME = &H7            '  abbreviated country name
    LOCALE_SNATIVECTRYNAME = &H8            '  native name of country
    LOCALE_IDEFAULTLANGUAGE = &H9           '  default language id
    LOCALE_IDEFAULTCOUNTRY = &HA            '  default country code
    LOCALE_IDEFAULTCODEPAGE = &HB           '  default code page
    '
    LOCALE_SLIST = &HC                      '  list item separator
    LOCALE_IMEASURE = &HD                   '  0 = metric, 1 = US
    '
    LOCALE_SDECIMAL = &HE                   '  decimal separator
    LOCALE_STHOUSAND = &HF                  '  thousand separator
    LOCALE_SGROUPING = &H10                 '  digit grouping
    LOCALE_IDIGITS = &H11                   '  number of fractional digits
    LOCALE_ILZERO = &H12                    '  leading zeros for decimal
    LOCALE_SNATIVEDIGITS = &H13             '  native ascii 0-9
    '
    LOCALE_SCURRENCY = &H14                 '  local monetary symbol
    LOCALE_SINTLSYMBOL = &H15               '  intl monetary symbol
    LOCALE_SMONDECIMALSEP = &H16            '  monetary decimal separator
    LOCALE_SMONTHOUSANDSEP = &H17           '  monetary thousand separator
    LOCALE_SMONGROUPING = &H18              '  monetary grouping
    LOCALE_ICURRDIGITS = &H19               '  # local monetary digits
    LOCALE_IINTLCURRDIGITS = &H1A           '  # intl monetary digits
    LOCALE_ICURRENCY = &H1B                 '  positive currency mode
    LOCALE_INEGCURR = &H1C                  '  negative currency mode
    '
    LOCALE_SDATE = &H1D                     '  date separator
    LOCALE_STIME = &H1E                     '  time separator
    LOCALE_SSHORTDATE = &H1F                '  short date format string
    LOCALE_SLONGDATE = &H20                 '  long date format string
    LOCALE_STIMEFORMAT = &H1003             '  time format string
    LOCALE_IDATE = &H21                     '  short date format ordering
    LOCALE_ILDATE = &H22                    '  long date format ordering
    LOCALE_ITIME = &H23                     '  time format specifier
    LOCALE_ICENTURY = &H24                  '  century format specifier
    LOCALE_ITLZERO = &H25                   '  leading zeros in time field
    LOCALE_IDAYLZERO = &H26                 '  leading zeros in day field
    LOCALE_IMONLZERO = &H27                 '  leading zeros in month field
    LOCALE_S1159 = &H28                     '  AM designator
    LOCALE_S2359 = &H29                     '  PM designator
    '
    LOCALE_SDAYNAME1 = &H2A                 '  long name for Monday
    LOCALE_SDAYNAME2 = &H2B                 '  long name for Tuesday
    LOCALE_SDAYNAME3 = &H2C                 '  long name for Wednesday
    LOCALE_SDAYNAME4 = &H2D                 '  long name for Thursday
    LOCALE_SDAYNAME5 = &H2E                 '  long name for Friday
    LOCALE_SDAYNAME6 = &H2F                 '  long name for Saturday
    LOCALE_SDAYNAME7 = &H30                 '  long name for Sunday
    LOCALE_SABBREVDAYNAME1 = &H31           '  abbreviated name for Monday
    LOCALE_SABBREVDAYNAME2 = &H32           '  abbreviated name for Tuesday
    LOCALE_SABBREVDAYNAME3 = &H33           '  abbreviated name for Wednesday
    LOCALE_SABBREVDAYNAME4 = &H34           '  abbreviated name for Thursday
    LOCALE_SABBREVDAYNAME5 = &H35           '  abbreviated name for Friday
    LOCALE_SABBREVDAYNAME6 = &H36           '  abbreviated name for Saturday
    LOCALE_SABBREVDAYNAME7 = &H37           '  abbreviated name for Sunday
    LOCALE_SMONTHNAME1 = &H38               '  long name for January
    LOCALE_SMONTHNAME2 = &H39               '  long name for February
    LOCALE_SMONTHNAME3 = &H3A               '  long name for March
    LOCALE_SMONTHNAME4 = &H3B               '  long name for April
    LOCALE_SMONTHNAME5 = &H3C               '  long name for May
    LOCALE_SMONTHNAME6 = &H3D               '  long name for June
    LOCALE_SMONTHNAME7 = &H3E               '  long name for July
    LOCALE_SMONTHNAME8 = &H3F               '  long name for August
    LOCALE_SMONTHNAME9 = &H40               '  long name for September
    LOCALE_SMONTHNAME10 = &H41              '  long name for October
    LOCALE_SMONTHNAME11 = &H42              '  long name for November
    LOCALE_SMONTHNAME12 = &H43              '  long name for December
    LOCALE_SABBREVMONTHNAME1 = &H44         '  abbreviated name for January
    LOCALE_SABBREVMONTHNAME2 = &H45         '  abbreviated name for February
    LOCALE_SABBREVMONTHNAME3 = &H46         '  abbreviated name for March
    LOCALE_SABBREVMONTHNAME4 = &H47         '  abbreviated name for April
    LOCALE_SABBREVMONTHNAME5 = &H48         '  abbreviated name for May
    LOCALE_SABBREVMONTHNAME6 = &H49         '  abbreviated name for June
    LOCALE_SABBREVMONTHNAME7 = &H4A         '  abbreviated name for July
    LOCALE_SABBREVMONTHNAME8 = &H4B         '  abbreviated name for August
    LOCALE_SABBREVMONTHNAME9 = &H4C         '  abbreviated name for September
    LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
    LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
    LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
    LOCALE_SABBREVMONTHNAME13 = &H100F
    '
    LOCALE_SPOSITIVESIGN = &H50             '  positive sign
    LOCALE_SNEGATIVESIGN = &H51             '  negative sign
    LOCALE_IPOSSIGNPOSN = &H52              '  positive sign position
    LOCALE_INEGSIGNPOSN = &H53              '  negative sign position
    LOCALE_IPOSSYMPRECEDES = &H54           '  mon sym precedes pos amt
    LOCALE_IPOSSEPBYSPACE = &H55            '  mon sym sep by space from pos amt
    LOCALE_INEGSYMPRECEDES = &H56           '  mon sym precedes neg amt
    LOCALE_INEGSEPBYSPACE = &H57            '  mon sym sep by space from neg amt
End Enum
' Time Flags for GetTimeFormatW.
Private Const TIME_NOMINUTESORSECONDS = &H1 '  do not use minutes or seconds
Private Const TIME_NOSECONDS = &H2          '  do not use seconds
Private Const TIME_NOTIMEMARKER = &H4       '  do not use time marker
Private Const TIME_FORCE24HOURFORMAT = &H8  '  always use 24 hour format

' Date Flags for GetDateFormatW.
Private Const DATE_SHORTDATE = &H1          '  use short date picture
Private Const DATE_LONGDATE = &H2           '  use long date picture

Public Function NativeLanguage() As String
    ' Devuelve el idioma actual, en el propio idioma
    NativeLanguage = LocaleInfo(LOCALE_SNATIVELANGNAME)
End Function
Public Function Language() As String
    ' Devuelve el idioma actual
    Language = LocaleInfo(LOCALE_SLANGUAGE)
End Function

Public Function Thousands() As String
    ' Devuelve el signo de los miles
    Thousands = LocaleInfo(LOCALE_STHOUSAND)
End Function

Public Function sDecimal() As String
    ' Devuelve el signo decimal
    sDecimal = Left$(LocaleInfo(LOCALE_SDECIMAL), 1)
End Function

Private Function TrimNull(ByVal sCadena As String) As String
    ' Devuelve la cadena hasta el primer null
    ' También quita los espacios extras
    Dim i As Long
    
    i = InStr(sCadena, Chr$(0))
    If i Then
        TrimNull = Left$(sCadena, i - 1)
    End If
    
    TrimNull = Trim$(sCadena)
End Function

Private Sub Class_Initialize()
    ' Al iniciar la clase, obtener el ID del idioma usado
    GetLCID
End Sub

Public Function GetLCID() As Long
    ' Obtener el ID del idioma usado
    ' Lo he puesto en un método público por si se quiere refrescar la información
    LCID = GetSystemDefaultLangID
    
    GetLCID = LCID
End Function

Public Function LocaleInfo(Optional ByVal nLocaleInfo As eLocaleInfo = LOCALE_SENGLANGUAGE) As String
    ' Devuelve la información indicada en el parámetro
    ' Si no se especifica el parámetro,
    ' será el nombre del idioma según la norma ISO (en inglés)
    Dim sBuf As String
    
    sBuf = Space$(MAX_LOCALE_BUF)
    Call GetLocaleInfo(LCID, nLocaleInfo, sBuf, Len(sBuf))
    LocaleInfo = TrimNull(sBuf)
End Function

Public Function ShortDateFormat() As String
    ' Devuelve el formato de la fecha corta
    ShortDateFormat = LocaleInfo(LOCALE_SSHORTDATE)
End Function

Public Function LongDateFormat() As String
    ' Devuelve el formato de la fecha larga
    LongDateFormat = LocaleInfo(LOCALE_SLONGDATE)
End Function

Public Function NativeCountryName() As String
    ' Devuelve el nombre del pais (en el idioma propio)
    NativeCountryName = LocaleInfo(LOCALE_SNATIVECTRYNAME)
End Function
Public Function CountryName() As String
    ' Devuelve el nombre del pais (en inglés)
    CountryName = LocaleInfo(LOCALE_SCOUNTRY)
End Function

Public Function LongDateFormatOrder() As String
    ' El orden de la fecha larga (MDY, DMY, YMD)
    Dim i As eDateOrder
    
    i = LocaleInfo(LOCALE_ILDATE)
    
    If i = DMY Then
        LongDateFormatOrder = "DMY"
    ElseIf i = MDY Then
        LongDateFormatOrder = "MDY"
    Else
        LongDateFormatOrder = "YMD"
    End If
End Function

Public Function DateFormatOrder() As String
    ' El orden de la fecha corta (MDY, DMY, YMD)
    Dim i As eDateOrder
    
    i = LocaleInfo(LOCALE_IDATE)
    
    If i = DMY Then
        DateFormatOrder = "DMY"
    ElseIf i = MDY Then
        DateFormatOrder = "MDY"
    Else
        DateFormatOrder = "YMD"
    End If
End Function





