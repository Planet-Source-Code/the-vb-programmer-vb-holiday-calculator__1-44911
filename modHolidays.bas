Attribute VB_Name = "modHolidays"
Option Explicit

' default day- and weeknumstyle
Public Const ISO_8601 = 1 '
'Calendar types
Public Const Gregorian = ISO_8601
'Miscellaneous public constants
Public Const Signed = -1
Public Const UnSigned = 1

'---------------------------------------------------------------------------
Public Function GetHolidayName(pdtmTestDate As Date) As String
'---------------------------------------------------------------------------

    Dim strHolidayName  As String
    
    strHolidayName = ""
    
    If IsNewYearsDay(pdtmTestDate) Then
        strHolidayName = "New Year's Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsMLKDay(pdtmTestDate) Then
        strHolidayName = "M. L. King Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsValentinesDay(pdtmTestDate) Then
        strHolidayName = "Valentine's Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsPresidentsDay(pdtmTestDate) Then
        strHolidayName = "President's Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsStPatsDay(pdtmTestDate) Then
        strHolidayName = "St. Patrick's Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsGoodFriday(pdtmTestDate) Then
        strHolidayName = "Good Friday"
        GoTo CheckJewishHoliday
    End If
    
    If IsEaster(pdtmTestDate) Then
        strHolidayName = "Easter"
        GoTo CheckJewishHoliday
    End If
    
    If IsMothersDay(pdtmTestDate) Then
        strHolidayName = "Mother's Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsMemorialDay(pdtmTestDate) Then
        strHolidayName = "Memorial Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsFlagDay(pdtmTestDate) Then
        strHolidayName = "Flag Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsFathersDay(pdtmTestDate) Then
        strHolidayName = "Father's Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsIndependenceDay(pdtmTestDate) Then
        strHolidayName = "Independence Day"
        GoTo CheckJewishHoliday
    End If

    If IsLaborDay(pdtmTestDate) Then
        strHolidayName = "Labor Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsColumbusDay(pdtmTestDate) Then
        strHolidayName = "Columbus Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsHalloween(pdtmTestDate) Then
        strHolidayName = "Halloween"
        GoTo CheckJewishHoliday
    End If
    
    If IsElectionDay(pdtmTestDate) Then
        strHolidayName = "Election Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsVeteransDay(pdtmTestDate) Then
        strHolidayName = "Veteran's Day"
        GoTo CheckJewishHoliday
    End If
    
    If IsThanksgiving(pdtmTestDate) Then
        strHolidayName = "Thanksgiving"
        GoTo CheckJewishHoliday
    End If
    
    If IsChristmas(pdtmTestDate) Then
        strHolidayName = "Christmas"
        GoTo CheckJewishHoliday
    End If

CheckJewishHoliday:
    If IsRoshHashanah(pdtmTestDate) Then
        If strHolidayName <> "" Then strHolidayName = strHolidayName & ",  "
        strHolidayName = strHolidayName & "Rosh Hashanah"
        GoTo SetReturnValue
    End If

    If IsYomKippur(pdtmTestDate) Then
        If strHolidayName <> "" Then strHolidayName = strHolidayName & ",  "
        strHolidayName = strHolidayName & "Yom Kippur"
        GoTo SetReturnValue
    End If
    
    If IsSukkot(pdtmTestDate) Then
        If strHolidayName <> "" Then strHolidayName = strHolidayName & ",  "
        strHolidayName = strHolidayName & "Sukkot"
        GoTo SetReturnValue
    End If

    If IsHannukah(pdtmTestDate) Then
        If strHolidayName <> "" Then strHolidayName = strHolidayName & ",  "
        strHolidayName = strHolidayName & "Hannukah"
        GoTo SetReturnValue
    End If
    
    If IsPurim(pdtmTestDate) Then
        If strHolidayName <> "" Then strHolidayName = strHolidayName & ",  "
        strHolidayName = strHolidayName & "Purim"
        GoTo SetReturnValue
    End If

    If IsPassover(pdtmTestDate) Then
        If strHolidayName <> "" Then strHolidayName = strHolidayName & ",  "
        strHolidayName = strHolidayName & "Passover"
        GoTo SetReturnValue
    End If
    
    If IsShavuot(pdtmTestDate) Then
        If strHolidayName <> "" Then strHolidayName = strHolidayName & " , "
        strHolidayName = strHolidayName & "Shavuot"
        GoTo SetReturnValue
    End If
    
SetReturnValue:
    GetHolidayName = strHolidayName
    
End Function

'---------------------------------------------------------------------------
Private Function IsNewYearsDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    ' New Year's Day - Jan 1
    
    If Month(pdtmTestDate) = 1 Then
        If Day(pdtmTestDate) = 1 Then
            IsNewYearsDay = True
        Else
            IsNewYearsDay = False
        End If
    Else
        IsNewYearsDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsMLKDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'MLK day = 3rd Monday in January
        
    Dim intMonCount As Integer
    Dim dtmMLK      As Date
    
    If Month(pdtmTestDate) <> 1 Then
        IsMLKDay = False
        Exit Function
    End If
        
    If Weekday(pdtmTestDate) <> vbMonday Then
        IsMLKDay = False
        Exit Function
    End If
        
    dtmMLK = DateSerial(Year(pdtmTestDate), 1, 1)
    intMonCount = 0
    
    Do
        If Weekday(dtmMLK) = vbMonday Then
            intMonCount = intMonCount + 1
            If intMonCount = 3 Then Exit Do
        End If
        dtmMLK = DateAdd("d", 1, dtmMLK)
    Loop

    If dtmMLK = pdtmTestDate Then
        IsMLKDay = True
    Else
        IsMLKDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsValentinesDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    ' Valentine's day - Feb 14
    
    If Month(pdtmTestDate) = 2 Then
        If Day(pdtmTestDate) = 14 Then
            IsValentinesDay = True
        Else
            IsValentinesDay = False
        End If
    Else
        IsValentinesDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsPresidentsDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'Pres day = 3rd Monday in February
        
    Dim intMonCount As Integer
    Dim dtmPres      As Date
    
    If Month(pdtmTestDate) <> 2 Then
        IsPresidentsDay = False
        Exit Function
    End If
    
    If Weekday(pdtmTestDate) <> vbMonday Then
        IsPresidentsDay = False
        Exit Function
    End If
    
    dtmPres = DateSerial(Year(pdtmTestDate), 2, 1)
    intMonCount = 0
    
    Do
        If Weekday(dtmPres) = vbMonday Then
            intMonCount = intMonCount + 1
            If intMonCount = 3 Then Exit Do
        End If
        dtmPres = DateAdd("d", 1, dtmPres)
    Loop

    If dtmPres = pdtmTestDate Then
        IsPresidentsDay = True
    Else
        IsPresidentsDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsStPatsDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' St. Patricks's day - Mar 17
    
    If Month(pdtmTestDate) = 3 Then
        If Day(pdtmTestDate) = 17 Then
            IsStPatsDay = True
        Else
            IsStPatsDay = False
        End If
    Else
        IsStPatsDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsGoodFriday(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Good Friday - the Friday before Easter Sunday
    
    Dim dtmEasterDate   As Date
    
    If Weekday(pdtmTestDate) <> vbFriday Then
        IsGoodFriday = False
        Exit Function
    End If
    
    dtmEasterDate = GetEasterDate(Year(pdtmTestDate))
    
    If dtmEasterDate = DateAdd("d", 2, pdtmTestDate) Then
        IsGoodFriday = True
    Else
        IsGoodFriday = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsEaster(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Easter - varies (anywhere from Mar 22 to Apr 25)
    
    Dim dtmEasterDate   As Date
    
    dtmEasterDate = GetEasterDate(Year(pdtmTestDate))
    
    If dtmEasterDate = pdtmTestDate Then
        IsEaster = True
    Else
        IsEaster = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsMothersDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'Mother's day = 2nd Sunday in May
        
    Dim intSunCount As Integer
    Dim dtmMom      As Date
    
    If Month(pdtmTestDate) <> 5 Then
        IsMothersDay = False
        Exit Function
    End If
    
    If Weekday(pdtmTestDate) <> vbSunday Then
        IsMothersDay = False
        Exit Function
    End If
    
    dtmMom = DateSerial(Year(pdtmTestDate), 5, 1)
    intSunCount = 0
    
    Do
        If Weekday(dtmMom) = vbSunday Then
            intSunCount = intSunCount + 1
            If intSunCount = 2 Then Exit Do
        End If
        dtmMom = DateAdd("d", 1, dtmMom)
    Loop

    If dtmMom = pdtmTestDate Then
        IsMothersDay = True
    Else
        IsMothersDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsMemorialDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'Mem day = last Monday in May
        
    Dim dtmMem      As Date
    
    If Month(pdtmTestDate) <> 5 Then
        IsMemorialDay = False
        Exit Function
    End If
    
    If Weekday(pdtmTestDate) <> vbMonday Then
        IsMemorialDay = False
        Exit Function
    End If
    
    dtmMem = DateSerial(Year(pdtmTestDate), 5, 31)
    
    Do
        If Weekday(dtmMem) = vbMonday Then Exit Do
        dtmMem = DateAdd("d", -1, dtmMem)
    Loop

    If dtmMem = pdtmTestDate Then
        IsMemorialDay = True
    Else
        IsMemorialDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsFlagDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Flag Day - June 14
    
    If Month(pdtmTestDate) = 6 Then
        If Day(pdtmTestDate) = 14 Then
            IsFlagDay = True
        Else
            IsFlagDay = False
        End If
    Else
        IsFlagDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsFathersDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'Father's day = 3rd Sunday in June
        
    Dim intSunCount As Integer
    Dim dtmDad      As Date
    
    If Month(pdtmTestDate) <> 6 Then
        IsFathersDay = False
        Exit Function
    End If
    
    If Weekday(pdtmTestDate) <> vbSunday Then
        IsFathersDay = False
        Exit Function
    End If
    
    dtmDad = DateSerial(Year(pdtmTestDate), 6, 1)
    intSunCount = 0
    
    Do
        If Weekday(dtmDad) = vbSunday Then
            intSunCount = intSunCount + 1
            If intSunCount = 3 Then Exit Do
        End If
        dtmDad = DateAdd("d", 1, dtmDad)
    Loop

    If dtmDad = pdtmTestDate Then
        IsFathersDay = True
    Else
        IsFathersDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsIndependenceDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Independence day - July 4
    
    If Month(pdtmTestDate) = 7 Then
        If Day(pdtmTestDate) = 4 Then
            IsIndependenceDay = True
        Else
            IsIndependenceDay = False
        End If
    Else
        IsIndependenceDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsLaborDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'Labor day = 1st Monday in Sept.
        
    Dim dtmLab      As Date
    
    If Month(pdtmTestDate) <> 9 Then
        IsLaborDay = False
        Exit Function
    End If
    
    If Weekday(pdtmTestDate) <> vbMonday Then
        IsLaborDay = False
        Exit Function
    End If

    dtmLab = DateSerial(Year(pdtmTestDate), 9, 1)
    
    Do
        If Weekday(dtmLab) = vbMonday Then Exit Do
        dtmLab = DateAdd("d", 1, dtmLab)
    Loop

    If dtmLab = pdtmTestDate Then
        IsLaborDay = True
    Else
        IsLaborDay = False
    End If

End Function


'---------------------------------------------------------------------------
Private Function IsColumbusDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'Columbus Day - 2nd Monday in October
        
    Dim intMonCount  As Integer
    Dim dtmColum     As Date
    
    If Month(pdtmTestDate) <> 10 Then
        IsColumbusDay = False
        Exit Function
    End If
    
    If Weekday(pdtmTestDate) <> vbMonday Then
        IsColumbusDay = False
        Exit Function
    End If
    
    dtmColum = DateSerial(Year(pdtmTestDate), 10, 1)
    intMonCount = 0
    
    Do
        If Weekday(dtmColum) = vbMonday Then
            intMonCount = intMonCount + 1
            If intMonCount = 2 Then Exit Do
        End If
        dtmColum = DateAdd("d", 1, dtmColum)
    Loop

    If dtmColum = pdtmTestDate Then
        IsColumbusDay = True
    Else
        IsColumbusDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsHalloween(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Halloween - Oct 31
    
    If Month(pdtmTestDate) = 10 Then
        If Day(pdtmTestDate) = 31 Then
            IsHalloween = True
        Else
            IsHalloween = False
        End If
    Else
        IsHalloween = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsElectionDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'Election day = 1st Tuesdat in Nov.
        
    Dim dtmElectionDay  As Date
    
    If Month(pdtmTestDate) <> 11 Then
        IsElectionDay = False
        Exit Function
    End If
    
    If Weekday(pdtmTestDate) <> vbTuesday Then
        IsElectionDay = False
        Exit Function
    End If

    dtmElectionDay = DateSerial(Year(pdtmTestDate), 11, 1)
    
    Do
        If Weekday(dtmElectionDay) = vbTuesday Then Exit Do
        dtmElectionDay = DateAdd("d", 1, dtmElectionDay)
    Loop

    If dtmElectionDay = pdtmTestDate Then
        IsElectionDay = True
    Else
        IsElectionDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsVeteransDay(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Veteran's Day - Nov 11
    
    If Month(pdtmTestDate) = 11 Then
        If Day(pdtmTestDate) = 11 Then
            IsVeteransDay = True
        Else
            IsVeteransDay = False
        End If
    Else
        IsVeteransDay = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsThanksgiving(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------
    
    'Thanksgiving = 4th Thurs in Nov
        
    Dim intThursCount As Integer
    Dim dtmThanks     As Date
    
    If Month(pdtmTestDate) <> 11 Then
        IsThanksgiving = False
        Exit Function
    End If
    
    If Weekday(pdtmTestDate) <> vbThursday Then
        IsThanksgiving = False
        Exit Function
    End If
    
    dtmThanks = DateSerial(Year(pdtmTestDate), 11, 1)
    intThursCount = 0
    
    Do
        If Weekday(dtmThanks) = vbThursday Then
            intThursCount = intThursCount + 1
            If intThursCount = 4 Then Exit Do
        End If
        dtmThanks = DateAdd("d", 1, dtmThanks)
    Loop

    If dtmThanks = pdtmTestDate Then
        IsThanksgiving = True
    Else
        IsThanksgiving = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsChristmas(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Christmas - Dec 25
    
    If Month(pdtmTestDate) = 12 Then
        If Day(pdtmTestDate) = 25 Then
            IsChristmas = True
        Else
            IsChristmas = False
        End If
    Else
        IsChristmas = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsRoshHashanah(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Rosh Hashanah - Tishri 1
    
    Dim lngDateNbr          As Long
    Dim dtmRoshHashanah     As Date
    
    lngDateNbr = HebrewDateInOrAfterCivilYear(Year(pdtmTestDate), 1, 1)
    dtmRoshHashanah = jdn_Gregorian(lngDateNbr)
    If dtmRoshHashanah = pdtmTestDate Then
        IsRoshHashanah = True
    Else
        IsRoshHashanah = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsYomKippur(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Yom Kippur - Tishri 10
    
    Dim lngDateNbr       As Long
    Dim dtmYomKippur     As Date
    
    lngDateNbr = HebrewDateInOrAfterCivilYear(Year(pdtmTestDate), 1, 10)
    dtmYomKippur = jdn_Gregorian(lngDateNbr)
    If dtmYomKippur = pdtmTestDate Then
        IsYomKippur = True
    Else
        IsYomKippur = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsSukkot(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Sukkot - Tishri 15
    
    Dim lngDateNbr    As Long
    Dim dtmSukkot     As Date
    
    lngDateNbr = HebrewDateInOrAfterCivilYear(Year(pdtmTestDate), 1, 15)
    dtmSukkot = jdn_Gregorian(lngDateNbr)
    If dtmSukkot = pdtmTestDate Then
        IsSukkot = True
    Else
        IsSukkot = False
    End If

End Function


'---------------------------------------------------------------------------
Private Function IsHannukah(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Hannukah - Kislev 25
    
    Dim lngDateNbr      As Long
    Dim dtmHannukah     As Date
    
    lngDateNbr = HebrewDateInOrAfterCivilYear(Year(pdtmTestDate), 3, 25)
    dtmHannukah = jdn_Gregorian(lngDateNbr)
    If dtmHannukah = pdtmTestDate Then
        IsHannukah = True
    Else
        IsHannukah = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsPurim(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Purim - Adar 14 (Adar II 14 in leap years)
    
    Dim lngDateNbr   As Long
    Dim dtmPurim     As Date
    
    lngDateNbr = HebrewDateInOrAfterCivilYear(Year(pdtmTestDate), -7, 14)
    dtmPurim = jdn_Gregorian(lngDateNbr)
    If dtmPurim = pdtmTestDate Then
        IsPurim = True
    Else
        IsPurim = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function IsPassover(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Passover - Nisan 15
    
    Dim lngDateNbr   As Long
    Dim dtmPassover     As Date
    
    lngDateNbr = HebrewDateInOrAfterCivilYear(Year(pdtmTestDate), -6, 15)
    dtmPassover = jdn_Gregorian(lngDateNbr)
    If dtmPassover = pdtmTestDate Then
        IsPassover = True
    Else
        IsPassover = False
    End If

End Function


'---------------------------------------------------------------------------
Private Function IsShavuot(pdtmTestDate As Date) As Boolean
'---------------------------------------------------------------------------

    ' Shavuot - Nisan 15
    
    Dim lngDateNbr   As Long
    Dim dtmShavuot     As Date
    
    lngDateNbr = HebrewDateInOrAfterCivilYear(Year(pdtmTestDate), -4, 6)
    dtmShavuot = jdn_Gregorian(lngDateNbr)
    If dtmShavuot = pdtmTestDate Then
        IsShavuot = True
    Else
        IsShavuot = False
    End If

End Function

'---------------------------------------------------------------------------
Private Function GetEasterDate(pintYear As Integer) As Date
'---------------------------------------------------------------------------

    Dim C           As Integer
    Dim n           As Integer
    Dim K           As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim l           As Integer
    Dim intMonth    As Integer
    Dim intDay      As Integer
    
    C = pintYear \ 100
    n = pintYear - 19 * (pintYear \ 19)
    K = (C - 17) \ 25
    i = C - C \ 4 - (C - K) \ 3 + 19 * n + 15
    i = i - 30 * (i \ 30)
    i = i - (i \ 28) * (1 - (i \ 28) * (29 \ (i + 1)) * ((21 - n) \ 11))
    j = pintYear + pintYear \ 4 + i + 2 - C + C \ 4
    j = j - 7 * (j \ 7)
    l = i - j
    intMonth = 3 + (l + 40) \ 44
    intDay = l + 28 - 31 * (intMonth \ 4)
    
    GetEasterDate = DateSerial(pintYear, intMonth, intDay)

End Function

'---------------------------------------------------------------------------
Sub jdn_hebrew(jdn As Long, _
               ByRef iYear As Integer, _
               ByRef iMonth As Integer, _
               ByRef iDay As Integer, _
               Optional monthcoding As Integer = UnSigned)
'---------------------------------------------------------------------------
               
    Dim InputJDN        As Long
    Dim tishri1         As Long
    Dim LeftOverDays    As Long
    
    If jdn <= 347997 Then
        iYear = 0
        iMonth = 0
        iDay = 0
    Else
        InputJDN = jdn - 347997
        iYear = (InputJDN \ 365) + 1
        tishri1 = Hebrew_ElapsedCalendarDays(iYear)
        While (tishri1 > InputJDN)
            iYear = iYear - 1
            tishri1 = Hebrew_ElapsedCalendarDays(iYear)
        Wend
        iMonth = 1
        LeftOverDays = InputJDN - tishri1
        While (LeftOverDays >= Hebrew_LastDayOfMonth(iYear, iMonth))
            LeftOverDays = LeftOverDays - Hebrew_LastDayOfMonth(iYear, iMonth)
            iMonth = iMonth + 1
        Wend
        If Sgn(monthcoding) = Signed Then
            If iMonth > 6 Then
                If Hebrew_LeapYear(iYear) Then
                    iMonth = iMonth - 14
                Else
                    iMonth = iMonth - 13
                End If
            End If
        End If
        iDay = LeftOverDays + 1
    End If
End Sub

'---------------------------------------------------------------------------
Function Hebrew_ShortKislev(iYear)
'---------------------------------------------------------------------------
    Hebrew_ShortKislev = ((Hebrew_DaysInYear(iYear) Mod 10) = 3)
End Function

'---------------------------------------------------------------------------
Function Hebrew_LongHeshvan(iYear) As Boolean
'---------------------------------------------------------------------------
    Hebrew_LongHeshvan = ((Hebrew_DaysInYear(iYear) Mod 10) = 5)
End Function

'---------------------------------------------------------------------------
Function Hebrew_LeapYear(iYear) As Boolean
'---------------------------------------------------------------------------
    If ((((7 * iYear) + 1) Mod 19) < 7) Then
       Hebrew_LeapYear = True
    Else
       Hebrew_LeapYear = False
    End If
End Function

'---------------------------------------------------------------------------
Function Hebrew_LastDayOfMonth(iYear, ByVal iMonth) As Integer
'---------------------------------------------------------------------------
    
    If ((iMonth > 6) And (Not (Hebrew_LeapYear(iYear)))) Then
       iMonth = iMonth + 1
    End If
    Select Case iMonth
    Case 2
        If Hebrew_LongHeshvan(iYear) Then
            Hebrew_LastDayOfMonth = 30
        Else
            Hebrew_LastDayOfMonth = 29
        End If
    Case 3
        If Hebrew_ShortKislev(iYear) Then
            Hebrew_LastDayOfMonth = 29
        Else
            Hebrew_LastDayOfMonth = 30
        End If
    Case 6
        If Hebrew_LeapYear(iYear) Then
            Hebrew_LastDayOfMonth = 30
        Else
            Hebrew_LastDayOfMonth = 29
        End If
    Case 4, 7, 9, 11, 13
        Hebrew_LastDayOfMonth = 29
    Case Else
        Hebrew_LastDayOfMonth = 30
    End Select
End Function

'---------------------------------------------------------------------------
Function hebrew_jdn(iYear, ByVal iMonth, iDay) As Long
'---------------------------------------------------------------------------
    
    Dim jdn     As Long
    Dim counter As Integer
    
    If iMonth < 0 Then
        If Hebrew_LeapYear(iYear) Then
            iMonth = 14 + iMonth
        Else
            iMonth = 13 + iMonth
        End If
    End If
    jdn = Hebrew_ElapsedCalendarDays(iYear)
    For counter = 1 To (iMonth - 1) Step 1
        jdn = jdn + Hebrew_LastDayOfMonth(iYear, counter)
    Next counter
    hebrew_jdn = jdn + (iDay - 1 + 347997)

End Function

'---------------------------------------------------------------------------
Function Hebrew_ElapsedCalendarDays(iYear) As Long
'---------------------------------------------------------------------------
    
    Dim MonthsElapsed As Long
    Dim PartsElapsed As Long
    Dim HoursElapsed As Long
    Dim ConjunctionDay As Long
    Dim ConjunctionParts As Long
    Dim AlternativeDay As Long

    MonthsElapsed = (235 * (((iYear - 1) \ 19))) + _
                    (12 * ((iYear - 1) Mod 19)) + _
                    (7 * ((iYear - 1) Mod 19) + 1) \ 19
    PartsElapsed = 204 + 793 * (MonthsElapsed Mod 1080)
    HoursElapsed = 5 + 12 * MonthsElapsed + _
                   793 * ((MonthsElapsed \ 1080)) + _
                   PartsElapsed \ 1080
    ConjunctionDay = 1 + 29 * MonthsElapsed + HoursElapsed \ 24
    ConjunctionParts = (1080 * (HoursElapsed Mod 24)) + _
                       PartsElapsed Mod 1080
    If ((ConjunctionParts >= 19440) Or _
        (((ConjunctionDay Mod 7) = 2) And _
        (ConjunctionParts >= 9924) And _
        (Not (Hebrew_LeapYear(iYear)))) Or _
       (((ConjunctionDay Mod 7) = 1) And _
        (ConjunctionParts >= 16789) And _
        (Hebrew_LeapYear(iYear - 1)))) _
    Then
        AlternativeDay = ConjunctionDay + 1
    Else
        AlternativeDay = ConjunctionDay
    End If
    If (((AlternativeDay Mod 7) = 0) Or _
        ((AlternativeDay Mod 7) = 3) Or _
        ((AlternativeDay Mod 7) = 5)) _
    Then
        AlternativeDay = AlternativeDay + 1
    End If
    Hebrew_ElapsedCalendarDays = AlternativeDay
End Function

'---------------------------------------------------------------------------
Function Hebrew_DaysInYear(iYear)
'---------------------------------------------------------------------------
    Hebrew_DaysInYear = Hebrew_ElapsedCalendarDays(iYear + 1) - _
                        Hebrew_ElapsedCalendarDays(iYear)
End Function

'---------------------------------------------------------------------------
Function HebrewDateInOrAfterCivilYear(ByVal civilYear As Integer, _
                                      ByVal HebrewMonth As Integer, _
                                      ByVal HebrewDay As Integer, _
                                      Optional ByVal CalendarType As Integer = Gregorian) As Long
'---------------------------------------------------------------------------
    
    Dim jdnJanuary1 As Long
    Dim jdnHoliday  As Long
    Dim hebrewYear  As Integer
    Dim dummy1      As Integer
    Dim dummy2      As Integer
    
    jdnJanuary1 = civil_jdn(civilYear, 1, 1, CalendarType)
    Call jdn_hebrew(jdnJanuary1, hebrewYear, dummy1, dummy2)
    jdnHoliday = hebrew_jdn(hebrewYear, HebrewMonth, HebrewDay)
    If jdnHoliday < jdnJanuary1 Then
        ' Oops! Wrong civil Year. Use next hebrewYear's in stead.
        jdnHoliday = hebrew_jdn(hebrewYear + 1, HebrewMonth, HebrewDay)
    End If
    
    HebrewDateInOrAfterCivilYear = jdnHoliday

End Function


'---------------------------------------------------------------------------
Public Function jdn_Gregorian(jdn As Long) As Date
'---------------------------------------------------------------------------
    
    Dim l As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim y As Long
    Dim m As Long
    Dim d As Long
    
    l = jdn + 68569
    n = ((4 * l) \ 146097)
    l = l - ((146097 * n + 3) \ 4)
    i = ((4000 * (l + 1)) \ 1461001)
    l = l - ((1461 * i) \ 4) + 31
    j = ((80 * l) \ 2447)
    d = l - ((2447 * j) \ 80)
    l = (j \ 11)
    m = j + 2 - 12 * l
    y = 100 * (n - 49) + i + l
    
    jdn_Gregorian = DateSerial(y, m, d)

End Function

'---------------------------------------------------------------------------
Function julian_jdn(iYear As Integer, _
                    iMonth As Integer, _
                    iDay As Integer) As Long
'---------------------------------------------------------------------------
                    
    Dim lYear As Long
    Dim lMonth As Long
    Dim lDay As Long

    lYear = CLng(iYear)
    lMonth = CLng(iMonth)
    lDay = CLng(iDay)

    julian_jdn = 367 * lYear - _
            ((7 * (lYear + 5001 + ((lMonth - 9) \ 7))) \ 4) _
            + ((275 * lMonth) \ 9) + lDay + 1729777

End Function

'---------------------------------------------------------------------------
Function civil_jdn(iYear As Integer, _
                   iMonth As Integer, _
                   iDay As Integer, _
                   Optional CalendarType As Integer = Gregorian) As Long
'---------------------------------------------------------------------------
                   
    Dim lYear   As Long
    Dim lMonth  As Long
    Dim lDay    As Long

    If CalendarType = Gregorian And ((iYear > 1582) Or _
        ((iYear = 1582) And (iMonth > 10)) Or _
        ((iYear = 1582) And (iMonth = 10) And (iDay > 14))) _
    Then
        lYear = CLng(iYear)
        lMonth = CLng(iMonth)
        lDay = CLng(iDay)
        civil_jdn = ((1461 * (lYear + 4800 + ((lMonth - 14) \ 12))) \ 4) _
            + ((367 * (lMonth - 2 - 12 * (((lMonth - 14) \ 12)))) \ 12) _
            - ((3 * (((lYear + 4900 + ((lMonth - 14) \ 12)) \ 100))) \ 4) _
            + lDay - 32075
    Else
        civil_jdn = julian_jdn(iYear, iMonth, iDay)
    End If

End Function
