Imports System.Math

Public Class EB_Eclipse

    Private Structure ObserverType
        'these are sent:
        Public latitude As Decimal          'degrees
        Public longitude As Decimal         'degrees
        Public elevation As Decimal         'elevation (feet)
        Public date_ As Decimal              'local date (JD for 12 noon on Calendar date)
        Public time As Decimal              'local civil time
        Public timezone As Integer         'time zone correction (West is +)
        Public dstflag As Integer          'observe DST?

        'these are returned:
        Public zonename As String    '* 5   'name of time zone "(XXX)"
        Public jd As Decimal                'Julian day.ut for Greenwich
        Public jd0 As Decimal               'Julian day.0 for Greenwich
        Public dst As Integer              'daylight savings correction
        Public lmt As Decimal               'local mean time
        Public ut As Decimal                'Greenwich mean time
        Public gst As Decimal               'Greenwich sidereal time
        Public lst As Decimal               'Local sidereal time
        Public eot As Decimal               'Equation of time
        Public ast As Decimal               'Apparent solar time
        Public dt As Decimal                'delta-t correction for TDT
        Public tdt As Decimal               'Terrestrial dynamic time
        Public et As Decimal                'Ephemeris time
        Public tai As Decimal               'Intenational Atomic Time
        Public utc As Decimal               'Coordinated Universal Time
        Public nutlong As Decimal           'nutation of ecliptic longitude (deg)
        Public nutob As Decimal             'nutation of obliquity (deg)
        Public obliquity As Decimal         'obliquity of ecliptic (deg)

        'to save time later:
        Public longhour As Decimal          'longitude/15 (deg -> hours)
        Public gst0 As Decimal              'GST at 0h UT, factor to convert UT <-> GST
        Public tanlat As Decimal
        Public sinlat As Decimal
        Public coslat As Decimal
        Public sinob As Decimal             'sin and cos of obliquity
        Public cosob As Decimal
        Public psinphi As Decimal           'used to figure parallax
        Public pcosphi As Decimal
    End Structure

    Private Structure PlanetType
        'TYPE PlanetType
        Public meananomaly As Decimal       'mean anomaly (rad)
        Public trueanomaly As Decimal           'true anomaly (rad)
        Public node As Decimal              'longitude of ascending node (rad)
        Public heliolong As Decimal         'heliocentric longitude (rad)
        Public heliolat As Decimal          'heliocentric latitude
        Public earthdist As Decimal         'distance from earth (AUs)
        Public sunangle As Decimal          'angle from sun (deg)
        Public sundist As Decimal           'distance from sun (AUs)
        Public eclong As Decimal            'ecliptic longitude (rad)
        Public eclat As Decimal             'ecliptic latitude (rad)
        Public ra As Decimal                'RA (hours)
        Public dec As Decimal               'DEC (deg)
        Public rise As Decimal              'local rise time
        Public set_ As Decimal               'local set time
        Public transit As Decimal           'local transit time
        Public azrise As Decimal            'rise azimuth (deg)
        Public azset As Decimal             'set azimuth (deg)
        Public altitude As Decimal          'altitude (deg)
        Public azimuth As Decimal           'azimuth (deg)
        Public phase As Decimal             'phase (0 to 1)
        Public lighttime As Decimal         'time for light to get to earth (hours)
        Public diameter As Decimal          'apparent diameter (arcsec)
        Public mag As Decimal               'apparent brightness
        Public brightlimb As Decimal        'position-angle of bright limb (deg)
        Public parallax As Decimal          'parallax from earth (deg)
    End Structure

    Private Structure OrbitType
        Public epoch As Decimal             'epoch (1990.0 for planets)
        Public period As Decimal            'period (tropical years)
        Public longitude As Decimal         'longitude at epoch (deg)
        Public perihelion As Decimal        'longitude of perihelion/perigee (deg)
        Public eccentric As Decimal         'eccentricity of the orbit
        Public axis As Decimal              'semi-major axis (AU/km)
        Public inclination As Decimal       'inclination of orbit (deg)
        Public node As Decimal              'longitude of ascending node (deg)
        Public diameter As Decimal          'angular diameter at 1 AU (arcsec)
        Public mag As Decimal               'visual magnitude at 1 AU
    End Structure


    Private Structure PlanetDataType
        Public motion As Decimal      'mean daily motion (sec/day)
        Public velocity As Decimal      'orbital velocity (miles/sec)
        Public sidyear As Decimal      'sidereal year (earth days)
        Public synyear As Decimal      'synodic year (earth days)
        Public meandist As Decimal       'mean dist to sun (millions of miles)
        Public maxdist As Decimal       'max dist to sun        "
        Public mindist As Decimal       'min dist to sun        "
        Public maxearth As Decimal      'max dist from earth    "
        Public minearth As Decimal      'min dist from earth    "
        Public diameter As Decimal      'at equator (miles)
        Public volume As Decimal      'earth=1
        Public mass As Decimal      'earth=1
        Public density As Decimal      'water=1
        Public edensity As Decimal      'earth=1
        Public day_renamed As Decimal      'earth=1
        Public gravity As Decimal      'earth=1
        Public albedo As Decimal      '100  reflection=1
        Public meantemp As Decimal      'fø
    End Structure


    Dim observer As ObserverType
    Dim planet(3) As PlanetType
    Dim orbit(3) As OrbitType

    Dim planetdata(3) As PlanetDataType

    Dim ur(5), sd(5) As Decimal
    Const lunar As Boolean = True
    Const solar As Boolean = False
    Const feb29 As Decimal = 60                        'for leap year
    Const epoch1900 As Decimal = 2415020.0          'epoch 1900.0
    Const epoch1990 As Decimal = 2447891.5          'epoch January 0.0
    Const epoch2000 As Decimal = 2451545.0          'epoch 2000.0
    Const yeardays As Decimal = 365.242191          'days in tropical year
    Const AU As Decimal = 92.956198                 'AU (millions of miles)
    Const earthradius As Decimal = 3963.34 * 5280.0 'radius of earth (feet)
    Const obliquity2000 As Decimal = 23.439292      'mean obliquity of ecliptic @2000.0
    Const lightspeed As Decimal = 186282.3976       'mps
    Const sun As Short = 0
    Const moon As Short = 1
    Const earth As Short = 2

    Const pi As Decimal = 3.14159265358979
    Const halfpi As Decimal = pi / 2
    Const pi2 As Decimal = pi + pi


    Dim dk As Decimal = 1
    Dim doy(12) As Int16       'count into almfile = doy (month )+day 
    Dim k, sk, lk As Decimal
    Dim Eclipse, etype As String
    Dim sm, mm, fm, jd, jw, s1, c1, jy, gy, gymag, mu As Decimal
    Dim nt, mg, pm, um, sc, mag, cnt, i As Decimal
    Dim getit As Boolean 'report
    Dim year, month, day As Double
    Dim date_ As Decimal
    Dim time_ As Decimal
    Dim min, sec, deg, ut, dt As Decimal
    Dim IsDaylightSaving As Boolean = True
    Dim planetname(3) As String
    Dim earthlong, earthdist, sunra As Decimal


    Private Sub for_N1()

        planetname(0) = "SUN"
        orbit(0).epoch = 1990
        orbit(0).period = 365.242191
        orbit(0).longitude = 279.403303
        orbit(0).perihelion = 282.768422
        orbit(0).eccentric = 0.016713
        orbit(0).axis = 0
        orbit(0).inclination = 0
        orbit(0).node = 0
        orbit(0).diameter = 0.533128
        orbit(0).mag = 0


        planetname(1) = "Moon"
        orbit(1).epoch = 1990
        orbit(1).period = 27.3217
        orbit(1).longitude = 318.351648
        orbit(1).perihelion = 36.34041
        orbit(1).eccentric = 0.0549
        orbit(1).axis = 238855.7
        orbit(1).inclination = 5.145396
        orbit(1).node = 318.510107
        orbit(1).diameter = 0.5181
        orbit(1).mag = 0


        planetname(2) = "Earth"
        orbit(2).epoch = 1990
        orbit(2).period = 1.00004
        orbit(2).longitude = 99.403308
        orbit(2).perihelion = 102.768413
        orbit(2).eccentric = 0.016713
        orbit(2).axis = 1.0
        orbit(2).inclination = 0
        orbit(2).node = 0
        orbit(2).diameter = 0
        orbit(2).mag = 0



        '''''''''''''''''''

        planetdata(0).motion = 0
        planetdata(0).velocity = 0
        planetdata(0).sidyear = 0
        planetdata(0).synyear = 0
        planetdata(0).meandist = 0
        planetdata(0).maxdist = 0
        planetdata(0).mindist = 0
        planetdata(0).maxearth = 94.6
        planetdata(0).minearth = 91.4
        planetdata(0).diameter = 870331
        planetdata(0).volume = 1299370
        planetdata(0).mass = 332946
        planetdata(0).density = 1.44
        planetdata(0).edensity = 0.26
        planetdata(0).day_renamed = 24.7
        planetdata(0).gravity = 27.9
        planetdata(0).albedo = 0
        planetdata(0).meantemp = 10000



        planetdata(1).motion = 0
        planetdata(1).velocity = 0
        planetdata(1).sidyear = 27.3217
        planetdata(1).synyear = 29.5306
        planetdata(1).meandist = 0
        planetdata(1).maxdist = 0
        planetdata(1).mindist = 0
        planetdata(1).maxearth = 252710
        planetdata(1).minearth = 221463
        planetdata(1).diameter = 2159.89
        planetdata(1).volume = 0.02
        planetdata(1).mass = 0.0123
        planetdata(1).density = 3.42
        planetdata(1).edensity = 0.62
        planetdata(1).day_renamed = 27.3217
        planetdata(1).gravity = 0.17
        planetdata(1).albedo = 0.15
        planetdata(1).meantemp = -10




        planetdata(2).motion = 3548
        planetdata(2).velocity = 18.51
        planetdata(2).sidyear = 365.242191
        planetdata(2).synyear = 0
        planetdata(2).meandist = 92.9
        planetdata(2).maxdist = 94.6
        planetdata(2).mindist = 91.4
        planetdata(2).maxearth = 0
        planetdata(2).minearth = 0
        planetdata(2).diameter = 7926
        planetdata(2).volume = 1
        planetdata(2).mass = 1
        planetdata(2).density = 5.52
        planetdata(2).edensity = 1.0
        planetdata(2).day_renamed = 0.9973
        planetdata(2).gravity = 1
        planetdata(2).albedo = 0.37
        planetdata(2).meantemp = 72

        orbit(sun).axis = AU * 1000000
        planetdata(sun).maxdist = 94600000
        planetdata(sun).mindist = 91400000

    End Sub


    Private Sub init_class()
        doy(1) = 0
        doy(2) = 31
        doy(3) = 60
        doy(4) = 91
        doy(5) = 121
        doy(6) = 152
        doy(7) = 182
        doy(8) = 213
        doy(9) = 244
        doy(10) = 274
        doy(11) = 305
        doy(12) = 335
    End Sub

    Private Function DayOfYear(ByVal d As Integer, ByVal m As Integer) As Integer

        Return doy(m) + d

    End Function

    Private Function IsLeapYear(ByVal year As Integer) As Boolean

        Dim y1 As Integer = year \ 100
        Dim y2 As Integer = year - 100 * y1
        'after Gregorian Calendar?
        If year > 1582 Then
            If y2 = 0 Then
                IsLeapYear = (y1 Mod 4) = 0
            Else
                IsLeapYear = (y2 Mod 4) = 0
            End If
        Else
            IsLeapYear = (y1 Mod 4) = 0
        End If

    End Function


    Private Function DegToDec(ByRef deg As Decimal, ByRef min As Decimal, ByRef sec As Decimal) As Decimal
        DegToDec = deg + min / 60.0 + sec / 3600.0
    End Function


    Private Function calcJD(ByVal year As Decimal, ByVal month As Decimal, ByVal day As Decimal) As Decimal

        '***********************************************************************/
        '* Name:    calcJD
        '* Type:    Function
        '* Purpose: Julian day from calendar day
        '* Arguments:
        '*   year : 4 digit year
        '*   month: January = 1
        '*   day  : 1 - 31
        '* Return value:
        '*   The Julian day corresponding to the date
        '* Note:
        '*   Number is returned for start of day.  Fractional days should be
        '*   added later.
        '***********************************************************************/

        Dim A As Decimal
        Dim B As Decimal
        Dim JD As Decimal


        If (month <= 2) Then
            year = year - 1
            month = month + 12
        End If

        A = Math.Floor(year / 100)
        B = 2 - A + Math.Floor(A / 4)

        JD = Math.Floor(365.25 * (year + 4716)) + _
             Math.Floor(30.6001 * (month + 1)) + day + B - 1524.5
        calcJD = JD

        'gp put the year and month back where they belong
        If month = 13 Then
            month = 1
            year = year + 1
        End If
        If month = 14 Then
            month = 2
            year = year + 1
        End If

    End Function


    Public Sub FindEclipse(ByVal latitude_deg As Double, ByVal latitude_min As Double, _
                               ByVal latitude_sec As Double, _
                               ByVal longitude_deg As Double, ByVal longitude_min As Double, _
                               ByVal longitude_sec As Double, _
                               ByVal LocationTimeZone As Short, ByVal dstflag As Short, _
                               ByVal LocationElevation As Double, ByVal LocationDay As Double, ByVal LocationMonth As Double, ByVal LocationYear As Double _
                               )
        init_class()
        for_N1()
        Dim Date__ As Decimal

        'Dim day, month, year As Integer

        observer.latitude = DegToDec(latitude_deg, latitude_min, latitude_sec)
        observer.longitude = DegToDec(longitude_deg, longitude_min, longitude_sec)
        observer.timezone = LocationTimeZone  '8
        observer.dstflag = dstflag '-1
        observer.elevation = LocationElevation ' 2500

        'for debug 
        day = LocationDay ' 4
        month = LocationMonth '1
        year = LocationYear '2011

        ''dk =1 forward  0 back
        Date__ = calcJD(year, month, day)


        'count lunations from 1 Jan 1900
        CalendarDay(Date__, day, month, year)

        Dim n As Decimal = DayOfYear(day, month)

        If IsLeapYear(year) = False And n > feb29 Then
            n = n - 1
        End If

        k = (year + n / 365.0 - 1900.0) * 12.3685
        n = Abs(k - Fix(k))

        If k < 0.0 Then
            n = n + 1
        End If

        'solar
        sk = k

        If n > 0.5 Then
            sk = sk + 0.5 * Sign(k)
        End If

        sk = Fix(sk)

        'lunar
        lk = Fix(k) + 0.5 * Sign(k)


        getit = False


        Do
            'search forward
            If dk > 0 Then
                If sk < lk Then
                    Call CheckSolar()

                    If Not getit Then
                        Call CheckLunar()
                    End If
                Else
                    Call CheckLunar()
                    If Not getit Then
                        Call CheckSolar()
                    End If

                End If
                'search backward
            Else
                If sk > lk Then
                    Call CheckSolar()
                    If Not getit Then
                        Call CheckLunar()
                    End If

                Else
                    Call CheckLunar()
                    If Not getit Then
                        Call CheckSolar()
                    End If
                End If
            End If

        Loop Until getit


    End Sub


    Private Sub CheckSolar()
        Eclipse = "solar"
        k = sk
        Call CheckEclipse()
        sk = sk + dk
        Return
    End Sub


    Private Sub CheckLunar()
        Eclipse = "lunar"
        k = lk
        Call CheckEclipse()
        lk = lk + dk
    End Sub

    Private Sub CheckEclipse()
        ''need to think

        Dim t As Decimal = k / 1236.85
        Dim t2 As Decimal = t * t
        Dim t3 As Decimal = t2 * t

        'sun mean anomaly
        sm = 359.2242 + 29.10535608 * k - 0.0000333 * t2 - 0.00000347 * t3

        Normalize(sm, 360.0)
        sm = DegToRad(sm)

        'moon mean anomaly
        mm = 306.0253 + 385.81691806 * k + 0.0107306 * t2 + 0.00001236 * t3
        Normalize(mm, 360.0)
        mm = DegToRad(mm)

        'arg of latitude
        fm = 21.2964 + 390.67050646 * k - 0.0016528 * t2 - 0.00000239 * t3
        Normalize(fm, 360.0)
        fm = DegToRad(fm)

        'check for eclipse and bail if nada
        If Abs(Sin(fm)) > 0.36 Then Return

        'julian day: jw=whole number, jd=decimal
        jd = 0.75933 + 0.53058868 * k + 0.0001178 * t2 - 0.000000155 * _
            t3 + 0.00033 * Sin(DegToRad((166.56 + 132.87 * t - 0.009173 * t2)))

        If Eclipse = "lunar" Then jd = jd + 0.5
        jw = Fix(epoch1900 + 29.0 * k)

        'added
        Dim n As Decimal
        'time of max eclipse
        n = (0.1734 - 0.000393 * t) * Sin(sm) + 0.0021 * Sin(sm + sm) - 0.4068 * _
            Sin(mm) + 0.0161 * Sin(mm + mm) - 0.0051 * Sin(sm + mm) - 0.0074 * Sin(sm - mm) - 0.0104 * Sin(fm + fm)
        jd = jd + n
        jw = jw + Fix(jd)
        jd = jd - Fix(jd)

        'check varius lunar radii
        s1 = 5.19595 - 0.0048 * Cos(sm) + 0.002 * Cos(sm + sm) - 0.3283 * Cos(mm) - 0.006 * Cos(sm + mm) + 0.0041 * Cos(sm - mm)
        c1 = 0.207 * Sin(sm) + 0.0024 * Sin(sm + sm) - 0.039 * Sin(mm) + 0.0115 * Sin(mm + mm) - 0.0073 * Sin(sm + mm) - 0.0067 * Sin(sm - mm) + 0.0117 * Sin(fm + fm)

        'least distance of shadow axis
        gy = s1 * Sin(fm) + c1 * Cos(fm)
        gymag = Abs(gy)

        'radius of umbral cone
        mu = 0.0059 + 0.0046 * Cos(sm) - 0.0182 * Cos(mm) + 0.0004 * Cos(mm + mm) - 0.0005 * Cos(sm + mm)

        'semi-durations of eclipse phases
        nt = 0.5458 + 0.04 * Cos(mm)
        ur(0) = 1.5572 + mu        'lunar
        ur(1) = 1.0129 - mu
        ur(2) = 0.4679 - mu
        ur(3) = 1.572 + mu         'solar
        ur(4) = 1.026 - mu

        'eclipse magnitudes: mg =solar, pm =penumbral, um =lunar umbral
        mg = (1.5432 + mu - gymag) / (0.546 + mu + mu)
        pm = (1.5572 + mu - gymag) / 0.545
        um = (1.0129 - mu - gymag) / 0.545

        'final check
        If Eclipse = "lunar" And pm >= 0.0 Then
            Call LunarEclipse()
        ElseIf Eclipse = "solar" And gymag <= 1.5432 + mu Then
            Call SolarEclipse()
        End If

    End Sub

    Private Sub LunarEclipse()

        sc = 0
        'Dim i As Integer

        Select Case um
            Case Is < 0.0
                etype = "Penumbral Lunar"
                mag = pm
                cnt = 0
            Case Is >= 1.0
                etype = "Total Lunar"
                mag = um
                cnt = 2
            Case Else
                etype = "Partial Lunar"
                mag = um
                cnt = 1
        End Select

        For i = 0 To cnt
            sd(i) = Sqrt(ur(i) * ur(i) - gy * gy) / nt
        Next i

        Call EclipseReport()

    End Sub


    Private Sub SolarEclipse()
        etype = "Solar"
        mag = mg
        If gymag < 0.9972 Then
            Call CheckTotal()
            cnt = 1
        ElseIf gymag < 0.9972 + Abs(mu) Then
            Call CheckTotal()
            etype = etype + " (non-central)"
            cnt = 0
        Else
            etype = "Partial Solar"
            cnt = 0
        End If
        For i = 0 To cnt
            sd(i) = Sqrt(ur(i + 3) * ur(i + 3) - gy * gy) / nt
        Next i

        Call EclipseReport()

    End Sub

    Private Sub CheckTotal()

        If mu < 0.0 Then
            etype = "Total Solar"
        ElseIf mu > 0.0047 Then
            etype = "Annular Solar"
        ElseIf mu < 0.00464 * Cos(Atan(gy / Sqrt(Abs(-gy * gy + 1)))) Then
            etype = "Annular/Total Solar"
        Else
            etype = "Annular Solar"
        End If

    End Sub

    Private Sub EclipseReport()
        Dim strn As String
        Dim n As Decimal

        Call EclipseJDtoCAL()
        observer.date_ = date_
        observer.time = time_
        CalcPlanets(True)
        CalcPlanetsRiseSet()
        CalcPlanetsAltAz()

        'Debug.WriteLine("eclipse type" & etype)
        _EclipseType = etype

        'Debug.WriteLine("Greenwich date" & day & "/" & month & "/" & year)
        _Greenwich_Date = Fix(day) & "/" & month & "/" & year


        If Cos(fm) < 0.0 Then strn = " descending" Else strn = " ascending "

        strn = "moon position " & strn & " node "

        If gy < 0.0 Then
            strn = strn + " below "
        Else
            strn = strn + " above "
        End If
        'Debug.WriteLine(strn)
        _Moon_Position_Node = strn

        ''n = mag : If n > 1 Then n = 1
        'Debug.WriteLine("sun  {RA, Dec} at max" + LTrim$(DecTime24(planet(sun).ra, True, True)) + ", " + LTrim(DecDegree(planet(sun).dec, True)))
        _SunRaAtMax = LTrim$(DecTime24(planet(sun).ra, True, True)) + ", " + LTrim(DecDegree(planet(sun).dec, True))


        'Debug.WriteLine("moon {RA, Dec} at max" + LTrim$(DecTime24(planet(moon).ra, True, True)) + ", " + LTrim(DecDegree(planet(moon).dec, True)))
        _MoonRaAtMax = LTrim$(DecTime24(planet(moon).ra, True, True)) + ", " + LTrim(DecDegree(planet(moon).dec, True))

        'figure longitude of ascending node and compare with moon at totality
        n = RadToDeg(planet(moon).node) * 24 / 360

        'Debug.WriteLine("moon node RA at max " + LTrim(DecTime24(n, True, True)))
        _MoonNodeRAatMax = LTrim(DecTime24(n, True, True))

        n = n - planet(moon).ra
        strn = LTrim$(DecTime24(n, True, True))
        If n > 0 Then strn = "+" + strn

        'Debug.WriteLine("  (" + strn + ")")
        _MoonNodeRAatMax = "  (" + strn + ")"


        'how far moon is from ecliptic at max
        strn = LTrim(DecDegree(planet(moon).eclat, True))
        'Debug.WriteLine("  Dec from ecliptic at max " + strn)
        _DecFromEclipticAtMax = strn



        ''print location and time 
        Dim flag As Boolean

        If Eclipse.ToUpper = "lunar".ToUpper Then

            'Debug.WriteLine("local altitude at max" & DecDegree(planet(moon).altitude, True))
            _LocalAltitudeAtMax = DecDegree(planet(moon).altitude, True)

            'Debug.WriteLine("local azimuth at max" & DecDegree(planet(moon).azimuth, True))
            _LocalAzimuthAtMax = DecDegree(planet(moon).azimuth, True)

            'Debug.WriteLine("eclipse begins")

            flag = True
            n = ut - sd(0)
            If flag = True Then
                'Debug.WriteLine("moon enters penumbra (P1)" + DecTime24(n, True, False))
                _MoonEntersPenumbra_P1_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _MoonEntersPenumbra_P1_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _MoonEntersPenumbra_P1_PST = DecTime12(n)
            End If


            flag = cnt > 0
            n = ut - sd(1)
            If flag = True Then
                'Debug.WriteLine("moon enters umbra  (U1)" + DecTime24(n, True, False))
                _MoonEntersUmbra_U1_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _MoonEntersUmbra_U1_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _MoonEntersPenumbra_P1_PST = DecTime12(n)
            End If

            flag = cnt > 1
            n = ut - sd(2)
            If flag = True Then
                'Debug.WriteLine("totality begins  (U2)" + DecTime24(n, True, False))
                _TotalityBegins_U2_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _TotalityBegins_U2_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _TotalityBegins_U2_PST = DecTime12(n)
            End If

            flag = True
            n = ut
            If flag = True Then
                'Debug.WriteLine("max eclipse    (UT)" + DecTime24(n, True, False))
                _MaxEclipse_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _MaxEclipse_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _MaxEclipse_PST = DecTime12(n)
            End If


            flag = cnt > 1
            n = ut + sd(2)
            If flag = True Then
                'Debug.WriteLine("    totality ends    (U3)" + DecTime24(n, True, False))
                _TotalityEnds_U3_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _TotalityEnds_U3_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _TotalityEnds_U3_PST = DecTime12(n)
            End If


            flag = cnt > 0
            n = ut + sd(1)
            If flag = True Then
                'Debug.WriteLine("  moon leaves umbra  (U4)" + DecTime24(n, True, False))
                _MoonLeavesUmbra_U4_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _MoonLeavesUmbra_U4_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _MoonLeavesUmbra_U4_PST = DecTime12(n)
            End If

            flag = True
            n = ut + sd(0)
            If flag = True Then
                'Debug.WriteLine("  moon leaves penumbra (P4)" + DecTime24(n, True, False))
                _MoonLeavesPenumbra_P4_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _MoonLeavesPenumbra_P4_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _MoonLeavesPenumbra_P4_PST = DecTime12(n)
            End If

        Else  ''solar

            'Debug.WriteLine("  local altitude at max" + LTrim$(DecDegree(planet(sun).altitude, True)))
            _LocalAltitudeAtMax = LTrim$(DecDegree(planet(sun).altitude, True))
            'Debug.WriteLine("  local azimuth at max" + LTrim$(DecDegree$(planet(sun).azimuth, True)))
            _LocalAzimuthAtMax = LTrim$(DecDegree$(planet(sun).azimuth, True))

            'Debug.WriteLine("eclipse begins......")

            flag = True
            n = ut - sd(0)
            If flag = True Then
                'Debug.WriteLine("eclipse begins" + DecTime24(n, True, False))
                _EclipseBegins_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _EclipseBegins_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _EclipseBegins_PST = DecTime12(n)
            End If

            flag = (cnt > 0)
            n = ut - sd(1)
            If flag = True Then
                'Debug.WriteLine("central eclipse begins" + DecTime24(n, True, False))
                _CentralEclipseBegins_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _CentralEclipseBegins_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _CentralEclipseBegins_PST = DecTime12(n)
            End If

            flag = True
            n = ut
            If flag = True Then
                'Debug.WriteLine("maximum eclipse" + DecTime24(n, True, False))
                _MaximumEclipse_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _MaximumEclipse_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _MaximumEclipse_PST = DecTime12(n)
            End If

            flag = (cnt > 0)
            n = ut + sd(1)
            If flag = True Then
                'Debug.WriteLine("central eclipse ends" + DecTime24(n, True, False))
                _CentralEclipseEnds_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _CentralEclipseEnds_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _CentralEclipseEnds_PST = DecTime12(n)
            End If

            flag = True
            n = ut + sd(0)
            If flag = True Then
                'Debug.WriteLine("eclipse ends" + DecTime24(n, True, False))
                _EclipseEnds_UT = DecTime24(n, True, False)
                EclipseUTtoLCT(n, dt)
                'Debug.WriteLine(DecTime24(n, True, False))
                _EclipseEnds_LCT = DecTime24(n, True, False)
                'Debug.WriteLine(DecTime12(n))
                _EclipseEnds_PST = DecTime12(n)
            End If

        End If

        '    lastdate = date_
        getit = True
    End Sub

    Private Function DecTime12$(ByVal T As Decimal)
        Dim n As Decimal = T
        Dim ampm As String
        Normalize(n, 24.0)
        Dim Hour As Short = Int(n)
        n = 60 * (n - Hour)
        Dim min As Short = Int(n)
        Dim sec As Short = 60 * (n - min)
        If sec >= 60 Then min = min + 1 : sec = sec - 60
        If min >= 60 Then Hour = Hour + 1 : min = min - 60
        If Hour < 12 Then ampm = " am" Else ampm = " pm"
        If Hour > 12 Then Hour = Hour - 12
        If Hour = 0 Then Hour = 12
        DecTime12 = Right$(" " + LTrim$(Str$(Hour)), 2) + ":" + Right$("0" + LTrim$(Str$(min)), 2) + ampm
    End Function

    Private Sub EclipseUTtoLCT(ByRef n As Decimal, ByRef dt As Decimal)
        n = n + dt
        If n < 0.0 Then n = n + 24 Else If n >= 24 Then n = n - 24
    End Sub

    Private Function DecTime24(ByVal T As Decimal, ByVal secflag As Boolean, ByVal unitflg As Boolean) As String
        Dim n As Decimal = T
        Dim strn As String
        Normalize(n, 24.0)
        Dim Hour As Short = Int(n)
        Dim min As Short
        Dim sec As Short

        n = 60 * (n - Hour)
        min = Int(n)
        sec = 60 * (n - min)
        If sec >= 60 Then min = min + 1 : sec = sec - 60
        If min >= 60 Then Hour = Hour + 1 : min = min - 60
        If unitflg Then
            strn = LTrim$(Str$(Hour)) + "h:" + Right$("0" + LTrim$(Str$(min)), 2) + "m"
            If secflag Then strn = strn + ":" + Right$("0" + LTrim$(Str$(sec)), 2) + "s"
            If Hour < 10 Then strn = " " + strn
        Else
            strn = LTrim$(Str$(Hour)) + ":" + Right$("0" + LTrim$(Str$(min)), 2)
            If secflag Then strn = strn + ":" + Right$("0" + LTrim$(Str$(sec)), 2)
            If Hour < 10 Then strn = " " + strn
        End If
        Return strn
    End Function


    Private Sub CalcPlanetsAltAz()
        ''Dim planet() As PlanetType
        For i = sun To earth
            CalcAltAz(planet(i).ra, planet(i).dec, planet(i).altitude, planet(i).azimuth)
        Next i
    End Sub

    Private Sub CalcPlanetsRiseSet()
        'Dim planet() As PlanetType
        CalcSunRiseSet()
        CalcMoonRiseSet()
    End Sub

    Private Sub EclipticToEquator(ByRef elong As Decimal, ByRef elat As Decimal, ByRef ra As Decimal, ByRef dec As Decimal)

        Dim rlon, rlat As Decimal

        rlon = DegToRad(elong)
        rlat = DegToRad(elat)
        dec = RadToDeg(ArcSin(Sin(rlat) * observer.cosob + Cos(rlat) * observer.sinob * Sin(rlon)))
        ra = RadToDeg(ArcTan(Sin(rlon) * observer.cosob - Tan(rlat) * observer.sinob, Cos(rlon))) / 15.0

    End Sub

    Private Sub CalcSunRiseSet()
        Dim x As Decimal = 0D
        Dim da As Decimal = 0D
        Dim psi As Double = 0.0R
        Dim cosdec As Double = 0.0R
        Dim dec As Decimal = 0D
        Dim gstset As Decimal = 0D
        Dim gstrise As Decimal = 0D
        Dim t00 As Decimal = 0D
        ''Dim observer As ObserverType, planet() As PlanetType
        Dim n, l As Decimal
        Dim ra1, ra2 As Decimal
        Dim dec1, dec2, gstset1, gstset2, gstrise1, gstrise2 As Decimal
        'position of sun at last next local midnight
        n = 0.985647       'mean change in ecliptic longitude per 24 hours
        l = RadToDeg(planet(sun).eclong) - observer.lmt * n / 24.0

        EclipticToEquator(l, 0.0, ra1, dec1)
        EclipticToEquator(l + n, 0.0, ra2, dec2)

        'figure GST for each rise/set
        Dim rdec As Decimal = DegToRad(dec1)
        n = RadToDeg(ArcCos(-observer.tanlat * Tan(rdec))) / 15.0
        gstrise1 = ra1 - n + observer.longhour
        gstset1 = ra1 + n + observer.longhour

        rdec = DegToRad(dec2)
        n = RadToDeg(ArcCos(-observer.tanlat * Tan(rdec))) / 15.0
        gstrise2 = ra2 - n + observer.longhour
        gstset2 = ra2 + n + observer.longhour

        Normalize24(gstrise1)
        Normalize24(gstset1)
        Normalize24(gstrise2)
        Normalize24(gstset2)
        If gstrise1 > gstrise2 Then gstrise2 = gstrise2 + 24.0
        If gstset1 > gstset2 Then gstset2 = gstset2 + 24.0


        t00 = observer.gst0 + observer.longhour * 1.002737909

        If gstrise1 < t00 Then gstrise1 = gstrise1 + 24.0 : gstrise2 = gstrise2 + 24
        If gstset1 < t00 Then gstset1 = gstset1 + 24.0 : gstset2 = gstset2 + 24.0

        gstrise = (24.07 * gstrise1 - t00 * (gstrise2 - gstrise1)) / (24.07 + gstrise1 - gstrise2)
        gstset = (24.07 * gstset1 - t00 * (gstset2 - gstset1)) / (24.07 + gstset1 - gstset2)

        'figure corrections for average RA+DEC (= local noon)
        dec = DegToRad((dec1 + dec2) / 2.0)
        cosdec = Cos(dec)
        psi = ArcCos(observer.sinlat / cosdec)

        'standard corrections:
        '         sun diameter= 5.33 deg
        '  horizontal parallax= 8.79 arcsec
        'refraction at horizon= 34 arcmin
        x = DegToRad(0.533 / 2.0 - 8.79 / 3600.0 + 34.0 / 60.0)

        'rise/set azimuths
        n = RadToDeg(ArcCos(Sin(dec)) / observer.coslat)
        da = RadToDeg(ArcSin(Tan(x) / Tan(psi)))
        planet(sun).azrise = n - da
        planet(sun).azset = 360.0 - n + da

        'rise/set times

        n = RadToDeg(ArcSin(Sin(x) / Sin(psi)))
        Dim dt As Decimal = (240.0 * n / cosdec) / 3600.0
        planet(sun).rise = GSTtoLCT(gstrise - dt)
        planet(sun).set_ = GSTtoLCT(gstset + dt)
        planet(sun).transit = LSTtoLCT((ra1 + ra2) / 2.0)

    End Sub
    Private Function GSTtoLCT(ByVal n As Decimal) As Decimal
        'Dim observer As ObserverType
        Dim gst As Double = n
        If gst >= 24.0 Then gst = gst - 24.0
        Dim ut As Decimal = gst - observer.gst0
        If ut < 0.0 Then ut = ut + 24.0#
        Dim lct As Decimal = ut * 0.9972695663 - observer.timezone + observer.dst
        If lct < 0.0 Then lct = lct + 24.0 Else If lct >= 24.0 Then lct = lct - 24.0
        GSTtoLCT = lct
    End Function

    Private Function LSTtoLCT(ByVal n As Decimal) As Decimal
        Dim lst As Double = 0.0R
        'Dim observer As ObserverType
        lst = n
        If lst < 24.0 Then lst = lst + 24.0 Else If lst >= 24.0 Then lst = lst - 24.0
        LSTtoLCT = GSTtoLCT(lst + observer.longhour)
    End Function

    Private Sub CalcMoonRiseSet()
        Dim da As Decimal = 0D
        Dim x As Decimal = 0D
        Dim psi As Double = 0.0R
        Dim cosdec As Double = 0.0R
        Dim dec As Decimal = 0D
        Dim ra As Decimal = 0D
        Dim gstset As Decimal = 0D
        Dim gstrise As Decimal = 0D
        Dim t00 As Decimal = 0D
        Dim gstset2 As Decimal = 0D
        Dim gstrise2 As Decimal = 0D
        Dim gstset1 As Decimal = 0D
        Dim gstrise1 As Decimal = 0D
        Dim n As Double = 0.0R
        Dim rdec As Decimal = 0D
        Dim dec2 As Decimal = 0D
        Dim ra2 As Decimal = 0D
        Dim dec1 As Decimal = 0D
        Dim ra1 As Decimal = 0D
        Dim elong As Decimal = 0D
        Dim elat As Decimal = 0D
        Dim dlong As Double = 0.0R
        Dim dlat As Double = 0.0R
        ''Dim observer As ObserverType, planet() As PlanetType

        'extrapolate position of mooon to last local midnight and local noon (+12h)
        'seems better if extrapolate midnight to midnight (+24 hours)
        Dim dt As Decimal = observer.lmt
        dlat = 0.05 * Cos(planet(moon).heliolong - planet(moon).node)
        dlong = 0.55 + 0.06 * Cos(planet(moon).meananomaly)
        elat = RadToDeg(planet(moon).eclat) - dt * dlat
        elong = RadToDeg(planet(moon).eclong) - dt * dlong
        EclipticToEquator(elong, elat, ra1, dec1)
        'EclipticToEquator elong  + 12  * dlong , elat  + 12  * dlat , ra2 , dec2 
        EclipticToEquator(elong + 24.0 * dlong, elat + 24.0 * dlat, ra2, dec2)

        'figure GST for each rise/set
        rdec = DegToRad(dec1)
        n = RadToDeg(ArcCos(-observer.tanlat * Tan(rdec))) / 15.0
        gstrise1 = ra1 - n + observer.longhour
        gstset1 = ra1 + n + observer.longhour

        rdec = DegToRad(dec2)
        n = RadToDeg(ArcCos(-observer.tanlat * Tan(rdec))) / 15.0
        gstrise2 = ra2 - n + observer.longhour
        gstset2 = ra2 + n + observer.longhour

        Normalize24(gstrise1)
        Normalize24(gstset1)
        Normalize24(gstrise2)
        Normalize24(gstset2)
        If gstrise1 > gstrise2 Then gstrise2 = gstrise2 + 24.0
        If gstset1 > gstset2 Then gstset2 = gstset2 + 24.0

        t00 = observer.gst0 + observer.longhour * 1.002737909
        If t00 < 0.0 Then t00 = t00 + 24.0

        If gstrise1 < t00 Then gstrise1 = gstrise1 + 24.0 : gstrise2 = gstrise2 + 24
        If gstset1 < t00 Then gstset1 = gstset1 + 24.0 : gstset2 = gstset2 + 24.0

        'gstrise  = (12.03  * gstrise1  - t00  * (gstrise2  - gstrise1 )) / (12.03  + gstrise1  - gstrise2 )
        'gstset  = (12.03  * gstset1  - t00  * (gstset2  - gstset1 )) / (12.03  + gstset1  - gstset2 )
        gstrise = (24.07 * gstrise1 - t00 * (gstrise2 - gstrise1)) / (24.07 + gstrise1 - gstrise2)
        gstset = (24.07 * gstset1 - t00 * (gstset2 - gstset1)) / (24.07 + gstset1 - gstset2)

        'figure corrections for average RA+DEC
        ra = (ra1 + ra2) / 2.0
        dec = DegToRad((dec1 + dec2) / 2.0)
        cosdec = Cos(dec)
        psi = ArcCos(observer.sinlat / cosdec)

        'correct for:
        '  apparent moon diameter
        '  horizontal parallax
        '  refraction at horizon= 34 arcmin
        x = DegToRad(planet(moon).diameter / 2.0 - planet(moon).parallax + 34.0 / 60.0)

        'rise/set azimuths
        n = RadToDeg(ArcCos(Sin(dec)) / observer.coslat)
        da = RadToDeg(ArcSin(Tan(x) / Tan(psi)))
        planet(moon).azrise = n - da
        planet(moon).azset = 360.0 - n + da

        'rise/set/transit times
        n = RadToDeg(ArcSin(Sin(x) / Sin(psi)))
        dt = (240.0 * n / cosdec) / 3600.0
        planet(moon).rise = GSTtoLCT(gstrise - dt)
        planet(moon).set_ = GSTtoLCT(gstset + dt)
        planet(moon).transit = LSTtoLCT(ra)

    End Sub

    Private Sub CalcPlanets(ByVal flag As Boolean)

        Dim pifac As Decimal = 360.0 / pi
        Dim n, mp, vp, E, d, l As Decimal

        '''observer.jd
        CalcObserver()
        Dim dayfac As Decimal = 360.0 / yeardays * (observer.jd - epoch1990)

        CalcSun(flag)

        For i = earth To earth
            n = dayfac / orbit(i).period
            Normalize(n, 360.0)

            'mean anomaly
            mp = DegToRad(n + orbit(i).longitude - orbit(i).perihelion)
            Normalize(mp, pi2)
            E = orbit(i).eccentric

            'true anomaly
            If flag Then
                'solve kepler's equation {E - e sin E = M} for E
                vp = mp
                Do
                    d = vp - E * Sin(vp) - mp
                    If Abs(d) <= 0.000001 Then Exit Do
                    vp = vp - d / (1.0 - E * Cos(vp))
                Loop
                vp = 2.0 * Atan(Sqrt((1.0 + E) / (1.0 - E)) * Tan(vp / 2.0))
                l = vp + DegToRad(orbit(i).perihelion)

            Else
                'approximate solution
                l = DegToRad(n + (pifac * E * Sin(mp)) + orbit(i).longitude)
                vp = l - DegToRad(orbit(i).perihelion)

            End If

            Normalize(vp, pi2)
            Normalize(l, pi2)
            planet(i).meananomaly = mp
            planet(i).trueanomaly = vp
            planet(i).heliolong = l
            planet(i).sundist = orbit(i).axis * (1.0 - (E * E)) / (1.0 + E * Cos(vp))

        Next i

        'data to use again
        earthlong = planet(earth).heliolong
        earthdist = planet(earth).sundist
        sunra = planet(sun).ra

        'figure RA and DEC
        For i = earth To earth
            If i <> earth Then
                PlanetRaDec()
            End If
        Next i

        CalcMoon()

    End Sub


    Private Sub PlanetRaDec()
        Dim r As Decimal = 0D
        Dim node, incline, n, heliolat, l, eclong, eclat As Decimal

        node = DegToRad(orbit(i).node)
        incline = DegToRad(orbit(i).inclination)
        n = planet(i).heliolong - node

        heliolat = ArcSin(Sin(n) * Sin(incline))
        planet(i).heliolat = heliolat
        l = ArcTan(Sin(n) * Cos(incline), Cos(n)) + node
        r = planet(i).sundist * Cos(heliolat)

        'ecliptic latitude and longitude
        If i < earth Then
            'inner planets
            n = earthlong - l
            eclong = ArcTan(r * Sin(n), earthdist - r * Cos(n)) + pi + earthlong

        Else
            'outer planets
            n = l - earthlong
            eclong = ArcTan(earthdist * Sin(n), r - earthdist * Cos(n)) + l
        End If

        Normalize(eclong, pi2)
        eclat = Atan(r * Tan(heliolat) * Sin(eclong - l) / (earthdist * Sin(l - earthlong)))

        planet(i).node = node
        planet(i).eclong = eclong
        planet(i).eclat = eclat
        planet(i).dec = RadToDeg(ArcSin(Sin(eclat) * observer.cosob + Cos(eclat) * observer.sinob * Sin(eclong)))
        planet(i).ra = RadToDeg(ArcTan(Sin(eclong) * observer.cosob - Tan(eclat) * observer.sinob, Cos(eclong))) / 15.0

    End Sub


    Private Function DecDegree(ByVal d As Decimal, ByVal secflag As Boolean) As String

        Dim n As Decimal
        Dim strn As String

        DecToDeg(d, deg, min, sec)
        If secflag = False And sec > 30 Then
            min = min + 1
            If min >= 60 Then
                min = min - 60
                If d > 0 Then deg = deg + 1 Else deg = deg - 1
            End If
        End If
        strn = Str(deg) + "ø " + Right("0" + LTrim(Str(min)), 2) + "'"
        If secflag Then strn = strn + sec.ToString().Substring(0, 2) + "''"
        If Abs(deg) < 100 Then
            strn = " " + strn
            If Abs(deg) < 10 Then strn = " " + strn
        End If
        DecDegree = strn
    End Function

    Private Sub DecToDeg(ByRef d As Decimal, ByRef deg As Decimal, _
                         ByRef min As Decimal, ByRef sec As Decimal)
        Dim n As Decimal

        n = d
        deg = Fix(d)
        n = 60 * Abs((n - deg))
        min = Int(n)
        sec = 60 * (n - min)
        If sec >= 60 Then min = min + 1 : sec = sec - 60

    End Sub


    Private Sub EclipseJDtoCAL()
        Dim n, n1, m As Decimal

        jd = jd + 0.5
        If jd >= 1 Then
            jw = jw + 1
            jd = jd - 1
        End If

        n = Fix(jw)
        n1 = Fix((n - 1867216.25) / 36524.25)
        If n >= 2299160.0 Then
            n = n + 1.0 + n1 - Fix(n1 / 4)
        End If

        m = n + 1524.0
        year = Fix((m - 122.1) / 365.25)
        n = Fix(365.25 * year)
        month = Fix((m - n) / 30.6001)
        n = m - n - Fix(30.6001 * month) + jd
        day = Fix(n)
        ut = (n - Fix(n)) * 24                       'ET time for max eclipse
        If month < 14 Then month = month - 1
        If month >= 14 Then month = month - 13
        If month > 2 Then year = year - 4716
        If month <= 2 Then year = year - 4715
        JulianDay(date_, day, month, year)

        'dummy call to set time corrections
        observer.date_ = date_
        observer.time = 12

        CalcObserver()
        ''calc planets???
        dt = -observer.timezone + observer.dst

        'local time and date
        time_ = ut + dt
        If time_ < 0.0 Then
            time_ = time_ + 24
            date_ = date_ - 1
        ElseIf time_ >= 24.0 Then
            time_ = time_ - 24
            date_ = date_ + 1
        End If
        ''EclipseReport()

    End Sub
    Private Function IsDaylightSavings(ByVal day As Decimal, ByVal month As Decimal, ByVal year As Double) As Boolean

        If month >= 4 And month <= 10 Then
            Return True
        Else
            Return False
        End If

        ''Return IsDaylightSaving
    End Function


    Private Sub CalendarDay(ByRef date_ As Decimal, ByRef day As Double, ByRef month As Double, ByRef year As Double)

        Dim ja, jb, jc, jd, je As Double
        'deal with cross-over to Gregorian Calendar
        Dim greg As Double = 2299161

        If date_ >= greg Then
            ja = (((date_ - 1867216) - 0.25) / 36524.25)
            ja = date_ + 1 + ja - Int(0.25 * ja)
        Else
            ja = date_
        End If

        jb = ja + 1524
        jc = Int(6680.0 + ((jb - 2439870) - 122.1) / 365.25)
        jd = 365 * jc + Int(0.25 * jc)
        je = Int((jb - jd) / 30.6001)

        day = jb - jd - Int(30.6001 * je)
        month = je - 1
        If month > 12 Then month = month - 12

        year = jc - 4715
        If month > 2 Then year = year - 1
        If year <= 0 Then year = year - 1


    End Sub



    Private Sub CalcObserver()
        Dim n, dt As Decimal
        Dim t, a, l, b As Decimal
        'Shared observer As ObserverType
        Dim pi2 As Double = Math.PI * 2

        date_ = observer.date_
        CalendarDay(date_, day, month, year)

        'daylight savings time?
        'How figure correctly?--figure day-of-year for start/end of DST
        If observer.dstflag And IsDaylightSavings(day, month, year) Then
            observer.dst = 1
        Else
            observer.dst = 0
        End If

        'figure name for time zone. How deal with longitudes?
        If observer.dst Then
            observer.zonename = "(PDT)"
        Else
            observer.zonename = "(PST)"
        End If

        'approximate correction for ET/TDT (sec)
        Select Case year
            Case 1600 To 1650 : dt = 80
            Case 1650 To 1700 : dt = 30
            Case 1700 To 1750 : dt = 16
            Case 1750 To 1800 : dt = 20
            Case 1800 To 1850 : dt = 10
            Case 1850 To 1900 : dt = 0
            Case 1900 To 1950 : dt = 30
            Case 1950 To 2000 : dt = 50
            Case Else : dt = 0
        End Select

        'time/date conversions
        observer.jd0 = date_ - 0.5
        observer.lmt = observer.time - observer.dst
        n = observer.lmt + observer.timezone

        If n < 0.0 Then
            observer.jd0 = observer.jd0 - 1.0
            n = n + 24.0
        ElseIf n >= 24.0 Then
            observer.jd0 = observer.jd0 + 1.0
            n = n - 24.0
        End If

        observer.jd = observer.jd0 + n / 24.0
        observer.ut = n

        observer.dt = dt / 3600.0
        observer.et = observer.ut + observer.dt       'how REALLY figure?
        observer.tdt = observer.et
        observer.tai = observer.et - 32.184 / 3600.0

        'Greenwich sidereal time
        n = (observer.jd0 - 2451545.0) / 36525.0
        n = 6.697374558 + 2400.051336 * n + 0.000025862 * n * n
        Normalize24(n)
        observer.gst0 = n

        n = n + observer.ut * 1.002737909
        Normalize24(n)
        observer.gst = n

        'correct for longitude
        observer.longhour = observer.longitude / 15.0

        'Local sidereal time
        n = n - observer.longhour
        If n < 0.0 Then n = n + 24.0 Else If n >= 24.0 Then n = n - 24.0

        observer.lst = n


        'nutation correction
        t = (observer.jd - epoch1900) / 36525.0
        a = t * 100.002136
        l = DegToRad(279.6967 + 360.0 * (a - Fix(a)))  'sun mean longitude
        Normalize(l, pi2)
        b = 5.372617 * t
        n = DegToRad(259.1833 - 360.0 * (b - Fix(b)))  'moon node longitude
        Normalize(n, pi2)
        observer.nutlong = (-17.2 * Sin(n) - 1.3 * Sin(l + l)) / 3600.0
        observer.nutob = (9.199999999999 * Cos(n) + 0.5 * Cos(l + l)) / 3600.0

        'obliquity of ecliptic
        t = (observer.jd - epoch2000) / 36525.0
        n = obliquity2000 - (46.815 * t + 5.99999999999 * t * t - 0.00181 * t * t * t) / 3600
        observer.obliquity = n
        observer.cosob = Cos(DegToRad(n))
        observer.sinob = Sin(DegToRad(n))

        'stuff to save time later
        n = DegToRad(observer.latitude)
        observer.tanlat = Tan(n)
        observer.sinlat = Sin(n)
        observer.coslat = Cos(n)

        'pre-figure for parallax calculation
        a = observer.elevation / earthradius
        n = Math.Atan(0.996647 * observer.tanlat)
        observer.psinphi = 0.996647 * Sin(n) + a * observer.sinlat
        observer.pcosphi = Cos(n) + a * observer.coslat

    End Sub



    Private Sub JulianDay(ByRef date_ As Decimal, _
                          ByRef day As Decimal, ByRef month As Decimal, ByRef year As Decimal)

        Dim greg As Decimal = 588829         'Gregorian calendar adopted 4 oct 1582
        Dim ty, jm, ja As Decimal

        'there was no year 0! the year after 1 BC was 1 AD.
        If year = 0 Then Exit Sub

        If year < 0 Then
            ty = year + 1
        Else
            ty = year
        End If


        If month > 2 Then
            jy = ty
            jm = month + 1
        Else
            jy = ty - 1
            jm = month + 13
        End If

        jd = Int(365.25 * jy) + Int(30.6001 * jm) + day + 1720995

        If day + 31 * (month + 12 * ty) >= greg Then
            ja = Int(0.01 * jy)
            jd = jd + 2 - ja + Int(0.25 * ja)
        End If

        date_ = jd

    End Sub



    Private Function DegToRad(ByVal deg As Decimal) As Decimal
        DegToRad = (deg * pi / 180)
    End Function

    Private Sub Normalize(ByRef n As Decimal, ByVal norm As Decimal)
        n = n - Int(n / norm) * norm
    End Sub

    Private Sub Normalize24(ByRef n As Decimal)
        n = n - Int(n / 24.0) * 24.0
    End Sub


    Private Sub CalcSun(ByVal flag As Boolean)
        'Shared observer As ObserverType, planet() As PlanetType
        'Shared orbit() As OrbitType
        Dim d, n, m, e, v, l, ec As Decimal
        Dim dec, ra As Decimal

        d = observer.jd - epoch1990
        n = 360.0 * d / yeardays
        Normalize(n, 360)

        m = DegToRad(n + orbit(sun).longitude - orbit(sun).perihelion)
        If m < 0.0 Then m = m + pi2
        e = orbit(sun).eccentric

        If flag Then
            'solve kepler's equation E- e sin E = M for E
            v = m
            Do
                d = v - e * Sin(v) - m
                If Abs(d) < 0.000001 Then Exit Do
                v = v - d / (1.0 - e * Cos(v))
            Loop

            v = 2.0 * Atan(Math.Sqrt((1.0 + e) / (1.0 - e)) * Tan(v / 2.0))
            l = v + DegToRad(orbit(sun).perihelion)

        Else
            'short cut
            ec = (360.0 / pi) * e * Sin(m)
            v = m + DegToRad(ec)
            l = DegToRad(n + ec + orbit(sun).longitude)
        End If

        Normalize(v, pi2)
        Normalize(l, pi2)
        planet(sun).meananomaly = m
        planet(sun).trueanomaly = v
        planet(sun).eclong = l
        planet(sun).eclat = 0


        'EclipticToEquator l , 0 , planet(sun).ra, planet(sun).dec
        dec = RadToDeg(ArcSin(observer.sinob * Sin(l)))
        ra = RadToDeg(ArcTan(Sin(l) * observer.cosob, Cos(l))) / 15.0
        planet(sun).ra = ra
        planet(sun).dec = dec
        CalcAltAz(ra, dec, planet(sun).altitude, planet(sun).azimuth)

    End Sub

    Private Function RadToDeg(ByVal rad As Decimal) As Decimal
        RadToDeg = 180 * rad / pi
    End Function

    Private Function ArcSin(ByVal n As Decimal) As Decimal
        Dim m As Decimal = -n * n + 1
        If m <= 0 Then
            ArcSin = halfpi
        Else
            ArcSin = Atan(n / Sqrt(m))
        End If

    End Function

    Private Function ArcCos(ByVal n As Decimal) As Decimal
        Dim m As Decimal

        m = -n * n + 1
        If m <= 0 Then
            ArcCos = 0
        Else
            ArcCos = -Atan(n / Sqrt(m)) + halfpi
        End If

    End Function

    Private Function ArcTan(ByVal y As Decimal, ByVal x As Decimal) As Decimal
        Dim n As Decimal
        Select Case Sign(x)
            Case -1 : n = pi + Atan(y / x)
            Case 0 : n = halfpi : If y < 0.0 Then n = n + pi
            Case 1 : n = Atan(y / x) : If n < 0.0 Then n = n + pi2
        End Select

        ArcTan = n
    End Function


    Private Sub CalcMoon()
        'Shared observer As ObserverType, planet() As PlanetType
        'Shared orbit() As OrbitType
        Dim incline, sunlong, sunanon, days, l, n As Decimal

        incline = DegToRad(orbit(moon).inclination)
        sunlong = planet(sun).eclong
        sunanon = Sin(planet(sun).meananomaly)

        days = observer.jd + observer.dt - epoch1990
        l = 13.1763966 * days + orbit(moon).longitude
        Normalize(l, 360.0)

        mm = l - 0.1114041 * days - orbit(moon).perihelion
        Normalize(mm, 360.0)

        n = orbit(moon).node - 0.0529539 * days
        Normalize(n, 360.0)

        Dim ev, ae, a3, ec, a4, v, p, elong, elat, ra, dec As Decimal

        'corrections
        ev = 1.2739 * Sin(2.0 * (DegToRad(l) - sunlong) - DegToRad(mm))
        ae = 0.1858 * sunanon
        a3 = 0.37 * sunanon
        mm = DegToRad(mm + ev - ae - a3)
        ec = 6.2886 * Sin(mm)
        a4 = 0.214 * Sin(mm + mm)
        l = l + ev + ec - ae + a4
        v = 0.6583 * Sin(2.0 * (DegToRad(l) - sunlong))
        l = DegToRad(l + v)
        n = DegToRad(n - 0.16 * sunanon)

        p = l - n
        elong = ArcTan(Sin(p * Cos(incline)), Cos(p)) + n
        Normalize(elong, pi2)
        elat = ArcSin(Sin(p) * Sin(incline))
        dec = RadToDeg(ArcSin(Sin(elat) * observer.cosob + Cos(elat) * observer.sinob * Sin(elong)))
        ra = RadToDeg(ArcTan(Sin(elong) * observer.cosob - Tan(elat) * observer.sinob, Cos(elong))) / 15.0

        planet(moon).ra = ra
        planet(moon).dec = dec
        planet(moon).eclat = elat
        planet(moon).eclong = elong
        planet(moon).meananomaly = mm
        planet(moon).node = n
        planet(moon).heliolong = l           'really earth-longitude
        planet(moon).heliolat = elat

        'do here for ease
        n = (1.0 - orbit(moon).eccentric * orbit(moon).eccentric) / (1.0 + orbit(moon).eccentric * Cos(mm + ec))
        planet(moon).earthdist = n
        planet(moon).diameter = orbit(moon).diameter / n
        planet(moon).parallax = 0.9507 / n
        n = elong - sunlong

        If n < 0 Then n = n + pi2 Else If n > pi2 Then n = n - pi2

        planet(moon).phase = 0.5 * (1.0 - Cos(n))

        If n <= pi Then planet(moon).phase = -planet(moon).phase

    End Sub


    Private Sub CalcAltAz(ByRef ra As Decimal, ByRef dec As Decimal, ByRef alt As Decimal, ByRef azi As Decimal)
        ''Shared observer As ObserverType

        Dim rdec, sindec, n As Decimal

        rdec = DegToRad(dec)
        sindec = Sin(rdec)

        n = observer.lst - ra

        If n < -12.0 Then n = n + 24.0 Else If n > 12.0 Then n = n - 24.0
        n = DegToRad(n * 15.0)

        alt = ArcSin(sindec * observer.sinlat + Cos(rdec) * observer.coslat * Cos(n))
        azi = ArcCos((sindec - observer.sinlat * Sin(alt)) / (observer.coslat * Cos(alt)))

        If Sin(n) > 0.0 Then azi = pi2 - azi

        alt = RadToDeg(alt)
        azi = RadToDeg(azi)

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Private _EclipseType As String
    Public ReadOnly Property EclipseType() As String
        Get
            Return _EclipseType
        End Get
    End Property


    Private _Greenwich_Date As String
    Public ReadOnly Property Greenwich_Date() As String
        Get
            Return _Greenwich_Date
        End Get
    End Property


    Private _Moon_Position_Node As String
    Public ReadOnly Property Moon_Position_Node() As String
        Get
            Return _Moon_Position_Node
        End Get
    End Property


    Private _SunRaAtMax As String
    Public ReadOnly Property SunRaAtMax() As String
        Get
            Return _SunRaAtMax
        End Get
    End Property

    Private _SunRaAtDec As String
    Public ReadOnly Property SunRaAtDec() As String
        Get
            Return _SunRaAtDec
        End Get
    End Property

    Private _MoonRaAtMax As String
    Public ReadOnly Property MoonRaAtMax() As String
        Get
            Return _MoonRaAtMax
        End Get
    End Property

    Private _MoonRaAtDec As String
    Public ReadOnly Property MoonRaAtDec() As String
        Get
            Return _MoonRaAtDec
        End Get
    End Property

    Private _MoonNodeRAatMax As String
    Public ReadOnly Property MoonNodeRAatMax() As String
        Get
            Return _MoonNodeRAatMax
        End Get
    End Property

    Private _DecFromEclipticAtMax As String
    Public ReadOnly Property DecFromEclipticAtMax() As String
        Get
            Return _DecFromEclipticAtMax
        End Get
    End Property


    Private _LocalAltitudeAtMax As String
    Public ReadOnly Property LocalAltitudeAtMax() As String
        Get
            Return _LocalAltitudeAtMax
        End Get
    End Property


    Private _LocalAzimuthAtMax As String
    Public ReadOnly Property LocalAzimuthAtMax() As String
        Get
            Return _LocalAzimuthAtMax
        End Get
    End Property

    Private _MoonEntersPenumbra_P1_UT As String
    Public ReadOnly Property MoonEntersPenumbra_P1_UT() As String
        Get
            Return _MoonEntersPenumbra_P1_UT
        End Get
    End Property

    Private _MoonEntersPenumbra_P1_LCT As String
    Public ReadOnly Property MoonEntersPenumbra_P1_LCT() As String
        Get
            Return _MoonEntersPenumbra_P1_LCT
        End Get
    End Property

    Private _MoonEntersPenumbra_P1_PST As String
    Public ReadOnly Property MoonEntersPenumbra_P1_PST() As String
        Get
            Return _MoonEntersPenumbra_P1_PST
        End Get
    End Property

    Private _MoonEntersUmbra_U1_UT As String
    Public ReadOnly Property MoonEntersUmbra_U1_UT() As String
        Get
            Return _MoonEntersUmbra_U1_UT
        End Get
    End Property


    Private _MoonEntersUmbra_U1_LCT As String
    Public ReadOnly Property MoonEntersUmbra_U1_LCT() As String
        Get
            Return _MoonEntersUmbra_U1_LCT
        End Get
    End Property

    Private _MoonEntersUmbra_U1_PST As String
    Public ReadOnly Property MoonEntersUmbra_U1_PST() As String
        Get
            Return _MoonEntersUmbra_U1_PST
        End Get
    End Property

    Private _TotalityBegins_U2_UT As String
    Public ReadOnly Property TotalityBegins_U2_UT() As String
        Get
            Return _TotalityBegins_U2_UT
        End Get
    End Property

    Private _TotalityBegins_U2_LCT As String
    Public ReadOnly Property TotalityBegins_U2_LCT() As String
        Get
            Return _TotalityBegins_U2_LCT
        End Get
    End Property


    Private _TotalityBegins_U2_PST As String
    Public ReadOnly Property TotalityBegins_U2_PST() As String
        Get
            Return _TotalityBegins_U2_PST
        End Get
    End Property


    Private _MaxEclipse_UT As String
    Public ReadOnly Property MaxEclipse_UT() As String
        Get
            Return _MaxEclipse_UT
        End Get
    End Property


    Private _MaxEclipse_LCT As String
    Public ReadOnly Property MaxEclipse_LCT() As String
        Get
            Return _MaxEclipse_LCT
        End Get
    End Property


    Private _MaxEclipse_PST As String
    Public ReadOnly Property MaxEclipse_PST() As String
        Get
            Return _MaxEclipse_PST
        End Get
    End Property

    Private _TotalityEnds_U3_UT As String
    Public ReadOnly Property TotalityEnds_U3_UT() As String
        Get
            Return _TotalityEnds_U3_UT
        End Get
    End Property

    Private _TotalityEnds_U3_LCT As String
    Public ReadOnly Property TotalityEnds_U3_LCT() As String
        Get
            Return _TotalityEnds_U3_LCT
        End Get
    End Property

    Private _TotalityEnds_U3_PST As String
    Public ReadOnly Property TotalityEnds_U3_PST() As String
        Get
            Return _TotalityEnds_U3_PST
        End Get
    End Property

    Private _MoonLeavesUmbra_U4_UT As String
    Public ReadOnly Property MoonLeavesUmbra_U4_UT() As String
        Get
            Return _MoonLeavesUmbra_U4_UT
        End Get
    End Property


    Private _MoonLeavesUmbra_U4_LCT As String
    Public ReadOnly Property MoonLeavesUmbra_U4_LCT() As String
        Get
            Return _MoonLeavesUmbra_U4_LCT
        End Get
    End Property


    Private _MoonLeavesUmbra_U4_PST As String
    Public ReadOnly Property MoonLeavesUmbra_U4_PST() As String
        Get
            Return _MoonLeavesUmbra_U4_PST
        End Get
    End Property


    Private _MoonLeavesPenumbra_P4_UT As String
    Public ReadOnly Property MoonLeavesPenumbra_P4_UT() As String
        Get
            Return _MoonLeavesPenumbra_P4_UT
        End Get
    End Property

    Private _MoonLeavesPenumbra_P4_LCT As String
    Public ReadOnly Property MoonLeavesPenumbra_P4_LCT() As String
        Get
            Return _MoonLeavesPenumbra_P4_LCT
        End Get
    End Property

    Private _MoonLeavesPenumbra_P4_PST As String
    Public ReadOnly Property MoonLeavesPenumbra_P4_PST() As String
        Get
            Return _MoonLeavesPenumbra_P4_PST
        End Get
    End Property


    'Private _LocalAltitudeAtMax As String
    'Public ReadOnly Property LocalAltitudeAtMax() As String
    '    Get
    '        Return _LocalAltitudeAtMax
    '    End Get
    'End Property


    '''''''''''''sun
    '''''''''''''eclips

    Private _LocalAzimuthAtMax_UT As String
    Public ReadOnly Property LocalAzimuthAtMax_UT() As String
        Get
            Return _LocalAzimuthAtMax_UT
        End Get
    End Property

    Private _LocalAzimuthAtMax_LCT As String
    Public ReadOnly Property LocalAzimuthAtMax_LCT() As String
        Get
            Return _LocalAzimuthAtMax_LCT
        End Get
    End Property

    Private _LocalAzimuthAtMax_PST As String
    Public ReadOnly Property LocalAzimuthAtMax_PST() As String
        Get
            Return _LocalAzimuthAtMax_PST
        End Get
    End Property


    Private _EclipseBegins_UT As String
    Public ReadOnly Property EclipseBegins_UT() As String
        Get
            Return _EclipseBegins_UT
        End Get
    End Property

    Private _EclipseBegins_LCT As String
    Public ReadOnly Property EclipseBegins_LCT() As String
        Get
            Return _EclipseBegins_LCT
        End Get
    End Property

    Private _EclipseBegins_PST As String
    Public ReadOnly Property EclipseBegins_PST() As String
        Get
            Return _EclipseBegins_PST
        End Get
    End Property


    Private _CentralEclipseBegins_UT As String
    Public ReadOnly Property CentralEclipseBegins_UT() As String
        Get
            Return _CentralEclipseBegins_UT
        End Get
    End Property

    Private _CentralEclipseBegins_LCT As String
    Public ReadOnly Property CentralEclipseBegins_LCT() As String
        Get
            Return _CentralEclipseBegins_LCT
        End Get
    End Property


    Private _CentralEclipseBegins_PST As String
    Public ReadOnly Property CentralEclipseBegins_PST() As String
        Get
            Return _CentralEclipseBegins_PST
        End Get
    End Property



    Private _MaximumEclipse_UT As String
    Public ReadOnly Property MaximumEclipse_UT() As String
        Get
            Return _MaximumEclipse_UT
        End Get
    End Property


    Private _MaximumEclipse_LCT As String
    Public ReadOnly Property MaximumEclipse_LCT() As String
        Get
            Return _MaximumEclipse_LCT
        End Get
    End Property

    Private _MaximumEclipse_PST As String
    Public ReadOnly Property MaximumEclipse_PST() As String
        Get
            Return _MaximumEclipse_PST
        End Get
    End Property




    Private _CentralEclipseEnds_UT As String
    Public ReadOnly Property CentralEclipseEnds_UT() As String
        Get
            Return _CentralEclipseEnds_UT
        End Get
    End Property

    Private _CentralEclipseEnds_LCT As String
    Public ReadOnly Property CentralEclipseEnds_LCT() As String
        Get
            Return _CentralEclipseEnds_LCT
        End Get
    End Property

    Private _CentralEclipseEnds_PST As String
    Public ReadOnly Property CentralEclipseEnds_PST() As String
        Get
            Return _CentralEclipseEnds_PST
        End Get
    End Property



    Private _EclipseEnds_UT As String
    Public ReadOnly Property EclipseEnds_UT() As String
        Get
            Return _EclipseEnds_UT
        End Get
    End Property

    Private _EclipseEnds_LCT As String
    Public ReadOnly Property EclipseEnds_LCT() As String
        Get
            Return _EclipseEnds_LCT
        End Get
    End Property

    Private _EclipseEnds_PST As String
    Public ReadOnly Property EclipseEnds_PST() As String
        Get
            Return _EclipseEnds_PST
        End Get
    End Property

End Class