Imports Microsoft.VisualBasic

Public Class EB_Moon

    Public Class Physical_Ephemeris

        '' This page works out some values of interest to lunar-tics
        '' like me. The central formula for Moon position is approximate.
        '' Finer details like physical (as opposed to optical)
        '' libration and the nutation have been neglected. Formulas have
        '' been simplified from Meeus 'Astronomical Algorithms' (1st Ed)
        '' Chapter 51 (sub-earth and sub-solar points, PA of pole and
        '' Bright Limb, Illuminated fraction). The libration figures
        '' are usually 0.05 to 0.2 degree different from the results 
        '' given by Harry Jamieson's 'Lunar Observer's Tool Kit' DOS 
        '' program. Some of the code is adapted from a BASIC program 
        '' by George Rosenberg (ALPO).
        ''
        '' I have coded this page in a 'BASIC like' way - I intend to make 
        '' far more use of appropriately defined global objects when I 
        '' understand how they work!
        ''
        '' Written while using Netscape Gold 3.04 to keep cross-platform,
        '' Tested on Navigator Gold 2.02, Communicator 4.6, MSIE 5
        ''
        '' Round doesn't seem to work on Navigator 2 - and if you use
        '' too many nested tables, you get the text boxes not showing
        '' up in Netscape 3, and you get 'undefined' errors for formbox
        '' names on Netscape 2. Current layout seems OK.
        ''
        '' You must put all the form.name.value = variable statements
        '' together at the _end_ of the function, as the order of these
        '' statements seems to be significant.
        ''

        ''
        Private mdaynumber As Double
        Private mjulday As Double
        Private mSunDistance As Double
        Private mSunRa As Double
        Private mSunDec As Double
        Private mMoonDist As Double
        Private mMoonRa As Double
        Private mMoonDec As Double
        Private mSelLatEarth As Double
        Private mSelLongEarth As Double
        Private mSelLatSun As Double
        Private mSelLongSun As Double
        Private mSelColongSun As Double
        Private mSelLongTerm As Double
        Private mSelIlum As Double
        Private mSelPaBl As Double
        Private mSelPaPole As Double

        Public Property daynumber() As Double
            Get
                Return mdaynumber
            End Get
            Set(ByVal Value As Double)
                mdaynumber = Value
            End Set
        End Property

        Public Property julday() As Double
            Get
                Return mjulday
            End Get
            Set(ByVal Value As Double)
                mjulday = Value
            End Set
        End Property

        Public Property SunDistance() As Double
            Get
                Return mSunDistance
            End Get
            Set(ByVal Value As Double)
                mSunDistance = Value
            End Set
        End Property

        Public Property SunRa() As Double
            Get
                Return mSunRa
            End Get
            Set(ByVal Value As Double)
                mSunRa = Value
            End Set
        End Property


        Public Property SunDec() As Double
            Get
                Return mSunDec
            End Get
            Set(ByVal Value As Double)
                mSunDec = Value
            End Set
        End Property

        Public Property MoonDist() As Double
            Get
                Return mMoonDist
            End Get
            Set(ByVal Value As Double)
                mMoonDist = Value
            End Set
        End Property

        Public Property MoonRa() As Double
            Get
                Return mMoonRa
            End Get
            Set(ByVal Value As Double)
                mMoonRa = Value
            End Set
        End Property

        Public Property MoonDec() As Double
            Get
                Return mMoonDec
            End Get
            Set(ByVal Value As Double)
                mMoonDec = Value
            End Set
        End Property

        Public Property SelLatEarth() As Double
            Get
                Return mSelLatEarth
            End Get
            Set(ByVal Value As Double)
                mSelLatEarth = Value
            End Set
        End Property

        Public Property SelLongEarth() As Double
            Get
                Return mSelLongEarth
            End Get
            Set(ByVal Value As Double)
                mSelLongEarth = Value
            End Set
        End Property

        Public Property SelLatSun() As Double
            Get
                Return mSelLatSun
            End Get
            Set(ByVal Value As Double)
                mSelLatSun = Value
            End Set
        End Property

        Private Property SelLongSun() As Double
            Get
                Return mSelLongSun
            End Get
            Set(ByVal Value As Double)
                mSelLongSun = Value
            End Set
        End Property

        Public Property SelColongSun() As Double
            Get
                Return mSelColongSun
            End Get
            Set(ByVal Value As Double)
                mSelColongSun = Value
            End Set
        End Property

        Public Property SelLongTerm() As Double
            Get
                Return mSelLongTerm
            End Get
            Set(ByVal Value As Double)
                mSelLongTerm = Value
            End Set
        End Property

        Public Property SelIlum() As Double
            Get
                Return mSelIlum
            End Get
            Set(ByVal Value As Double)
                mSelIlum = Value
            End Set
        End Property


        Public Property SelPaBl() As Double
            Get
                Return mSelPaBl
            End Get
            Set(ByVal Value As Double)
                mSelPaBl = Value
            End Set
        End Property


        Public Property SelPaPole() As Double
            Get
                Return mSelPaPole
            End Get
            Set(ByVal Value As Double)
                mSelPaPole = Value
            End Set
        End Property

        Public Sub setUTime(ByVal iYear As Short, ByVal iMonth As Short, _
                                 ByVal iDay As Short, Optional ByVal iHour As Short = 0, _
                                 Optional ByVal iMinute As Short = 0)

            Dim s, sm, sd, smi As String

            If Len(iMonth.ToString) = 1 Then
                sm = "0" & iMonth.ToString
            Else
                sm = iMonth.ToString
            End If


            If Len(iDay.ToString) = 1 Then
                sd = "0" & iDay.ToString
            Else
                sd = iDay.ToString
            End If

            If Len(iMinute.ToString) = 1 Then
                smi = "0" & iMinute.ToString
            Else
                smi = iMinute.ToString
            End If


            s = iYear.ToString & sm & sd & "." & iHour.ToString & smi

            doCalcs(CType(s, Double))

        End Sub


        Private Sub doCalcs(ByVal num As Double)

            Dim g, days, t, L1, M1, C1, V1, Ec1, R1, Th1, Om1, Lam1, Obl, Ra1, Dec1 As Double
            Dim F, L2, Om2, M2, DD, R2, R3, Bm, Lm, HLm, HBm, Ra2, Dec2, EL, EB, W, XX, YY, A As Double
            Dim Co, SLt, Psi, Il, K, P1, P2, y, m, d, bit, h, min, bk As Double
            Dim D2, I, SL, SB As Double
            Dim alert As String


            ''	Get date and time code from user, isolate the year, month, day and hours
            ''	and minutes, and do some basic error checking! This only works for AD years

            g = num
            y = Math.Floor(g / 10000)
            m = Math.Floor((g - y * 10000) / 100)
            d = Math.Floor(g - y * 10000 - m * 100)
            bit = (g - Math.Floor(g)) * 100
            h = Math.Floor(bit)
            min = Math.Floor(bit * 100 - h * 100 + 0.5)

            ''primative error checking - accounting for right number of
            ''days per month including leap years. Using bk variable to 
            ''prevent multiple alerts. See functions isleap(y) 
            ''and goodmonthday(y, m, d).

            bk = 0
            If (g < 16000000) Then
                bk = 1
                alert = "Routines are not accurate enough to work back that" + " far - answers are meaningless!"
                Throw New Exception(alert)
            End If
            If (g > 23000000) Then
                bk = 1
                alert = "Routines are not accurate enough to work far into the future" + " - answers are meaningless!"
                Throw New Exception(alert)
            End If
            If (((m < 1) OrElse (m > 12)) AndAlso (bk <> 1)) Then
                bk = 1
                alert = "Months are not right - type date again"
                Throw New Exception(alert)
            End If
            If ((goodmonthday(y, m, d) = 0) AndAlso (bk <> 1)) Then
                bk = 1
                alert = "Wrong number of days for the month or not a leap year - type date again"
                Throw New Exception(alert)
            End If
            If ((h > 23) AndAlso (bk <> 1)) Then
                bk = 1
                alert = "Hours are not right - type date again"
                Throw New Exception(alert)
            End If
            If ((min > 59) AndAlso (bk <> 1)) Then
                alert = "Minutes are not right - type date again"
                Throw New Exception(alert)
            End If



            ''	Get the number of days since J2000.0 using day2000() function
            days = day2000(y, m, d, h + min / 60)
            t = days / 36525


            ''	Sun formulas

            ''	L1	- Mean longitude
            ''	M1	- Mean anomaly
            ''	C1	- Equation of centre
            ''	V1	- True anomaly
            ''	Ec1	- Eccentricity 
            ''	R1	- Sun distance
            ''	Th1	- Theta (true longitude)
            ''	Om1	- Long Asc Node (Omega)
            '' 	Lam1- Lambda (apparent longitude)
            ''	Obl	- Obliquity of ecliptic
            ''	Ra1	- Right Ascension
            ''	Dec1- Declination
            ''

            L1 = range(280.466 + 36000.8 * t)
            M1 = range(357.529 + 35999 * t - 0.0001536 * t * t + t * t * t / 24490000)
            C1 = (1.915 - 0.004817 * t - 0.000014 * t * t) * dsin(M1)
            C1 = C1 + (0.01999 - 0.000101 * t) * dsin(2 * M1)
            C1 = C1 + 0.00029 * dsin(3 * M1)
            V1 = M1 + C1
            Ec1 = 0.01671 - 0.00004204 * t - 0.0000001236 * t * t
            R1 = 0.99972 / (1 + Ec1 * dcos(V1))
            Th1 = L1 + C1
            Om1 = range(125.04 - 1934.1 * t)
            Lam1 = Th1 - 0.00569 - 0.00478 * dsin(Om1)
            Obl = (84381.448 - 46.815 * t) / 3600
            Ra1 = datan2(dsin(Th1) * dcos(Obl) - dtan(0) * dsin(Obl), dcos(Th1))
            Dec1 = dasin(dsin(0) * dcos(Obl) + dcos(0) * dsin(Obl) * dsin(Th1))

            ''	Moon formulas
            ''
            ''	F 	- Argument of latitude (F)
            ''	L2 	- Mean longitude (L')
            ''	Om2 - Long. Asc. Node (Om')
            ''	M2	- Mean anomaly (M')
            ''	D	- Mean elongation (D)
            ''  D2(-2 * D)
            ''	R2	- Lunar distance (Earth - Moon distance)
            ''	R3	- Distance ratio (Sun / Moon)
            ''	Bm	- Geocentric Latitude of Moon
            ''	Lm	- Geocentric Longitude of Moon
            ''	HLm	- Heliocentric longitude
            ''	HBm	- Heliocentric latitude
            ''	Ra2	- Lunar Right Ascension
            ''  Dec2(-Declination)



            F = range(93.2721 + 483202 * t - 0.003403 * t * t - t * t * t / 3526000)
            L2 = range(218.316 + 481268 * t)
            Om2 = range(125.045 - 1934.14 * t + 0.002071 * t * t + t * t * t / 450000)
            M2 = range(134.963 + 477199 * t + 0.008997 * t * t + t * t * t / 69700)
            DD = range(297.85 + 445267 * t - 0.00163 * t * t + t * t * t / 545900)
            D2 = 2 * DD
            R2 = 1 + (-20954 * dcos(M2) - 3699 * dcos(D2 - M2) - 2956 * dcos(D2)) / 385000
            R3 = (R2 / R1) / 379.168831168831
            Bm = 5.128 * dsin(F) + 0.2806 * dsin(M2 + F)
            Bm = Bm + 0.2777 * dsin(M2 - F) + 0.1732 * dsin(D2 - F)
            Lm = 6.289 * dsin(M2) + 1.274 * dsin(D2 - M2) + 0.6583 * dsin(D2)
            Lm = Lm + 0.2136 * dsin(2 * M2) - 0.1851 * dsin(M1) - 0.1143 * dsin(2 * F)
            Lm = Lm + 0.0588 * dsin(D2 - 2 * M2)
            Lm = Lm + 0.0572 * dsin(D2 - M1 - M2) + 0.0533 * dsin(D2 + M2)
            Lm = Lm + L2
            Ra2 = datan2(dsin(Lm) * dcos(Obl) - dtan(Bm) * dsin(Obl), dcos(Lm))
            Dec2 = dasin(dsin(Bm) * dcos(Obl) + dcos(Bm) * dsin(Obl) * dsin(Lm))
            HLm = range(Lam1 + 180 + (180 / Math.PI) * R3 * dcos(Bm) * dsin(Lam1 - Lm))
            HBm = R3 * Bm



            ''Selenographic coords of the sub Earth point
            ''This gives you the (geocentric) libration 
            ''approximating to that listed in most almanacs
            ''Topocentric libration can be up to a degree
            ''different either way

            ''Physical libration ignored, as is nutation.

            ''I	- Inclination of (mean) lunar orbit to ecliptic
            ''EL	- Selenographic longitude of sub Earth point
            ''EB	- Sel Lat of sub Earth point
            ''W	- angle variable
            ''XX	- Rectangular coordinate
            ''YY	- Rectangular coordinate
            ''A	- Angle variable (see Meeus ch 51 for notation)



            I = 1.54242
            W = Lm - Om2
            YY = dcos(W) * dcos(Bm)
            XX = dsin(W) * dcos(Bm) * dcos(I) - dsin(Bm) * dsin(I)
            A = datan2(XX, YY)
            EL = A - F
            EB = dasin(-dsin(W) * dcos(Bm) * dsin(I) - dsin(Bm) * dcos(I))


            ''Selenographic coords of sub-solar point. This point is
            ''the() 'pole' of the illuminated hemisphere of the Moon
            ''and so describes the position of the terminator on the 
            ''lunar surface. The information is communicated through
            ''numbers like the colongitude, and the longitude of the
            ''terminator.

            ''SL	- Sel Long of sub-solar point
            ''SB	- Sel Lat of sub-solar point
            ''W, YY, XX, A	- temporary variables as for sub-Earth point
            ''Co	- Colongitude of the Sun
            ''SLt	- Selenographic longitude of terminator 
            ''riset - Lunar sunrise or set


            W = range(HLm - Om2)
            YY = dcos(W) * dcos(HBm)
            XX = dsin(W) * dcos(HBm) * dcos(I) - dsin(HBm) * dsin(I)
            A = datan2(XX, YY)
            SL = range(A - F)
            SB = dasin(-dsin(W) * dcos(HBm) * dsin(I) - dsin(HBm) * dcos(I))



            If (SL < 90) Then
                Co = 90 - SL
            Else
                Co = 450 - SL
            End If


            If ((Co > 90) AndAlso (Co < 270)) Then
                SLt = 180 - Co
            Else
                If (Co < 90) Then
                    SLt = 0 - Co
                Else
                    SLt = 360 - Co
                End If
            End If




            ''Calculate the illuminated fraction, the position angle of the bright
            ''limb, and the position angle of the Moon's rotation axis. All position
            ''angles relate to the North Celestial Pole - you need to work out the
            ''       'Parallactic angle' to calculate the orientation to your local zenith.

            ''	Iluminated fraction
            A = dcos(Bm) * dcos(Lm - Lam1)
            Psi = 90 - datan(A / Math.Sqrt(1 - A * A))
            XX = R1 * dsin(Psi)
            YY = R3 - R1 * A
            Il = datan2(XX, YY)
            K = (1 + dcos(Il)) / 2

            ''	PA bright limb
            XX = dsin(Dec1) * dcos(Dec2) - dcos(Dec1) * dsin(Dec2) * dcos(Ra1 - Ra2)
            YY = dcos(Dec1) * dsin(Ra1 - Ra2)
            P1 = datan2(YY, XX)

            ''	PA Moon's rotation axis
            ''	Neglects nutation and physical libration, so Meeus' angle
            ''	V is just Om2
            XX = dsin(I) * dsin(Om2)
            YY = dsin(I) * dcos(Om2) * dcos(Obl) - dcos(I) * dsin(Obl)
            W = datan2(XX, YY)
            A = Math.Sqrt(XX * XX + YY * YY) * dcos(Ra2 - W)
            P2 = dasin(A / dcos(EB))


            ''	Write Sun numbers to form
            Me.daynumber = round(days, 4)
            Me.julday = round(days + 2451545.0, 4)
            Me.SunDistance = round(R1, 4)

            Me.SunRa = round(Ra1 / 15, 3)
            Me.SunDec = round(Dec1, 2)


            ''	Write Moon numbers to form
            Me.MoonDist = round(R2 * 60.268511, 2)

            Me.MoonRa = round(Ra2 / 15, 3)
            Me.MoonDec = round(Dec2, 2)


            ''	Print the libration numbers
            Me.SelLatEarth = round(EB, 1)
            Me.SelLongEarth = round(EL, 1)


            ''	Print the Sub-solar numbers
            Me.SelLatSun = round(SB, 1)
            Me.SelLongSun = round(SL, 1)

            Me.SelColongSun = round(Co, 2)
            Me.SelLongTerm = round(SLt, 1)

            ''	Print the rest - position angles and illuminated fraction
            Me.SelIlum = round(K, 3)
            Me.SelPaBl = round(P1, 1)
            Me.SelPaPole = round(P2, 1)

        End Sub





        '' this is the usual days since J2000 function
        Private Function day2000(ByVal y As Double, ByVal m As Double, ByVal d As Double, _
                                 ByVal h As Double) As Double
            Dim d1, b, c, greg, a As Double
            greg = y * 10000 + m * 100 + d

            If (m = 1 OrElse m = 2) Then
                y = y - 1
                m = m + 12
            End If

            ''  reverts to Julian calendar before 4th Oct 1582
            ''  no good for UK, America or Sweeden!

            If (greg > 15821004) Then
                a = Math.Floor(y / 100)
                b = 2 - a + Math.Floor(a / 4)
            Else
                b = 0
            End If

            c = Math.Floor(365.25 * y)
            d1 = Math.Floor(30.6001 * (m + 1))

            Return (b + c + d1 - 730550.5 + d + h / 24)
        End Function


        ''	Leap year detecting function (gregorian calendar)
        ''  returns 1 for leap year and 0 for non-leap year
        Private Function isleap(ByVal y As Double) As Double
            Dim a As Double
            ''	assume not a leap year...
            a = 0
            ''	...flag leap year candidates...

            If (y Mod 4 = 0) Then
                a = 1
            End If


            ''	...if year is a century year then not leap...
            If (y Mod 100 = 0) Then
                a = 0
            End If

            ''	...except if century year divisible by 400...
            If (y Mod 400 = 0) Then
                a = 1
            End If

            ''	...and so done according to Gregory's wishes 
            Return a
        End Function



        ''Month and day number checking function
        ''This will work OK for Julian or Gregorian
        ''providing isleap() is defined appropriately
        ''Returns 1 if Month and Day combination OK,
        ''and 0 if month and day combination impossible

        Private Function goodmonthday(ByVal y As Double, ByVal m As Double, ByVal d As Double)
            Dim a, leap As Double
            leap = isleap(y)
            ''	assume OK
            a = 1
            ''	first deal with zero day number!
            If (d = 0) Then
                a = 0
            End If
            ''	Sort Feburary next
            If ((m = 2) AndAlso (leap = 1) AndAlso (d > 29)) Then
                a = 0
            End If


            If ((m = 2) AndAlso (d > 28) AndAlso (leap = 0)) Then
                a = 0
            End If

            ''	then the rest of the months - 30 days...

            If (((m = 4) OrElse (m = 6) OrElse (m = 9) OrElse (m = 11)) AndAlso d > 30) Then
                a = 0
            End If

            ''	...31 days...	
            If (d > 31) Then
                a = 0
            End If

            ''	...and so done
            Return a
        End Function

        ''Trigonometric functions working in degrees - this just
        ''makes implementing the formulas in books easier at the
        ''cost of some wasted multiplications.
        ''The 'range' function brings angles into range 0 to 360,
        ''and an atan2(x,y) function returns arctan in correct
        ''quadrant. ipart(x) returns smallest integer nearest zero


        Private Function dsin(ByVal x As Double)
            Return Math.Sin(Math.PI / 180 * x)
        End Function

        Private Function dcos(ByVal x As Double)
            Return Math.Cos(Math.PI / 180 * x)
        End Function

        Private Function dtan(ByVal x As Double)
            Return Math.Tan(Math.PI / 180 * x)
        End Function

        Private Function dasin(ByVal x As Double)
            Return 180 / Math.PI * Math.Asin(x)
        End Function

        Private Function dacos(ByVal x As Double)
            Return 180 / Math.PI * Math.Acos(x)
        End Function

        Private Function datan(ByVal x As Double)
            Return 180 / Math.PI * Math.Atan(x)
        End Function

        Private Function datan2(ByVal y As Double, ByVal x As Double)
            Dim a As Double

            If ((x = 0) AndAlso (y = 0)) Then
                Return 0
            Else
                a = datan(y / x)
                If (x < 0) Then
                    a = a + 180
                End If
                If (y < 0 AndAlso x > 0) Then
                    a = a + 360
                End If
                Return a
            End If
        End Function

        Private Function ipart(ByVal x As Double)
            Dim a As Double
            If (x > 0) Then
                a = Math.Floor(x)
            Else
                a = Math.Ceiling(x)
            End If

            Return a
        End Function

        Private Function range(ByVal x As Double)
            Dim a, b As Double
            b = x / 360
            a = 360 * (b - ipart(b))
            If (a < 0) Then
                a = a + 360
            End If
            Return a
        End Function


        ''round rounds the number num to dp decimal places
        ''the second line is some C like jiggery pokery I
        ''found in an O'Reilly book which means if dp is null
        ''you get 2 decimal places.

        Private Function round(ByVal num As Double, ByVal dp As Double)
            ''   dp = (!dp ? 2: dp);
            Return Math.Round(num * Math.Pow(10, dp)) / Math.Pow(10, dp)
        End Function


    End Class



    '''


    Public Class Phases

        Public Enum EB_MoonLight
            EB_Moon_New = 1
            EB_Moon_WaxingCrescent = 2
            EB_Moon_FirstQuarter = 3
            EB_Moon_WaxingGibbous = 4
            EB_Moon_Full = 5
            EB_Moon_WaningGibbous = 6
            EB_Moon_LastQuarter = 7
            EB_Moon_MorningCrescent = 8
        End Enum

        Public Enum EB_Moon_Zodiac
            EB_Pisces = 1
            EB_Aries = 2
            EB_Taurus = 3
            EB_Gemini = 4
            EB_Cancer = 5
            EB_Leo = 6
            EB_Virgo = 7
            EB_Libra = 8
            EB_Scorpio = 9
            EB_Sagittarius = 10
            EB_Capricorn = 11
            EB_Aquarius = 12
        End Enum


        Private n0 As Integer = 0
        Private f0 As Double = 0
        Private AG As Double = f0    '// Moon's age
        Private DI As Double = f0    '// Moon's distance in earth radii
        Private LA As Double = f0    '// Moon's ecliptic latitude
        Private LO As Double = f0    '// Moon's ecliptic longitude
        Private TZ As Double = 12
        Private Phase As EB_MoonLight
        Private Zodiac As EB_Moon_Zodiac



        Public ReadOnly Property Moon_EclipticLongitude() As Double
            Get
                Return LO
            End Get
        End Property

        Public ReadOnly Property Moon_EclipticLatitude() As Double
            Get
                Return LA
            End Get
        End Property

        Public ReadOnly Property Moon_DistanceInEarthRadii() As Double
            Get
                Return DI
            End Get
        End Property
        Public ReadOnly Property Moon_AgeFromNew() As Double
            Get
                Return AG
            End Get
        End Property



        Public Sub calculate_phase(ByVal iYear As Double, ByVal iMonth As Double, _
                                        ByVal iDay As Double, ByVal iHour As Double, ByVal iMinute As Double, _
                                        ByVal iSec As Double, ByVal TimeZone As Double)
            Moon_positition(iYear, iMonth, iDay, iHour, iMinute, iSec, TimeZone)
            ''''''''''''''''''''''''''''''''''''''''''''''''
        End Sub

        Public ReadOnly Property Moon_Phase() As EB_MoonLight
            Get
                Return Phase
            End Get
        End Property

        Public ReadOnly Property Moon_Zodiac() As EB_Moon_Zodiac
            Get
                Return Zodiac
            End Get
        End Property


        Private Sub Moon_positition(ByVal Y As Integer, ByVal M As Integer, ByVal D As Integer, _
                                    ByVal hour As Integer, ByVal min As Integer, ByVal sec As Integer, ByVal zt6 As Double)

            Dim YY As Integer = n0
            Dim MM As Integer = n0
            Dim K1 As Integer = n0
            Dim K2 As Integer = n0
            Dim K3 = n0
            Dim JD = n0
            Dim IP = f0
            Dim DP = f0
            Dim NP = f0
            Dim RP = f0

            Dim im As Double
            im = min - (60 * zt6) '+ (DayLightSaveingTime * 60)

            '// calculate the Julian date at 12h UT
            'YY = Y - Math.Floor((12 - M) / 10)
            'MM = M + 9
            'If MM >= 12 Then MM = MM - 12


            'K1 = Math.Floor(365.25 * (YY + 4712))
            'K2 = Math.Floor(30.6 * MM + 0.5)
            'K3 = Math.Floor(Math.Floor((YY / 100) + 49) * 0.75) - 38

            ''Dim Hour As Integer
            ''Dim min As Integer
            ''Dim sec As Integer
            Dim extra As Double = 100.0 * Y + M - 190002.5
            Dim rjd As Double = 367.0 * Y

            rjd -= Math.Floor(7.0 * (Y + Math.Floor((M + 9.0) / 12.0)) / 4.0)
            rjd += Math.Floor(275.0 * M / 9.0)
            rjd += D
            rjd += (hour + (im + sec / 60) / 60) / 24
            rjd += 1721013.5
            rjd -= 0.5 * extra / Math.Abs(extra)
            rjd += 0.5
            JD = rjd
            'JD = K1 + K2 + D + 59                  '// for dates in Julian calendar
            'If JD > 2299160 Then
            'JD = JD - K3        '// for Gregorian calendar
            'End If
            '// calculate moon's age in days

            Dim v As Double
            ' Calculate illumination (synodic) phase
            v = (JD - 2451550.1) / 29.530588853
            v = v - CInt(v)
            If v < 0 Then
                v = v + 1
            End If
            IP = v

            ' Moon's age in days
            AG = IP * 29.53


            ' Convert phase to radians
            IP = IP * (Math.PI * 2)

            ' Calculate distance from anomalistic phase
            v = (JD - 2451562.2) / 27.55454988
            v = v - CInt(v)
            If v < 0 Then
                v = v + 1
            End If
            DP = v
            DP = DP * (Math.PI * 2) ' Convert to radians
            DI = 60.4 - 3.3 * Math.Cos(DP) - 0.6 * Math.Cos(2 * IP - DP) - 0.5 * Math.Cos(2 * IP)



            ' Calculate latitude from nodal (draconic) phase
            v = (JD - 2451565.2) / 27.212220817
            v = v - Int(v)
            If v < 0 Then
                v = v + 1
            End If
            NP = v


            ' Convert to radians
            NP = NP * (Math.PI * 2)
            LA = 5.1 * Math.Sin(NP)

            ' Calculate longitude from sidereal motion
            v = (JD - 2451555.8) / 27.321582241
            ' Normalize values to range 0 to 1
            v = v - CInt(v)
            If v < 0 Then
                v = v + 1
            End If

            RP = v
            LO = 360 * RP + 6.3 * Math.Sin(DP) + 1.3 * Math.Sin(2 * IP - DP) + 0.7 * Math.Sin(2 * IP)



            '''

            'AG 
            Select Case Math.Floor(AG)
                Case 0, 29
                    Phase = EB_MoonLight.EB_Moon_New  '"NEW"
                Case 1 To 6
                    Phase = EB_MoonLight.EB_Moon_WaxingCrescent
                Case 7
                    Phase = EB_MoonLight.EB_Moon_FirstQuarter
                Case 8, 9, 10, 11, 12, 13
                    Phase = EB_MoonLight.EB_Moon_WaxingGibbous
                Case 14
                    Phase = EB_MoonLight.EB_Moon_Full  '"FULL"
                Case 15 To 21
                    Phase = EB_MoonLight.EB_Moon_WaningGibbous  '"Waning gibbous"
                Case 22
                    Phase = EB_MoonLight.EB_Moon_LastQuarter '"Last quarter"
                Case 23 To 28
                    Phase = EB_MoonLight.EB_Moon_MorningCrescent '"Morning crescent"
            End Select




            If LO < 33.18 Then
                Zodiac = EB_Moon_Zodiac.EB_Pisces   '"Pisces"
            ElseIf LO < 51.16 Then
                Zodiac = EB_Moon_Zodiac.EB_Aries  '"Aries"
            ElseIf LO < 93.44 Then
                Zodiac = EB_Moon_Zodiac.EB_Taurus  '"Taurus"
            ElseIf LO < 119.48 Then
                Zodiac = EB_Moon_Zodiac.EB_Gemini  '"Gemini"
            ElseIf LO < 135.3 Then
                Zodiac = EB_Moon_Zodiac.EB_Cancer  '"Cancer"
            ElseIf LO < 173.34 Then
                Zodiac = EB_Moon_Zodiac.EB_Leo  '"Leo"
            ElseIf LO < 224.17 Then
                Zodiac = EB_Moon_Zodiac.EB_Virgo  '"Virgo"
            ElseIf LO < 242.57 Then
                Zodiac = EB_Moon_Zodiac.EB_Libra  '"Libra"
            ElseIf LO < 271.26 Then
                Zodiac = EB_Moon_Zodiac.EB_Scorpio  '"Scorpio"
            ElseIf LO < 302.49 Then
                Zodiac = EB_Moon_Zodiac.EB_Sagittarius  '"Sagittarius"
            ElseIf LO < 311.72 Then
                Zodiac = EB_Moon_Zodiac.EB_Capricorn  '"Capricorn"
            ElseIf LO < 348.58 Then
                Zodiac = EB_Moon_Zodiac.EB_Aquarius  '"Aquarius"
            Else
                Zodiac = EB_Moon_Zodiac.EB_Pisces  '"Pisces"
            End If

            '// so longitude is not greater than 360!
            If (LO > 360) Then
                LO = LO - 360
            End If


        End Sub


        Private Function normalize(ByVal v) As Double

            v = v - Math.Floor(v)
            If (v < 0) Then
                v = v + 1
            End If
            Return v


        End Function


    End Class







    Public Class Rise_Set

        Private bol_below_horizon_all_day As Boolean = False
        Private bol_above_horizon_all_day As Boolean = False
        Private bol_norise As Boolean = False
        Private bol_noset As Boolean = False
        Private strUpTime As String
        Private strDownTime As String


        Public Sub Find_Rise_Set(ByVal iYear As Integer, ByVal iMonth As Integer, ByVal iDay As Integer, _
                                        ByVal TimeZone As Double, ByVal longitude As Double _
                                      , ByVal latitude As Double)

            bol_above_horizon_all_day = False
            bol_above_horizon_all_day = False
            bol_norise = False
            bol_noset = False
            strUpTime = String.Empty
            strDownTime = String.Empty


            find_moonrise_set(mjd(iDay, iMonth, iYear, 0), TimeZone, longitude, latitude)

        End Sub
        ''''''''''''''''''''''''''

        Public ReadOnly Property UpTime() As String
            Get
                Return strUpTime
            End Get
        End Property

        Public ReadOnly Property DownTime() As String
            Get
                Return strDownTime
            End Get
        End Property

        Public ReadOnly Property IsBelowHorizonAllAay() As Boolean
            Get
                Return bol_below_horizon_all_day
            End Get
        End Property

        Public ReadOnly Property IsAboveHorizonAllAay() As Boolean
            Get
                Return bol_above_horizon_all_day
            End Get
        End Property

        Public ReadOnly Property HasNoRise() As Boolean
            Get
                Return bol_norise
            End Get
        End Property

        Public ReadOnly Property HasNoSet() As Boolean
            Get
                Return bol_noset
            End Get
        End Property


        '


        Private Function mjd(ByVal day As Double, ByVal month As Double, _
                                ByVal year As Double, ByVal hour As Double) As Double
            '//
            '//	Takes the day, month, year and hours in the day and returns the
            '//  modified julian day number defined as mjd = jd - 2400000.5
            '//  checked OK for Greg era dates - 26th Dec 02
            '//
            Dim a, b As Double
            If (month <= 2) Then
                month = month + 12
                year = year - 1
            End If

            a = 10000.0 * year + 100.0 * month + day
            If (a <= 15821004.1) Then
                b = -2 * Math.Floor((year + 4716) / 4) - 1179

            Else
                b = Math.Floor(year / 400) - Math.Floor(year / 100) + Math.Floor(year / 4)
            End If
            a = 365.0 * year - 679004.0
            Return (a + b + Math.Floor(30.6001 * (month + 1)) + day + hour / 24.0)

        End Function




        Private Sub find_moonrise_set(ByVal mjd As Double, ByVal tz As Double, ByVal glong As Double _
                                            , ByVal glat As Double)

            '//	Im using a separate function for moonrise/set to allow for different tabulations
            '//  of moonrise and sun events ie weekly for sun and daily for moon. The logic of
            '//  the function is identical to find_sun_and_twi_events_for_date_()
            '//
            'KeyString = "\nKey\n.... means sun or moon below horizon all day or\n     twilight never begins\n" +
            '    "**** means sun or moon above horizon all day\n     or twilight never ends\n" +
            '    "---- in rise column means no rise that day and\n" +
            '    "     in set column - no set that day\n";


            Dim sglong, sglat, cglat, date_, ym, yz, utrise, utset, j As Double
            Dim yp, nz, hour, z1, z2, iobj, xe, ye As Double
            Dim rads As Double = 0.0174532925
            Dim rise, sett, above, stt As Boolean
            Dim quadout(5) As Double
            Dim sinho As Double
            Dim always_up As String = " ****"
            Dim always_down As String = " ...."
            Dim outstring As String = ""

            sinho = Math.Sin(rads * 8 / 60) ';		//moonrise taken as centre of moon at +8 arcmin
            sglat = Math.Sin(rads * glat)
            cglat = Math.Cos(rads * glat)
            date_ = mjd - tz / 24
            rise = False
            sett = False
            above = False
            hour = 1.0
            ym = sin_alt(1, date_, hour - 1.0, glong, cglat, sglat) - sinho

            If (ym > 0.0) Then above = True

            While (hour < 25 And (sett = False Or rise = False))
                yz = sin_alt(1, date_, hour, glong, cglat, sglat) - sinho
                yp = sin_alt(1, date_, hour + 1.0, glong, cglat, sglat) - sinho
                quadout = quad(ym, yz, yp)
                nz = quadout(0)
                z1 = quadout(1)
                z2 = quadout(2)
                xe = quadout(3)
                ye = quadout(4)

                ''// case when one event is found in the interval
                If (nz = 1) Then
                    If (ym < 0.0) Then
                        utrise = hour + z1
                        rise = True
                    Else
                        utset = hour + z1
                        sett = True
                    End If

                End If

                '} // end of nz = 1 case

                '// case where two events are found in this interval
                '// (rare but whole reason we are not using simple iteration)
                If (nz = 2) Then
                    If (ye < 0.0) Then
                        utrise = hour + z2
                        utset = hour + z1
                    Else
                        utrise = hour + z1
                        utset = hour + z2
                    End If

                End If


                '// set up the next search interval
                ym = yp
                hour += 1.0
                'hour += 2.0

            End While '(hour < 25 And (sett = False Or rise = False))  '} // end of while loop



            If (rise = True Or sett = True) Then
                If (rise = True) Then
                    outstring += " " + (hrsmin(utrise)).ToString
                    strUpTime = toNormalTime(hrsmin(utrise))
                    If strUpTime.Length > 0 Then
                        bol_norise = False
                    Else
                        bol_norise = True
                    End If

                Else
                    outstring += " ----"
                    bol_norise = True
                End If

                If (sett = True) Then
                    outstring += " " + (hrsmin(utset)).ToString
                    strDownTime = toNormalTime(hrsmin(utset))
                    If strDownTime.Length > 0 Then
                        bol_noset = False
                    Else
                        bol_noset = True
                    End If

                Else
                    outstring += " ----"
                    bol_noset = True
                End If

            Else
                If (above = True) Then
                    outstring += always_up + always_up
                    bol_above_horizon_all_day = True
                Else
                    outstring += always_down + always_down
                    bol_below_horizon_all_day = True
                End If
            End If

            'Return outstring
        End Sub




        Private Function minimoon(ByVal t As Double) As Double()

            '// takes t and returns the geocentric ra and dec in an array mooneq
            '// claimed good to 5' (angle) in ra and 1' in dec
            '// tallies with another approximate method and with ICE for a couple of dates
            '//
            Dim p2 As Double = 6.283185307
            Dim arc As Double = 206264.8062
            Dim coseps As Double = 0.91748
            Dim sineps As Double = 0.39778
            Dim L0, L, LS, F, D, H, S, N, DL, CB, L_moon, B_moon, V, W, X, Y, Z, RHO, Dec, Ra As Double
            Dim mooneq(5) As Double      ' = New Array)

            L0 = frac(0.606433 + 1336.855225 * t) '// mean longitude of moon
            L = p2 * frac(0.374897 + 1325.55241 * t)  '//mean anomaly of Moon
            LS = p2 * frac(0.993133 + 99.997361 * t) '//mean anomaly of Sun
            D = p2 * frac(0.827361 + 1236.853086 * t) '//difference in longitude of moon and sun
            F = p2 * frac(0.259086 + 1342.227825 * t) '//mean argument of latitude

            '// corrections to mean longitude in arcsec
            DL = 22640 * Math.Sin(L)
            DL += -4586 * Math.Sin(L - 2 * D)
            DL += +2370 * Math.Sin(2 * D)
            DL += +769 * Math.Sin(2 * L)
            DL += -668 * Math.Sin(LS)
            DL += -412 * Math.Sin(2 * F)
            DL += -212 * Math.Sin(2 * L - 2 * D)
            DL += -206 * Math.Sin(L + LS - 2 * D)
            DL += +192 * Math.Sin(L + 2 * D)
            DL += -165 * Math.Sin(LS - 2 * D)
            DL += -125 * Math.Sin(D)
            DL += -110 * Math.Sin(L + LS)
            DL += +148 * Math.Sin(L - LS)
            DL += -55 * Math.Sin(2 * F - 2 * D)

            '// simplified form of the latitude terms
            S = F + (DL + 412 * Math.Sin(2 * F) + 541 * Math.Sin(LS)) / arc
            H = F - 2 * D
            N = -526 * Math.Sin(H)
            N += +44 * Math.Sin(L + H)
            N += -31 * Math.Sin(-L + H)
            N += -23 * Math.Sin(LS + H)
            N += +11 * Math.Sin(-LS + H)
            N += -25 * Math.Sin(-2 * L + F)
            N += +21 * Math.Sin(-L + F)

            '	// ecliptic long and lat of Moon in rads
            L_moon = p2 * frac(L0 + DL / 1296000)
            B_moon = (18520.0 * Math.Sin(S) + N) / arc

            '// equatorial coord conversion - note fixed obliquity
            CB = Math.Cos(B_moon)
            X = CB * Math.Cos(L_moon)
            V = CB * Math.Sin(L_moon)
            W = Math.Sin(B_moon)
            Y = coseps * V - sineps * W
            Z = sineps * V + coseps * W
            RHO = Math.Sqrt(1.0 - Z * Z)
            Dec = (360.0 / p2) * Math.Atan(Z / RHO)
            Ra = (48.0 / p2) * Math.Atan(Y / (X + RHO))
            If (Ra < 0) Then Ra += 24

            mooneq(1) = Dec
            mooneq(2) = Ra

            Return mooneq
        End Function

        Private Function frac(ByVal x As Double) As Double
            '//
            '//	returns the fractional part of x as used in minimoon and minisun
            '//
            Dim a As Double
            a = x - Math.Floor(x)
            If (a < 0) Then a += 1
            Return a
        End Function

        Private Function hrsmin(ByVal hours As Double) As Double
            '//
            '//	takes decimal hours and returns a string in hhmm format
            '//
            Dim hrs, h, m, dum, l As Double
            hrs = Math.Floor(hours * 60 + 0.5) / 60.0
            h = Math.Floor(hrs)
            m = Math.Floor(60 * (hrs - h) + 0.5)
            dum = h * 100 + m
            '//
            '// the jiggery pokery below is to make sure that two minutes past midnight
            '// comes out as 0002 not 2. Javascript does not appear to have 'format codes'
            '// like C
            '//
            If (dum < 1000) Then dum = "0" + dum
            If (dum < 100) Then dum = "0" + dum
            If (dum < 10) Then dum = "0" + dum


            Return dum
        End Function

        Private Function toNormalTime(ByVal x As Double) As String
            Dim l As Int16
            Dim t, xx As String
            xx = x.ToString
            l = xx.Length

            If l = 3 Then
                t = xx.Substring(0, 1) & ":" & xx.Substring(1, 2)
            ElseIf l = 4 Then
                t = xx.Substring(0, 2) & ":" & xx.Substring(2, 2)
            Else
                t = String.Empty
            End If


            Return t
        End Function


        Private Function quad(ByVal ym As Double, ByVal yz As Double, ByVal yp As Double) As Double()


            '//
            '//	finds the parabola throuh the three points (-1,ym), (0,yz), (1, yp)
            '//  and returns the coordinates of the max/min (if any) xe, ye
            '//  the values of x where the parabola crosses zero (roots of the quadratic)
            '//  and the number of roots (0, 1 or 2) within the interval [-1, 1]
            '//
            '//	well, this routine is producing sensible answers
            '//
            '//  results passed as array [nz, z1, z2, xe, ye]
            '//
            Dim nz, a, b, c, dis, dx, xe, ye, z1, z2 As Double
            Dim quadout(5) As Double ' = new Array;

            nz = 0
            a = 0.5 * (ym + yp) - yz
            b = 0.5 * (yp - ym)
            c = yz
            xe = -b / (2 * a)
            ye = (a * xe + b) * xe + c
            dis = b * b - 4.0 * a * c
            If (dis > 0) Then
                dx = 0.5 * Math.Sqrt(dis) / Math.Abs(a)
                z1 = xe - dx
                z2 = xe + dx

                If (Math.Abs(z1) <= 1.0) Then nz += 1
                If (Math.Abs(z2) <= 1.0) Then nz += 1
                If (z1 < -1.0) Then z1 = z2

            End If

            quadout(0) = nz
            quadout(1) = z1
            quadout(2) = z2
            quadout(3) = xe
            quadout(4) = ye
            Return quadout

        End Function




        Private Function sin_alt(ByVal iobj As Double, ByVal mjd0 As Double, ByVal hour As Double _
                        , ByVal glong As Double, ByVal cglat As Double, ByVal sglat As Double) As Double

            ''	this rather mickey mouse function takes a lot of
            '//  arguments and then returns the sine of the altitude of
            '//  the object labelled by iobj. iobj = 1 is moon, iobj = 2 is sun
            '//
            Dim mjd, t, ra, dec, tau, salt As Double
            Dim rads As Double = 0.0174532925
            Dim objpos(5) As Double  '= New Array
            mjd = mjd0 + hour / 24.0
            t = (mjd - 51544.5) / 36525.0
            If (iobj = 1) Then
                objpos = minimoon(t)
            Else
                'objpos = minisun(t)  comment by programmer
            End If

            ra = objpos(2)
            dec = objpos(1)
            '// hour angle of object
            tau = 15.0 * (lmst(mjd, glong) - ra)
            '// sin(alt) of object using the conversion formulas
            salt = sglat * Math.Sin(rads * dec) + cglat * Math.Cos(rads * dec) * Math.Cos(rads * tau)

            Return salt

        End Function



        Private Function lmst(ByVal mjd As Double, ByVal glong As Double) As Double
            '//
            '//	Takes the mjd and the longitude (west negative) and then returns
            '//  the local sidereal time in hours. Im using Meeus formula 11.4
            '//  instead of messing about with UTo and so on
            '//
            Dim lst, t, d As Double
            d = mjd - 51544.5
            t = d / 36525.0
            lst = range(280.46061837 + 360.98564736629 * d + 0.000387933 * t * t - t * t * t / 38710000)
            Return (lst / 15.0 + glong / 15)
        End Function


        Private Function range(ByVal x As Double) As Double
            '//
            '//	returns an angle in degrees in the range 0 to 360
            '//
            Dim a, b As Double
            b = x / 360
            a = 360 * (b - ipart(b))
            If (a < 0) Then
                a = a + 360
            End If
            Return a
        End Function


        Private Function ipart(ByVal x As Double) As Double
            '//
            '//	returns the integer part - like int() in basic
            '//
            Dim a As Double
            If (x > 0) Then
                a = Math.Floor(x)
            Else
                a = Math.Ceiling(x)
            End If
            Return a
        End Function

    End Class


End Class
