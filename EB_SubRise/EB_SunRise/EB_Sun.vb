Option Explicit On 

Imports Microsoft.VisualBasic


Public Class EB_Sun

    Const ConAstronomical = 18
    Const conNautical = 12
    Const conCivil = 6

    ' Calculation of local times of sunrise, solar noon, and sunset
    ' based on the calculation procedure by NOAA in the javascript in
    ' http:''www.srrb.noaa.gov/highlights/sunrise/sunrise.html and
    ' http:''www.srrb.noaa.gov/highlights/sunrise/azel.html
    '
    ' Five functions are available for use from Excel worksheets:
    '
    '   - sunrise(lat, lon, year, month, day, timezone, dlstime)
    '   - solarnoon(lat, lon, year, month, day, timezone, dlstime)
    '   - sunset(lat, lon, year, month, day, timezone, dlstime)
    '   - solarazimuth(lat, lon, year, month, day, hour, minute, second, timezone, dlstime)
    '   - solarelevation(lat, lon, year, month, day, hour, minute, second, timezone, dlstime)
    '
    ' The sign convention for inputs to the functions named sunrise, solarnoon,
    ' sunset, solarazimuth, and solarelevationis:
    '
    '   - positive latitude decimal degrees for northern hemisphere
    '   - negative longitude degrees for western hemisphere
    '   - negative time zone hours for western hemisphere
    '
    ' The other functions in the VBA module use the original
    ' NOAA sign convention of positive longitude in the western hemisphere.
    '
    ' The calculations in the NOAA Sunrise/Sunset and Solar Position
    ' Calculators are based on equations from Astronomical Algorithms,
    ' by Jean Meeus. NOAA also included atmospheric refraction effects.
    ' The sunrise and sunset results were reported by NOAA
    ' to be accurate to within +/- 1 minute for locations between +/- 72°
    ' latitude, and within ten minutes outside of those latitudes.
    '
    ' This translation was tested for selected locations
    ' and found to provide results within +/- 1 minute of the
    ' original Javascript code.
    '
    ' This translation does not include calculation of prior or next
    ' susets for locations above the Arctic Circle and below the
    ' Antarctic Circle, when a sunrise or sunset does not occur.
    '





    Private Function radToDeg(ByVal angleRad As Double) As Double
        ''' Convert radian angle to degrees

        radToDeg = (180.0  * angleRad / Math.PI)

    End Function


    Private Function degToRad(ByVal angleDeg As Double) As Double
        ''' Convert degree angle to radians

        degToRad = (Math.PI * angleDeg / 180.0 )

    End Function


    Private Function calcJD(ByVal year As Double, ByVal month As Double, ByVal day As Double) As Double

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

        Dim A As Double
        Dim B As Double
        Dim JD As Double


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


    Private Function calcTimeJulianCent(ByVal JD As Double) As Double

        '***********************************************************************/
        '* Name:    calcTimeJulianCent
        '* Type:    Function
        '* Purpose: convert Julian Day to centuries since J2000.0.
        '* Arguments:
        '*   jd : the Julian Day to convert
        '* Return value:
        '*   the T value corresponding to the Julian Day
        '***********************************************************************/

        Dim t As Double

        t = (JD - 2451545.0 ) / 36525.0 
        calcTimeJulianCent = t

    End Function


    Private Function calcJDFromJulianCent(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcJDFromJulianCent
        '* Type:    Function
        '* Purpose: convert centuries since J2000.0 to Julian Day.
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   the Julian Day corresponding to the t value
        '***********************************************************************/

        Dim JD As Double

        JD = t * 36525.0  + 2451545.0 
        calcJDFromJulianCent = JD

    End Function


    Private Function calcGeomMeanLongSun(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calGeomMeanLongSun
        '* Type:    Function
        '* Purpose: calculate the Geometric Mean Longitude of the Sun
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   the Geometric Mean Longitude of the Sun in degrees
        '***********************************************************************/

        Dim l0 As Double

        l0 = 280.46646 + t * (36000.76983 + 0.0003032 * t)
        Do
            If (l0 <= 360) And (l0 >= 0) Then Exit Do
            If l0 > 360 Then l0 = l0 - 360
            If l0 < 0 Then l0 = l0 + 360
        Loop

        calcGeomMeanLongSun = l0

    End Function


    Private Function calcGeomMeanAnomalySun(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calGeomAnomalySun
        '* Type:    Function
        '* Purpose: calculate the Geometric Mean Anomaly of the Sun
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   the Geometric Mean Anomaly of the Sun in degrees
        '***********************************************************************/

        Dim m As Double

        m = 357.52911 + t * (35999.05029 - 0.0001537 * t)
        calcGeomMeanAnomalySun = m

    End Function


    Private Function calcEccentricityEarthOrbit(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcEccentricityEarthOrbit
        '* Type:    Function
        '* Purpose: calculate the eccentricity of earth's orbit
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   the unitless eccentricity
        '***********************************************************************/

        Dim e As Double

        e = 0.016708634 - t * (0.000042037 + 0.0000001267 * t)
        calcEccentricityEarthOrbit = e

    End Function


    Private Function calcSunEqOfCenter(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcSunEqOfCenter
        '* Type:    Function
        '* Purpose: calculate the equation of center for the sun
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   in degrees
        '***********************************************************************/

        Dim m As Double, mrad As Double, sinm As Double, sin2m As Double, sin3m As Double
        Dim c As Double

        m = calcGeomMeanAnomalySun(t)

        mrad = degToRad(m)
        sinm = Math.Sin(mrad)
        sin2m = Math.Sin(mrad + mrad)
        sin3m = Math.Sin(mrad + mrad + mrad)

        c = sinm * (1.914602 - t * (0.004817 + 0.000014 * t)) _
            + sin2m * (0.019993 - 0.000101 * t) + sin3m * 0.000289

        calcSunEqOfCenter = c

    End Function


    Private Function calcSunTrueLong(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcSunTrueLong
        '* Type:    Function
        '* Purpose: calculate the true longitude of the sun
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   sun's true longitude in degrees
        '***********************************************************************/

        Dim l0 As Double, c As Double, O As Double

        l0 = calcGeomMeanLongSun(t)
        c = calcSunEqOfCenter(t)

        O = l0 + c
        calcSunTrueLong = O

    End Function


    Private Function calcSunTrueAnomaly(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcSunTrueAnomaly (not used by sunrise, solarnoon, sunset)
        '* Type:    Function
        '* Purpose: calculate the true anamoly of the sun
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   sun's true anamoly in degrees
        '***********************************************************************/

        Dim m As Double, c As Double, v As Double

        m = calcGeomMeanAnomalySun(t)
        c = calcSunEqOfCenter(t)

        v = m + c
        calcSunTrueAnomaly = v

    End Function


    Private Function calcSunRadVector(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcSunRadVector (not used by sunrise, solarnoon, sunset)
        '* Type:    Function
        '* Purpose: calculate the distance to the sun in AU
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   sun radius vector in AUs
        '***********************************************************************/

        Dim v As Double, e As Double, R As Double

        v = calcSunTrueAnomaly(t)
        e = calcEccentricityEarthOrbit(t)

        R = (1.000001018 * (1 - e * e)) / (1 + e * Math.Cos(degToRad(v)))
        calcSunRadVector = R

    End Function


    Private Function calcSunApparentLong(ByVal t) As Double

        '***********************************************************************/
        '* Name:    calcSunApparentLong (not used by sunrise, solarnoon, sunset)
        '* Type:    Function
        '* Purpose: calculate the apparent longitude of the sun
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   sun's apparent longitude in degrees
        '***********************************************************************/

        Dim O As Double, omega As Double, lambda As Double

        O = calcSunTrueLong(t)

        omega = 125.04 - 1934.136 * t
        lambda = O - 0.00569 - 0.00478 * Math.Sin(degToRad(omega))
        calcSunApparentLong = lambda

    End Function


    Private Function calcMeanObliquityOfEcliptic(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcMeanObliquityOfEcliptic
        '* Type:    Function
        '* Purpose: calculate the mean obliquity of the ecliptic
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   mean obliquity in degrees
        '***********************************************************************/

        Dim seconds As Double, e0 As Double

        seconds = 21.448 - t * (46.815 + t * (0.00059 - t * (0.001813)))
        e0 = 23.0  + (26.0  + (seconds / 60.0 )) / 60.0 
        calcMeanObliquityOfEcliptic = e0

    End Function


    Private Function calcObliquityCorrection(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcObliquityCorrection
        '* Type:    Function
        '* Purpose: calculate the corrected obliquity of the ecliptic
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   corrected obliquity in degrees
        '***********************************************************************/

        Dim e0 As Double, omega As Double, e As Double

        e0 = calcMeanObliquityOfEcliptic(t)

        omega = 125.04 - 1934.136 * t
        e = e0 + 0.00256 * Math.Cos(degToRad(omega))
        calcObliquityCorrection = e

    End Function


    Private Function calcSunRtAscension(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcSunRtAscension (not used by sunrise, solarnoon, sunset)
        '* Type:    Function
        '* Purpose: calculate the right ascension of the sun
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   sun's right ascension in degrees
        '***********************************************************************/

        Dim e As Double, lambda As Double, tananum As Double, tanadenom As Double
        Dim alpha As Double

        e = calcObliquityCorrection(t)
        lambda = calcSunApparentLong(t)

        tananum = (Math.Cos(degToRad(e)) * Math.Sin(degToRad(lambda)))
        tanadenom = (Math.Cos(degToRad(lambda)))

        'original NOAA code using javascript Math.Atan2(y,x) convention:
        '        var alpha = radToDeg(Math.atan2(tananum, tanadenom));
        '        alpha = radToDeg(math.Atan2(tananum, tanadenom))

        'translated using Excel VBA math.Atan2(x,y) convention:
        alpha = radToDeg(Math.Atan2(tanadenom, tananum))

        calcSunRtAscension = alpha

    End Function


    Private Function calcSunDeclination(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcSunDeclination
        '* Type:    Function
        '* Purpose: calculate the declination of the sun
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   sun's declination in degrees
        '***********************************************************************/

        Dim e As Double, lambda As Double, sint As Double, theta As Double

        e = calcObliquityCorrection(t)
        lambda = calcSunApparentLong(t)

        sint = Math.Sin(degToRad(e)) * Math.Sin(degToRad(lambda))
        theta = radToDeg(Math.Asin(sint))
        calcSunDeclination = theta

    End Function


    Private Function calcEquationOfTime(ByVal t As Double) As Double

        '***********************************************************************/
        '* Name:    calcEquationOfTime
        '* Type:    Function
        '* Purpose: calculate the difference between true solar time and mean
        '*     solar time
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '* Return value:
        '*   equation of time in minutes of time
        '***********************************************************************/

        Dim epsilon As Double, l0 As Double, e As Double, m As Double
        Dim y As Double, sin2l0 As Double, sinm As Double
        Dim cos2l0 As Double, sin4l0 As Double, sin2m As Double, Etime As Double

        epsilon = calcObliquityCorrection(t)
        l0 = calcGeomMeanLongSun(t)
        e = calcEccentricityEarthOrbit(t)
        m = calcGeomMeanAnomalySun(t)

        y = Math.Tan(degToRad(epsilon) / 2.0 )
        y = y ^ 2

        sin2l0 = Math.Sin(2.0  * degToRad(l0))
        sinm = Math.Sin(degToRad(m))
        cos2l0 = Math.Cos(2.0  * degToRad(l0))
        sin4l0 = Math.Sin(4.0  * degToRad(l0))
        sin2m = Math.Sin(2.0  * degToRad(m))

        Etime = y * sin2l0 - 2.0  * e * sinm + 4.0  * e * y * sinm * cos2l0 _
                - 0.5 * y * y * sin4l0 - 1.25 * e * e * sin2m

        calcEquationOfTime = radToDeg(Etime) * 4.0 

    End Function


    Private Function calcHourAngleDawn(ByVal lat As Double, ByVal solarDec As Double, _
                  ByVal solardepression As Double) As Double

        '***********************************************************************/
        '* Name:    calcHourAngleDawn
        '* Type:    Function
        '* Purpose: calculate the hour angle of the sun at dawn for the
        '*         latitude
        '*         for user selected solar depression below horizon
        '* Arguments:
        '*   lat : latitude of observer in degrees
        '*   solarDec : declination angle of sun in degrees
        '*   solardepression: angle of the sun below the horizion in degrees
        '* Return value:
        '*   hour angle of dawn in radians
        '***********************************************************************/

        Dim latRad As Double, sdRad As Double, HAarg As Double, HA As Double

        latRad = degToRad(lat)
        sdRad = degToRad(solarDec)

        HAarg = (Math.Cos(degToRad(90 + solardepression)) / (Math.Cos(latRad) * Math.Cos(sdRad)) - Math.Tan(latRad) * Math.Tan(sdRad))

        HA = (Math.Acos(Math.Cos(degToRad(90 + solardepression)) _
              / (Math.Cos(latRad) * Math.Cos(sdRad)) - Math.Tan(latRad) * Math.Tan(sdRad)))

        calcHourAngleDawn = HA

    End Function


    Private Function calcHourAngleSunrise(ByVal lat As Double, ByVal solarDec As Double) As Double

        '***********************************************************************/
        '* Name:    calcHourAngleSunrise
        '* Type:    Function
        '* Purpose: calculate the hour angle of the sun at sunrise for the
        '*         latitude
        '* Arguments:
        '*   lat : latitude of observer in degrees
        '* solarDec : declination angle of sun in degrees
        '* Return value:
        '*   hour angle of sunrise in radians
        '*
        '* Note: For sunrise and sunset calculations, we assume 0.833° of atmospheric refraction
        '* For details about refraction see http:''www.srrb.noaa.gov/highlights/sunrise/calcdetails.html
        '*
        '***********************************************************************/

        Dim latRad As Double, sdRad As Double, HAarg As Double, HA As Double

        latRad = degToRad(lat)
        sdRad = degToRad(solarDec)

        HAarg = (Math.Cos(degToRad(90.833)) / (Math.Cos(latRad) * Math.Cos(sdRad)) - Math.Tan(latRad) * Math.Tan(sdRad))

        HA = (Math.Acos(Math.Cos(degToRad(90.833)) _
              / (Math.Cos(latRad) * Math.Cos(sdRad)) - Math.Tan(latRad) * Math.Tan(sdRad)))

        calcHourAngleSunrise = HA

    End Function


    Private Function calcHourAngleSunset(ByVal lat As Double, ByVal solarDec As Double) As Double

        '***********************************************************************/
        '* Name:    calcHourAngleSunset
        '* Type:    Function
        '* Purpose: calculate the hour angle of the sun at sunset for the
        '*         latitude
        '* Arguments:
        '*   lat : latitude of observer in degrees
        '* solarDec : declination angle of sun in degrees
        '* Return value:
        '*   hour angle of sunset in radians
        '*
        '* Note: For sunrise and sunset calculations, we assume 0.833° of atmospheric refraction
        '* For details about refraction see http:''www.srrb.noaa.gov/highlights/sunrise/calcdetails.html
        '*
        '***********************************************************************/

        Dim latRad As Double, sdRad As Double, HAarg As Double, HA As Double

        latRad = degToRad(lat)
        sdRad = degToRad(solarDec)

        HAarg = (Math.Cos(degToRad(90.833)) / (Math.Cos(latRad) * Math.Cos(sdRad)) - Math.Tan(latRad) * Math.Tan(sdRad))

        HA = (Math.Acos(Math.Cos(degToRad(90.833)) _
               / (Math.Cos(latRad) * Math.Cos(sdRad)) - Math.Tan(latRad) * Math.Tan(sdRad)))

        calcHourAngleSunset = -HA

    End Function


    Private Function calcHourAngleDusk(ByVal lat As Double, ByVal solarDec As Object, _
                                  ByVal solardepression As Double) As Double

        '***********************************************************************/
        '* Name:    calcHourAngleDusk
        '* Type:    Function
        '* Purpose: calculate the hour angle of the sun at dusk for the
        '*         latitude
        '*         for user selected solar depression below horizon
        '* Arguments:
        '*   lat : latitude of observer in degrees
        '*   solarDec : declination angle of sun in degrees
        '*   solardepression: angle of sun below horizon in degrees
        '* Return value:
        '*   hour angle of dusk in radians
        '***********************************************************************/

        Dim latRad As Double, sdRad As Double, HAarg As Double, HA As Double

        latRad = degToRad(lat)
        sdRad = degToRad(solarDec)

        HAarg = (Math.Cos(degToRad(90 + solardepression)) / (Math.Cos(latRad) * Math.Cos(sdRad)) - Math.Tan(latRad) * Math.Tan(sdRad))

        HA = (Math.Acos(Math.Cos(degToRad(90 + solardepression)) _
               / (Math.Cos(latRad) * Math.Cos(sdRad)) - Math.Tan(latRad) * Math.Tan(sdRad)))

        calcHourAngleDusk = -HA

    End Function


    Private Function calcDawnUTC(ByVal JD As Double, ByVal latitude As Double, ByVal longitude As Double _
                        , ByVal solardepression As Double) As Double

        '***********************************************************************/
        '* Name:    calcDawnUTC
        '* Type:    Function
        '* Purpose: calculate the Universal Coordinated Time (UTC) of dawn
        '*         for the given day at the given location on earth
        '*         for user selected solar depression below horizon
        '* Arguments:
        '*   JD  : julian day
        '*   latitude : latitude of observer in degrees
        '*   longitude : longitude of observer in degrees
        '*   solardepression: angle of sun below the horizon in degrees
        '* Return value:
        '*   time in minutes from zero Z
        '***********************************************************************/

        Dim t As Double, eqtime As Double, solarDec As Double, hourangle As Double
        Dim delta As Double, timeDiff As Double, timeUTC As Double
        Dim newt As Double

        t = calcTimeJulianCent(JD)

        '        '' *** First pass to approximate sunrise

        eqtime = calcEquationOfTime(t)
        solarDec = calcSunDeclination(t)
        hourangle = calcHourAngleSunrise(latitude, solarDec)

        delta = longitude - radToDeg(hourangle)
        timeDiff = 4 * delta
        ' in minutes of time
        timeUTC = 720 + timeDiff - eqtime
        ' in minutes

        ' *** Second pass includes fractional jday in gamma calc

        newt = calcTimeJulianCent(calcJDFromJulianCent(t) + timeUTC / 1440.0 )
        eqtime = calcEquationOfTime(newt)
        solarDec = calcSunDeclination(newt)
        hourangle = calcHourAngleDawn(latitude, solarDec, solardepression)
        delta = longitude - radToDeg(hourangle)
        timeDiff = 4 * delta
        timeUTC = 720 + timeDiff - eqtime
        ' in minutes

        calcDawnUTC = timeUTC

    End Function


    Private Function calcSunriseUTC(ByVal JD As Double, ByVal latitude As Double, ByVal longitude As Double) As Double

        '***********************************************************************/
        '* Name:    calcSunriseUTC
        '* Type:    Function
        '* Purpose: calculate the Universal Coordinated Time (UTC) of sunrise
        '*         for the given day at the given location on earth
        '* Arguments:
        '*   JD  : julian day
        '*   latitude : latitude of observer in degrees
        '*   longitude : longitude of observer in degrees
        '* Return value:
        '*   time in minutes from zero Z
        '***********************************************************************/

        Dim t As Double, eqtime As Double, solarDec As Double, hourangle As Double
        Dim delta As Double, timeDiff As Double, timeUTC As Double
        Dim newt As Double

        t = calcTimeJulianCent(JD)

        '        '' *** First pass to approximate sunrise

        eqtime = calcEquationOfTime(t)
        solarDec = calcSunDeclination(t)
        hourangle = calcHourAngleSunrise(latitude, solarDec)

        delta = longitude - radToDeg(hourangle)
        timeDiff = 4 * delta
        ' in minutes of time
        timeUTC = 720 + timeDiff - eqtime
        ' in minutes

        ' *** Second pass includes fractional jday in gamma calc

        newt = calcTimeJulianCent(calcJDFromJulianCent(t) + timeUTC / 1440.0 )
        eqtime = calcEquationOfTime(newt)
        solarDec = calcSunDeclination(newt)
        hourangle = calcHourAngleSunrise(latitude, solarDec)
        delta = longitude - radToDeg(hourangle)
        timeDiff = 4 * delta
        timeUTC = 720 + timeDiff - eqtime
        ' in minutes

        calcSunriseUTC = timeUTC

    End Function


    Private Function calcSolNoonUTC(ByVal t As Double, ByVal longitude As Double) As Double

        '***********************************************************************/
        '* Name:    calcSolNoonUTC
        '* Type:    Function
        '* Purpose: calculate the Universal Coordinated Time (UTC) of solar
        '*     noon for the given day at the given location on earth
        '* Arguments:
        '*   t : number of Julian centuries since J2000.0
        '*   longitude : longitude of observer in degrees
        '* Return value:
        '*   time in minutes from zero Z
        '***********************************************************************/

        Dim newt As Double, eqtime As Double, solarNoonDec As Double, solNoonUTC As Double

        newt = calcTimeJulianCent(calcJDFromJulianCent(t) + 0.5 + longitude / 360.0 )

        eqtime = calcEquationOfTime(newt)
        solarNoonDec = calcSunDeclination(newt)
        solNoonUTC = 720 + (longitude * 4) - eqtime

        calcSolNoonUTC = solNoonUTC

    End Function


    Private Function calcSunsetUTC(ByVal JD As Double, ByVal latitude As Double, ByVal longitude As Double) As Double

        '***********************************************************************/
        '* Name:    calcSunsetUTC
        '* Type:    Function
        '* Purpose: calculate the Universal Coordinated Time (UTC) of sunset
        '*         for the given day at the given location on earth
        '* Arguments:
        '*   JD  : julian day
        '*   latitude : latitude of observer in degrees
        '*   longitude : longitude of observer in degrees
        '* Return value:
        '*   time in minutes from zero Z
        '***********************************************************************/

        Dim t As Double, eqtime As Double, solarDec As Double, hourangle As Double
        Dim delta As Double, timeDiff As Double, timeUTC As Double
        Dim newt As Double

        t = calcTimeJulianCent(JD)

        '        '' First calculates sunrise and approx length of day

        eqtime = calcEquationOfTime(t)
        solarDec = calcSunDeclination(t)
        hourangle = calcHourAngleSunset(latitude, solarDec)

        delta = longitude - radToDeg(hourangle)
        timeDiff = 4 * delta
        timeUTC = 720 + timeDiff - eqtime

        '        '' first pass used to include fractional day in gamma calc

        newt = calcTimeJulianCent(calcJDFromJulianCent(t) + timeUTC / 1440.0 )
        eqtime = calcEquationOfTime(newt)
        solarDec = calcSunDeclination(newt)
        hourangle = calcHourAngleSunset(latitude, solarDec)

        delta = longitude - radToDeg(hourangle)
        timeDiff = 4 * delta
        timeUTC = 720 + timeDiff - eqtime
        '        '' in minutes

        calcSunsetUTC = timeUTC

    End Function


    Private Function calcDuskUTC(ByVal JD As Double, ByVal latitude As Double, _
                         ByVal longitude As Double, ByVal solardepression As Double)

        '***********************************************************************/
        '* Name:    calcDuskUTC
        '* Type:    Function
        '* Purpose: calculate the Universal Coordinated Time (UTC) of dusk
        '*         for the given day at the given location on earth
        '*         for user selected solar depression below horizon
        '* Arguments:
        '*   JD  : julian day
        '*   latitude : latitude of observer in degrees
        '*   longitude : longitude of observer in degrees
        '*   solardepression: angle of sun below horizon
        '* Return value:
        '*   time in minutes from zero Z
        '***********************************************************************/

        Dim t As Double, eqtime As Double, solarDec As Double, hourangle As Double
        Dim delta As Double, timeDiff As Double, timeUTC As Double
        Dim newt As Double

        t = calcTimeJulianCent(JD)

        '        '' First calculates sunrise and approx length of day

        eqtime = calcEquationOfTime(t)
        solarDec = calcSunDeclination(t)
        hourangle = calcHourAngleSunset(latitude, solarDec)

        delta = longitude - radToDeg(hourangle)
        timeDiff = 4 * delta
        timeUTC = 720 + timeDiff - eqtime

        '        '' first pass used to include fractional day in gamma calc

        newt = calcTimeJulianCent(calcJDFromJulianCent(t) + timeUTC / 1440.0 )
        eqtime = calcEquationOfTime(newt)
        solarDec = calcSunDeclination(newt)
        hourangle = calcHourAngleDusk(latitude, solarDec, solardepression)

        delta = longitude - radToDeg(hourangle)
        timeDiff = 4 * delta
        timeUTC = 720 + timeDiff - eqtime
        '        '' in minutes

        calcDuskUTC = timeUTC

    End Function


    Public Function dawn(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
            ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
            ByVal dlstime As Double, ByVal solardepression As Double) As Date

        '***********************************************************************/
        '* Name:    dawn
        '* Type:    Main Function called by spreadsheet
        '* Purpose: calculate time of dawn  for the entered date
        '*     and location.
        '* For latitudes greater than 72 degrees N and S, calculations are
        '* accurate to within 10 minutes. For latitudes less than +/- 72°
        '* accuracy is approximately one minute.
        '* Arguments:
        '   latitude = latitude (decimal degrees)
        '   longitude = longitude (decimal degrees)
        '    NOTE: longitude is negative for western hemisphere for input cells
        '          in the spreadsheet for calls to the functions named
        '          sunrise, solarnoon, and sunset. Those functions convert the
        '          longitude to positive for the western hemisphere for calls to
        '          other functions using the original sign convention
        '          from the NOAA javascript code.
        '   year = year
        '   month = month
        '   day = day
        '   timezone = time zone hours relative to GMT/UTC (hours)
        '   dlstime = daylight savings time (0 = no, 1 = yes) (hours)
        '   solardepression = angle of sun below horizon in degrees
        '* Return value:
        '*   dawn time in local time (days)
        '***********************************************************************/

        Dim longitude As Double, latitude As Double, JD As Double
        Dim riseTimeGMT As Double, riseTimeLST As Double

        ' change sign convention for longitude from negative to positive in western hemisphere
        longitude = lon * -1
        latitude = lat
        If (latitude > 89.8) Then latitude = 89.8
        If (latitude < -89.8) Then latitude = -89.8

        JD = calcJD(year, month, day)

        '            '' Calculate sunrise for this date
        riseTimeGMT = calcDawnUTC(JD, latitude, longitude, solardepression)

        '            ''  adjust for time zone and daylight savings time in minutes
        riseTimeLST = riseTimeGMT + (60 * timezone) + (dlstime * 60)


        '            ''  convert to days
        dawn = Format(Date.FromOADate(riseTimeLST / 1440), "Short Time")
        'dawn = riseTimeLST / 1440

        'Debug.Print Format((riseTimeLST / 1440), "Short Time")

    End Function


    Public Function sunrise(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
                     ByVal month As Double, ByVal day As Double, _
                     ByVal timezone As Double, ByVal dlstime As Double) As Date

        '***********************************************************************/
        '* Name:    sunrise
        '* Type:    Main Function called by spreadsheet
        '* Purpose: calculate time of sunrise  for the entered date
        '*     and location.
        '* For latitudes greater than 72 degrees N and S, calculations are
        '* accurate to within 10 minutes. For latitudes less than +/- 72°
        '* accuracy is approximately one minute.
        '* Arguments:
        '   latitude = latitude (decimal degrees)
        '   longitude = longitude (decimal degrees)
        '    NOTE: longitude is negative for western hemisphere for input cells
        '          in the spreadsheet for calls to the functions named
        '          sunrise, solarnoon, and sunset. Those functions convert the
        '          longitude to positive for the western hemisphere for calls to
        '          other functions using the original sign convention
        '          from the NOAA javascript code.
        '   year = year
        '   month = month
        '   day = day
        '   timezone = time zone hours relative to GMT/UTC (hours)
        '   dlstime = daylight savings time (0 = no, 1 = yes) (hours)
        '* Return value:
        '*   sunrise time in local time (days)
        '***********************************************************************/

        Dim longitude As Double, latitude As Double, JD As Double
        Dim riseTimeGMT As Double, riseTimeLST As Double

        ' change sign convention for longitude from negative to positive in western hemisphere
        longitude = lon * -1
        latitude = lat
        If (latitude > 89.8) Then latitude = 89.8
        If (latitude < -89.8) Then latitude = -89.8

        JD = calcJD(year, month, day)

        '            '' Calculate sunrise for this date
        riseTimeGMT = calcSunriseUTC(JD, latitude, longitude)

        '            ''  adjust for time zone and daylight savings time in minutes
        riseTimeLST = riseTimeGMT + (60 * timezone) + (dlstime * 60)

        '            ''  convert to days
        'sunrise = riseTimeLST / 1440
        sunrise = Format(Date.FromOADate(riseTimeLST / 1440), "Short Time")

    End Function


    Public Function solarnoon(ByVal lat As Double, ByVal lon As Double, _
                        ByVal year As Double, ByVal month As Double, ByVal day As Double, _
                        ByVal timezone As Double, ByVal dlstime As Double) As Date

        '***********************************************************************/
        '* Name:    solarnoon
        '* Type:    Main Function called by spreadsheet
        '* Purpose: calculate the Universal Coordinated Time (UTC) of solar
        '*     noon for the given day at the given location on earth
        '* Arguments:
        '    year
        '    month
        '    day
        '*   longitude : longitude of observer in degrees
        '    NOTE: longitude is negative for western hemisphere for input cells
        '          in the spreadsheet for calls to the functions named
        '          sunrise, solarnoon, and sunset. Those functions convert the
        '          longitude to positive for the western hemisphere for calls to
        '          other functions using the original sign convention
        '          from the NOAA javascript code.
        '* Return value:
        '*   time of solar noon in local time days
        '***********************************************************************/

        Dim longitude As Double, latitude As Double, JD As Double
        Dim t As Double, newt As Double
        Dim eqtime, SoNo As Double
        Dim solarNoonDec As Double, solNoonUTC As Double

        ' change sign convention for longitude from negative to positive in western hemisphere
        longitude = lon * -1
        latitude = lat
        If (latitude > 89.8) Then latitude = 89.8
        If (latitude < -89.8) Then latitude = -89.8

        JD = calcJD(year, month, day)
        t = calcTimeJulianCent(JD)

        newt = calcTimeJulianCent(calcJDFromJulianCent(t) + 0.5 + longitude / 360.0 )

        eqtime = calcEquationOfTime(newt)
        solarNoonDec = calcSunDeclination(newt)
        solNoonUTC = 720 + (longitude * 4) - eqtime

        '            ''  adjust for time zone and daylight savings time in minutes
        'solarnoon = solNoonUTC + (60 * timezone) + (dlstime * 60)
        SoNo = solNoonUTC + (60 * timezone) + (dlstime * 60)


        '            ''  convert to days
        'solarnoon = solarnoon / 1440
        SoNo = SoNo / 1440

        Return Format(Date.FromOADate(SoNo), "Short Time")

    End Function


    Public Function sunset(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
                    ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
                    ByVal dlstime As Double) As Date

        '***********************************************************************/
        '* Name:    sunset
        '* Type:    Main Function called by spreadsheet
        '* Purpose: calculate time of sunrise and sunset for the entered date
        '*     and location.
        '* For latitudes greater than 72 degrees N and S, calculations are
        '* accurate to within 10 minutes. For latitudes less than +/- 72°
        '* accuracy is approximately one minute.
        '* Arguments:
        '   latitude = latitude (decimal degrees)
        '   longitude = longitude (decimal degrees)
        '    NOTE: longitude is negative for western hemisphere for input cells
        '          in the spreadsheet for calls to the functions named
        '          sunrise, solarnoon, and sunset. Those functions convert the
        '          longitude to positive for the western hemisphere for calls to
        '          other functions using the original sign convention
        '          from the NOAA javascript code.
        '   year = year
        '   month = month
        '   day = day
        '   timezone = time zone hours relative to GMT/UTC (hours)
        '   dlstime = daylight savings time (0 = no, 1 = yes) (hours)
        '* Return value:
        '*   sunset time in local time (days)
        '***********************************************************************/

        Dim longitude As Double, latitude As Double, JD As Double
        Dim setTimeGMT As Double, setTimeLST As Double

        ' change sign convention for longitude from negative to positive in western hemisphere
        longitude = lon * -1
        latitude = lat
        If (latitude > 89.8) Then latitude = 89.8
        If (latitude < -89.8) Then latitude = -89.8

        JD = calcJD(year, month, day)

        '           '' Calculate sunset for this date
        setTimeGMT = calcSunsetUTC(JD, latitude, longitude)

        '            ''  adjust for time zone and daylight savings time in minutes
        setTimeLST = setTimeGMT + (60 * timezone) + (dlstime * 60)

        '            ''  convert to days
        sunset = Format(Date.FromOADate(setTimeLST / 1440), "Short Time")

    End Function


    Public Function dusk(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
                  ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
                  ByVal dlstime As Double, ByVal solardepression As Double) As Date

        '***********************************************************************/
        '* Name:    dusk
        '* Type:    Main Function called by spreadsheet
        '* Purpose: calculate time of sunrise and sunset for the entered date
        '*     and location.
        '* For latitudes greater than 72 degrees N and S, calculations are
        '* accurate to within 10 minutes. For latitudes less than +/- 72°
        '* accuracy is approximately one minute.
        '* Arguments:
        '   latitude = latitude (decimal degrees)
        '   longitude = longitude (decimal degrees)
        '    NOTE: longitude is negative for western hemisphere for input cells
        '          in the spreadsheet for calls to the functions named
        '          sunrise, solarnoon, and sunset. Those functions convert the
        '          longitude to positive for the western hemisphere for calls to
        '          other functions using the original sign convention
        '          from the NOAA javascript code.
        '   year = year
        '   month = month
        '   day = day
        '   timezone = time zone hours relative to GMT/UTC (hours)
        '   dlstime = daylight savings time (0 = no, 1 = yes) (hours)
        '   solardepression = angle of sun below horizon in degrees
        '* Return value:
        '*   dusk time in local time (days)
        '***********************************************************************/

        Dim longitude As Double, latitude As Double, JD As Double
        Dim setTimeGMT As Double, setTimeLST As Double

        ' change sign convention for longitude from negative to positive in western hemisphere
        longitude = lon * -1
        latitude = lat
        If (latitude > 89.8) Then latitude = 89.8
        If (latitude < -89.8) Then latitude = -89.8

        JD = calcJD(year, month, day)

        '           '' Calculate sunset for this date
        setTimeGMT = calcDuskUTC(JD, latitude, longitude, solardepression)

        '            ''  adjust for time zone and daylight savings time in minutes
        setTimeLST = setTimeGMT + (60 * timezone) + (dlstime * 60)

        '            ''  convert to days
        dusk = Format(Date.FromOADate(setTimeLST / 1440), "Short Time")

    End Function


    Public Function solarazimuth(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
                          ByVal month As Double, ByVal day As Double, ByVal hours As Double, _
                          ByVal minutes As Double, ByVal seconds As Double, ByVal timezone As Double _
                          , ByVal dlstime As Double) As Double

        '***********************************************************************/
        '* Name:    solarazimuth
        '* Type:    Main Function
        '* Purpose: calculate solar azimuth (deg from north) for the entered
        '*          date, time and location. Returns -999999 if darker than twilight
        '*
        '* Arguments:
        '*   latitude, longitude, year, month, day, hour, minute, second,
        '*   timezone, daylightsavingstime
        '* Return value:
        '*   solar azimuth in degrees from north
        '*
        '* Note: solarelevation and solarazimuth functions are identical
        '*       and could be converted to a VBA subroutine that would return
        '*       both values.
        '*
        '***********************************************************************/

        Dim longitude As Double, latitude As Double
        Dim zone As Double, daySavings As Double
        Dim hh As Double, mm As Double, ss As Double, timenow As Double
        Dim JD As Double, t As Double, R As Double
        Dim alpha As Double, theta As Double, Etime As Double, eqtime As Double
        Dim solarDec As Double, earthRadVec As Double, solarTimeFix As Double
        Dim trueSolarTime As Double, hourangle As Double, harad As Double
        Dim csz As Double, zenith As Double, azDenom As Double, azRad As Double
        Dim azimuth As Double, exoatmElevation As Double
        Dim step1 As Double, step2 As Double, step3 As Double
        Dim refractionCorrection As Double, te As Double, solarzen As Double

        ' change sign convention for longitude from negative to positive in western hemisphere
        longitude = lon * -1
        latitude = lat
        If (latitude > 89.8) Then latitude = 89.8
        If (latitude < -89.8) Then latitude = -89.8

        'change time zone to ppositive hours in western hemisphere
        zone = timezone * -1
        daySavings = dlstime * 60
        hh = hours - (daySavings / 60)
        mm = minutes
        ss = seconds

        '''    timenow is GMT time for calculation in hours since 0Z
        timenow = hh + mm / 60 + ss / 3600 + zone

        JD = calcJD(year, month, day)
        t = calcTimeJulianCent(JD + timenow / 24.0 )
        R = calcSunRadVector(t)
        alpha = calcSunRtAscension(t)
        theta = calcSunDeclination(t)
        Etime = calcEquationOfTime(t)

        eqtime = Etime
        solarDec = theta '''    in degrees
        earthRadVec = R

        solarTimeFix = eqtime - 4.0  * longitude + 60.0  * zone
        trueSolarTime = hh * 60.0  + mm + ss / 60.0  + solarTimeFix
        '''    in minutes

        Do While (trueSolarTime > 1440)
            trueSolarTime = trueSolarTime - 1440
        Loop

        hourangle = trueSolarTime / 4.0  - 180.0 
        '''    Thanks to Louis Schwarzmayr for the next line:
        If (hourangle < -180) Then hourangle = hourangle + 360.0 

        harad = degToRad(hourangle)

        csz = Math.Sin(degToRad(latitude)) * _
              Math.Sin(degToRad(solarDec)) + _
              Math.Cos(degToRad(latitude)) * _
              Math.Cos(degToRad(solarDec)) * Math.Cos(harad)

        If (csz > 1.0 ) Then
            csz = 1.0 
        ElseIf (csz < -1.0 ) Then
            csz = -1.0 
        End If

        zenith = radToDeg(Math.Acos(csz))

        azDenom = (Math.Cos(degToRad(latitude)) * Math.Sin(degToRad(zenith)))

        If (Math.Abs(azDenom) > 0.001) Then
            azRad = ((Math.Sin(degToRad(latitude)) * _
                Math.Cos(degToRad(zenith))) - _
                Math.Sin(degToRad(solarDec))) / azDenom
            If (Math.Abs(azRad) > 1.0 ) Then
                If (azRad < 0) Then
                    azRad = -1.0 
                Else
                    azRad = 1.0 
                End If
            End If

            azimuth = 180.0  - radToDeg(Math.Acos(azRad))

            If (hourangle > 0.0 ) Then
                azimuth = -azimuth
            End If
        Else
            If (latitude > 0.0 ) Then
                azimuth = 180.0 
            Else
                azimuth = 0.0 
            End If
        End If
        If (azimuth < 0.0 ) Then
            azimuth = azimuth + 360.0 
        End If

        exoatmElevation = 90.0  - zenith

        'beginning of complex expression commented out
        '            If (exoatmElevation > 85 ) Then
        '                refractionCorrection = 0 
        '            Else
        '                te = math.tan(degToRad(exoatmElevation))
        '                If (exoatmElevation > 5 ) Then
        '                    refractionCorrection = 58.1 / te - 0.07 / (te * te * te) + _
        '                        0.000086 / (te * te * te * te * te)
        '                ElseIf (exoatmElevation > -0.575) Then
        '                    refractionCorrection = 1735  + exoatmElevation * _
        '                        (-518.2 + exoatmElevation * (103.4 + _
        '                        exoatmElevation * (-12.79 + _
        '                        exoatmElevation * 0.711)))
        '                Else
        '                    refractionCorrection = -20.774 / te
        '                End If
        '                refractionCorrection = refractionCorrection / 3600 
        '            End If
        'end of complex expression

        'beginning of simplified expression
        If (exoatmElevation > 85.0 ) Then
            refractionCorrection = 0.0 
        Else
            te = Math.Tan(degToRad(exoatmElevation))
            If (exoatmElevation > 5.0 ) Then
                refractionCorrection = 58.1 / te - 0.07 / (te * te * te) + _
                    0.000086 / (te * te * te * te * te)
            ElseIf (exoatmElevation > -0.575) Then
                step1 = (-12.79 + exoatmElevation * 0.711)
                step2 = (103.4 + exoatmElevation * (step1))
                step3 = (-518.2 + exoatmElevation * (step2))
                refractionCorrection = 1735.0  + exoatmElevation * (step3)
            Else
                refractionCorrection = -20.774 / te
            End If
            refractionCorrection = refractionCorrection / 3600.0 
        End If
        'end of simplified expression

        solarzen = zenith - refractionCorrection

        '            If (solarZen < 108 ) Then
        solarazimuth = azimuth

        '              solarelevation = 90  - solarZen
        '              If (solarZen < 90 ) Then
        '                coszen = math.cos(degToRad(solarZen))
        '              Else
        '                coszen = 0 
        '              End If
        '            Else    ''' do not report az & el after astro twilight
        '              solarazimuth = -999999
        '              solarelevation = -999999
        '              coszen = -999999
        '            End If

    End Function


    Public Function solarelevation(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
                            ByVal month As Double, ByVal day As Double, _
                            ByVal hours As Double, ByVal minutes As Double, ByVal seconds As Double, _
                            ByVal timezone As Double, ByVal dlstime As Double) As Double

        '***********************************************************************/
        '* Name:    solarazimuth
        '* Type:    Main Function
        '* Purpose: calculate solar azimuth (deg from north) for the entered
        '*          date, time and location. Returns -999999 if darker than twilight
        '*
        '* Arguments:
        '*   latitude, longitude, year, month, day, hour, minute, second,
        '*   timezone, daylightsavingstime
        '* Return value:
        '*   solar azimuth in degrees from north
        '*
        '* Note: solarelevation and solarazimuth functions are identical
        '*       and could converted to a VBA subroutine that would return
        '*       both values.
        '*
        '***********************************************************************/

        Dim longitude As Double, latitude As Double
        Dim zone As Double, daySavings As Double
        Dim hh As Double, mm As Double, ss As Double, timenow As Double
        Dim JD As Double, t As Double, R As Double
        Dim alpha As Double, theta As Double, Etime As Double, eqtime As Double
        Dim solarDec As Double, earthRadVec As Double, solarTimeFix As Double
        Dim trueSolarTime As Double, hourangle As Double, harad As Double
        Dim csz As Double, zenith As Double, azDenom As Double, azRad As Double
        Dim azimuth As Double, exoatmElevation As Double
        Dim step1 As Double, step2 As Double, step3 As Double
        Dim refractionCorrection As Double, te As Double, solarzen As Double

        ' change sign convention for longitude from negative to positive in western hemisphere
        longitude = lon * -1
        latitude = lat
        If (latitude > 89.8) Then latitude = 89.8
        If (latitude < -89.8) Then latitude = -89.8

        'change time zone to ppositive hours in western hemisphere
        zone = timezone * -1
        daySavings = dlstime * 60
        hh = hours - (daySavings / 60)
        mm = minutes
        ss = seconds

        '''    timenow is GMT time for calculation in hours since 0Z
        timenow = hh + mm / 60 + ss / 3600 + zone

        JD = calcJD(year, month, day)
        t = calcTimeJulianCent(JD + timenow / 24.0 )
        R = calcSunRadVector(t)
        alpha = calcSunRtAscension(t)
        theta = calcSunDeclination(t)
        Etime = calcEquationOfTime(t)

        eqtime = Etime
        solarDec = theta '''    in degrees
        earthRadVec = R

        solarTimeFix = eqtime - 4.0  * longitude + 60.0  * zone
        trueSolarTime = hh * 60.0  + mm + ss / 60.0  + solarTimeFix
        '''    in minutes

        Do While (trueSolarTime > 1440)
            trueSolarTime = trueSolarTime - 1440
        Loop

        hourangle = trueSolarTime / 4.0  - 180.0 
        '''    Thanks to Louis Schwarzmayr for the next line:
        If (hourangle < -180) Then hourangle = hourangle + 360.0 

        harad = degToRad(hourangle)

        csz = Math.Sin(degToRad(latitude)) * _
              Math.Sin(degToRad(solarDec)) + _
              Math.Cos(degToRad(latitude)) * _
              Math.Cos(degToRad(solarDec)) * Math.Cos(harad)

        If (csz > 1.0 ) Then
            csz = 1.0 
        ElseIf (csz < -1.0 ) Then
            csz = -1.0 
        End If

        zenith = radToDeg(Math.Acos(csz))

        azDenom = (Math.Cos(degToRad(latitude)) * Math.Sin(degToRad(zenith)))

        If (Math.Abs(azDenom) > 0.001) Then
            azRad = ((Math.Sin(degToRad(latitude)) * _
                Math.Cos(degToRad(zenith))) - _
                Math.Sin(degToRad(solarDec))) / azDenom
            If (Math.Abs(azRad) > 1.0 ) Then
                If (azRad < 0) Then
                    azRad = -1.0 
                Else
                    azRad = 1.0 
                End If
            End If

            azimuth = 180.0  - radToDeg(Math.Acos(azRad))

            If (hourangle > 0.0 ) Then
                azimuth = -azimuth
            End If
        Else
            If (latitude > 0.0 ) Then
                azimuth = 180.0 
            Else
                azimuth = 0.0 
            End If
        End If
        If (azimuth < 0.0 ) Then
            azimuth = azimuth + 360.0 
        End If

        exoatmElevation = 90.0  - zenith

        'beginning of complex expression commented out
        '            If (exoatmElevation > 85 ) Then
        '                refractionCorrection = 0 
        '            Else
        '                te = math.tan(degToRad(exoatmElevation))
        '                If (exoatmElevation > 5 ) Then
        '                    refractionCorrection = 58.1 / te - 0.07 / (te * te * te) + _
        '                        0.000086 / (te * te * te * te * te)
        '                ElseIf (exoatmElevation > -0.575) Then
        '                    refractionCorrection = 1735  + exoatmElevation * _
        '                        (-518.2 + exoatmElevation * (103.4 + _
        '                        exoatmElevation * (-12.79 + _
        '                        exoatmElevation * 0.711)))
        '                Else
        '                    refractionCorrection = -20.774 / te
        '                End If
        '                refractionCorrection = refractionCorrection / 3600 
        '            End If
        'end of complex expression

        'beginning of simplified expression
        If (exoatmElevation > 85.0 ) Then
            refractionCorrection = 0.0 
        Else
            te = Math.Tan(degToRad(exoatmElevation))
            If (exoatmElevation > 5.0 ) Then
                refractionCorrection = 58.1 / te - 0.07 / (te * te * te) + _
                    0.000086 / (te * te * te * te * te)
            ElseIf (exoatmElevation > -0.575) Then
                step1 = (-12.79 + exoatmElevation * 0.711)
                step2 = (103.4 + exoatmElevation * (step1))
                step3 = (-518.2 + exoatmElevation * (step2))
                refractionCorrection = 1735.0  + exoatmElevation * (step3)
            Else
                refractionCorrection = -20.774 / te
            End If
            refractionCorrection = refractionCorrection / 3600.0 
        End If
        'end of simplified expression

        solarzen = zenith - refractionCorrection

        '            If (solarZen < 108 ) Then
        '              solarazimuth = azimuth
        solarelevation = 90.0  - solarzen
        '              If (solarZen < 90 ) Then
        '                coszen = math.cos(degToRad(solarZen))
        '              Else
        '                coszen = 0 
        '              End If
        '            Else    ''' do not report az & el after astro twilight
        '              solarazimuth = -999999
        '              solarelevation = -999999
        '              coszen = -999999
        '            End If

    End Function


    Private Sub solarposition(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, ByVal month As Double _
                , ByVal day As Double, ByVal hours As Double, ByVal minutes As Double, _
                ByVal seconds As Double, ByVal timezone As Double, ByVal dlstime As Double, _
                ByVal solarazimuth As Double, ByVal solarelevation As Double)

        '***********************************************************************/
        '* Name:    solarazimuth
        '* Type:    Main Function
        '* Purpose: calculate solar azimuth (deg from north) for the entered
        '*          date, time and location. Returns -999999 if darker than twilight
        '*
        '* Arguments:
        '*   latitude, longitude, year, month, day, hour, minute, second,
        '*   timezone, daylightsavingstime
        '* Return value:
        '*   solar azimuth in degrees from north
        '*
        '* Note: solarelevation and solarazimuth functions are identical
        '*       and could converted to a VBA subroutine that would return
        '*       both values.
        '*
        '***********************************************************************/

        Dim longitude As Double, latitude As Double
        Dim zone As Double, daySavings As Double
        Dim hh As Double, mm As Double, ss As Double, timenow As Double
        Dim JD As Double, t As Double, R As Double
        Dim alpha As Double, theta As Double, Etime As Double, eqtime As Double
        Dim solarDec As Double, earthRadVec As Double, solarTimeFix As Double
        Dim trueSolarTime As Double, hourangle As Double, harad As Double
        Dim csz As Double, zenith As Double, azDenom As Double, azRad As Double
        Dim azimuth As Double, exoatmElevation As Double
        Dim step1 As Double, step2 As Double, step3 As Double
        Dim refractionCorrection As Double, te As Double, solarzen As Double

        ' change sign convention for longitude from negative to positive in western hemisphere
        longitude = lon * -1
        latitude = lat
        If (latitude > 89.8) Then latitude = 89.8
        If (latitude < -89.8) Then latitude = -89.8

        'change time zone to ppositive hours in western hemisphere
        zone = timezone * -1
        daySavings = dlstime * 60
        hh = hours - (daySavings / 60)
        mm = minutes
        ss = seconds

        '''    timenow is GMT time for calculation in hours since 0Z
        timenow = hh + mm / 60 + ss / 3600 + zone

        JD = calcJD(year, month, day)
        t = calcTimeJulianCent(JD + timenow / 24.0 )
        R = calcSunRadVector(t)
        alpha = calcSunRtAscension(t)
        theta = calcSunDeclination(t)
        Etime = calcEquationOfTime(t)

        eqtime = Etime
        solarDec = theta '''    in degrees
        earthRadVec = R

        solarTimeFix = eqtime - 4.0  * longitude + 60.0  * zone
        trueSolarTime = hh * 60.0  + mm + ss / 60.0  + solarTimeFix
        '''    in minutes

        Do While (trueSolarTime > 1440)
            trueSolarTime = trueSolarTime - 1440
        Loop

        hourangle = trueSolarTime / 4.0  - 180.0 
        '''    Thanks to Louis Schwarzmayr for the next line:
        If (hourangle < -180) Then hourangle = hourangle + 360.0 

        harad = degToRad(hourangle)

        csz = Math.Sin(degToRad(latitude)) * _
              Math.Sin(degToRad(solarDec)) + _
              Math.Cos(degToRad(latitude)) * _
              Math.Cos(degToRad(solarDec)) * Math.Cos(harad)

        If (csz > 1.0 ) Then
            csz = 1.0 
        ElseIf (csz < -1.0 ) Then
            csz = -1.0 
        End If

        zenith = radToDeg(Math.Acos(csz))

        azDenom = (Math.Cos(degToRad(latitude)) * Math.Sin(degToRad(zenith)))

        If (Math.Abs(azDenom) > 0.001) Then
            azRad = ((Math.Sin(degToRad(latitude)) * _
                Math.Cos(degToRad(zenith))) - _
                Math.Sin(degToRad(solarDec))) / azDenom
            If (Math.Abs(azRad) > 1.0 ) Then
                If (azRad < 0) Then
                    azRad = -1.0 
                Else
                    azRad = 1.0 
                End If
            End If

            azimuth = 180.0  - radToDeg(Math.Acos(azRad))

            If (hourangle > 0.0 ) Then
                azimuth = -azimuth
            End If
        Else
            If (latitude > 0.0 ) Then
                azimuth = 180.0 
            Else
                azimuth = 0.0 
            End If
        End If
        If (azimuth < 0.0 ) Then
            azimuth = azimuth + 360.0 
        End If

        exoatmElevation = 90.0  - zenith

        'beginning of complex expression commented out
        '            If (exoatmElevation > 85 ) Then
        '                refractionCorrection = 0 
        '            Else
        '                te = math.tan(degToRad(exoatmElevation))
        '                If (exoatmElevation > 5 ) Then
        '                    refractionCorrection = 58.1 / te - 0.07 / (te * te * te) + _
        '                        0.000086 / (te * te * te * te * te)
        '                ElseIf (exoatmElevation > -0.575) Then
        '                    refractionCorrection = 1735  + exoatmElevation * _
        '                        (-518.2 + exoatmElevation * (103.4 + _
        '                        exoatmElevation * (-12.79 + _
        '                        exoatmElevation * 0.711)))
        '                Else
        '                    refractionCorrection = -20.774 / te
        '                End If
        '                refractionCorrection = refractionCorrection / 3600 
        '            End If
        'end of complex expression


        'beginning of simplified expression
        If (exoatmElevation > 85.0 ) Then
            refractionCorrection = 0.0 
        Else
            te = Math.Tan(degToRad(exoatmElevation))
            If (exoatmElevation > 5.0 ) Then
                refractionCorrection = 58.1 / te - 0.07 / (te * te * te) + _
                    0.000086 / (te * te * te * te * te)
            ElseIf (exoatmElevation > -0.575) Then
                step1 = (-12.79 + exoatmElevation * 0.711)
                step2 = (103.4 + exoatmElevation * (step1))
                step3 = (-518.2 + exoatmElevation * (step2))
                refractionCorrection = 1735.0  + exoatmElevation * (step3)
            Else
                refractionCorrection = -20.774 / te
            End If
            refractionCorrection = refractionCorrection / 3600.0 
        End If
        'end of simplified expression


        solarzen = zenith - refractionCorrection

        '            If (solarZen < 108 ) Then
        solarazimuth = azimuth
        solarelevation = 90.0  - solarzen
        '              If (solarZen < 90 ) Then
        '                coszen = math.cos(degToRad(solarZen))
        '              Else
        '                coszen = 0 
        '              End If
        '            Else    ''' do not report az & el after astro twilight
        '              solarazimuth = -999999
        '              solarelevation = -999999
        '              coszen = -999999
        '            End If

    End Sub


    '''''''''''''''''''''halper functions ''''''''''''''''''''''''''''''' 

    Public Function astronomical_dawn(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
            ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
            ByVal dlstime As Double) As Date

        Return dawn(lat, lon, year, month, day, timezone, dlstime, ConAstronomical)
    End Function

    'nautical
    Public Function nautical_dawn(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
            ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
            ByVal dlstime As Double) As Date

        Return dawn(lat, lon, year, month, day, timezone, dlstime, conNautical)
    End Function


    'civil
    Public Function civil_dawn(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
        ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
        ByVal dlstime As Double) As Date

        Return dawn(lat, lon, year, month, day, timezone, dlstime, conCivil)
    End Function

    '''''

    Public Function astronomical_dusk(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
                      ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
                      ByVal dlstime As Double) As Date
        Return dusk(lat, lon, year, month, day, timezone, dlstime, ConAstronomical)
    End Function


    Public Function nautical_dusk(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
                  ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
                  ByVal dlstime As Double) As Date
        Return dusk(lat, lon, year, month, day, timezone, dlstime, conNautical)
    End Function


    Public Function civil_dusk(ByVal lat As Double, ByVal lon As Double, ByVal year As Double, _
                  ByVal month As Double, ByVal day As Double, ByVal timezone As Double, _
                  ByVal dlstime As Double) As Date
        Return dusk(lat, lon, year, month, day, timezone, dlstime, conCivil)
    End Function

End Class