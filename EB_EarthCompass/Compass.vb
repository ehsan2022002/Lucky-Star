Public Class Compass

    Public Function LocateInDegree(ByVal sourceLat As Double, ByVal sourceLon As Double, _
                              ByVal DestLat As Double, ByVal DestLon As Double) As Double

        Dim a, b, c, s, t, u, v, w, g, h, i, j As Double

        a = 90 - sourceLat
        b = 90 - DestLat
        c = sourceLon - DestLon
        s = Math.Cos(DToR((a - b) / 2))
        t = Math.Cos(DToR((a + b) / 2))
        u = 1 / Math.Tan(DToR(c / 2))

        v = s * u / t
        w = Math.Atan(v) * 180 / Math.PI

        g = Math.Sin(DToR((a - b) / 2))
        h = Math.Sin(DToR((a + b) / 2))


        i = g * u / h
        j = Math.Atan(i) * 180 / Math.PI

        Return w - j
        'so if we face to north and turn to west in return degress we get it
        'west of south location will get negative

    End Function


    Public Function DistanceInKM(ByVal sourceLat As Double, ByVal sourceLon As Double, _
                              ByVal DestLat As Double, ByVal DestLon As Double) As Double



        Dim a, b, c, o, p, cx, dx As Double
        Const consDgreeLong As Decimal = 111.1949


        a = 90 - sourceLat
        b = 90 - DestLat
        c = sourceLon - DestLon

        o = Math.Cos(DToR(a)) * Math.Cos(DToR(b))
        p = Math.Sin(DToR(a)) * Math.Sin(DToR(b)) * Math.Cos(DToR(c))
        cx = RToD(Math.Acos(o + p))


        dx = cx * consDgreeLong
        Return dx

    End Function



    Private Function DToR(ByVal d As Double) As Double
        Return (d * Math.PI / 180)
    End Function

    Private Function RToD(ByVal d As Double) As Double
        Return (180 / Math.PI * d)
    End Function

End Class
