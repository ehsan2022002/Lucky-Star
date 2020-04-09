Imports Microsoft.VisualBasic

Public Class EB_SolarSystem





    '    Dim two As Double = 2.0
    '    Dim zeroPointTwo As Double = 0.2
    '    Dim quotient As Double = two / zeroPointTwo
    '    Dim doubleRemainder As Double = two Mod zeroPointTwo

    'MsgBox("2.0 is represented as " & two.ToString("G17") _
    '  & vbCrLf & "0.2 is represented as " & zeroPointTwo.ToString("G17") _
    '  & vbCrLf & "2.0 / 0.2 generates " & quotient.ToString("G17") _
    '  & vbCrLf & "2.0 Mod 0.2 generates " _
    '  & doubleRemainder.ToString("G17"))

    '    Dim decimalRemainder As Decimal = 2D Mod 0.2D
    ''MsgBox("2.0D Mod 0.2D generates " & CStr(decimalRemainder))





    '''class main 
    Public MustInherit Class UTDayAndTime
        Private mDay As Short
        Private mMonth As Short
        Private mYear As Short
        Private mHour As Short
        Private mMinute As Short
        Private mRad As Double
        Private mJDofElements As Double
        Private mInclination As Double
        Private mAscNode As Double
        Private mperihelion As Double
        Private mMeanDistance As Double
        Private mdailyMotion As Double
        Private meccentricity As Double
        Private mMeanlongitude As Double

        ''Public Enum Planet As Integer
        ''    JDofelements = 2450680.5
        ''End Enum

        Public Property year() As Short
            Get
                Return mYear
            End Get
            Set(ByVal Value As Short)
                mYear = Value
            End Set
        End Property

        Public Property Month() As Short
            Get
                Return mMonth
            End Get
            Set(ByVal Value As Short)
                mMonth = Value
            End Set
        End Property

        Public Property Day() As Short
            Get
                Return mDay
            End Get
            Set(ByVal Value As Short)
                mDay = Value
            End Set
        End Property

        Public Property UTHour() As Short
            Get
                Return mHour
            End Get
            Set(ByVal Value As Short)
                mHour = Value
            End Set
        End Property

        Public Property Minute() As Short
            Get
                Return mMinute
            End Get
            Set(ByVal Value As Short)
                mMinute = Value
            End Set
        End Property


        Public Property JDofElements() As Double
            Get
                Return mJDofElements
            End Get
            Set(ByVal Value As Double)
                mJDofElements = Value
            End Set
        End Property

        Public Property inclination() As Double
            Get
                Return mInclination
            End Get
            Set(ByVal Value As Double)
                mInclination = Value
            End Set
        End Property

        Public Property ascNode() As Double
            Get
                Return mAscNode
            End Get
            Set(ByVal Value As Double)
                mAscNode = Value
            End Set
        End Property


        Public Property perihelion() As Double
            Get
                Return mperihelion
            End Get
            Set(ByVal Value As Double)
                mperihelion = Value
            End Set
        End Property

        Public Property meanDistance() As Double
            Get
                Return mMeanDistance
            End Get
            Set(ByVal Value As Double)
                mMeanDistance = Value
            End Set
        End Property

        Public Property dailyMotion() As Double
            Get
                Return mdailyMotion
            End Get
            Set(ByVal Value As Double)
                mdailyMotion = Value
            End Set
        End Property

        Public Property eccentricity() As Double
            Get
                Return meccentricity
            End Get
            Set(ByVal Value As Double)
                meccentricity = Value
            End Set
        End Property


        Public Property meanLongitude() As Double
            Get
                Return mMeanlongitude
            End Get
            Set(ByVal Value As Double)
                mMeanlongitude = Value
            End Set
        End Property

        ''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''

        Public ReadOnly Property earthInclination() As Double
            Get
                Return 0
            End Get
        End Property

        Public ReadOnly Property earthAscNode() As Double
            Get
                Return 349.2
            End Get
        End Property

        Public ReadOnly Property earthPerihelion() As Double
            Get
                Return 102.8517
            End Get
        End Property

        Public ReadOnly Property earthMeanDistance() As Double
            Get
                Return 1
            End Get
        End Property

        Public ReadOnly Property earthDailyMotion() As Double
            Get
                Return 0.9855796
            End Get
        End Property

        Public ReadOnly Property earthEccentricity() As Double
            Get
                Return 0.0166967
            End Get
        End Property

        Public ReadOnly Property earthMeanLongitude() As Double
            Get
                Return 328.40353
            End Get
        End Property


        ''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''


        Public Overridable Function Days_to_J2000() As Double
            Dim x As Double

            x = 367 * Me.year - Int(7 * (Me.year + Int((Me.Month + 9) / 12)) / 4) + _
                Int(275 * Me.Month / 9) + Me.Day - 730531.5 + (Me.UTHour + Me.Minute / 60)

            Return x
        End Function

        Private Function DToR(ByVal d As Double) As Double
            Return (d * Math.PI / 180)
        End Function

        Private Function RToD(ByVal d As Double) As Double
            Return (180 / Math.PI * d)
        End Function


        Public Overridable Function Days_from_elements() As Double
            Return (Days_to_J2000()) - (Me.JDofElements - 2451545)
        End Function






        Public Overridable Function Mean_anomaly() As Double
            Return RToD(Mean_anomalyRad())
        End Function
        Public Overridable Function Mean_anomalyRad() As Double
            Return ((DToR(Me.dailyMotion) * Days_from_elements() + _
                        DToR(Me.meanLongitude) - DToR(Me.perihelion)) Mod (2 * Math.PI))
        End Function



        Public Overridable Function True_anomaly() As Double
            Return RToD(True_anomalyRad())
        End Function
        Public Overridable Function True_anomalyRad() As Double
            Dim dblMAR As Double = Mean_anomalyRad()

            Return (dblMAR + (2 * Me.eccentricity - Me.eccentricity ^ 3 / 4 + _
                   5 / 96 * Me.eccentricity ^ 5) * Math.Sin(dblMAR) + _
                   (5 * Me.eccentricity ^ 2 / 4 - 11 / 24 * Me.eccentricity ^ 4) * _
                   Math.Sin(2 * Me.eccentricity) + (13 * Me.eccentricity ^ 3 / 12 - 43 / 64 * _
                   Me.eccentricity ^ 5) * Math.Sin(3 * dblMAR) + 103 / 96 * Me.eccentricity ^ 4 * _
                   Math.Sin(4 * dblMAR) + 1097 / 960 * Me.eccentricity ^ 5 * Math.Sin(5 * dblMAR))

        End Function



        Public Overridable Function Longitude() As Double
            Return RToD(LongitudeRad())
        End Function
        Public Overridable Function LongitudeRad() As Double
            Return ((True_anomalyRad() + DToR(Me.perihelion)) Mod (2 * Math.PI))
        End Function



        Public Overridable Function Radius_Vector() As Double
            Return Me.meanDistance * (1 - Me.eccentricity ^ 2) / (1 + Me.eccentricity * Math.Cos(True_anomalyRad()))
        End Function
        Public Overridable Function Earth_Radius_Vector() As Double
            Return earthMeanDistance * (1 - Me.eccentricity ^ 2) / _
                   (1 + Me.eccentricity * Math.Cos(Earth_True_anomaly))
        End Function


        Public Overridable Function Earth_Mean_anomaly() As Double
            Return RToD(Earth_Mean_anomalyRad())
        End Function

        Public Overridable Function Earth_Mean_longitude() As Double
            Return 328.40353
        End Function

        Public Overridable Function Earth_Mean_longitudeRad() As Double
            Return DToR(Earth_Mean_longitude)
        End Function


        Public Overridable Function Earth_Mean_anomalyRad() As Double

            Dim z As Double
            z = (DToR(Me.earthDailyMotion) * Days_from_elements() + _
                Earth_Mean_longitudeRad() - DToR(Me.earthPerihelion))
            Return z Mod (2 * Math.PI)

        End Function


        Public Overridable Function Earth_True_anomaly() As Double
            Return RToD(Earth_True_anomalyRad())
        End Function
        Public Overridable Function Earth_True_anomalyRad() As Double
            Dim dblEMAR As Double
            dblEMAR = Earth_Mean_anomalyRad()

            Return (dblEMAR + (2 * Me.earthEccentricity - Me.earthEccentricity ^ 3 / 4) * _
                    Math.Sin(dblEMAR) + 5 * Me.earthEccentricity ^ 2 / 4 * Math.Sin(2 * dblEMAR) + 13 _
                    * Me.earthEccentricity ^ 3 / 12 * Math.Sin(3 * dblEMAR))

        End Function


        Public Overridable Function Earth_Longitude() As Double
            Return RToD(Earth_LongitudeRad())
        End Function
        Public Overridable Function Earth_LongitudeRad() As Double
            Return ((Earth_True_anomalyRad() + DToR(Me.earthPerihelion)) Mod (2 * Math.PI))
        End Function


        Public Overridable Function Helio_long() As Double
            Return RToD(Helio_longRad())
        End Function
        Public Overridable Function Helio_longRad() As Double
            Return Math.Atan2((Math.Sin(LongitudeRad() - DToR(Me.ascNode)) * Math.Cos(DToR(Me.inclination))), _
                    (Math.Cos(LongitudeRad() - DToR(Me.dailyMotion)))) + DToR(Me.ascNode)

            ''Return Math.Atan2(Math.Cos(LongitudeRad() - (DToR(Me.dailyMotion))) _
            ''    , (Math.Sin(LongitudeRad() - (DToR(Me.dailyMotion))) * _
            ''    Math.Cos(DToR(Me.inclination)))) + DToR(Me.ascNode)
        End Function


        Public Overridable Function Helio_lat() As Double
            Return RToD(Helio_latRad())
        End Function
        Public Overridable Function Helio_latRad() As Double
            'Helio lat (phi)''=ASIN(SIN(G15-G6)*SIN(G5))
            Return Math.Asin(Math.Sin(LongitudeRad() - DToR(Me.ascNode)) * Math.Sin(DToR(Me.inclination)))
        End Function


        'dist from Sun (au)
        Public Overridable Function dist_from_Sun() As Double
            Return (Radius_Vector() * Math.Cos(Helio_latRad()))
        End Function


        Public Overridable Function lambda() As Double
            Return RToD(lambdaRad())
        End Function
        Public Overridable Function lambdaRad() As Double
            ''Return (Math.Atan(Me.Earth_Radius_Vector * Math.Sin(Helio_longRad() - LongitudeRad()) / _
            ''        (dist_from_Sun() - Earth_Radius_Vector() * Math.Cos(Helio_longRad() - LongitudeRad()))) + _
            ''        Helio_longRad()) Mod (2 * Math.PI)

            '=MOD(ATAN(B11* SIN(I15-C9)/(H16-B11*COS(I15-C9)))+PI()+I15;2*PI())
            Return (Math.Atan(Me.dist_from_Sun * Math.Sin(Earth_LongitudeRad() - Helio_longRad()) / _
                    (Earth_Radius_Vector() - dist_from_Sun() * Math.Cos(Earth_LongitudeRad() - Helio_longRad()))) + _
                     Math.PI + Earth_LongitudeRad()) Mod (2 * Math.PI)

        End Function


        Public Overridable Function bata() As Double
            Return RToD(bataRad())
        End Function
        Public Overridable Function bataRad() As Double
            ''Return Math.Atan(dist_from_Sun() * Math.Tan(Helio_latRad) * Math.Sin(lambdaRad() - Helio_latRad()) / _
            ''            (Earth_Radius_Vector() * Math.Sin(Helio_longRad() - LongitudeRad())))
            ''=ATAN(B11*TAN(C10)*SIN(C12-C9)/(H16*SIN(C9-I15)))

            Return Math.Atan(dist_from_Sun() * Math.Tan(Helio_latRad) * Math.Sin(lambdaRad() - Helio_longRad()) / _
                        (Earth_Radius_Vector() * Math.Sin(Helio_longRad() - Earth_LongitudeRad())))

        End Function



        Public Overridable Function obliquity_of_ecliptic() As Double
            Return 23.429292
        End Function
        Public Overridable Function obliquity_of_eclipticRad() As Double
            Return DToR(obliquity_of_ecliptic())
        End Function


        Public Overridable Function alpha() As Double
            Return RToD(alphaRad())
        End Function
        Public Overridable Function alphaRad() As Double

            Dim x, y, z As Double

            '=MOD(ATAN2(COS(C12);SIN(C12)*COS(C14)-TAN(C13)*SIN(C14));2*PI())

            z = Math.Sin(lambdaRad) * Math.Cos(obliquity_of_eclipticRad) - Math.Tan(bataRad) * Math.Sin(obliquity_of_eclipticRad)

            y = Math.Cos(lambdaRad)
            x = (2 * Math.PI)

            Return (Math.Atan2(z, y)) Mod x

        End Function






        'delta
        Public Overridable Function delta() As Double
            Return RToD(deltaRad())
        End Function
        Public Overridable Function deltaRad() As Double

            '=SIN(C13)*COS(C14)+COS(C13)*SIN(C14)*SIN(C12)

            Return Math.Asin(Math.Sin(bataRad()) * Math.Cos(obliquity_of_eclipticRad) + Math.Cos(bataRad) * _
                   Math.Sin(obliquity_of_eclipticRad) * Math.Sin(lambdaRad))
        End Function


    End Class
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



    Public Class Mercury
        Inherits UTDayAndTime

        Public Sub New()
            MyBase.JDofElements = 2450680.5
            MyBase.inclination = 7.00507
            MyBase.ascNode = 48.3339
            MyBase.perihelion = 77.454
            MyBase.meanDistance = 0.3870978
            MyBase.dailyMotion = 4.092353
            MyBase.eccentricity = 0.2056324
            MyBase.meanLongitude = 314.42369
        End Sub




    End Class


    Public Class Venus
        Inherits UTDayAndTime

        Public Sub New()
            MyBase.JDofElements = 2450680.5
            MyBase.inclination = 3.39472
            MyBase.ascNode = 76.6889
            MyBase.perihelion = 131.761
            MyBase.meanDistance = 0.7233238
            MyBase.dailyMotion = 1.602158
            MyBase.eccentricity = 0.0067933
            MyBase.meanLongitude = 236.94045
        End Sub



    End Class


    Public Class Mars
        Inherits UTDayAndTime
        Public Sub New()
            MyBase.JDofElements = 2450680.5
            MyBase.inclination = 1.84992
            MyBase.ascNode = 49.5664
            MyBase.perihelion = 336.0882
            MyBase.meanDistance = 1.5236365
            MyBase.dailyMotion = 0.5240613
            MyBase.eccentricity = 0.0934231
            MyBase.meanLongitude = 262.42784
        End Sub




    End Class


    Public Class Jupiter
        Inherits UTDayAndTime
        Public Sub New()
            MyBase.JDofElements = 2450680.5
            MyBase.inclination = 1.30463
            MyBase.ascNode = 100.4713
            MyBase.perihelion = 15.6978
            MyBase.meanDistance = 5.202597
            MyBase.dailyMotion = 0.08309618
            MyBase.eccentricity = 0.0484646
            MyBase.meanLongitude = 322.55983
        End Sub




    End Class


    Public Class Saturn
        Inherits UTDayAndTime
        Public Sub New()
            MyBase.JDofElements = 2450681
            MyBase.inclination = 2.48524
            MyBase.ascNode = 113.6358
            MyBase.perihelion = 88.863
            MyBase.meanDistance = 9.5719
            MyBase.dailyMotion = 0.03328656
            MyBase.eccentricity = 0.0531651
            MyBase.meanLongitude = 20.95759
        End Sub





    End Class



    Public Class Uranus
        Inherits UTDayAndTime
        Public Sub New()
            MyBase.JDofElements = 2450681
            MyBase.inclination = 0.77343
            MyBase.ascNode = 74.0954
            MyBase.perihelion = 175.6807
            MyBase.meanDistance = 19.30181
            MyBase.dailyMotion = 0.01162295
            MyBase.eccentricity = 0.0428959
            MyBase.meanLongitude = 303.18967
        End Sub




    End Class


    Public Class Neptune
        Inherits UTDayAndTime
        Public Sub New()
            MyBase.JDofElements = 2450680.5
            MyBase.inclination = 1.7681
            MyBase.ascNode = 131.7925
            MyBase.perihelion = 7.206
            MyBase.meanDistance = 30.26664
            MyBase.dailyMotion = 0.005919282
            MyBase.eccentricity = 0.0102981
            MyBase.meanLongitude = 299.8641
        End Sub




    End Class


    Public Class Pluto
        Inherits UTDayAndTime
        Public Sub New()
            MyBase.JDofElements = 2450680.5
            MyBase.inclination = 17.12137
            MyBase.ascNode = 110.3833
            MyBase.perihelion = 224.8025
            MyBase.meanDistance = 39.5804
            MyBase.dailyMotion = 0.003958072
            MyBase.eccentricity = 0.2501272
            MyBase.meanLongitude = 235.7656
        End Sub


    End Class

End Class
