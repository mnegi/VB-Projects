Attribute VB_Name = "basMathsFormulae"


'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'   THIS MODULE IS TOTALLY BASED ON THE THE         *
'   MATHEMATICAL FUNCTIONS. I WROTE THIS FUNCTION   *
'   BECAUSE I FOUND SOME PROBLEMS WHILE USING THE   *
'   TRIGONOMETRIC FUNCTIONS. WHAT HAPPENS WHEN WE   *
'   CALL ANY FUNCTION SUCH AS Sin(x),Cos(x) OR ANY  *
'   OTHER, IT CONSIDER ANGLE IN RADIANS , BUT WE    *
'   ALWAYS ASSUME THAT, TO BE IN DEGREES. THIS      *
'   MODULE DOES THE REQUIRED CONVERSION.            *
'   AND ALSO IT GIVES YOU SOME DERIVED FUNCYIONS    *
'   SUCH AS INVERSE FUNCTIONS.                      *
'                                                   *
'****************************************************


Public Const Pi = 3.14159265358979
'Sin
Public Function Sine(x As Double) As Double
Sine = Sin((Pi / 180) * CDbl(x))
End Function
'Cos
Public Function CosTheta(x As Double) As Double
CosTheta = Cos((Pi / 180) * CDbl(x))
End Function
'Tangent
Public Function Tangent(x As Double) As Double
Tangent = Tan((Pi / 180) * CDbl(x))
End Function
'Cosecant
Public Function Cosecant(x As Double) As Double
Cosecant = CDbl(1 / Sin((Pi / 180) * CDbl(x)))
End Function
'Secant
Public Function Secant(x As Double) As Double
Secant = CDbl(1 / Cos((Pi / 180) * CDbl(x)))
End Function
'Cotangent
Public Function Cotangent(x As Double) As Double
Cotangent = CDbl(1 / Tan((Pi / 180) * CDbl(x)))
End Function
'Inverse Tangent
Public Function ITan(x As Double) As Double
ITan = CDbl((180 / Pi) * Atn(x))
End Function
'Inverse Sin
Public Function ISin(x As Double) As Double
ISin = CDbl((180 / Pi) * Atn(x / Sqr(-x * x + 1)))
End Function
'Inverse Cos
Public Function ICos(x As Double) As Double
ICos = CDbl((180 / Pi) * Atn(-x / Sqr(-x * x + 1))) + 2 * CDbl((180 / Pi) * Atn(1))
End Function
'Inverse  Cosecant
Public Function ICosec(x As Double) As Double
ICosec = CDbl((180 / Pi) * Atn(x / Sqr(x * x - 1))) + Sgn((x) - 1) * (2 * CDbl((180 / Pi) * Atn(1)))
End Function
'Inverse Secant
Public Function ISec(x As Double) As Double
ISec = CDbl((180 / Pi) * Atn(x / Sqr(x * x - 1))) + Sgn((x) - 1) * (2 * CDbl((180 / Pi) * Atn(1)))
End Function
'Inverse Cotangent
Public Function ICot(x As Double) As Double
ICot = CDbl((180 / Pi) * Atn(x)) + 2 * CDbl((180 / Pi) * Atn(1))
End Function
'Hyperbolic Sin
Public Function HSin(x As Double) As Double
HSin = CDbl((Exp(x) - Exp(-x)) / 2)
End Function
'Hyperbolic Cos
Public Function HCos(x As Double) As Double
HCos = CDbl((Exp(x) + Exp(-x)) / 2)
End Function
'Hyperbolic Tangent
Public Function HTan(x As Double) As Double
HTan = CDbl((Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x)))
End Function
'Hyperbolic Cosecant
Public Function HCosec(x As Double) As Double
HCosec = CDbl(2 / (Exp(x) + Exp(-x)))
End Function
'Hyperbolic Secant
Public Function HSec(x As Double) As Double
HSec = CDbl(2 / (Exp(x) - Exp(-x)))
End Function
'Hyperbolic Cotangent
Public Function HCotan(x As Double) As Double
HCotan = CDbl((Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x)))
End Function
'Inverse Hyperbolic Sine
Public Function IHSin(x As Double) As Double
IHSin = CDbl(Log(x + Sqr(x * x + 1)))
End Function
'Inverse Hyperbolic Cos
Public Function IHCos(x As Double) As Double
IHCos = CDbl(Log(x + Sqr(x * x - 1)))
End Function
'Inverse Hyperbolic Tangent
Public Function IHTan(x As Double) As Double
IHTan = CDbl(Log((1 + x) / (1 - x)) / 2)
End Function
'Inverse Hyperbolic Secant
Public Function IHSec(x As Double) As Double
IHSec = CDbl(Log((Sqr(-x * x + 1) + 1) / x))
End Function
'Inverse Hyperbolic Cosecant
Public Function IHCosec(x As Double) As Double
IHCosec = CDbl(Log((Sgn(x) * Sqr(x * x + 1) + 1) / x))
End Function
'Inverse Hyperbolic Cotangent
Public Function IHCot(x As Double) As Double
IHCot = CDbl(Log((Sgn(x) * Sqr(x * x + 1) + 1) / x))
End Function


'********************************************************************
'
'   OTHER USEFUL FUNCTIONS
'
'********************************************************************
Public Function Power(x As Double, Y As Double) As Double
Power = x ^ Y
End Function

Public Function LogN(Base As Double, x As Double) As Double
LogN = Log(x) / Log(Base)
End Function



