With these function you can calculate delta, gamma, vega, theta, rho

Option Explicit

 

Public Function delta_put(s As Double, k As Double, r As Double, t As Double, vol As Double) As Double

 

Dim d1 As Double

Dim N_d1 As Double

 

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))

N_d1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, True)

 

delta_put = N_d1 - 1

 

End Function

 

Public Function delta_call(s As Double, k As Double, r As Double, t As Double, vol As Double) As Double

 

Dim d1 As Double

Dim N_d1 As Double

 

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))

N_d1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, True)

 

delta_call = N_d1

 

End Function

 

Public Function vega_call_put(s As Double, k As Double, r As Double, t As Double, vol As Double) As Double

 

Dim d1 As Double

 

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))

vega_call_put = (s * Sqr(t) * 1 / Sqr(2 * WorksheetFunction.Pi()) * Exp(-d1 ^ (2) / 2)) / 100

 

End Function

 

Public Function theta_call(s As Double, k As Double, r As Double, t As Double, vol As Double) As Double

 

Dim d1 As Double

Dim d2 As Double

Dim N_derivative_d1 As Double

Dim N_d2 As Double

 

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))

d2 = d1 - vol * Sqr(t)

N_derivative_d1 = 1 / Sqr(2 * WorksheetFunction.Pi()) * Exp(-d1 ^ (2) / 2)

N_d2 = Application.WorksheetFunction.Norm_Dist(d2, 0, 1, True)

 

theta_call = -(s * N_derivative_d1 * vol) / (2 * Sqr(t)) - r * k * Exp(-r * t) * N_d2

 

End Function

 

Public Function theta_put(s As Double, k As Double, r As Double, t As Double, vol As Double) As Double

 

Dim d1 As Double

Dim d2 As Double

Dim N_derivative_d1 As Double

Dim N_negative_d2 As Double

 

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))

d2 = d1 - vol * Sqr(t)

N_derivative_d1 = 1 / Sqr(2 * WorksheetFunction.Pi()) * Exp(-d1 ^ (2) / 2)

N_negative_d2 = Application.WorksheetFunction.Norm_Dist(-d2, 0, 1, True)

 

theta_put = -(s * N_derivative_d1 * vol) / (2 * Sqr(t)) + r * k * Exp(-r * t) * N_negative_d2

 

End Function

 

Public Function rho_put(s As Double, k As Double, r As Double, t As Double, vol As Double) As Double

 

Dim d1 As Double

Dim d2 As Double

Dim N_negative_d2 As Double

 

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))

d2 = d1 - vol * Sqr(t)

N_negative_d2 = Application.WorksheetFunction.Norm_Dist(-d2, 0, 1, True)

 

rho_put = (-k * t * Exp(-r * t) * N_negative_d2) / 100

 

End Function

 

Public Function rho_call(s As Double, k As Double, r As Double, t As Double, vol As Double) As Double

 

Dim d1 As Double

Dim d2 As Double

Dim N_d2 As Double

 

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))

d2 = d1 - vol * Sqr(t)

N_d2 = Application.WorksheetFunction.Norm_Dist(d2, 0, 1, True)

 

rho_call = (k * t * Exp(-r * t) * N_d2) / 100

 

End Function

 

Public Function gamma_call_put(s As Double, k As Double, r As Double, t As Double, vol As Double) As Double

 

Dim d1 As Double

Dim N_derivative_d1 As Double

 

d1 = (Application.WorksheetFunction.Ln(s / k) + (r + vol ^ (2) / 2) * t) / (vol * Sqr(t))

N_derivative_d1 = 1 / Sqr(2 * WorksheetFunction.Pi()) * Exp(-d1 ^ (2) / 2)

 

gamma_call_put = N_derivative_d1 / (s * vol * Sqr(t))

 

End Function