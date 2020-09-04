Attribute VB_Name = "alpha_perlin_math"




'https://mrl.nyu.edu/~perlin/noise/
'https://adrianb.io/2014/08/09/perlinnoise.html
'https://www.codeproject.com/Articles/785084/A-generic-lattice-noise-algorithm-an-evolution-of
'https://www.khanacademy.org/computing/computer-programming/programming-natural-simulations/programming-noise/a/perlin-noise

Sub tst()

Dim x As Double

x = 2.1

x = floor(x)

End Sub


'perlin
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/



'math
'-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/-/

Public Function floor(ByVal value As Double)
    floor = Application.WorksheetFunction.RoundDown(value, 0)
End Function



