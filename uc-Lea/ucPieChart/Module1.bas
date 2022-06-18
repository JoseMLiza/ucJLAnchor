Attribute VB_Name = "Module1"
Option Explicit

Private Const PI As Single = 3.14159265358979
  
'Calcula en angulo entre dos puntos
Public Function MATH_GetAngle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
      
    MATH_GetAngle = CSng(MATH_Atan2(CDbl(Y2) - CDbl(Y1), CDbl(X2) - CDbl(X1)) * 180 / 3.14)
      
End Function
  
  
  

  
'Función que Devuelve el arco tangente en radianes de dos números
Public Function MATH_Atan2(X As Double, Y As Double) As Double
    On Error GoTo ErrOut
    Dim Theta As Double
  
    If (Abs(X) < 0.0000001) Then
         If (Abs(Y) < 0.0000001) Then
               Theta = 0#
  
         ElseIf (Y > 0#) Then
                Theta = 1.5707963267949
  
         Else
                Theta = -1.5707963267949
  
         End If
  
    Else
         Theta = Atn(Y / X)
    
         If (X < 0) Then
               If (Y >= 0#) Then
                    Theta = 3.14159265358979 + Theta
  
                Else
                   Theta = Theta - 3.14159265358979
  
                End If
  
             End If
  
     End If
      
    MATH_Atan2 = Theta
  
ErrOut:
End Function

