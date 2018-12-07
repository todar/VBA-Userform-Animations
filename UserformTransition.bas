Attribute VB_Name = "UserformTransition"
Option Explicit

'Sleep FUNCTIONLITY
#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'USED FOR MICRO TIMER
#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias _
    "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias _
    "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" Alias _
    "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias _
    "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If


Public Sub Transition(ParamArray Elements() As Variant)
        
        Dim index As Integer
        Dim allComplete As Boolean
        'Dim CurrentTime As Double
        
        Dim form As MSForms.UserForm
        
        Set form = Elements(LBound(Elements, 1))("form")
        
        MicroTimer True
        Do
            For index = LBound(Elements, 1) To UBound(Elements, 1)
                
                'CHECK TO SEE IF IT IS AT IT'S DESTINATION ALREADY
                If TransitionComplete(Elements(index)) = False Then
                    
                    'INCREMENT DESTINATION
                    IncrementEl Elements(index), MicroTimer
                    
                End If
                
            Next index
            
            Sleep 40 'SLOWS DOWN THE FRAMERATE, OTHERWISE FLASHES
            form.Repaint
            
            'CHECK TO SEE IF ALL ARE COMPLETE
            allComplete = True
            For index = LBound(Elements, 1) To UBound(Elements, 1)
                
                If TransitionComplete(Elements(index)) = False Then
                    allComplete = False
                End If
                
            Next index
          
            If allComplete Then Exit Do

        Loop

End Sub

Private Function TransitionComplete(ByVal El As Scripting.Dictionary) As Boolean
    
    'DO ACTIONS WHEN GOING NEGATIVE
    TransitionComplete = El("destination") = CallByName(El("obj"), El("property"), VbGet)
    
End Function

Private Function IncrementEl(ByVal El As Scripting.Dictionary, CurrentTime As Double) As Boolean
    
    Dim IncrementValue As Double
    Dim CurrentValue As Double
    
    IncrementValue = easeInAndOut(CurrentTime, El("startValue"), El("travel"), El("milSec"))
    
    If El("travel") < 0 Then
        
        If Math.Round(IncrementValue, 4) < El("destination") Then
            CallByName El("obj"), El("property"), VbLet, El("destination")
        Else
            CallByName El("obj"), El("property"), VbLet, IncrementValue
        End If
        
    Else
    
        If Math.Round(IncrementValue, 4) > El("destination") Then
            CallByName El("obj"), El("property"), VbLet, El("destination")
        Else
            CallByName El("obj"), El("property"), VbLet, IncrementValue
        End If
        
    End If
    
End Function

Public Function Effect(Obj As Object, Property As String, Destination As Double, MilSecs As Double) As Scripting.Dictionary
    
    Dim Temp As New Scripting.Dictionary
    
    Set Temp("obj") = Obj
    Temp("property") = Property
    Temp("startValue") = CallByName(Obj, Property, VbGet)
    Temp("destination") = Destination
    Temp("travel") = Destination - Temp("startValue")
    Temp("milSec") = MilSecs
    Temp("complete") = False
    
    On Error GoTo Catch:
    Set Temp("form") = Obj.Parent
    
    Set Effect = Temp
    Exit Function
Catch:
    Set Temp("form") = Obj
    Resume Next
End Function


Public Function MicroTimer(Optional StartTime As Boolean = False) As Double
   
    ' uses Windows API calls to the high resolution timer
    Static dTime As Double
    
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    
    MicroTimer = 0

    ' get frequency
    If cyFrequency = 0 Then getFrequency cyFrequency
    
    ' get ticks
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    
    ' calc seconds
    If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
    
    If StartTime = True Then
        dTime = MicroTimer
        MicroTimer = 0
    Else
        'MicroTimer = RoundUp((MicroTimer - dTime) * 1000)  'CONVERT TO MILSECS
        MicroTimer = (MicroTimer - dTime) * 1000  'CONVERT TO MILSECS
    End If
    
End Function


'--legend
'-b start value
'-c DESTINATION - START
'-d total time
'-t current time (the only one that changes)
Private Function easeInAndOut(ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double

    'cubic
    d = d / 2
    t = t / d
    If (t < 1) Then
        easeInAndOut = c / 2 * t * t * t + b
    Else
        t = t - 2
        easeInAndOut = c / 2 * (t * t * t + 2) + b
    End If

End Function

'Function easeInAndOut(ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
'
'    'quartic
'    d = d / 2
'    t = t / d
'
'    If (t < 1) Then
'        easeInAndOut = c / 2 * t * t * t * t + b
'    Else
'        t = t - 2
'        easeInAndOut = -c / 2 * (t * t * t * t - 2) + b
'    End If
'
'
'End Function
'
'Function easeInAndOut3(ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
'
'    'quintic
'    d = d / 2
'    t = t / d
'
'    If (t < 1) Then
'        easeInAndOut3 = c / 2 * t * t * t * t * t + b
'    Else
'        t = t - 2
'        easeInAndOut3 = c / 2 * (t * t * t * t * t + 2) + b
'    End If
'
'
'End Function
'
'Function easeInAndOut4(ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
'
'    'sinusoidal
'    easeInAndOut4 = -c / 2 * (Math.Cos(Application.WorksheetFunction.pi * t / d) - 1) + b
'
'
'End Function
'
'Function easeInAndOut5(ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
'
'    'circular
'    d = d / 2
'    t = t / d
'
'    If (t < 1) Then
'        easeInAndOut5 = -c / 2 * (Math.Sqr(1 - t * t) - 1) + b
'    Else
'        t = t - 2
'        easeInAndOut5 = c / 2 * (Math.Sqr(1 - t * t) + 1) + b
'    End If
'
'End Function


'Function Bezier4(p1 As XYZ, p2 As XYZ, p3 As XYZ, p4 As XYZ, mu#) As XYZ
''   Four control point Bezier interpolation
''   mu ranges from 0 to 1, start to end of curve
'Dim mum1#, mum13#, mu3#
'Dim p As XYZ
'
'   mum1 = 1 - mu
'   mum13 = mum1 * mum1 * mum1
'   mu3 = mu * mu * mu
'
'   p.x = mum13 * p1.x + 3 * mu * mum1 * mum1 * p2.x + 3 * mu * mu * mum1 * p3.x + mu3 * p4.x
'   p.y = mum13 * p1.y + 3 * mu * mum1 * mum1 * p2.y + 3 * mu * mu * mum1 * p3.y + mu3 * p4.y
''   p.z = mum13 * p1.z + 3 * mu * mum1 * mum1 * p2.z + 3 * mu * mu * mum1 * p3.z + mu3 * p4.z
'
'   Bezier4 = p
'End Function
