Attribute VB_Name = "Animations"
Option Explicit
Option Compare Text
Option Private Module

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


'@AUTHOR: ROBERT TODAR
'@LICENCE: MIT

'DEPENDENCIES
' - REFERENCE SET FOR 'MICROSOFT SCRIPTING RUNTIME' FOR Scripting.Dictionary
' - MUST HAVE API'S ABOVE.
' - MOST OF THESE FUNCTIONS RELY ON EACH OTHER, THIS MODULE SHOULD STAY IN TACT.

'PUBLIC METHODS/FUNCTIONS
' - Transition
' - Effect
' - MicroTimer

'PRIVATE METHODS/FUNCTIONS
' - AllTransitionsComplete
' - TransitionComplete
' - IncrementElement
' - easeInAndOut

'NOTES:
' - CREATED THIS TO MAKE TRANSITIONS AND ANIMATIONS IN USERFORMS.
' - A LITTLE SIMIALR TO CSS. KINDA :)
' - ALSO CAN ANIMATE THE USERFORM AS WELL (SUCH AS A SLIDE IN EFFECT)

'TODO:
' - CHANGE THE WAY THAT THE EFFECT IS CALLED, THAT WAY THERE MIGHT BE AN OPTION OF WHAT TYPE
' - OF EFFECT, SUCH AS LINER, EASE-OUT, EASE-IN, POSSIBLY EVEN Bezier CURVE.
'
' - NEED TO ADD A FUNCTION FOR FINDING THE USERFORM FOR REFRESH, CURRENTLY JUST GRABS IT FROM
' - THE FIRST ELEMENT THAT IS ADDED. WORKS FOR NOW, BUT NOT VERY DYNAMIC.
'
' - CURRENTLY HAVE A SLEEP HARDCODED, SHOULD LOOK INTO TESTING DIFFERENT THINGS TO SEE IF IT
' - CAN HELP REDUCE FLASHING. AGAIN, WORKS FOR NOW, BUT SHOULD BE A BETTER WAY.

'EXAMPLES (IN A USERFORM):
'
' 'SINGLE EFFECT
'  Transition Effect(box, "left", 200, 300)
'
' 'MULTIPLE EFFECTS AT ONCE
'  Transition Effect(sidebar, "width", 0, 500) _
'           , Effect(box, "Top", Me.InsideHeight - box.Height, 1000) _
'           , Effect(box2, "Top", 0, 100) _
'           , Effect(GoButton, "fontsize", 12, 1000) _
'           , Effect(me, "Top", 20, 2000)
'


'******************************************************************************************
' PUBLIC METHODS/FUNCTIONS
'******************************************************************************************
Public Sub Transition(ParamArray Elements() As Variant)
        
        'CAPTURE THE FORM
        Dim form As MSForms.UserForm
        Set form = Elements(LBound(Elements, 1))("form")
        
        MicroTimer True
        Do
            
            'INCREMENT EACH ELEMENTS DESTINATION
            Dim Index As Integer
            For Index = LBound(Elements, 1) To UBound(Elements, 1)

                IncrementElement Elements(Index), MicroTimer

            Next Index
            
            'SLEEP NEEDED TO SLOW DOWN THE FRAMERATE, OTHERWISE FLASHES A LOT
            Sleep 40
            form.Repaint
            
            'CHECK TO SEE IF ALL ARE COMPLETE
        Loop Until AllTransitionsComplete(CVar(Elements))

End Sub

Public Function Effect(obj As Object, Property As String, Destination As Double, MilSecs As Double) As Scripting.Dictionary
    
    Dim Temp As New Scripting.Dictionary
    
    Set Temp("obj") = obj
    Temp("property") = Property
    Temp("startValue") = CallByName(obj, Property, VbGet)
    Temp("destination") = Destination
    Temp("travel") = Destination - Temp("startValue")
    Temp("milSec") = MilSecs
    Temp("complete") = False
    
    On Error GoTo Catch:
    Set Temp("form") = obj.Parent
    
    Set Effect = Temp
    Exit Function
Catch:
    Set Temp("form") = obj
    Resume Next
    
End Function


Public Function MicroTimer(Optional StartTime As Boolean = False) As Double
   
    ' uses Windows API calls to the high resolution timer
    Static dTime As Double
    
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    
    MicroTimer = 0

    'get frequency
    If cyFrequency = 0 Then getFrequency cyFrequency
    
    'get ticks
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    
    'calc seconds
    If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
    
    If StartTime = True Then
        dTime = MicroTimer
        MicroTimer = 0
    Else
        MicroTimer = (MicroTimer - dTime) * 1000  'CONVERT TO MILSECS
    End If
    
End Function

'******************************************************************************************
' PRIVATE METHODS/FUNCTIONS
'******************************************************************************************
Private Function AllTransitionsComplete(Elements As Variant) As Boolean
    
    Dim El As Object
    Dim Index As Integer
    
    For Index = LBound(Elements, 1) To UBound(Elements, 1)
                
        Set El = Elements(Index)
                
        If Not TransitionComplete(El) Then
            AllTransitionsComplete = False
            Exit Function
        End If
                
    Next Index
    
    AllTransitionsComplete = True
    
End Function

Private Function TransitionComplete(ByVal El As Scripting.Dictionary) As Boolean
    
    If El("destination") = CallByName(El("obj"), El("property"), VbGet) Then
        TransitionComplete = True
    End If
    
End Function

Private Function IncrementElement(ByVal El As Scripting.Dictionary, CurrentTime As Double) As Boolean

    Dim IncrementValue As Double
    Dim CurrentValue As Double
    
    If TransitionComplete(El) Then
        Exit Function
    End If
    
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
