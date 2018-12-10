# VBA Userform Transitions and Animations
Tools to have transitions and animations with Userforms and Userform controls
Similar to CSS Transitions and Animations (kinda lol).

# Example
```vb
'EXAMPLE (IN A USERFORM WITH TWO LABELS AND A COMMAND BUTTON)
Private Sub CommandButton1_Click()
    
    'APPLYING SINGLE EFFECT
    Transition Effect(Label1, "left", 200, 300)
    
    'SINGLE EFFECT, SHOWING USERFORMS CAN BE EFFECTED AS WELL
    Transition Effect(Me, "TOP", 500, 1000)
    
    'CAN ALSO WORK WITH FONTSIZE
    Transition Effect(CommandButton1, "fontsize", 16, 200)
    
    'APPLYING MULTIPLE EFFECTS AT ONCE
    Transition Effect(Label1, "Top", Me.InsideHeight - Label1.Height, 200) _
             , Effect(Label2, "left", Me.Width - Label2.Width, 500) _
             , Effect(me, "Top", 0, 2000)
             
    'CONSECTIVE EFFECTS IN A ROW MAKE IT WORK LIKE AN ANIMATION
    Transition Effect(Label1, "Top", 20, 1000)
    Transition Effect(Label1, "Left", 16, 200)
    Transition Effect(Label1, "Top", 0, 200)
   
End Sub
```

# PUBLIC METHODS/FUNCTIONS
- Transition
- Effect
- MicroTimer

# PRIVATE METHODS/FUNCTIONS
- AllTransitionsComplete
- TransitionComplete
- IncrementElement
- easeInAndOut

# TODO:
- CHANGE THE WAY THAT THE EFFECT IS CALLED, THAT WAY THERE MIGHT BE AN OPTION OF WHAT TYPE
OF EFFECT, SUCH AS LINER, EASE-OUT, EASE-IN, POSSIBLY EVEN Bezier CURVE.

- NEED TO ADD A FUNCTION FOR FINDING THE USERFORM FOR REFRESH, CURRENTLY JUST GRABS IT FROM
THE FIRST ELEMENT THAT IS ADDED. WORKS FOR NOW, BUT NOT VERY DYNAMIC.

- CURRENTLY HAVE A SLEEP HARDCODED, SHOULD LOOK INTO TESTING DIFFERENT THINGS TO SEE IF IT
CAN HELP REDUCE FLASHING. AGAIN, WORKS FOR NOW, BUT SHOULD BE A BETTER WAY.

