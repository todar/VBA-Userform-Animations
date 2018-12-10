# VBA Userform Transitions and Animations
Tools to have transitions and animations with Userforms and Userform controls
Similar to CSS Transitions and Animations (kinda lol).

# Usage
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
# Output
![](UserformAnimation.gif)

# Public Methods/Functions
- Transition
- Effect
- MicroTimer

# Private Methods/Functions
- AllTransitionsComplete
- TransitionComplete
- IncrementElement
- easeInAndOut

# TODO:
- Change The Way That The Effect Is Called, That Way There Might Be An Option Of What Type Of Effect, Such As Liner, Ease-Out, Ease-In, Possibly Even Bezier Curve.

- Need To Add A Function For Finding The Userform For Refresh, Currently Just Grabs It From The First Element That Is Added. Works For Now, But Not Very Dynamic.


- Currently Have A Sleep Hardcoded, Should Look Into Testing Different Things To See If It can Help Reduce Flashing. Again, Works For Now, But Should Be A Better Way.


