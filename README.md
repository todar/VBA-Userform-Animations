# VBA-Userform-Transitions
Tools to have transitions and animations with Userforms and Userform controls

`'EXAMPLE (IN A USERFORM WITH TWO LABELS AND A COMMAND BUTTON)
Private Sub CommandButton1_Click()
    
    'SIMPLE EXAMPLES MOVING LABELS
    Transition Effect(Label1, "Top", Me.InsideHeight - Label1.Height, 200), _
               Effect(Label2, "left", Me.Width - Label2.Width, 500)
             
    'ADDING MORE TRANSITION FUNCTIONS CREATES AN ANIMATION EFFECT
    Transition Effect(Label1, "width", 25, 200), Effect(Label1, "left", 0, 1000)
    
    'CAN ALSO WORK WITH FONTSIZE
    Transition Effect(CommandButton1, "fontsize", 16, 200)
    
    'AS WELL AS USERFORMS
    Transition Effect(Me, "TOP", 20, 1000)
    
End Sub`
