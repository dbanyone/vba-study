VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub forNextStep()

    For i = 2 To 20 Step 2
    
        MsgBox "Repetiu " & i & " vezes!"
    
    Next i

End Sub

Sub forNextStepII()

    For i = 20 To 1 Step -2
    
        MsgBox "Numero: " & i
    
    Next i

End Sub
