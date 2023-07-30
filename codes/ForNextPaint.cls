VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub forNextPintarCelular()

    For i = 1 To 20 Step 1
    
        If Range("A" & i).Value >= 10 Then
        
            Range("A" & i).Font.Color = vbWhite
            Range("A" & i).Font.Size = 20
            Range("A" & i).Interior.Color = vbBlue
            Range("A" & i).HorizontalAlignment = xlCenter
            
        ElseIf Range("A" & i).Value < 10 Then
        
            Range("A" & i).Font.Color = vbRed
            Range("A" & i).Font.Size = 20
            Range("A" & i).Interior.Color = vbYellow
            Range("A" & i).HorizontalAlignment = xlCenter
        
        End If
    
    Next i

End Sub
