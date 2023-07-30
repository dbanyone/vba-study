VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub forNextcalcularMedia()

    Dim nota1, nota2, nota3, nota4 As Double
    Dim ultimaLinha As Integer
    
    ultimaLinha = Range("A1").End(xlDown).Row
    
    For i = 2 To ultimaLinha Step 1
    
        nota1 = CDbl(Range("B" & i).Value)
        nota2 = CDbl(Range("C" & i).Value)
        nota3 = CDbl(Range("D" & i).Value)
        nota4 = CDbl(Range("E" & i).Value)
   
        media = (nota1 + nota2 + nota3 + nota4) / 4
    
        If media < 2 Then
        
            Range("F" & i).Value = "Reprovado"
            Range("G" & i).Value = media
            Range("F" & i).Font.ColorIndex = 3
            Range("F" & i).Interior.ColorIndex = 30
            Range("F" & i).Font.Bold = True
            Range("G" & i).Font.Bold = True
            
        ElseIf media < 5 Then
        
            Range("F" & i).Value = "Recuperacao"
            Range("G" & i).Value = media
            Range("F" & i).Font.ColorIndex = 6
            Range("F" & i).Interior.ColorIndex = 44
            Range("F" & i).Font.Bold = True
            Range("G" & i).Font.Bold = True

            
        ElseIf media <= 10 Then
        
            Range("F" & i).Value = "Aprovado"
            Range("G" & i).Value = media
            Range("F" & i).Font.ColorIndex = 4
            Range("F" & i).Interior.ColorIndex = 10
            Range("F" & i).Font.Bold = True
            Range("G" & i).Font.Bold = True
            
        End If
    
    Next i

End Sub

Sub limparTudo()

    
    
    Dim ultimaLinha As Integer
    
    ultimaLinha = Range("A1").End(xlDown).Row
    
    For i = 2 To ultimaLinha Step 1
        
        Range("F" & i).Value = ""
        Range("G" & i).Value = ""
        Range("F" & i).Font.ColorIndex = 1
        Range("F" & i).Interior.ColorIndex = 2
        Range("G" & i).Font.ColorIndex = 1
        Range("G" & i).Interior.ColorIndex = 2

            
    Next i

End Sub

