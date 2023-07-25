VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub verificaNotaAluno()

    Dim media As Double
    'Dim nota1 As Double
    'Dim nota2 As Double
    'Dim nota3 As Double
    'Dim nota4 As Double
    
    
    Sheets("Ex").Range("B5") = InputBox("Digite a primeira nota")
    Sheets("Ex").Range("C5") = InputBox("Digite a segunda nota")
    Sheets("Ex").Range("D5") = InputBox("Digite a terceira nota")
    Sheets("Ex").Range("E5") = InputBox("Digite a quarta nota")
    
    media = (Range("B5") + Range("C5") + Range("D5") + Range("E5")) / 4
    
    Range("F5") = media
    
    If media > 7 Then
    
        Range("G5").Font.ColorIndex = 4
        Range("G5").Interior.ColorIndex = 43
        Range("G5") = "Aprovado"
   
    ElseIf media > 3 Then
    
         
        Range("G5").Font.ColorIndex = 28
        Range("G5").Interior.ColorIndex = 41
        Range("G5") = "Recuperação"
    
    
    Else
    
        Range("G5").Font.ColorIndex = 3
        Range("G5").Interior.ColorIndex = 53
        Range("G5") = "Reprovado"
    
    End If

End Sub
