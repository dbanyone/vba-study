VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub forNextOutrasPlanilhas()

    Dim ultimaLinha As Integer
    Dim ultimaLinhaOutraPlanilha As Integer
    ultimaLinha = Range("A1").End(xlDown).Row
    
    For i = 2 To ultimaLinha Step 1
    
        If Range("A" & i).Value = "Produto A" Then
        
            'pula o cabeçalho e pega apartir da segunda linha
            ultimaLinhaOutraPlanilha = 2
            
            'ignora o erro que vai dar (se estiver tudo vazio)
            On Error Resume Next
            
            'vai na ultima linha e pula mais uma com o +1
            ultimaLinhaOutraPlanilha = Sheets("Produtos A").Range("A1").End(xlDown).Row + 1
        
            'copia tudo de A até C da Planilha Produtos, e cola na Planilha Produtos A
            Sheets("Produtos").Range("A" & i & ":" & "C" & i).Copy Sheets("Produtos A").Range("A" & ultimaLinhaOutraPlanilha & ":" & "C" & ultimaLinhaOutraPlanilha)
            
        ElseIf Range("A" & i).Value = "Produto B" Then
        
            ultimaLinhaOutraPlanilha = 2
            On Error Resume Next
            ultimaLinhaOutraPlanilha = Sheets("Produtos B").Range("A1").End(xlDown).Row + 1
        
            Sheets("Produtos").Range("A" & i & ":" & "C" & i).Copy Sheets("Produtos B").Range("A" & ultimaLinhaOutraPlanilha & ":" & "C" & ultimaLinhaOutraPlanilha)
        
        ElseIf Range("A" & i).Value = "Produto C" Then
        
            ultimaLinhaOutraPlanilha = 2
            On Error Resume Next
            ultimaLinhaOutraPlanilha = Sheets("Produtos C").Range("A1").End(xlDown).Row + 1
        
            Sheets("Produtos").Range("A" & i & ":" & "C" & i).Copy Sheets("Produtos C").Range("A" & ultimaLinhaOutraPlanilha & ":" & "C" & ultimaLinhaOutraPlanilha)
        
        End If
    
    Next i
    
    MsgBox "Relatorio separado!"

End Sub
