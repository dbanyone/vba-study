VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub forNextSeparaNomePorNome()

    Dim ultimaLinhaOutraTabela
    Dim ultimaLinha As Integer
    ultimaLinha = Range("A1").End(xlDown).Row
    
    For i = 1 To ultimaLinha
    
        If Left(Range("A" & i).Value, 1) = "A" Then
        
            ultimaLinhaOutraTabela = 2
            
            On Error Resume Next
            
            ultimaLinhaOutraTabela = Sheets("Nomes").Range("E1").End(xlDown).Row + 1
            
            Sheets("Nomes").Range("E" & i & ":" & "C" & i).Copy Sheets("Nome").Range("E" & ultimaLinhaOutraTabela & ":" & "G" & ultimaLinhaOutraTabela)
            
            
        
        End If
        
    
    Next i

End Sub
