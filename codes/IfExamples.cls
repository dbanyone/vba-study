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

Sub addAmigo()

    If MsgBox(Range("A10") & " quer ser seu amigo, deseja adicionar? ", vbOKCancel + vbInformation, "Deseja adicionar a sua lista de amigos") = vbOK Then
    
        Sheets("Amigos").Select
        
        If Sheets("Amigos").Range("A2") = "" Then
        
            Sheets("Amigos").Range("A2") = Sheets("Professor").Range("A10")
            
            Sheets("Professor").Select
            
            MsgBox "Amigo adicionado com sucesso!", Title = "Mensagem"
        
        Else
        
            'Encontra a última célula e pula uma linha
            Sheets("Amigos").Range("A1").End(xlDown).Offset(1, 0).Select
            
            Dim linha As String
            Dim colunaELinha As String
            linha = ActiveCell.Row 'pegando a posição da célula
            colunaELinha = "A" & linha
            
            'Seleciona a última linha em branco
            Sheets("Amigos").Range(colunaELinha).Select
            
            'Imprime o nome do aluno embaixo da lista
            ActiveCell.Value = Sheets("Professor").Range("A10")
            
            MsgBox "Amigo adicionado com sucesso!", Title = "Mensagem"
        
        End If
    
    Else
    
        MsgBox "Recusado com sucesso!", Title = "Mensagem"
    
    End If
    

End Sub


Sub desafioENumero()

    'Range("A15").Select
    
    If IsNumeric(Range("A15").Value) Then
        
        MsgBox "É um número"
    
    Else
    
        MsgBox "Não é um número"
    
    End If
    

End Sub


Sub ifAninhado()

    Dim sexo, roupa As String
    
    sexo = InputBox("Digite M para masculino e F para feminino")
    
    If sexo = "M" Then
    
        roupa = InputBox("Escolha para C para camisa ou S para Sapato")
        
        If roupa = "C" Then
        
            MsgBox "Voce escolheu Calça masculina"
        
        ElseIf roupa = "S" Then
        
            MsgBox "Voce escolheu Sapato masculino"
        
        Else
        
            MsgBox "Opção inválida!"
        
        End If
        
    
    ElseIf sexo = "F" Then
    
        roupa = InputBox("Escolha para C para camisa ou S para Sapato")
        
        If roupa = "S" Then
        
            MsgBox "Voce escolheu Sapato feminino"
        
        ElseIf roupa = "C" Then
        
            MsgBox "Voce escolheu Calça feminina"
            
        Else
        
            MsgBox "Opção inválida!"
        
        End If
        
    
    Else
    
        MsgBox "Gênero inválido!"
    
    End If
    
    
    
End Sub
