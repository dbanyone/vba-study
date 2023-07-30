VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub mediaAlunoSelectCase()

    Dim nota1 As Double
    Dim nota2 As Double
    Dim nota3 As Double
    Dim nota4 As Double
    Dim situacao As String
    
    nota1 = InputBox("Digite a nota 1 do aluno")
    nota2 = InputBox("Digite a nota 1 do aluno")
    nota3 = InputBox("Digite a nota 1 do aluno")
    nota4 = InputBox("Digite a nota 1 do aluno")
    
    media = (nota1 + nota2 + nota3 + nota4) / 4
    
    Select Case media
    
        Case Is < 2
            
            MsgBox "Reprovado"
            situacao = "Reprovado"
            
        Case Is >= 4
            
            MsgBox "Recuperação"
            situacao = "Recuperação"
            
        Case Is <= 10
        
            MsgBox "Aprovado"
            situacao = "Aprovado"
            
        Case Else
        
            MsgBox "Digite valores entre 1 e 10.", vbCritical
            
    End Select
    
    MsgBox ("Nota 1: " & nota1 & vbCrLf & _
            "Nota 2: " & nota2 & vbCrLf & _
            "Nota 3: " & nota3 & vbCrLf & _
            "Nota 4: " & nota4) & vbCrLf & _
            "Situação: " & situacao, vbInformation


End Sub


Sub selectCaseRange()

    Select Case Range("A2")
    
        Case "Masculino"
        
            Select Case Range("B2")
            
                Case "Sapato"
                
                    MsgBox "Masculino - Sapato"
                
                Case "Camisa"
        
                     MsgBox "Masculino - Camisa"
                     
                Case Else
                
                    MsgBox "Valor inválido"
                    
            End Select
            
            
        Case "Feminino"
        
            Select Case Range("B2")
            
                Case "Vestido"
                
                    MsgBox "Feminino - Vestido"
                
                Case "Bolsa"
            
                     MsgBox "Feminino - Bolsa"
                     
                Case Else
                
                    MsgBox "Valor inválido"
                    
            End Select

    
    End Select

End Sub
