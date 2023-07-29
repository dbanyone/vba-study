VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub numeros()

    Dim numero As Byte
    
    numero = InputBox("Digite um número entre 1 e 100")
    
    If numero >= 90 And numero <= 100 Then
    
        MsgBox "O número está entre 90 e 100"
        
    ElseIf numero >= 60 And numero < 90 Then
    
        MsgBox "O número está dentro de 60 e 89"
        
    ElseIf numero >= 30 And numero < 60 Then
    
        MsgBox "O número está entre 30 e 59"
        
    ElseIf numero >= 10 And numero < 30 Then
    
        MsgBox "O número está entre 10 e 29"
    
    ElseIf numero < 10 Then
    
        MsgBox "Número menor que 10"
        
    Else
    
        MsgBox "Número inválido"
        
    End If
        
    

End Sub
