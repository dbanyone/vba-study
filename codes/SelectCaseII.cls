VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub selectCaseII()

    Dim numero As Integer
    
    numero = InputBox("Digite um número.")
    
    Select Case numero
    
        Case Is >= 10
            
            MsgBox "Numero maior que 10"
            
        Case Is < 10
        
            MsgBox "Menor que 10"
            
    End Select

End Sub
