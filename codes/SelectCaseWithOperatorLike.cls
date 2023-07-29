VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub selectCaseOperatorLike()

    Dim texto As String
    texto = InputBox("Digite uma palavra ou frase")
    
    Select Case texto
    
        Case texto Like "*a*"
        
            MsgBox "Texto contem letra A"
            
        Case Else
            
            MsgBox "Texto sem letra A"
            
    End Select
    

End Sub
