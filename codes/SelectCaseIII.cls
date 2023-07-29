VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub selectCaseIII()

    Dim resultado As VbMsgBoxResult
    resultado = MsgBox(("Sim") & vbCrLf & _
                        "Não" & vbCrLf & _
                        "Cancelar", vbYesNoCancel, "Escolha uma opção")
                        
                        
    Select Case resultado
    
        Case vbYes
        
            MsgBox "Voce escolheu sim"
            
        Case vbNo
        
            MsgBox "Voce escolheu nao"
            
        Case vbCancel
        
            MsgBox "Voce escolheu cancelar"
            
    End Select
    
    
End Sub
