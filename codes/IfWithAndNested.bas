Attribute VB_Name = "Módulo1"
Sub numerosAMais()

    Dim numeroDownload As Integer
    
    numeroDownload = Range("A13").Value
    
    If numeroDownload >= 100 Then
    
        Range("B13").Value = "Download concluido"
    
    Else
    
        If numeroDownload >= 90 And numeroDownload < 99 Then
        
            Range("B13").Value = "90 a 99%..."
    
        ElseIf numeroDownload >= 60 And numeroDownload < 90 Then
            
            Range("B13").Value = "60 a 89%..."
            
        ElseIf numeroDownload >= 40 And numeroDownload < 60 Then
            
            Range("B13").Value = "40 a 69%..."
            
        ElseIf numeroDownload >= 30 And numeroDownload < 40 Then
            
            Range("B13").Value = "30 a 39%..."
            
        ElseIf numeroDownload >= 10 And numeroDownload < 30 Then
            
            Range("B13").Value = "10 a 29%..."
            
        Else
            
            Range("B13") = "Iniciando download..."
            
        End If
        
    
    End If
    

End Sub
