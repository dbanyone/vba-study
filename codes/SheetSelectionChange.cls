VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

    Application.StatusBar = "Planilha " & Sh.Name & " - " & Target.Address
    
    Select Case Sh.Name
    
    Case "Planilha1"
        
        Range("A1").Value = "Essa é a planilha 1"
    
    Case "Planilha2"
        
        Range("A1").Value = "Essa é a planilha 2"
        
    Case Else
        
        Range("A1").Value = "Essa é a " + Sh.Name
        
    End Select

End Sub
