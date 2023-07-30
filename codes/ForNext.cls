VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub forNext()

    Dim i As Integer
    
    For i = 1 To 100
    
        Range("A" & i).Value = i
    
    Next i

End Sub

Sub forNextII()

    Dim i As Integer
    
    For i = 1 To 10
    
        ActiveCell.Value = i
        
        ActiveCell.Offset(1, 0).Activate
    
    Next i

End Sub
