Public Sub Workbook_Open()

    With wsProblems
        .Range("Question1").ClearContents
        .Range("Question2").ClearContents
        .Range("Question3").ClearContents
        .Range("Question4").ClearContents
    End With
    With wsControls
        .Range("RachelControls").ClearContents
        .Range("KellieControls").ClearContents
        .Range("ChloeControls").ClearContents
        .Range("AnyaControls").ClearContents
    End With
    
End Sub