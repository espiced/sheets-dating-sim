Public Sub Workbook_Open()

    With wsProblems
        .Activate
        .Range("D12").Select
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
    wsRachel.txtRachel.Text = ""
    With wsKellie
        .cmdKContinue.Visible = True
        .txtKellie.Text = ""
        .cmdKTriv1.Visible = False
        .cmdKTriv2.Visible = False
        .cmdKTriv3.Visible = False
    End With
    With wsChloe
        .cmdCContinue.Visible = True
        .txtChloe.Value = ""
        .cmdCTriv1.Visible = False
        .cmdCTriv2.Visible = False
        .cmdCTriv3.Visible = False
    End With
    With wsAnya
        .cmdAContinue.Visible = True
        .txtAnya.Value = ""
        .cmdATriv1.Visible = False
        .cmdATriv2.Visible = False
        .cmdATriv3.Visible = False
    End With
    
End Sub
