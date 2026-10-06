Attribute VB_Name = "CloseSeason"
Option Explicit

Public Sub CloseSeason()
    Dim answer As VbMsgBoxResult

    answer = MsgBox( _
        "Провели ли вы синхронизацию изменений?", _
        vbQuestion + vbYesNo, _
        "Закрытие сезона" _
    )

    If answer <> vbYes Then
        Exit Sub
    End If

    ThisWorkbook.Save
    RunPython "from inventory_app.season_management import season_service; season_service.main()"
End Sub
