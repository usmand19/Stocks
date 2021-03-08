Attribute VB_Name = "Module1"


Sub Auto_Open()
' Upon opening the workbook, this macro schedules itself every 10 seconds to update the stock prices through the 'RefreshAll' command

    ActiveWorkbook.RefreshAll
    Application.OnTime Now + TimeValue("00:00:10"), "Auto_Open"
End Sub


