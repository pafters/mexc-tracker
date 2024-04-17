Attribute VB_Name = "UpdateTable"
Dim TimerStopped As Boolean

Sub StartTimer()
    TimerStopped = False
    Call UpdateTable
End Sub

Sub TimerStop()
    TimerStopped = True
End Sub


Sub UpdateTable()
    If TimerStopped Then Exit Sub

    Dim request As Object
    Set request = CreateObject("MSXML2.XMLHTTP")

   
    request.Open "GET", "https://api.mexc.com/api/v3/ticker/24hr/", False
    request.send
    

    Dim json As Object
    Set json = JsonConverter.ParseJson(request.responseText)

  
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Dim symbol As String
    Dim lastPrice As String


    For Each Item In json
        symbol = UCase(Replace(Item("symbol"), "/", ""))
        For I = 2 To 40 ' список пар A2:A40
            If symbol = UCase(Replace(ws.Cells(I, 1).Value, "/", "")) Then
                lastPrice = Item("lastPrice")
                bidPrice = Item("bidPrice")
                askPrice = Item("lastPrice")
                ws.Cells(I, 2).Value = lastPrice
                ws.Cells(I, 3).Value = bidPrice
                ws.Cells(I, 4).Value = askPrice
            End If
        Next I
    Next Item

    Application.OnTime Now + TimeValue("00:00:10"), "UpdateTable"
End Sub
