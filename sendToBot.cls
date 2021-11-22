VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Worksheets("Master").Range("A:A")) Is Nothing Then
    Call test
    End If
End Sub

Sub test()
    Dim lastrow As Long, res As String
    With Worksheets("Master")
        If WorksheetFunction.CountA(.Range("A10:A20")) = 0 Then
            MsgBox "Belum ada data"
            Exit Sub
        End If
        lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        res = Join(Application.Transpose(.Range("A10:A" & lastrow + 1).Value), vbCrLf)
        Do While InStr(1, res, vbCrLf & vbCrLf) ' remove extra vbCrLf (if there're empty rows)
            res = Replace(res, vbCrLf & vbCrLf, vbCrLf)
        Loop
        KirimPesanTelegram Left(res, Len(res) - 1)
    End With
End Sub

Sub KirimPesanTelegram(ByRef Pesan As String)
    Dim Token As String, ChatID As String
    Dim sURL As String, oHttp As Object, sHTML As String
    
Token = "yourToken" 'Token Bot dari @BotFather
ChatID = "YourChatID" 'ID User Telegram yang ingin di kirim pesan
    sURL = "https://api.telegram.org/bot" & Token & "/sendMessage?chat_id=" & ChatID & "&text=" & Pesan
    
    Set oHttp = CreateObject("Msxml2.XMLHTTP")
    oHttp.Open "POST", sURL, False
    oHttp.Send

    sHTML = oHttp.ResponseText

    Debug.Print sHTML
End Sub



