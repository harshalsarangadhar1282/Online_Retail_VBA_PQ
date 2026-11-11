Sub RefreshAndFormat()
    ThisWorkbook.RefreshAll
    Application.Wait Now + TimeValue("0:00:05")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Cleaned_Data")
    ws.Range("A1:I1").Font.Bold = True
    ws.Range("A1:I1").Interior.Color = RGB(31, 78, 120)
    ws.Range("A1:I1").Font.Color = vbWhite
    ws.Columns("A:I").AutoFit
    MsgBox "Online Retail data refreshed and formatted successfully!", vbInformation
End Sub
