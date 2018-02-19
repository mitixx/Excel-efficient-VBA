'Excel sheet copy VBA(monthly)'
Sub copysheet()
Application.ScreenUpdating = False

'show question dialog'
Dim year As String
year = Application.InputBox(Prompt:="作成したい年を入力してください",Type:=1)
If year = "False" Then Exit Sub
Dim month As String
month = Application.InputBox(Prompt:="作成したい月を入力してください",Type:=1)
If month = "False" Then Exit Sub

Dim Start_Day As Long
Dim End_Day As Long
Dim i As Long
Start_Day = 1
End_Day = 31 '月の最終日を入力'
For i = Start_Day To End_Day
    Sheets("原本").Copy After:=ActiveSheet 'コピー元を指定する'
    ActiveSheet.Name = year & "年" & month & "月" & i & "日"
Next i

MsgBox "コピー完了しました！", vbInformation

Application.ScreenUpdating = True
End Sub