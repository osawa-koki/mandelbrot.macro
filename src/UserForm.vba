Option Explicit

Private Sub CommandButton_Click()
  ' 値をセットするための変数の定義。
  Dim Width As Long
  Dim Height As Long
  Dim XMin As Double
  Dim XMax As Double
  Dim YMin As Double
  Dim YMax As Double
  Dim MaxIterations As Long
  Dim SheetName As String

  ' ユーザフォームの値を取得。
  Width = UserForm.TextBoxWidth.Value
  Height = UserForm.TextBoxHeight.Value
  XMin = UserForm.TextBoxXMin.Value
  XMax = UserForm.TextBoxXMax.Value
  YMin = UserForm.TextBoxYMin.Value
  YMax = UserForm.TextBoxYMax.Value
  MaxIterations = UserForm.TextBoxMaxIterations.Value
  SheetName = UserForm.TextBoxSheetName.Value

  ' ユーザフォームを非表示。
  UserForm.Hide

  ' イミディエイトウィンドウの出力をクリア。
  Debug.Print String(100, vbCrLf)

  ' デバグ用に値を出力。
  Debug.Print "Width: " & Width
  Debug.Print "Height: " & Height
  Debug.Print "XMin: " & XMin
  Debug.Print "XMax: " & XMax
  Debug.Print "YMin: " & YMin
  Debug.Print "YMax: " & YMax
  Debug.Print "MaxIterations: " & MaxIterations
  Debug.Print "SheetName: " & SheetName

  ' シートの削除
  Application.DisplayAlerts = False ' メッセージを非表示
  Dim ws As Worksheet
  For Each ws In Worksheets
    If ws.Name = SheetName Then
      ws.Delete
    End If
  Next ws
  Application.DisplayAlerts = True  ' メッセージを表示

  ' シートの追加
  Dim sheet As Worksheet
  Set sheet = Worksheets.Add
  sheet.Name = SheetName
  sheet.Activate

  ' 行と列のサイズを設定
  sheet.Range(Rows(1), Rows(height)).RowHeight = 7.5
  sheet.Range(Columns(1), Columns(width)).ColumnWidth = 0.77

End Sub
