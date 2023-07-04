Option Explicit

Sub Mandelbrot()
  ' デフォルト値をセットするための変数の定義。
  Dim DefaultWidth As Long
  Dim DefaultHeight As Long
  Dim DefaultXMin As Double
  Dim DefaultXMax As Double
  Dim DefaultYMin As Double
  Dim DefaultYMax As Double
  Dim DefaultMaxIterations As Long
  Dim DefaultSheetName As String

  ' デフォルト値の設定。
  DefaultWidth = 50
  DefaultHeight = 50
  DefaultXMin = -2.0
  DefaultXMax = 1.0
  DefaultYMin = -1.5
  DefaultYMax = 1.5
  DefaultMaxIterations = 100
  DefaultSheetName = "Mandelbrot"

  ' デフォルト値をユーザフォームにセット。
  UserForm.TextBoxWidth.Value = DefaultWidth
  UserForm.TextBoxHeight.Value = DefaultHeight
  UserForm.TextBoxXMin.Value = DefaultXMin
  UserForm.TextBoxXMax.Value = DefaultXMax
  UserForm.TextBoxYMin.Value = DefaultYMin
  UserForm.TextBoxYMax.Value = DefaultYMax
  UserForm.TextBoxMaxIterations.Value = DefaultMaxIterations
  UserForm.TextBoxSheetName.Value = DefaultSheetName

  ' ユーザフォームを表示。
  UserForm.Show
End Sub
