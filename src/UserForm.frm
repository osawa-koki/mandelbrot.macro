VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "UserForm"
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
  sheet.Range(Rows(1), Rows(Height)).RowHeight = 7.5
  sheet.Range(Columns(1), Columns(Width)).ColumnWidth = 0.77

  ' ピクセル情報を格納する配列の定義。
  Dim PixelInfoArray() As PixelInfo
  ReDim PixelInfoArray(Width * Height)

  ' ピクセル情報を格納する配列に初期値をセット。
  Dim i As Long
  For i = 1 To Width * Height
    Set PixelInfoArray(i) = New PixelInfo
    PixelInfoArray(i).Color = RGB(0, 0, 0)
    PixelInfoArray(i).x = (i - 1) \ Width
    PixelInfoArray(i).y = (i - 1) Mod Width
  Next i

  Dim x As Integer
  Dim y As Integer
  For x = 1 To Width
    For y = 1 To Height
      Dim cReal As Double
      Dim cImag As Double
      cReal = XMin + (XMax - XMin) * (x - 1) / (Width - 1)
      cImag = YMin + (YMax - YMin) * (y - 1) / (Height - 1)
      Dim zReal As Double
      Dim zImag As Double
      zReal = 0
      zImag = 0
      Dim n As Long
      For n = 1 To MaxIterations
        Dim zRealTemp As Double
        Dim zImagTemp As Double
        zRealTemp = zReal * zReal - zImag * zImag + cReal
        zImagTemp = 2 * zReal * zImag + cImag
        zReal = zRealTemp
        zImag = zImagTemp
        If zReal * zReal + zImag * zImag > 4 Then
          Exit For
        End If
      Next n
      Debug.Print n
      If n = MaxIterations Then
        PixelInfoArray((y - 1) * Width + x).Color = RGB(0, 0, 0)
      Else
        PixelInfoArray((y - 1) * Width + x).Color = RGB(255 * n / MaxIterations, 0, 0)
      End If
    Next y
  Next x

  ' ピクセル情報をシートに出力。
  For i = 1 To Width * Height
    sheet.Cells(PixelInfoArray(i).x + 1, PixelInfoArray(i).y + 1).Interior.Color = PixelInfoArray(i).Color
  Next i

End Sub

