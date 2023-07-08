VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "UserForm"
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_Click()
  ' �l���Z�b�g���邽�߂̕ϐ��̒�`�B
  Dim Width As Long
  Dim Height As Long
  Dim XMin As Double
  Dim XMax As Double
  Dim YMin As Double
  Dim YMax As Double
  Dim MaxIterations As Long
  Dim SheetName As String

  ' ���[�U�t�H�[���̒l���擾�B
  Width = UserForm.TextBoxWidth.Value
  Height = UserForm.TextBoxHeight.Value
  XMin = UserForm.TextBoxXMin.Value
  XMax = UserForm.TextBoxXMax.Value
  YMin = UserForm.TextBoxYMin.Value
  YMax = UserForm.TextBoxYMax.Value
  MaxIterations = UserForm.TextBoxMaxIterations.Value
  SheetName = UserForm.TextBoxSheetName.Value

  ' ���[�U�t�H�[�����\���B
  UserForm.Hide

  ' �C�~�f�B�G�C�g�E�B���h�E�̏o�͂��N���A�B
  Debug.Print String(100, vbCrLf)

  ' �f�o�O�p�ɒl���o�́B
  Debug.Print "Width: " & Width
  Debug.Print "Height: " & Height
  Debug.Print "XMin: " & XMin
  Debug.Print "XMax: " & XMax
  Debug.Print "YMin: " & YMin
  Debug.Print "YMax: " & YMax
  Debug.Print "MaxIterations: " & MaxIterations
  Debug.Print "SheetName: " & SheetName

  ' �V�[�g�̍폜
  Application.DisplayAlerts = False ' ���b�Z�[�W���\��
  Dim ws As Worksheet
  For Each ws In Worksheets
    If ws.Name = SheetName Then
      ws.Delete
    End If
  Next ws
  Application.DisplayAlerts = True  ' ���b�Z�[�W��\��

  ' �V�[�g�̒ǉ�
  Dim sheet As Worksheet
  Set sheet = Worksheets.Add
  sheet.Name = SheetName
  sheet.Activate

  ' �s�Ɨ�̃T�C�Y��ݒ�
  sheet.Range(Rows(1), Rows(Height)).RowHeight = 7.5
  sheet.Range(Columns(1), Columns(Width)).ColumnWidth = 0.77

  ' �s�N�Z�������i�[����z��̒�`�B
  Dim PixelInfoArray() As PixelInfo
  ReDim PixelInfoArray(Width * Height)

  ' �s�N�Z�������i�[����z��ɏ����l���Z�b�g�B
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

  ' �s�N�Z�������V�[�g�ɏo�́B
  For i = 1 To Width * Height
    sheet.Cells(PixelInfoArray(i).x + 1, PixelInfoArray(i).y + 1).Interior.Color = PixelInfoArray(i).Color
  Next i

End Sub

