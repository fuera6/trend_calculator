# trend_calculator
트렌드 분석을 위한 βdiff, rOR 계산기

## Usage
* beta difference 계산
* ratio of odds ratio 계산

## Time
* 2024.03

## Directories & Files
* 분석틀.xlsx: βdiff, rOR을 계산해주는 계산기

## Notice (24.04.09)
βdiff와 rOR 결과가 조건부서식으로 나오기 때문에 이를 다른 곳에 복붙하면 서식복사가 불가능하다. 해결하려면 엑셀 VBA 써서 해당 부분 서식 유지해야 한다. 아래는 그 코드이다.

(VBA 들어가는 법: 엑셀에서 원하는 영역 드래그 → Alt+F11 → 삽입 → 모듈 → 코드 복붙 → F5 → 해당영역인지 확인 → 끄고 엑셀로 돌아가기)
```
Sub Keep_Format()
'UpdatebyExtendoffice20181128
    Dim xRg As Range
    Dim xTxt As String
    Dim xCell As Range
    On Error Resume Next
    If ActiveWindow.RangeSelection.Count > 1 Then
      xTxt = ActiveWindow.RangeSelection.AddressLocal
    Else
      xTxt = ActiveSheet.UsedRange.AddressLocal
    End If
    Set xRg = Application.InputBox("Select range:", "Kutools for Excel", xTxt, , , , , 8)
    If xRg Is Nothing Then Exit Sub
    For Each xCell In xRg
        With xCell
            .Font.FontStyle = .DisplayFormat.Font.FontStyle
            .Font.Strikethrough = .DisplayFormat.Font.Strikethrough
            .Interior.Pattern = .DisplayFormat.Interior.Pattern
            If .Interior.Pattern <> xlNone Then
                 .Interior.PatternColorIndex = .DisplayFormat.Interior.PatternColorIndex
                .Interior.Color = .DisplayFormat.Interior.Color
            End If
            .Interior.TintAndShade = .DisplayFormat.Interior.TintAndShade
            .Interior.PatternTintAndShade = .DisplayFormat.Interior.PatternTintAndShade

        End With
    Next
    xRg.FormatConditions.Delete
End Sub
```

