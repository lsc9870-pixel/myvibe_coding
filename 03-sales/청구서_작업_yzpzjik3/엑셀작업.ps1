
param([string]$설정파일)
$ErrorActionPreference = 'Stop'
$설정 = Get-Content -LiteralPath $설정파일 -Raw -Encoding UTF8 | ConvertFrom-Json
$엑셀 = $null
try {
    $엑셀 = New-Object -ComObject Excel.Application
    $엑셀.Visible = $false
    $엑셀.DisplayAlerts = $false
    $엑셀.AskToUpdateLinks = $false
    foreach ($항목 in $설정.목록) {
        $책 = $null
        try {
            $책 = $엑셀.Workbooks.Open($설정.양식, 0, $true)
            $시트 = $책.Worksheets.Item('청구서')
            $추가 = [Math]::Max(0, $항목.내역.Count - 16)
            if ($추가 -gt 0) {
                for ($i = 0; $i -lt $추가; $i++) {
                    $시트.Rows.Item(35).Copy()
                    [void]$시트.Rows.Item(36).Insert(-4121)
                }
            }
            $마지막 = 35 + $추가
            $합계행 = 36 + $추가
            $세금행 = 37 + $추가
            $시트.Range('B5').NumberFormat = '@'
            $시트.Range('B5').Value2 = [string]$항목.우편번호
            $시트.Range('B6').Value2 = [string]$항목.주소
            $시트.Range('B7').Value2 = [string]$항목.이름
            for ($행 = 20; $행 -le $마지막; $행++) {
                $시트.Cells.Item($행, 1).ClearContents() | Out-Null
                $시트.Cells.Item($행, 2).MergeArea.ClearContents() | Out-Null
                $시트.Cells.Item($행, 6).ClearContents() | Out-Null
                $시트.Cells.Item($행, 7).ClearContents() | Out-Null
                $시트.Cells.Item($행, 8).Formula = "=F$행*G$행"
            }
            $행 = 20
            foreach ($내역 in $항목.내역) {
                $시트.Cells.Item($행, 1).Value2 = ([datetime]::ParseExact($내역.날짜, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)).ToOADate()
                $시트.Cells.Item($행, 2).Value2 = [string]$내역.상품명
                $시트.Cells.Item($행, 6).Value2 = [double]$내역.수량
                $시트.Cells.Item($행, 7).Value2 = [double]$내역.단가
                $행++
            }
            $시트.Cells.Item($합계행, 8).Formula = "=SUM(H20:H$마지막)"
            $시트.Cells.Item($세금행, 8).Formula = "=INT(H$합계행*0.1)"
            $시트.Range('E17').Formula = "=H$합계행"
            $시트.Range('G17').Formula = "=H$세금행"
            $시트.Range('I17').Formula = "=H$합계행+H$세금행"
            $엑셀.CalculateFull()
            $기대 = 0.0
            foreach ($내역 in $항목.내역) { $기대 += [double]$내역.수량 * [double]$내역.단가 }
            if ([Math]::Abs([double]$시트.Cells.Item($합계행, 8).Value2 - $기대) -gt 0.00001) { throw '청구서 합계가 검산 결과와 다릅니다.' }
            if ([Math]::Abs([double]$시트.Range('I17').Value2 - ($기대 + [Math]::Floor($기대 * 0.1))) -gt 0.00001) { throw '청구 금액이 검산 결과와 다릅니다.' }
            $시트.PageSetup.PrintArea = '$A$1:$I$' + (41 + $추가)
            $시트.PageSetup.Zoom = $false
            $시트.PageSetup.FitToPagesWide = 1
            $시트.PageSetup.FitToPagesTall = $false
            $시트.ExportAsFixedFormat(0, [string]$항목.임시PDF)
        } finally {
            if ($책) { $책.Close($false) }
        }
    }
} finally {
    if ($엑셀) { $엑셀.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($엑셀) }
}
