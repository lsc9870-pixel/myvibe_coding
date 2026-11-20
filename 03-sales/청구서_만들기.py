"""월별 매출을 거래처별 청구서 PDF로 만듭니다."""
import argparse
import calendar
import datetime as dt
import json
import math
import os
import re
import subprocess
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("오류: openpyxl이 필요합니다. python -m pip install openpyxl 명령으로 설치해 주세요.")

기준 = Path(__file__).resolve().parent


def 인수_확인(인수=None):
    분석기 = argparse.ArgumentParser(description="거래처별 월간 청구서 PDF를 만듭니다.")
    분석기.add_argument("연도", nargs="?", type=int)
    분석기.add_argument("월", nargs="?", type=int)
    값 = 분석기.parse_args(인수)
    if 값.연도 is None and 값.월 is None:
        return 2026, 10
    if 값.연도 is None or 값.월 is None or not (1900 <= 값.연도 <= 9998 and 1 <= 값.월 <= 12):
        분석기.error("연도와 월을 함께 입력하세요. 예: 2026 10")
    return 값.연도, 값.월


def 제목_찾기(시트, 필수):
    제목들 = [셀.value for 셀 in 시트[1]]
    결과 = {}
    for 제목 in 필수:
        if 제목들.count(제목) != 1:
            raise ValueError(f"'{시트.title}' 시트에 '{제목}' 제목이 없거나 중복됩니다.")
        결과[제목] = 제목들.index(제목)
    return 결과


def 코드_정리(값):
    if 값 is None or isinstance(값, bool):
        return ""
    if isinstance(값, (int, float)) and float(값).is_integer():
        return str(int(값))
    return str(값).strip()


def 숫자_확인(값, 위치):
    if isinstance(값, bool) or not isinstance(값, (int, float)) or not math.isfinite(값):
        raise ValueError(f"{위치}: 숫자가 없거나 올바르지 않습니다.")
    return 값


def 자료_읽기(경로, 연도, 월):
    책 = openpyxl.load_workbook(경로, data_only=True, read_only=True)
    try:
        if '매출' not in 책 or '거래처' not in 책:
            raise ValueError("매출 또는 거래처 시트가 없습니다.")
        거래처 = 책['거래처']
        열 = 제목_찾기(거래처, ['거래처코드', '거래처명', '우편번호', '주소'])
        목록 = {}
        for 번호, 행 in enumerate(거래처.iter_rows(min_row=2, values_only=True), 2):
            if all(값 is None for 값 in 행):
                continue
            코드 = 코드_정리(행[열['거래처코드']])
            if not 코드 or 코드 in 목록:
                raise ValueError(f"거래처 시트 {번호}행: 거래처코드가 없거나 중복됩니다.")
            이름, 우편, 주소 = (행[열[이름]] for 이름 in ['거래처명', '우편번호', '주소'])
            if any(값 is None or not str(값).strip() for 값 in (이름, 우편, 주소)):
                raise ValueError(f"거래처 시트 {번호}행: 이름, 우편번호 또는 주소가 없습니다.")
            우편셀 = 거래처.cell(번호, 열['우편번호'] + 1)
            if isinstance(우편, (int, float)) and re.fullmatch(r'0+', 우편셀.number_format or ''):
                우편 = str(int(우편)).zfill(len(우편셀.number_format))
            목록[코드] = {'코드': 코드, '이름': str(이름), '우편번호': str(우편), '주소': str(주소)}
        시트 = 책['매출']
        열 = 제목_찾기(시트, ['매출일', '상품명', '거래처코드', '매출수량', '단가'])
        결과 = defaultdict(list)
        for 번호, 행 in enumerate(시트.iter_rows(min_row=2, values_only=True), 2):
            if all(값 is None for 값 in 행):
                continue
            날짜 = 행[열['매출일']]
            if isinstance(날짜, dt.datetime):
                날짜 = 날짜.date()
            if not isinstance(날짜, dt.date):
                raise ValueError(f"매출 시트 {번호}행: 매출일이 올바른 날짜가 아닙니다.")
            if (날짜.year, 날짜.month) != (연도, 월):
                continue
            코드 = 코드_정리(행[열['거래처코드']])
            if 코드 not in 목록:
                raise ValueError(f"매출 시트 {번호}행: 거래처코드 '{코드}'를 찾을 수 없습니다.")
            상품 = 행[열['상품명']]
            if 상품 is None or not str(상품).strip():
                raise ValueError(f"매출 시트 {번호}행: 상품명이 없습니다.")
            수량 = 숫자_확인(행[열['매출수량']], f"매출 시트 {번호}행 수량")
            단가 = 숫자_확인(행[열['단가']], f"매출 시트 {번호}행 단가")
            결과[코드].append({'날짜': 날짜.isoformat(), '상품명': str(상품), '수량': 수량, '단가': 단가})
        for 행들 in 결과.values():
            행들.sort(key=lambda 행: 행['날짜'])
        return 목록, 결과
    finally:
        책.close()


def 파일이름(연도, 월, 거래처):
    이름 = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', 거래처['이름']).rstrip(' .') or '이름없음'
    코드 = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', 거래처['코드']).rstrip(' .')
    return f"{연도:04d}-{월:02d}_{코드}_{이름}.pdf"


# Excel이 양식을 열어 수식을 다시 계산하고 PDF로 저장합니다.
엑셀_작업 = r'''
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
'''


def 실행(연도, 월):
    양식 = 기준 / '청구서.xlsx'
    매출 = 기준 / '매출.xlsx'
    for 경로 in (양식, 매출):
        if not 경로.is_file():
            raise ValueError(f"파일이 없습니다: {경로.name}")
    양식책 = openpyxl.load_workbook(양식, read_only=True)
    try:
        if '청구서' not in 양식책:
            raise ValueError("청구서.xlsx에 '청구서' 시트가 없습니다.")
    finally:
        양식책.close()
    목록, 내역 = 자료_읽기(매출, 연도, 월)
    if not 내역:
        print(f"{연도}년 {월}월의 매출이 없어 PDF를 만들지 않았습니다.")
        return
    이름들 = [파일이름(연도, 월, 목록[코드]) for 코드 in 내역]
    if len({이름.casefold() for 이름 in 이름들}) != len(이름들):
        raise ValueError("PDF 파일 이름이 중복됩니다.")
    출력 = 기준 / 'PDF'
    출력.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='청구서_작업_', dir=기준) as 임시폴더:
        임시 = Path(임시폴더)
        작업목록 = []
        for 번호, 코드 in enumerate(sorted(내역)):
            작업목록.append({**목록[코드], '내역': 내역[코드], '임시PDF': str(임시 / f'{번호}.pdf'), '파일이름': 파일이름(연도, 월, 목록[코드])})
        설정 = 임시 / '설정.json'
        설정.write_text(json.dumps({'양식': str(양식), '목록': 작업목록}, ensure_ascii=False), encoding='utf-8-sig')
        스크립트 = 임시 / '엑셀작업.ps1'
        스크립트.write_text(엑셀_작업, encoding='utf-8-sig')
        결과 = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(스크립트), str(설정)], capture_output=True)
        if 결과.returncode:
            오류 = 결과.stderr.decode('utf-8', errors='replace').strip()
            raise RuntimeError(f"Excel 계산 또는 PDF 저장에 실패했습니다. Excel 설치와 실행 권한을 확인하세요.\n{오류}")
        for 항목 in 작업목록:
            파일 = Path(항목['임시PDF'])
            if not 파일.is_file() or 파일.stat().st_size == 0:
                raise RuntimeError(f"PDF 생성에 실패했습니다: {항목['파일이름']}")
        for 항목 in 작업목록:
            os.replace(항목['임시PDF'], 출력 / 항목['파일이름'])
    print(f"완료: {연도}년 {월}월 청구서 PDF {len(작업목록)}개를 저장했습니다. ({출력})")


if __name__ == '__main__':
    try:
        실행(*인수_확인())
    except Exception as 오류:
        print(f"오류: {오류}", file=sys.stderr)
        sys.exit(1)


