"""CSV 파일을 검사한 뒤 매출 엑셀에 입력합니다."""
import csv
import os
import re
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("오류: openpyxl이 필요합니다. python -m pip install openpyxl 명령으로 설치해 주세요.")
    sys.exit(1)

기준 = Path(__file__).resolve().parent


def 값_변환(값):
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", 값):
        try:
            return date.fromisoformat(값)
        except ValueError:
            return 값
    if re.fullmatch(r"-?\d+(?:\.\d+)?", 값):
        정수부분 = 값.lstrip("-").split(".")[0]
        if len(정수부분) > 1 and 정수부분.startswith("0"):
            return 값
        return float(값) if "." in 값 else int(값)
    return 값


def csv_읽기(경로):
    원본 = 경로.read_bytes()
    try:
        내용 = 원본.decode("utf-8-sig")
    except UnicodeDecodeError:
        내용 = 원본.decode("cp949")
    import io
    try:
        return list(csv.reader(io.StringIO(내용, newline=""), strict=True))
    except csv.Error as 오류:
        raise ValueError(f"{경로.name}: CSV 내용을 읽을 수 없습니다: {오류}") from 오류


def 임시_경로(접미사):
    번호, 이름 = tempfile.mkstemp(prefix="매출_임시_", suffix=접미사, dir=기준)
    os.close(번호)
    return Path(이름)


def 실행():
    엑셀 = 기준 / "매출.xlsx"
    입력 = 기준 / "csv"
    완료 = 기준 / "completed"
    if not 입력.is_dir():
        raise ValueError("csv 폴더를 찾을 수 없습니다.")
    파일들 = sorted((파일 for 파일 in 입력.iterdir() if 파일.is_file() and 파일.suffix.lower() == ".csv"), key=lambda 파일: 파일.name)
    if not 파일들:
        raise ValueError("csv 폴더에 CSV 파일이 없습니다.")
    if 완료.exists() and not 완료.is_dir():
        raise ValueError("completed가 폴더가 아닙니다.")

    책 = openpyxl.load_workbook(엑셀)
    if "매출" not in 책.sheetnames:
        책.close()
        raise ValueError("엑셀에 '매출' 시트가 없습니다.")
    시트 = 책["매출"]
    제목 = [시트.cell(1, 열).value for 열 in range(1, 시트.max_column + 1)]
    while 제목 and 제목[-1] is None:
        제목.pop()
    행들 = []
    for 파일 in 파일들:
        내용 = csv_읽기(파일)
        if not 내용 or 내용[0] != 제목:
            raise ValueError(f"{파일.name}: 첫 행의 제목 또는 열 순서가 엑셀과 다릅니다.")
        for 번호, 행 in enumerate(내용[1:], 2):
            if len(행) != len(제목):
                raise ValueError(f"{파일.name} {번호}행: 열 개수가 제목과 다릅니다.")
            행들.append([값_변환(값) for 값 in 행])
    for 파일 in 파일들:
        if (완료 / 파일.name).exists():
            raise ValueError(f"completed에 같은 이름의 파일이 있습니다: {파일.name}")

    기존_마지막 = 시트.max_row
    for 행 in 시트.iter_rows(min_row=2, max_row=기존_마지막):
        for 셀 in 행:
            if not isinstance(셀, openpyxl.cell.cell.MergedCell):
                셀.value = None
    for 행번호, 행 in enumerate(행들, 2):
        for 열번호, 값 in enumerate(행, 1):
            셀 = 시트.cell(행번호, 열번호)
            기존_서식 = 셀._style
            셀.value = 값
            if isinstance(값, str):
                셀.data_type = "s"
            셀._style = 기존_서식

    백업 = None
    임시 = None
    반영됨 = False
    이동됨 = []
    복구오류 = []
    try:
        백업 = 임시_경로(".xlsx")
        shutil.copy2(엑셀, 백업)
        임시 = 임시_경로(".xlsx")
        책.save(임시)
        책.close()
        확인 = openpyxl.load_workbook(임시)
        확인.close()
        os.replace(임시, 엑셀)
        임시 = None
        반영됨 = True
        완료.mkdir(exist_ok=True)
        for 파일 in 파일들:
            대상 = 완료 / 파일.name
            # 같은 이름이 생겼더라도 기존 파일을 덮어쓰지 않습니다.
            with 대상.open("xb") as 출력:
                이동됨.append((파일, 대상))
                with 파일.open("rb") as 원본파일:
                    shutil.copyfileobj(원본파일, 출력)
            파일.unlink()
    except Exception as 오류:
        for 원본, 대상 in reversed(이동됨):
            try:
                if not 원본.exists():
                    with 원본.open("xb") as 출력, 대상.open("rb") as 입력파일:
                        shutil.copyfileobj(입력파일, 출력)
                대상.unlink()
            except OSError as 복구:
                복구오류.append(f"{대상.name}: {복구}")
        if 반영됨 and 백업:
            try:
                shutil.copy2(백업, 엑셀)
            except OSError as 복구:
                복구오류.append(f"엑셀 복구: {복구}")
        if 복구오류:
            raise RuntimeError(f"작업 실패: {오류}\n복구 실패: {'; '.join(복구오류)}\n엑셀 임시 사본: {백업}") from 오류
        raise RuntimeError(f"작업 실패: {오류}. 변경된 파일은 원래 상태로 되돌렸습니다.") from 오류
    finally:
        책.close()
        if 임시 and 임시.exists():
            임시.unlink()
        if 백업 and 백업.exists() and not 복구오류:
            백업.unlink()
    print(f"완료: CSV {len(파일들)}개, 데이터 {len(행들)}행을 처리했습니다.")


if __name__ == "__main__":
    try:
        실행()
    except Exception as 오류:
        print(f"오류: {오류}")
        sys.exit(1)

