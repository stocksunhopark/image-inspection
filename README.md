# Excel 이미지 육안 검사기

Excel 통합 문서에 삽입된 이미지를 추출해, 같은 시트·셀 위치의 이미지를
나란히 넘겨 보는 PyQt6 프로그램입니다.

이 프로그램은 이미지 유사도 점수나 `PASS/FAIL`을 계산하지 않습니다.
판정 임계값, 비교 알고리즘, 캐시, 결과 리포트도 사용하지 않으며 사용자가
화면을 보면서 직접 검사하는 용도에만 집중합니다.

## 주요 기능

- **Double 모드**: Excel Ref와 비교A, 총 2개 파일을 나란히 표시
- **Triple 모드**: Excel Ref, 비교A, 비교B, 총 3개 파일을 나란히 표시
- **Quadra 모드**: Excel Ref, 비교A, 비교B, 비교C, 총 4개 파일을 나란히 표시
- 기존 테스트 양식과 같은 시트 탭 순서 및 이미지 앵커 구조 지원
- 동일 셀 → 동일 병합영역 → 인접 셀(±1행/열) 순으로 표시 위치 정렬
- 시트 선택, 처음/이전/다음/마지막 이동, 전체 위치 슬라이더
- `리스트실행` 별도 창에서 전체 이미지 위치 목록과 이미지를 동시에 확인
- 리스트 창 상단에서 `전체 시트` 또는 이미지가 있는 개별 시트 선택
- 목록 행 클릭 또는 `↑/↓` 키 이동 시 Double/Triple/Quadra 이미지 즉시 갱신
- 서로 다른 해상도·종횡비도 같은 크기의 letterbox 미리보기로 표시
- 리스트 창은 상단 전체 목록, 하단 Double/Triple/Quadra 이미지 구조
- 방향키, PageUp/PageDown, Home/End, Space 키 탐색
- 이미지 클릭 시 창맞춤·확대·축소 가능한 원본 보기
- 한 Excel에만 이미지가 있는 위치도 누락하지 않고 빈 칸과 함께 표시

## 요구 사항

- Python 3.10 이상
- Windows 10/11 권장
- `PyQt6`, `openpyxl`, `Pillow`

설치:

```bash
pip install -r requirements.txt
```

개발/테스트 의존성까지 설치하려면:

```bash
pip install -r requirements-dev.txt
```

## 실행

Windows에서는 `run_gui.bat`을 실행하거나 다음 명령을 사용합니다.

```bash
python excel_image_inspector_gui.py
```

1. `Double`, `Triple` 또는 `Quadra` 모드를 선택합니다.
2. Excel Ref와 비교A를 선택하고, Triple이면 비교B, Quadra이면 비교B와 비교C도 선택합니다.
3. `이미지 불러오기`를 누릅니다.
4. 버튼이나 키보드로 이미지 위치를 넘기며 육안 검사합니다.
5. `리스트실행`을 누르면 전체 위치 목록과 이미지를 한 창에서 볼 수 있습니다.
6. 이미지를 클릭하면 확대 창이 열립니다.

지원 파일은 `.xlsx`, `.xlsm`입니다. 비교 파일은 기존 테스트에 사용한 것처럼
같은 양식과 시트 순서를 사용하는 것을 권장합니다.

## 위치 정렬 방식

픽셀 내용은 전혀 비교하지 않습니다. 화면에서 함께 보여 줄 위치만 아래 순서로
맞춥니다.

1. 앵커가 정확히 같은 셀
2. 앵커가 같은 병합영역 안에 있는 셀
3. 행과 열이 각각 1칸 이내인 인접 셀
4. 어느 파일에만 있는 이미지는 다른 칸을 비워 독립 위치로 표시

## 프로젝트 구조

```text
excel_image_inspector_gui.py  # 프로그램 진입점
models.py                     # 추출 이미지와 검사 위치 모델
excel_manager.py              # Excel 이미지 위치/원본 bytes 추출
inspection_service.py         # Double/Triple 위치 정렬 및 로딩
app_state.py                  # 현재 시트/위치 탐색 상태
workers.py                    # 백그라운드 Excel 로딩
ui/main_window.py             # 메인 검사 화면
ui/dialogs.py                 # 단일 이미지 확대 창과 전체 이미지 리스트 창
ui/widgets.py                 # 파일 드롭 및 이미지 클릭 위젯
tests/                        # 위치 정렬·추출·탐색 회귀 테스트
```

추출된 원본 이미지는 실행 중 임시폴더에 저장되고, 새 작업을 불러오거나 프로그램을
종료하면 자동으로 정리됩니다.
