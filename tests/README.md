# 테스트 안내

테스트는 자동 이미지 비교가 아니라 육안 검사기의 데이터 흐름을 보호합니다.

- `test_image_pairing.py`: 정확 셀, 병합영역, ±1행/열 위치 정렬
- `test_inspection_service.py`: Double/Triple/Quadra 시트 합집합, 역할 상태, 빈 시트, Source ID 무결성, 취소/정리
- `test_app_state.py`: 처음/이전/다음/마지막 및 시트 이동 경계
- `test_sheet_names.py`: Excel 시트 탭 순서 읽기
- `test_image_viewer_dialog.py`: 확대 창의 창맞춤/확대 동작
- `test_image_list_window.py`: 리스트의 `↑/↓` 이동, 시트/이미지 없음 구분, Double/Triple/Quadra 이미지 갱신, 하이퍼링크 모드
- `test_excel_jump.py`: 리스트 칸에서 Excel 이동 대상을 고르는 순수 로직
- `test_preview_layout.py`: Triple 동일 크기 letterbox와 상단 목록·하단 이미지 배치
- `test_main_window.py`: 비교 점수 GUI가 없는 Double/Triple/Quadra 화면과 시트 역할 상태 탐색

실행:

```bash
pytest
```
