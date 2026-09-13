---
name: "consulting-pptx-to-html"
description: "완성된 PPTX를 문구·장수·순서·좌표·글꼴·색상·마스터·Header·Footer를 최대한 보존하는 단일 HTML 발표 자료로 변환한다. 슬라이드 진입 시 내용을 짧게 자동 표시하며 편집 가능한 텍스트·표·SVG를 유지한다. PPTX 디자인 유지, PPT를 HTML 강의안으로 변환, consulting-pptx 결과의 웹 발표 버전 제작 요청에 사용한다."
---

# PPTX에서 HTML로

## 변환 원칙

- 승인된 PPTX가 내용과 디자인의 기준이다. 원본은 읽기 전용으로 취급한다.
- PPTX 한 장은 HTML 한 장이다. 요약, 문체 교정, 페이지 분할, 장수 변경, 폰트 축소, 디자인 재구성을 하지 않는다.
- 제목·결론형 Header·본문·Footer는 원본 좌표와 서식에 놓는다. 플레이어 조작부는 캔버스 밖에 둔다.
- 텍스트는 HTML, 표는 HTML 표, 기본 도형·연결선은 SVG로 만든다. 전체 슬라이드를 이미지로 바꾸지 않는다.
- 기본 결과는 오프라인 단일 HTML이다. 슬라이드가 바뀌면 내용이 짧게 자동 등장하며, 문구·최종 좌표·글꼴 크기는 유지한다. 항목 클릭 강조나 수동 단계 공개는 넣지 않는다.
- 절차를 순서대로 보여 달라는 요청에는 [절차별 누적 등장](references/procedural-reveal.md)을 적용한다. 해당 슬라이드와 객체를 명시적으로 매핑하고, 기본 0.3초 간격으로 앞 단계가 남은 채 다음 단계를 표시한다.
- 지원하지 않는 요소는 조용히 버리지 않는다. 보고서의 슬라이드 번호·object_id로 식별하고, 사용자 승인 후 해당 요소만 대체한다.
- PPTX 애니메이션·SmartArt·차트·OLE의 완전 호환을 주장하지 않는다. 지원 범위는 `references/compatibility.md`를 먼저 확인한다.

## 실행 순서

1. 입력 PPTX와 결과 경로를 확인한다. 소스 장수·비율·폰트·특수 요소를 확인한다. 공개 저장소에는 실제 강의·고객 자료를 넣지 않는다.
2. 스킬 폴더를 기준으로 변환한다. Python 3.9 이상, 표준 라이브러리만 필요하다.

   ```text
   python scripts/convert.py input.pptx -o output.html --motion auto
   ```

3. 함께 생성되는 `output.audit.json`을 읽는다. `serialization_check`, `slide_count`, `issues`, `stats`, `fonts_requested`, `notes_included`, 자동 등장 설정 `motion`, 글꼴 행간 보정 `line_height_calibration`을 확인한다. `html_sha256`은 실제 출력 파일 바이트의 해시다.
   - 절차 매핑을 사용했다면 선택한 슬라이드·단계·객체 ID와 간격도 확인한다. 매핑 없는 슬라이드의 기존 등장 방식은 유지한다.
   - 종료 0: 변환 오류 없음. 구조 검사 통과이지 시각적 동등성 보증은 아니다.
   - 종료 2: 미지원 요소가 있는 부분 결과. 원래 위치의 검토 표시와 플레이어 경고를 확인한다. 완료본으로 전달하지 않는다.
   - `--allow-unsupported`는 명시적으로 부분 미리보기를 요청한 때만 쓴다.
4. 미지원 요소는 원본 PowerPoint에서 해당 요소를 PNG 또는 안전한 SVG로 내보내고, object_id별 JSON을 작성한다. 전체 페이지 이미지로 우회하지 않는다.

   ```text
   python scripts/convert.py input.pptx -o output.html --fallback-manifest fallbacks.json --overwrite
   ```

   `fallbacks.json` 예: `{"s003_slide_7":"fallbacks/chart.png"}`. 대체 이미지가 원래 요소와 일치하는지 별도로 대조한다. 해당 요소는 편집 가능한 텍스트가 아니다.
5. 글꼴을 확인한다. 기본 출력은 CDN/네트워크 요청이 없고 시스템 글꼴을 쓴다. 글꼴이 없으면 설치하거나, 재배포 권한이 있는 파일만 포함한다. 폰트 이름 지정만으로 실제 설치를 확인했다고 말하지 않는다. Pretendard의 백분율 행간에는 자연 행높이 계수 1.2를 적용한다. 다른 글꼴이나 점 단위 행간에 일괄 적용하지 않는다. 원본 실측에 따른 `--line-height-factor FAMILY=FACTOR` 조정은 `references/compatibility.md`를 따른다.

   ```text
   python scripts/convert.py input.pptx -o output.html --font "Pretendard,400,fonts/Pretendard-Regular.woff2" --font "Pretendard,700,fonts/Pretendard-Bold.woff2" --overwrite
   ```

6. HTML을 브라우저로 열고 모든 슬라이드의 텍스트 넘침·경계, 자동 등장, 이동·전체화면·편집·저장·인쇄를 확인한다. 허용된 브라우저 도구나 다음 검사기를 사용한다. 접근 정책상 확인할 수 없는 항목은 미실시로 남기며, 정책 우회를 검수 요건으로 삼지 않는다.

   ```text
   python scripts/check_html.py output.html --report browser-qa.json
   ```

7. PowerPoint 원본과 HTML을 같은 크기로 나란히 비교한다. 자동 등장이 끝난 화면으로 비교한다. 소스 렌더러가 없으면 `PowerPoint 시각 대조 미실시`라고 분명히 알린다. 구조·단위/DOM 검사, 실제 HTML 화면 검사, PowerPoint 일치 검사를 구분한다. 경고를 없애기 위해 문구를 고치거나 원본 글꼴 크기를 줄이지 않는다.
   - 절차별 등장은 실제 시간 흐름에서 중간 단계의 누적과 마지막 전체 노출을 확인한다. `finish()`, `currentTime` 변경, `clear()` 호출, 애니메이션 비활성화로 만든 최종 화면은 등장 완료의 증거가 아니다.
8. 결과 HTML, 구조 보고서, 실제 검사 범위와 미해결 항목을 전달한다. HTML 편집은 PPTX에 역반영되지 않는다.

## 발표 동작과 노트

- `--motion auto`가 기본이다. 일반 슬라이드는 요소별 재생 420ms, 지연 최대 500ms로 전체 내용이 약 0.92초 안에 나타난다. 명시적으로 매핑한 절차 슬라이드는 별도 간격으로 누적 표시하고 결론을 마지막에 보여 준다. 슬라이드 자체는 자동으로 넘기지 않는다.
- 방향키·Space·이전/다음 버튼은 한 번 조작할 때 한 슬라이드 이동한다. 다음 항목 공개에 이동 키를 소모하지 않는다. 클릭 강조, 수동 단계 버튼, 별도 재생 버튼을 기본으로 추가하지 않는다.
- 정적인 결과를 원하면 `--motion none`을 쓴다. 절차 매핑과 함께 지정하면 오류다. 모션 감소 설정, 편집·저장·인쇄·검수와 페이지 복원에서는 내용이 숨은 채 남지 않도록 한다.
- 노트는 기본 제외다. 사용자가 원할 때만 `--include-notes`를 쓴다. HTML에 포함된 노트는 파일을 받은 사람이 모두 읽을 수 있다.
- 발표 중 간단한 문구 수정은 E 또는 편집 버튼, 수정본 저장은 Ctrl/Cmd+S다. 원본 PPTX를 보존한다.
- 퀴즈, 계산기, 수동 단계별 공개, 항목 클릭 강조 같은 추가 상호작용은 사용자가 요청할 때만 넣는다.
- 저장소 사용법과 설치 옵션은 `README.md`, 회귀 검사는 `python -m unittest discover -s tests -v`를 따른다.
