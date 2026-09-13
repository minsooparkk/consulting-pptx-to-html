# consulting-pptx-to-html

**완성된 PPTX를 다시 디자인하지 않고, 원래 배치에 가까운 HTML 발표 자료로 옮기는 AI 스킬과 변환 도구입니다.**

`consulting-pptx`로 만든 Pretendard·네이비 컨설팅 자료를 주요 대상으로 합니다. PPTX의 텍스트, 기본 도형, 표, 마스터와 레이아웃의 장식, 제목·Header·Footer를 읽어 HTML/SVG로 재현합니다. 원본 PPTX는 수정하지 않습니다.

> 범용 PowerPoint 렌더러나 픽셀 단위 완전 호환 제품은 아닙니다. 처리하지 못한 요소는 보고서에 남깁니다. 구조 검사 통과와 PowerPoint 화면 일치는 별개입니다.

## 무엇이 다른가요

| 항목 | 동작 |
|---|---|
| 내용과 장수 | PPTX 1장 = HTML 1장. 자동 요약·문체 교정·페이지 분할 없음 |
| 디자인 기준 | 별도 HTML 테마가 아니라 입력 PPTX의 좌표·색상·글꼴·서식 |
| 텍스트와 표 | 선택 가능한 HTML 텍스트와 HTML 표 |
| 기본 도형과 연결선 | SVG. 원래 순서와 위치를 유지 |
| 마스터·레이아웃 | 공통 장식과 Footer를 포함하고 사용된 플레이스홀더 서식을 상속 |
| 발표 조작 | 이전/다음, 번호 이동, 해시 링크, 키보드, 터치, 전체 화면 API |
| 자동 등장 | 슬라이드 진입 후 약 0.92초 안에 내용 표시. 한 번 조작하면 한 슬라이드 이동 |
| 간단 편집 | 글자 단위 서식 구간을 편집하고 수정한 HTML 저장 |
| 네트워크 | 기본 출력에 CDN·분석 코드·외부 스크립트 없음 |
| 미지원 요소 | 오류와 object_id를 기록. 해당 요소만 이미지로 대체 가능 |

텍스트·표·SVG와 플레이어를 오프라인 단일 HTML에 담습니다. 퀴즈, 실시간 계산기, 항목 클릭 강조, 수동 단계 공개, PPTX 애니메이션 재생, HTML 수정의 PPTX 역반영은 기본 기능에 포함되지 않습니다.

## 빠른 시작

### 1. 복제

저장소 접근 권한이 있는 계정으로 복제합니다.

```text
git clone https://github.com/minsooparkk/consulting-pptx-to-html.git
cd consulting-pptx-to-html
```

ZIP으로 받았다면 폴더 이름을 `consulting-pptx-to-html`로 바꾸고 전체 구조를 유지하세요.

### 2. 준비

- **Python 3.9 이상**. 변환·설치·Python 회귀 테스트는 표준 라이브러리만 사용합니다. `pip install`은 필요 없습니다.
- **현대적인 브라우저**. 자동 검사기는 Chrome·Chromium·Edge 실행 파일을 사용합니다.
- **원본 PPTX의 글꼴**. 기본 결과는 시스템에 설치된 글꼴을 사용합니다.
- **PowerPoint 또는 별도의 신뢰할 수 있는 렌더러**. 원본과의 시각 대조용이며 이 스킬에 포함되지 않습니다.

macOS/Linux에서는 아래 명령의 `python`을 `python3`으로 바꿀 수 있습니다. Windows에서는 `py -3`도 사용할 수 있습니다. Mac 전용 PowerShell이나 Node.js 의존성은 없습니다.

### 3. 샘플 변환

```text
python scripts/convert.py examples/sample.pptx -o preview.html
python scripts/check_html.py preview.html --report browser-qa.json
```

`preview.html`을 브라우저로 열면 됩니다. 변환과 함께 `preview.audit.json`이 생성됩니다. 저장소의 `examples/sample.html`도 다운로드해 직접 열 수 있습니다. GitHub 파일 보기 화면은 발표용 HTML을 실행하지 않습니다.

### 4. 실제 PPTX 변환

```text
python scripts/convert.py "완성된_발표자료.pptx" -o "발표자료.html" --motion auto
```

`--motion auto`는 생략해도 기본 적용됩니다. 움직임 없이 즉시 전체 내용을 표시하려면 `--motion none`을 사용하세요. 기존 결과를 의도적으로 교체할 때만 `--overwrite`를 사용하세요. 이미 존재하는 파일은 기본적으로 덮어쓰지 않습니다.

```text
python scripts/convert.py "완성된_발표자료.pptx" -o "발표자료.html" --overwrite
```

HTML 파일이 PPTX보다 반드시 작지는 않습니다. SVG·텍스트의 양과 폰트 포함 여부에 따라 달라지며, 실제 크기는 변환 보고서의 `source_bytes`와 `html_bytes`로 확인합니다.

## Codex / Claude Code에 스킬 설치

복제한 저장소 폴더에서 실행합니다.

```text
python scripts/install.py --target all --dry-run
python scripts/install.py --target all
```

| 옵션 | 설치 경로 |
|---|---|
| `--target codex` | `~/.agents/skills/consulting-pptx-to-html` |
| `--target claude` | `~/.claude/skills/consulting-pptx-to-html` |
| `--target all` | 위 두 위치. 기본값 |

다른 내용의 설치본이 있으면 중단합니다. 업데이트는 원본 저장소에서 먼저 `git pull --ff-only`를 실행한 뒤 진행합니다.

```text
python scripts/install.py --target all --replace --dry-run
python scripts/install.py --target all --replace
```

이전 설치본은 `~/.skill-backups/consulting-pptx-to-html/<target>/<timestamp>/`에 백업합니다. 심볼릭 링크인 설치 경로는 자동 교체하지 않습니다. `--home <경로>`로 별도 홈을 지정할 수 있습니다. 설치 후 각 앱에서 스킬이 인식되는지는 따로 확인하세요.

## 에이전트에 요청하는 예시

```text
consulting-pptx-to-html 스킬로 첨부한 완성 PPTX를 HTML 강의안으로 변환해줘.
원본 문구, 장수, 순서, 글꼴 크기, 표와 도식의 배치를 유지해줘.
제목·Header·Footer는 원본 기준으로 두고, 발표 조작부만 캔버스 밖에 추가해줘.
슬라이드를 넘기면 내용을 짧게 자동 표시하고, 항목 클릭 강조나 수동 단계 공개는 넣지 마.
노트는 포함하지 말고 미지원 요소는 먼저 보고해줘.
HTML과 구조 검사 결과를 만들고, 브라우저 검사와 PowerPoint 대조 결과를 구분해서 알려줘.
```

`consulting-pptx`는 PPTX 제작·디자인 변경용이고, 이 스킬은 **승인된 PPTX의 HTML 버전 제작용**입니다. 원고에서 새 덱을 만드는 도구로 사용하지 않습니다.

## 발표와 편집

기본 자동 등장은 슬라이드가 바뀔 때 실행됩니다. 요소별 재생 시간은 420ms, 시작 지연은 최대 500ms로, 전체 내용이 약 0.92초 안에 나타납니다. 내용의 최종 위치·크기·문구는 원본대로 유지합니다. 슬라이드는 자동으로 넘어가지 않으며, 방향키·Space·이전/다음은 한 번 조작할 때 한 슬라이드 이동합니다. 항목 클릭이나 추가 단계 버튼은 필요 없습니다.

`--motion none`과 운영체제의 모션 감소 설정에서는 내용을 즉시 표시합니다. 편집·저장·인쇄·검수 중에는 등장 효과 때문에 내용이 숨지 않도록 처리합니다. 이 효과는 HTML 플레이어의 발표 기능이며, 원본 PPTX 애니메이션을 재생하는 기능은 아닙니다.

| 조작 | 기능 |
|---|---|
| 방향키, Space, PageUp/PageDown | 이전·다음 슬라이드 |
| Home / End | 처음·마지막 슬라이드 |
| 번호 입력 후 Enter | 특정 슬라이드로 이동 |
| 주소 뒤 `#3` | 3번 슬라이드에서 시작 |
| F / 전체 화면 버튼 | 브라우저 전체 화면 API. 브라우저 권한·실제 사용자 입력 필요 |
| 가로 스와이프 | 이전·다음 슬라이드 |
| E / 편집 버튼 | 텍스트 편집 전환 |
| Ctrl/Cmd+S | 현재 수정 상태의 HTML 다운로드 |
| N / 노트 버튼 | 포함된 발표자 노트 보기 |
| 브라우저 인쇄 | 슬라이드당 한 페이지. 원본 Footer 유지, 조작부 제외 |

편집 중에는 방향키가 글자 커서를 이동하며, 일반 Undo/Redo를 가로채지 않습니다. 붙여넣기는 일반 텍스트로 처리합니다. **HTML 편집은 PPTX를 수정하지 않습니다.** 구조·수치 변경은 원본 PPTX에서 먼저 확정하고 다시 변환하는 것이 안전합니다. 다시 변환하면 HTML에서 한 수정이 자동으로 병합되지 않습니다.

## 글꼴과 오프라인 사용

기본 HTML은 외부 리소스를 요청하지 않지만, 원본 글꼴이 없는 컴퓨터에서는 브라우저가 대체 글꼴을 사용할 수 있습니다. 동일한 글꼴을 설치하거나, 재배포 권한이 있는 글꼴을 포함하세요.

```text
python scripts/convert.py input.pptx -o output.html --font "Pretendard,400,fonts/Pretendard-Regular.woff2" --font "Pretendard,700,fonts/Pretendard-Bold.woff2"
```

- `--font` 형식은 `글꼴이름,굵기,파일경로`이며 반복 지정할 수 있습니다.
- WOFF2·WOFF·TTF·OTF 파일을 지원합니다. 원본에서 사용한 이름과 굵기를 맞추세요.
- 이름을 지정했다고 설치 여부를 검증한 것은 아닙니다. 포함하지 않은 다른 글꼴은 여전히 시스템 환경에 의존합니다.
- 폰트 포함 시 파일이 커집니다. 라이선스와 필요한 저작권 고지를 확인하세요.
- 이 저장소에는 폰트 바이너리가 들어 있지 않습니다.

### PowerPoint와 브라우저의 행간 맞추기

Pretendard의 백분율 행간은 PowerPoint에서 실측한 자연 행높이 계수 **1.2**를 기본 적용합니다. 예를 들어 원본 1.1배는 CSS `line-height: 1.32`, 1.05배는 `1.26`이 됩니다. 글자 크기·텍스트 상자·원본 문구를 바꾸는 보정이 아닙니다. 점 단위로 지정된 행간은 그대로 단위 변환하며, 이 계수를 다른 글꼴에 일괄 적용하지 않습니다.

다른 글꼴 또는 다른 메트릭이 확인된 환경에서는 원본 PowerPoint의 실제 줄 사이 거리를 측정한 뒤 글꼴별로 지정할 수 있습니다.

```text
python scripts/convert.py input.pptx -o output.html --line-height-factor "Pretendard=1.2"
```

이 옵션은 `FAMILY=FACTOR` 형식으로 반복 지정합니다. 적용된 보정은 감사 보고서에 남습니다. 계수는 보편적인 글꼴 공식이 아니며, 실제 원본과 비교해 결정합니다. 측정 방법과 혼합 글꼴의 한계는 [호환 범위](references/compatibility.md)를 확인하세요.

## 노트와 개인정보

노트는 기본 제외입니다. 포함이 필요한 경우에만 다음을 사용합니다.

```text
python scripts/convert.py input.pptx -o output.html --include-notes
```

포함된 노트는 암호화된 발표자 전용 데이터가 아닙니다. **HTML을 받은 사람이 모두 읽을 수 있습니다.** 숨겨진 PPTX 슬라이드는 장수 보존을 위해 출력에 포함되며 경고가 기록됩니다. 배포 전 원본의 비공개·숨김 콘텐츠를 확인하세요. 댓글과 첨부된 OLE 파일은 결과에 포함하지 않습니다.

## 미지원 요소 처리

오류가 있으면 HTML은 검토용 부분 결과이며, 기본 종료 코드는 `2`입니다. 보고서의 `conversion_status`가 `PARTIAL_REVIEW_REQUIRED`이면 완성본으로 배포하지 마세요. `issues`에는 슬라이드 번호와 `object_id`가 있습니다.

차트처럼 직접 재현하기 어려운 요소는 원본 PowerPoint에서 **해당 요소만** PNG 또는 안전한 SVG로 내보낸 뒤 매핑합니다.

`fallbacks.json` 예시:

```json
{
  "s003_slide_7": "fallbacks/chart.png"
}
```

```text
python scripts/convert.py input.pptx -o output.html --fallback-manifest fallbacks.json --overwrite
```

경로는 매핑 JSON 위치를 기준으로 해석합니다. 이 대체 요소 안의 글자는 HTML 텍스트로 편집되지 않습니다. 잘못된 이미지를 지정했는지는 자동으로 판정할 수 없으므로 원본과 대조하세요.

`--allow-unsupported`는 미지원 항목이 있어도 부분 미리보기를 남기고 종료 0을 반환하는 명시적 예외입니다. 오류를 해결하거나 품질을 보증하는 옵션이 아닙니다.

## 검수 방법

### 구조·내용 검사

`output.audit.json`에서 다음을 확인합니다.

- `slide_count`, 슬라이드별 요소 수와 텍스트 구간 수
- `serialization_check: PASS`, 추출한 텍스트와 HTML 직렬화 텍스트의 해시 일치
- `issues`의 오류와 경고, 포함한 노트·글꼴
- `motion`의 자동 등장 모드·시간, 수동 단계 및 항목 클릭 비활성화 설정
- `line_height_calibration`의 기본 계수·재정의 값·실제 적용 문단, `exact_point_spacing_unchanged`
- 실제 출력 파일 바이트 기준의 `html_sha256`, `html_bytes`
- `powerpoint_visual_comparison: NOT_PERFORMED`

`serialization_check`의 텍스트 해시는 변환기가 읽어낸 지원 텍스트의 직렬화 검사입니다. 미지원 개체 내부의 모든 내용을 읽었다는 보증이 아닙니다. `html_sha256`은 이와 별도로 실제 HTML 파일 전체의 바이트를 식별합니다.

`motion.browser_verified: false`는 변환만으로 브라우저 실행을 확인하지 않았다는 뜻입니다. 별도 실제 조작 결과를 기록하기 전에는 동작 검증 완료로 해석하지 않습니다.

### 브라우저 검사

```text
python scripts/check_html.py output.html --report browser-qa.json
```

브라우저를 자동으로 찾지 못하면 실행 파일을 지정합니다.

```text
python scripts/check_html.py output.html --browser "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
```

환경에서 허용하는 브라우저 도구를 사용하세요. 직접 콘솔 실행이 가능한 환경에서는 생성된 HTML을 연 뒤 다음 감사 API도 사용할 수 있습니다.

```javascript
await window.PPTPlayer.audit()
```

검사기는 모든 페이지의 텍스트 경계와 캔버스 밖 요소를 측정하고 원래 페이지로 돌아갑니다. 기본 허용 오차는 2.5 CSS px입니다. 회전된 텍스트 등은 수동 검수 항목으로 남깁니다. 의도적인 겹침, 잘못된 연결 관계, 의미 손실, 글꼴 실제 설치 여부를 완전히 판정하지는 않습니다.

자동 등장이 끝난 화면에서 경계와 원본 배치를 비교하세요. 슬라이드 진입 후 내용이 자동으로 나타나는지, 방향키 한 번에 다음 슬라이드로 이동하는지, 편집·저장·인쇄 시 내용이 빠지지 않는지 실제 조작으로 확인합니다. 접근 정책으로 실행할 수 없는 검사는 `미실시` 또는 `환경 제한`으로 남깁니다. 정책을 우회해 실행하는 것은 검수 요건이 아닙니다.

### 회귀 테스트

```text
python -m unittest discover -s tests -v
```

일반화된 4장 샘플과 임시 OOXML 변형으로 문구 일치, 마스터·레이아웃, 노트 제외, 그룹 좌표, HTML 문자 이스케이프, 이미지·크롭·대체 요소, 미지원 차트 경고, 덮어쓰기 방지를 검사합니다. 테스트용 임시 폴더는 `tests/` 아래에 만들고 제거합니다.

플레이어 상태 회귀는 `jsdom`을 사용할 수 있는 Node.js 환경에서 선택적으로 실행합니다. 자동 등장, 페이지 이동, 편집·인쇄 상태, 저장 후 재열기를 검사하며 실제 브라우저의 레이아웃·다운로드·인쇄 결과 검사를 대신하지 않습니다.

```text
node tests/test_player.cjs examples/sample.html
```

원본 PowerPoint와 HTML을 같은 크기로 나란히 비교하고 실제 비교 범위를 기록하세요. 구조·단위/DOM 검사 통과와 실제 브라우저 동작, PowerPoint 시각 대조는 별도 결과입니다. 렌더러가 없어 비교하지 못했다면 명시하며, HTML에서 넘침이 없다는 사실만으로 원본과 같다고 판단하지 않습니다.

## 초기 검수 범위

다음은 초기 버전의 검사 기록입니다. 현재 변경의 자동 등장·행간 보정 검증 결과를 대신하지 않으며, 각 변환 결과는 별도로 검수합니다.

- 회귀 테스트 7개 통과. 별도 테스트 홈에서 Codex·Claude Code 설치 파일 복사·해시 검사 통과.
- 일반화한 4장 샘플: 82개 배치 요소, 72개 텍스트 구간, 브라우저 넘침·캔버스 밖 요소 0건.
- 저장소에 포함하지 않은 35장 자료: 원본 XML의 표시 텍스트 706개 구간을 독립 대조, 표 11개 변환, 35장 브라우저 경계 검사 통과.
- 35장 HTML을 PDF로 인쇄해 35페이지 생성과 첫·마지막 페이지의 Footer 텍스트를 확인.
- 번호 이동·Home/End·해시 초기화·문구 편집·수정본 직렬화와 재열기 확인.
- **미실시/환경 제한:** PowerPoint 원본과 픽셀 대조, macOS/Linux 실기 실행, 터치 실기와 브라우저 다운로드 UI. 전체 화면 API 요청은 자동화 환경에서 거부됨. 별도 Chrome 실행 파일을 찾을 수 없어 CLI 브라우저 구동 대신 동일한 인페이지 감사 API로 측정함.
- 그림자 근사 경고는 유지. 이 검수 결과를 범용 PPTX 완전 호환이나 모든 환경의 동작 보증으로 해석하지 않음.

## 구성

```text
consulting-pptx-to-html/
├── SKILL.md
├── README.md
├── assets/
│   ├── player.html
│   ├── player.css
│   ├── player.js
│   ├── auto-motion.js
│   └── auto-motion.css
├── scripts/
│   ├── convert.py
│   ├── pptx_html.py
│   ├── check_html.py
│   └── install.py
├── references/
│   └── compatibility.md
├── examples/
│   ├── sample.pptx
│   ├── sample.html
│   ├── sample.audit.json
│   └── sample.browser-qa.json
└── tests/
    ├── test_conversion.py
    └── test_player.cjs
```

## 설계 참고

[kernel-html-slides](https://github.com/hany202507/kernel-html-slides)의 단일 HTML 발표·편집 경험을 기능적으로 참고했습니다. 해당 프로젝트의 템플릿·디자인 강제 규칙·소스 코드를 복사한 배포본은 아닙니다. 이 플레이어와 변환기는 별도로 구현했으며, 디자인의 기준은 입력 PPTX입니다.

[consulting-pptx](https://github.com/minsooparkk/consulting-pptx)의 일반화된 템플릿을 회귀 샘플로 사용합니다. 실제 강의 자료·고객 자료·글꼴 파일은 저장소에 넣지 않습니다. 이 저장소에는 별도의 오픈소스 라이선스를 부여하지 않았습니다.
