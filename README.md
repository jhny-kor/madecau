# madecau

Windows의 공식 MarkAny 복호화 화면을 사용자 권한 범위에서 순차 조작하기 위한 별도 도구입니다.

첫 단계는 버튼과 팝업의 Windows UI Automation 식별자를 수집하는 것입니다. 화면 좌표를 먼저 사용하지 않아 창 위치나 해상도가 달라져도 안정적으로 동작하게 합니다.

## UI 정보 수집

Windows PowerShell에서 실행합니다.

```powershell
cd C:\path\to\gHwpx\markany-ui-automation
powershell.exe -NoProfile -File .\Capture-MarkAnyUi.ps1 -SelfTest
powershell.exe -NoProfile -File .\Capture-MarkAnyUi.ps1 -OutputPath .\captures\markany-ui.json
```

두 번째 명령을 실행한 직후 MarkAny 창을 클릭해 전면에 두면 5초 뒤 해당 창의 UI 구조를 저장합니다. 팝업마다 한 번씩 실행해 별도 JSON 파일로 저장합니다.

```powershell
powershell.exe -NoProfile -File .\Capture-MarkAnyUi.ps1 -OutputPath .\captures\decrypt-popup.json
powershell.exe -NoProfile -File .\Capture-MarkAnyUi.ps1 -OutputPath .\captures\folder-popup.json
```

JSON에는 입력값 자체를 수집하지 않지만 창 제목과 버튼 이름에 문서명 등이 포함될 수 있습니다. 공유하기 전에 민감정보를 확인하고 가립니다.

## 빌드

Windows PC에서 Python 3.10 이상을 설치한 뒤 실행합니다. 나머지는 스크립트가 알아서 합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build-madecau.ps1 -SelfTest
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build-madecau.ps1
```

- `dist\madecau.exe` — 단일 실행 파일. 그대로 복사해서 바로 실행할 수 있습니다.
- `dist\madecau-Setup.exe` — [Inno Setup 6](https://jrsoftware.org/isdl.php)이 설치되어 있을 때만 함께 생성됩니다. 이 파일 하나만 옮겨서 설치하면 시작 메뉴와 바탕화면에 등록됩니다. 관리자 권한은 필요 없습니다.

설치 파일이 필요 없으면 `-SkipInstaller` 를 붙입니다.

설치 화면과 제어판 프로그램 목록에는 이름이 `madecau`, 개발자가 `김지현` 으로 뜹니다.
창 제목과 사이드바에도 같은 이름이 나옵니다 (`markany_auto.py` 의 `APP_NAME`).
`madecau.iss` 의 `AppName` / `AppPublisher` 입니다.

### 아이콘

`madecau.ico` 는 저장소에 들어 있습니다. 빌드 스크립트가 exe 에 박고 설치 파일에도 씁니다.
모양을 고치려면 `tools/make_icon.py` 를 손보고 다시 굽습니다 (Pillow 필요, 빌드에는 안 쓰입니다).

```bash
python3 tools/make_icon.py
```

MarkAny 마크의 각진 아스테리스크를 변형해 가운데를 열쇠구멍으로 뚫은 모양입니다.
16/32px 는 구멍이 메워지므로 구멍을 키운 별도 도안에서 뽑습니다.

> `Build-madecau.ps1` 과 `madecau.iss` 는 **UTF-8 BOM**으로 저장해야 합니다.
> BOM이 없으면 Windows PowerShell 5.1이 파일을 CP949로 읽어 한글이 깨지고, Inno Setup이 `[Setup]` 지시자를 인식하지 못합니다.

## 자동화 실행

빌드한 exe를 실행하거나, 개발 중이라면 `pip install pywinauto tkinterdnd2` 후 직접 실행합니다.
(`tkinterdnd2` 는 드래그앤드롭 전용입니다. 없으면 `파일 추가` / `폴더 추가` 버튼만 쓸 수 있고 나머지는 그대로 동작합니다.)

```powershell
python markany_auto.py            # GUI
python markany_auto.py --selftest # 파일 수집/배치 로직만 검증
python markany_auto.py --dump     # 지금 떠 있는 ESAgent 창 구조 출력
```

진행 중 **ESC** 를 누르면 중지합니다. 자동화가 마우스와 포커스를 가져가므로 GUI를 클릭할 필요 없이 어디서든 누르면 됩니다. GUI의 `중지` 버튼도 같은 동작입니다. 현재 진행 중인 단계가 끝나는 시점에 멈춥니다.

GUI의 `창 구조 저장` 버튼은 `--dump` 와 같은 내용을 저장 폴더의 `markany_dump.txt` 로 남깁니다. 모르는 팝업이 떠서 멈췄을 때 이 파일을 확인합니다.

창은 왼쪽 사이드바(대상 파일 수, 배치, 진행률, 상태)와 오른쪽 본문(파일 목록 / 저장 설정 / 진행 기록)으로 나뉩니다.
사이드바 숫자는 `시작` 을 누른 뒤부터 채워집니다.

GUI에서 파일 또는 폴더를 추가하고 저장 위치를 고른 뒤 시작합니다. 시작하면
이미 떠 있는 `MADRMAgent` 창을 재사용하고, 없으면
`C:\MarkAny\Common\ESAgent.exe`를 실행해 10초 후 실제 창에 연결해 자동화합니다.
제목과 사유의 기본값은 `자동복호화사유`입니다.
목록 칸에 파일이나 폴더를 **끌어다 놓아도** 추가됩니다.

**압축파일**(`.zip`, `.7z`, `.rar`, `.alz` 등)은 어차피 거부되므로 첨부 대상에서 미리 제외합니다. 로그에 `압축파일 제외` 로 남습니다.

이미 복호화된 파일은 **헤더로 미리 판별**합니다 (`이미 복호화된 파일 미리 제외` 체크, 기본 켜짐). 복호화된 파일은 원래 형식의 시그니처가 그대로 보입니다.

| 확장자 | 원본 시그니처 |
| --- | --- |
| `.hwp` `.doc` `.xls` `.ppt` | `D0 CF 11 E0 A1 B1 1A E1` (OLE2) |
| `.hwpx` `.docx` `.xlsx` `.pptx` | `PK\x03\x04` (ZIP) |
| `.pdf` | `%PDF` |

시그니처가 없는 확장자(`.txt` 등)는 판단하지 않고 그냥 첨부해 MarkAny가 결정하게 둡니다. 확실할 때만 제외하므로, 놓치면 첨부됐다가 거부될 뿐입니다.

> 첫 실행에서 로그의 `이미 복호화됨 제외` 목록을 탐색기의 **열쇠 아이콘**과 대조해 보세요. 열쇠가 붙은(=암호화된) 파일이 제외 목록에 있으면 이 체크를 끄시면 됩니다. MarkAny가 원본 헤더를 남기는 방식이라면 헤더 검사로는 구분되지 않습니다.

체크를 끄면 이미 복호화된 **일반 파일**은 첨부 시점에야 알 수 있습니다. `일반파일은 첨부할 수 없습니다` 팝업이 파일마다 하나씩 뜨는데, 전부 확인을 눌러 닫고 **그 파일만 건너뛴 채 계속** 진행합니다. 건너뛴 파일은 `건너뜀 (첨부 거부됨)` 으로 남고 마지막 줄에 집계됩니다. 팝업이 없으면 곧바로 제목 입력으로 넘어갑니다.

### 대기 시간

느린 단계는 고정 시간을 자는 게 아니라 **끝난 것을 확인할 때까지** 기다립니다. 기다리는 동안에도 팝업은 계속 처리합니다.

| 단계 | 기준 | 최대 |
| --- | --- | --- |
| 첨부 거부 팝업 | 3초 동안 새 팝업이 없으면 통과 | 60초 |
| 신청 | 반출 신청 창이 닫힐 때까지 | 300초 |
| 다운로드 | 대화상자/팝업이 15초간 없고, 파일의 수정 시각이 바뀔 때까지 | 180초 |

저장 위치는 둘 중 하나입니다.

- **지정 폴더** — 고른 폴더 한 곳에 모두 저장합니다.
- **원본 폴더** — 각 파일을 원래 있던 폴더에 같은 이름으로 저장합니다. **원본 파일을 덮어쓰며 되돌릴 수 없습니다.** 다운로드는 배치마다 폴더를 한 번만 지정할 수 있어, 이 모드에서는 배치가 폴더 단위로도 나뉩니다.

파일을 15개씩 나눠 배치마다 반출 신청 → 목록 최신 건 더블클릭 → 파일다운 → 지정 폴더 저장까지 반복합니다.
진행 기록은 저장 폴더에 `markany_<시각>.log` 로 남습니다.

## 자동화 범위

- 사용자가 선택한 파일만 순차 처리
- 공식 복호화 버튼과 저장 대화상자 조작
- 파일별 성공, 실패, 재시도 기록
- 저장 결과의 HWP, HWPX, PDF 헤더 확인
- 승인, OTP, 관리자 권한 화면에서는 자동 중단

Windows가 잠기거나 로그아웃된 상태와 UAC 보안 화면은 조작하지 않습니다.
