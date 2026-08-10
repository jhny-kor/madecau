# MarkAny UI 자동화

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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build-MarkAnyAuto.ps1 -SelfTest
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build-MarkAnyAuto.ps1
```

- `dist\MarkAnyAuto.exe` — 단일 실행 파일. 그대로 복사해서 바로 실행할 수 있습니다.
- `dist\MarkAnyAuto-Setup.exe` — [Inno Setup 6](https://jrsoftware.org/isdl.php)이 설치되어 있을 때만 함께 생성됩니다. 이 파일 하나만 옮겨서 설치하면 시작 메뉴와 바탕화면에 등록됩니다. 관리자 권한은 필요 없습니다.

설치 파일이 필요 없으면 `-SkipInstaller` 를 붙입니다.

> `Build-MarkAnyAuto.ps1` 과 `MarkAnyAuto.iss` 는 **UTF-8 BOM**으로 저장해야 합니다.
> BOM이 없으면 Windows PowerShell 5.1이 파일을 CP949로 읽어 한글이 깨지고, Inno Setup이 `[Setup]` 지시자를 인식하지 못합니다.

## 자동화 실행

빌드한 exe를 실행하거나, 개발 중이라면 `pip install pywinauto` 후 직접 실행합니다.

```powershell
python markany_auto.py            # GUI
python markany_auto.py --selftest # 파일 수집/배치 로직만 검증
python markany_auto.py --dump     # 지금 떠 있는 ESAgent 창 구조 출력
```

진행 중 **ESC** 를 누르면 중지합니다. 자동화가 마우스와 포커스를 가져가므로 GUI를 클릭할 필요 없이 어디서든 누르면 됩니다. GUI의 `중지` 버튼도 같은 동작입니다. 현재 진행 중인 단계가 끝나는 시점에 멈춥니다.

GUI의 `창 구조 저장` 버튼은 `--dump` 와 같은 내용을 저장 폴더의 `markany_dump.txt` 로 남깁니다. 모르는 팝업이 떠서 멈췄을 때 이 파일을 확인합니다.

MADRMAgent(문서보안) 창을 열어둔 채 GUI에서 파일 또는 폴더를 추가하고 저장 위치를 고른 뒤 시작합니다.

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
