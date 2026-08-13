"""MarkAny(ESAgent) 문서반출 복호화 자동화.

captures/markany-1..5.json 에서 확인한 Win32 컨트롤 ID로 조작한다.
좌표가 아닌 컨트롤 ID를 쓰므로 창 위치/해상도가 바뀌어도 동작한다.

사용:  python markany_auto.py            # GUI
       python markany_auto.py --selftest # 순수 로직 검증 (Windows 불필요)
       python markany_auto.py --dump     # 현재 떠 있는 ESAgent 창 구조 출력
"""

from __future__ import annotations

import logging
import os
import queue
import sys
import threading
import time
from pathlib import Path, PureWindowsPath

BATCH_SIZE = 15  # 반출 신청 팝업의 "(개수 : 15 / 사이즈 : 무제한)"

# ponytail: 파일명 칸은 MAX_PATH(260자) 근처에서 입력을 잘라낸다. 여유를 두고
# 나눠 넣고, 실제로 잘렸는지는 입력 후 읽어서 확인한다.
NAME_FIELD_LIMIT = 250

# captures/*.json 에서 읽은 컨트롤 ID
MAIN_TITLE = "MADRMAgent"
ID_MAIN_REQUEST = 1014      # 반출 신청 버튼
ID_MAIN_SEARCH = 1010       # 검색 버튼
ID_MAIN_LIST = 1002         # 신청 목록 SysListView32

REQ_TITLE = "반출 신청"
ID_REQ_SUBJECT = 1036       # 제목 Edit
ID_REQ_PREPOST = 1031       # 사전/사후 콤보
PREPOST_INDEX = 1           # 0-based. 사전=0, 사후=1
ID_REQ_REASON = 1034        # 사유 Edit
ID_REQ_ATTACH = 1048        # 파일첨부 버튼
ID_REQ_FILELIST = 1047      # 첨부 목록 SysListView32
ID_REQ_SUBMIT = 1037        # 신청 버튼

DETAIL_TITLE_RE = r"문서반출.*"
ID_DETAIL_DOWNLOAD = 1057   # 파일다운 버튼

# 열기/저장 대화상자 제목. 구조로도 판별하므로 목록에 없어도 동작한다.
FILE_DIALOG_TITLES = ("열기", "Open", "다른 이름으로 저장", "Save",
                      "파일 선택", "폴더", "Folder", "찾아보기", "Browse")
APPLY_ALL_TEXT = "이하 동일"  # "이하 동일 파일에 적용" 체크박스
OK_TEXTS = ("확인", "예", "&예", "OK", "&Yes")

# 압축파일은 MarkAny 가 거부하므로 아예 첨부하지 않는다.
ARCHIVE_EXTS = {".zip", ".7z", ".rar", ".alz", ".egg", ".tar", ".gz", ".bz2",
                ".xz", ".cab", ".iso", ".lzh", ".arj", ".ace"}

# 확장자별 원본 시그니처. 헤더가 이대로면 암호화가 풀린 상태라 MarkAny 가
# "일반파일은 첨부할 수 없습니다" 로 거부한다. 모르는 확장자(.txt 등)는
# 판단하지 않고 첨부해서 MarkAny 가 결정하게 둔다.
_OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # hwp, doc, xls, ppt
_ZIP = b"PK\x03\x04"                         # hwpx, docx, xlsx, pptx
FILE_MAGIC = {
    ".hwp": _OLE2, ".doc": _OLE2, ".xls": _OLE2, ".ppt": _OLE2,
    ".hwpx": _ZIP, ".docx": _ZIP, ".xlsx": _ZIP, ".pptx": _ZIP,
    ".pdf": b"%PDF",
}

# 거부된 파일마다 "일반파일은 첨부할 수 없습니다" 팝업이 하나씩 뜬다.
REJECT_POPUP_WAIT = 3.0
# 신청은 첨부 파일을 서버로 올리므로 오래 걸린다.
SUBMIT_TIMEOUT = 300.0
# 다운로드는 복호화까지 끝나야 파일이 나타난다.
DOWNLOAD_TIMEOUT = 180.0

# 자동 중단 키워드 (README: 승인/OTP/관리자 화면은 건드리지 않는다)
ABORT_KEYWORDS = ("OTP", "관리자 권한", "본인 인증", "인증서")

log = logging.getLogger("markany")


# --------------------------------------------------------------------------
# 순수 로직 (Windows 없이도 테스트 가능)
# --------------------------------------------------------------------------
def collect_files(paths) -> list[Path]:
    """파일/폴더 목록을 실제 파일 목록으로 펼친다. 폴더는 재귀."""
    out: list[Path] = []
    for p in paths:
        p = Path(p)
        if p.is_dir():
            out.extend(f for f in sorted(p.rglob("*")) if f.is_file())
        elif p.is_file():
            out.append(p)
    seen, uniq = set(), []
    for f in out:
        key = str(f.resolve()).lower()
        if key not in seen and not f.name.startswith("~$"):
            seen.add(key)
            uniq.append(f)
    return uniq


def split_archives(files: list[Path]) -> tuple[list[Path], list[Path]]:
    """(첨부할 파일, 제외할 압축파일).

    압축파일은 MarkAny 가 첨부를 거부하며 파일마다 팝업을 띄운다. 어차피
    안 되는 것을 넣어 팝업을 만들 이유가 없으므로 미리 뺀다.
    """
    keep: list[Path] = []
    archives: list[Path] = []
    for f in files:
        (archives if f.suffix.lower() in ARCHIVE_EXTS else keep).append(f)
    return keep, archives


def looks_decrypted(path: Path) -> bool:
    """헤더가 원본 형식 그대로면 True (이미 복호화됨). 판단 불가면 False.

    한쪽으로만 틀리게 만들었다. 확실할 때만 True 를 주므로, 놓치면 첨부됐다가
    MarkAny 가 거부할 뿐이다. 반대로 틀리면 복호화가 필요한 파일을 건너뛴다.
    """
    magic = FILE_MAGIC.get(path.suffix.lower())
    if magic is None:
        return False
    try:
        with path.open("rb") as fh:
            return fh.read(len(magic)) == magic
    except OSError:
        return False


def split_decrypted(files: list[Path]) -> tuple[list[Path], list[Path]]:
    """(첨부할 파일, 이미 복호화된 것으로 보이는 파일)."""
    todo: list[Path] = []
    done: list[Path] = []
    for f in files:
        (done if looks_decrypted(f) else todo).append(f)
    return todo, done


def chunks(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def group_by_folder(files: list[Path]) -> dict[Path, list[Path]]:
    """파일선택 대화상자는 폴더 단위로 다중 선택하므로 원본 폴더별로 묶는다."""
    out: dict[Path, list[Path]] = {}
    for f in files:
        out.setdefault(f.parent, []).append(f)
    return out


def make_batches(files: list[Path], dest: Path | None,
                 size: int = BATCH_SIZE) -> list[tuple[list[Path], Path]]:
    """(파일들, 저장폴더) 목록. dest 가 None 이면 각 파일의 원본 폴더에 저장한다.

    다운로드는 배치마다 폴더를 한 번만 고를 수 있다. 원본 폴더에 저장하려면
    한 배치에 한 폴더의 파일만 담아야 한다.
    """
    if dest is not None:
        return [(part, dest) for part in chunks(files, size)]
    out = []
    for folder, group in group_by_folder(files).items():
        out.extend((part, folder) for part in chunks(group, size))
    return out


def quoted_names(files: list[Path]) -> str:
    """파일명 칸에 넣을 문자열. 한 개면 따옴표 없이 이름만."""
    if len(files) == 1:
        return files[0].name
    return " ".join(f'"{f.name}"' for f in files)


def pack_names(files: list[Path], limit: int = NAME_FIELD_LIMIT) -> list[list[Path]]:
    """파일명 칸 길이 제한에 맞춰 파일을 나눈다.

    제한을 넘기면 대화상자가 이름을 잘라버리고 "파일 이름이 올바르지 않습니다"
    를 띄운다. 잘리기 전에 나눠서 여러 번 첨부한다.
    """
    out: list[list[Path]] = []
    current: list[Path] = []
    for f in files:
        if current and len(quoted_names(current + [f])) > limit:
            out.append(current)
            current = []
        current.append(f)
    if current:
        out.append(current)
    return out


# --------------------------------------------------------------------------
# 자동화
# --------------------------------------------------------------------------
class NeedsCapture(RuntimeError):
    """UI 구조를 몰라서 진행할 수 없음. 캡처를 더 떠야 한다."""


class Aborted(RuntimeError):
    """승인/OTP 등 사람이 처리해야 하는 화면이 떠서 중단."""


class MarkAny:
    def __init__(self, stop: "threading.Event | None" = None):
        from pywinauto import Application  # Windows 전용이라 지연 import

        self.stop = stop
        self.app = Application(backend="win32").connect(title=MAIN_TITLE, timeout=15)
        self.main = self.app.window(title=MAIN_TITLE)
        self.main.wait("visible ready", timeout=15)
        # 우리가 여는 창들. 팝업 청소 대상에서 뺀다. 제목으로 거르면 확인
        # 메시지박스가 상세 창과 제목이 같아("문서반출") 같이 걸러진다.
        self._own = {self.main.wrapper_object().handle}

    # ---- 공용 헬퍼 -------------------------------------------------------
    def _win(self, **kw):
        return self.app.window(**kw)

    def _wait(self, timeout=20, **kw):
        w = self._win(**kw)
        w.wait("visible ready", timeout=timeout)
        return w

    def _click(self, parent, control_id, name=""):
        self._guard()
        btn = parent.child_window(control_id=control_id)
        btn.wait("visible enabled", timeout=15)
        btn.click_input()
        log.info("클릭: %s (id=%s)", name or control_id, control_id)
        time.sleep(0.3)

    def _set_text(self, parent, control_id, text, name=""):
        self._guard()
        edit = parent.child_window(control_id=control_id, class_name="Edit")
        edit.wait("visible enabled", timeout=15)
        edit.set_edit_text(text)
        log.info("입력: %s = %r", name or control_id, text)

    def _visible_dialogs(self):
        """ESAgent 프로세스의 보이는 최상위 창들."""
        out = []
        for w in self.app.windows(visible_only=True, top_level_only=True):
            try:
                out.append((w.window_text(), w))
            except Exception:
                continue
        return out

    def _guard(self):
        if self.stop is not None and self.stop.is_set():
            raise Aborted("사용자가 중지했습니다.")
        for title, _ in self._visible_dialogs():
            for kw in ABORT_KEYWORDS:
                if kw in title:
                    raise Aborted(f"사람이 처리해야 하는 화면입니다: {title!r}")

    def _handle_popups(self, tries: dict) -> bool:
        """확인/예만 누르면 되는 팝업을 한 번 처리한다. 눌렀으면 True.

        신청 확인, 덮어쓰기 확인, "이하 동일 파일에 적용", 다운로드 완료 알림이
        모두 여기로 온다. 같은 창은 두 번까지만 시도해서, 안 닫히는 창 하나에
        영원히 매달리지 않는다.
        """
        for title, w in self._visible_dialogs():
            if w.handle in self._own or tries.get(w.handle, 0) >= 2:
                continue
            apply_all = self._find_by_text(w, APPLY_ALL_TEXT)
            if apply_all is not None:
                log.info("'이하 동일 파일에 적용' 체크")
                apply_all.click_input()
                time.sleep(0.2)
            btn = self._find_button(w, OK_TEXTS)
            if btn is None:
                if apply_all is not None:
                    log.warning("%r 팝업에서 확인 버튼을 못 찾았습니다.", title)
                continue
            log.info("팝업 확인: %r", title)
            btn.click_input()
            tries[w.handle] = tries.get(w.handle, 0) + 1
            time.sleep(0.5)
            if self._alive(w):
                log.warning("%r 창이 닫히지 않았습니다 (%d회).", title, tries[w.handle])
            return True
        return False

    def _dismiss_messageboxes(self, seconds=3.0, hard_limit=60.0) -> int:
        """더 뜰 게 없을 때까지 팝업을 닫고, 누른 횟수를 돌려준다.

        거부된 파일마다 팝업이 하나씩 뜨므로 하나 처리할 때마다 마감을 늘려
        다음 팝업을 기다린다.
        """
        deadline = time.time() + seconds
        hard_end = time.time() + hard_limit
        tries: dict[int, int] = {}
        clicked = 0
        while time.time() < deadline and time.time() < hard_end:
            self._guard()
            if self._handle_popups(tries):
                clicked += 1
                deadline = time.time() + seconds
            else:
                time.sleep(0.3)
        return clicked

    @staticmethod
    def _find_button(win, texts):
        """텍스트가 일치하는 버튼. Button > CheckBox/Pane > Static 순.

        이 앱은 버튼을 오너 드로우로 만들어서 pywinauto 가 CheckBox 로 부르는
        경우가 있고, 라벨(Static)로만 그려진 것도 있다. Static 을 누르면 보통
        아무 일도 안 일어나므로 마지막에만 쓴다.
        """
        wanted = {x.replace("&", "") for x in texts}
        rank = {"Button": 0, "CheckBox": 1, "Pane": 1, "Static": 2}
        best, best_rank = None, 99
        for c in win.descendants():
            try:
                t = (c.window_text() or "").strip().replace("&", "")
                cls = c.friendly_class_name()
            except Exception:
                continue
            if t not in wanted:
                continue
            r = rank.get(cls, 99)
            if r < best_rank:
                best, best_rank = c, r
                if r == 0:
                    break
        return best

    @staticmethod
    def _find_by_text(win, needle):
        for c in win.descendants():
            try:
                if needle in (c.window_text() or ""):
                    return c
            except Exception:
                continue
        return None

    # ---- 콤보 -----------------------------------------------------------
    @staticmethod
    def _alive(win):
        try:
            return bool(win.is_visible())
        except Exception:
            return False

    @staticmethod
    def _toplevel():
        from pywinauto import Desktop

        out = set()
        for w in Desktop(backend="win32").windows(visible_only=True, top_level_only=True):
            try:
                out.add((w.handle, w.class_name()))
            except Exception:
                continue
        return out

    def _select_combo(self, parent, control_id, index, name=""):
        """커스텀 드로우 콤보(class=Button)에서 index번째(0-based) 항목을 고른다.

        항목 텍스트로는 찾을 수 없다. 리스트 항목은 별도 HWND 가 아니고,
        텍스트로 뒤지면 콤보 버튼 자신의 캡션("사전/사후 콤보")이 먼저 걸린다.
        그래서 드롭다운이 실제로 뜬 자리를 좌표로 짚어 어떤 창인지 확인한 뒤,
        그 안에서 index번째 줄을 누른다. 사람이 하는 동작 그대로다.
        """
        from pywinauto import Desktop

        combo = parent.child_window(control_id=control_id).wrapper_object()
        rect = combo.rectangle()
        before = self._toplevel()

        combo.click_input()
        log.info("클릭: %s (id=%s)", name or control_id, control_id)
        time.sleep(0.6)

        new = self._toplevel() - before
        if new:
            log.info("새로 뜬 창: %s", sorted(new))

        # 콤보 바로 아래를 짚으면 드롭다운이 열렸는지, 무엇인지 알 수 있다.
        probe = (rect.left + rect.width() // 2, rect.bottom + 8)
        try:
            drop = Desktop(backend="win32").from_point(*probe)
        except Exception as e:  # noqa: BLE001
            drop = None
            log.warning("드롭다운 지점 조회 실패: %s", e)

        if drop is None or drop.handle == combo.handle:
            raise NeedsCapture(
                f"{name or control_id} 드롭다운이 열리지 않았습니다.\n"
                f"짚은 지점 {probe}, 새 창 {sorted(new) or '없음'}\n"
                f"보인 창: {self._window_summary()}"
            )
        log.info("드롭다운: class=%r rect=%s", drop.class_name(), drop.rectangle())

        # 진짜 리스트박스면 항목 사각형을 그대로 쓴다.
        target = None
        try:
            log.info("드롭다운 항목: %s", drop.item_texts())
            r = drop.item_rect(index)
            target = (r.left + 8, (r.top + r.bottom) // 2)
        except Exception as e:  # noqa: BLE001
            log.info("리스트박스 API 사용 불가(%s), 줄 높이로 계산합니다.", e)
            item_h = rect.height()
            target = (20, item_h * index + item_h // 2)

        drop.click_input(coords=target)
        time.sleep(0.5)

        # ponytail: 콤보 캡션은 항상 "사전/사후 콤보"라 고른 값을 되읽을 수 없다.
        # 대신 드롭다운이 닫혔는지로 클릭이 항목에 맞았는지 확인한다.
        if self._alive(drop):
            raise NeedsCapture(
                f"{name or control_id} 항목을 눌렀지만 드롭다운이 닫히지 않았습니다.\n"
                f"누른 좌표 {target} (드롭다운 {drop.rectangle()}, 콤보 {rect})"
            )
        log.info("선택: %s = %d번째 항목", name or control_id, index + 1)

    # ---- 파일 선택 대화상자 -----------------------------------------------
    @staticmethod
    def _scan(win):
        """대화상자 안의 Edit 목록과 버튼 컨트롤 ID를 한 번에 훑는다.

        child_window(...).exists() 는 같은 클래스가 여러 개면
        ElementAmbiguousError 를 던진다. 열기 대화상자에는 파일명 칸과
        주소 표시줄 두 개의 Edit 이 있어 그대로 걸린다. 그래서 자식을
        직접 훑는다.
        """
        edits, buttons = [], set()
        for c in win.descendants():
            try:
                cls, cid = c.class_name(), c.control_id()
            except Exception:
                continue
            if cls == "Edit":
                edits.append(c)
            elif cls == "Button":
                buttons.add(cid)
        return edits, buttons

    @staticmethod
    def _filename_edit(edits):
        """파일명 입력 칸. 주소 표시줄 Edit 과 헷갈리지 않게 ID를 먼저 본다."""
        visible = []
        for e in edits:
            try:
                if e.is_visible():
                    visible.append(e)
            except Exception:
                continue
        pool = visible or edits
        for cid in (1148, 1001):  # 클래식 / 모던(Vista+) 파일명 칸
            for e in pool:
                try:
                    if e.control_id() == cid:
                        return e
                except Exception:
                    continue
        return pool[0] if pool else None

    @staticmethod
    def _button(win, control_id):
        for c in win.descendants():
            try:
                if c.class_name() == "Button" and c.control_id() == control_id:
                    return c
            except Exception:
                continue
        return None

    def _find_file_dialog(self):
        """열기/저장 대화상자를 데스크톱 전체에서 찾는다. 없으면 None.

        ESAgent 프로세스 안에서만 찾으면 대화상자를 다른 프로세스가 띄울 때
        놓친다. 제목은 환경마다 다르므로 '파일명 Edit + 확인 버튼(id=1)'
        구조로 판별하고, 제목까지 맞으면 우선한다.
        """
        from pywinauto import Desktop

        fallback = None
        for w in Desktop(backend="win32").windows(class_name="#32770",
                                                  visible_only=True,
                                                  top_level_only=True):
            try:
                title = w.window_text()
            except Exception:
                continue
            if title == MAIN_TITLE or title == REQ_TITLE or title.startswith("문서반출"):
                continue
            edits, buttons = self._scan(w)
            if 1 not in buttons or self._filename_edit(edits) is None:
                continue
            if any(k in title for k in FILE_DIALOG_TITLES):
                return title, w
            fallback = fallback or (title, w)
        return fallback

    def _wait_file_dialog(self, timeout=25):
        end = time.time() + timeout
        while time.time() < end:
            self._guard()
            found = self._find_file_dialog()
            if found:
                log.info("파일 대화상자: %r", found[0])
                return found
            time.sleep(0.3)
        raise TimeoutError(
            "파일 대화상자를 찾지 못했습니다.\n보인 창: " + self._window_summary()
        )

    @staticmethod
    def _window_summary():
        """실패했을 때 무엇이 떠 있었는지 로그에 남긴다."""
        from pywinauto import Desktop

        rows = []
        for w in Desktop(backend="win32").windows(visible_only=True, top_level_only=True):
            try:
                rows.append(f"{w.window_text()!r}({w.class_name()})")
            except Exception:
                continue
        return ", ".join(rows) or "(없음)"

    def _fill_file_dialog(self, dlg, text: str) -> bool:
        """파일명 칸에 넣고 확인 버튼을 누른다. 입력이 잘리면 누르지 않고 False."""
        edits, _ = self._scan(dlg)
        edit = self._filename_edit(edits)
        btn = self._button(dlg, 1)  # 열기(O) / 저장(S)
        if edit is None or btn is None:
            raise NeedsCapture("파일 대화상자에서 파일명 칸이나 확인 버튼을 찾지 못했습니다.")
        edit.set_edit_text(text)
        time.sleep(0.2)
        got = edit.window_text()
        if got != text:
            # 그대로 누르면 "파일 이름이 올바르지 않습니다" 가 뜬다.
            log.warning("파일명 칸이 %d/%d자로 잘렸습니다.", len(got), len(text))
            return False
        btn.click_input()
        time.sleep(0.6)
        return True

    def _cancel_dialog(self, dlg):
        self._dismiss_messageboxes(seconds=2)
        cancel = self._button(dlg, 2)
        if cancel is not None:
            cancel.click_input()
            time.sleep(0.5)

    @staticmethod
    def _wait_dialog_closed(dlg, timeout=15.0):
        end = time.time() + timeout
        while time.time() < end:
            try:
                if not dlg.is_visible():
                    return True
            except Exception:
                return True  # 창이 사라져 핸들이 무효해진 경우
            time.sleep(0.3)
        return False

    # ---- 파일첨부 --------------------------------------------------------
    def _attach_from_folder(self, req, lv, folder: Path, files: list[Path]) -> int:
        """대화상자를 원본 폴더로 옮긴 뒤 그 폴더의 파일들을 골라 첨부한다.

        파일명 칸에 폴더 경로를 넣고 열면 그 폴더로 이동하고, 이어서 파일명들을
        넣으면 한 번에 선택된다. 첨부된 개수를 돌려준다.
        """
        before = lv.item_count()
        self._click(req, ID_REQ_ATTACH, "파일첨부")
        _, dlg = self._wait_file_dialog()

        if not self._fill_file_dialog(dlg, str(folder)):      # 폴더로 이동
            self._cancel_dialog(dlg)
            return 0
        if not self._fill_file_dialog(dlg, quoted_names(files)):  # 그 안에서 선택
            self._cancel_dialog(dlg)
            return 0

        if not self._wait_dialog_closed(dlg):
            # 다중 선택을 막는 대화상자면 "파일을 찾을 수 없습니다" 류가 뜬다.
            self._cancel_dialog(dlg)
            return 0

        # 이미 복호화된 일반 파일이면 "일반파일은 첨부할 수 없습니다" 팝업이
        # 파일마다 하나씩 뜬다. 전부 확인을 눌러 닫고, 붙은 개수만 세서 계속한다.
        clicked = self._dismiss_messageboxes(seconds=REJECT_POPUP_WAIT)
        if clicked:
            log.info("첨부 거부 팝업 %d개 확인", clicked)
        return lv.item_count() - before

    @staticmethod
    def _attached_names(lv) -> set[str]:
        """첨부 목록에 실제로 올라간 파일명(소문자).

        목록이 전체 경로를 보여줄 수도 있어 이름만 떼어낸다. 윈도우 앱이 준
        문자열이므로 역슬래시를 구분자로 읽는다.
        """
        names = set()
        for i in range(lv.item_count()):
            try:
                text = (lv.get_item(i).text() or "").strip()
            except Exception:
                continue
            if text:
                names.add(PureWindowsPath(text).name.lower())
        return names

    def _attach_all(self, req, files: list[Path]) -> list[Path]:
        """첨부하고 실제로 붙은 파일만 돌려준다.

        압축파일이나 이미 복호화된 일반 파일은 MarkAny 가 거부하면서 오류
        팝업을 띄운다. 확인을 눌러 닫고 그 파일만 건너뛴 채 계속한다.
        """
        lv = req.child_window(control_id=ID_REQ_FILELIST,
                              class_name="SysListView32").wrapper_object()
        for folder, group in group_by_folder(files).items():
            for part in pack_names(group):
                added = self._attach_from_folder(req, lv, folder, part)
                log.info("첨부 %d/%d개: %s", added, len(part), folder)
                if added or len(part) == 1:
                    continue
                # 하나도 안 붙었다. 다중 선택을 안 받는 대화상자일 수 있으니
                # 한 개씩 넣어본다. 거부당한 파일이면 여기서도 안 붙고 넘어간다.
                # ponytail: 느리지만 확실하다.
                log.warning("한 개씩 다시 첨부합니다: %s", folder)
                for f in part:
                    self._attach_from_folder(req, lv, folder, [f])

        count = lv.item_count()
        attached = self._attached_names(lv)
        ok = [f for f in files if f.name.lower() in attached]
        if count and not ok:
            raise NeedsCapture(
                f"첨부 목록 {count}건을 읽었지만 파일명을 맞추지 못했습니다: "
                f"{sorted(attached)[:3]}"
            )
        for f in files:
            if f.name.lower() not in attached:
                log.warning("건너뜀 (첨부 거부됨): %s", f.name)
        return ok

    # ---- 1) 반출 신청 -----------------------------------------------------
    def request_batch(self, files: list[Path], subject: str, reason: str) -> list[Path]:
        """신청하고 실제로 첨부된 파일 목록을 돌려준다. 하나도 없으면 빈 목록."""
        self.main.set_focus()
        self._click(self.main, ID_MAIN_REQUEST, "반출 신청")
        req = self._wait(title=REQ_TITLE, timeout=20)
        req_handle = req.wrapper_object().handle
        self._own.add(req_handle)

        attached = self._attach_all(req, files)
        if not attached:
            log.warning("첨부된 파일이 없어 이 배치를 건너뜁니다.")
            try:
                req.close()
                time.sleep(0.5)
            except Exception as e:  # noqa: BLE001
                log.warning("반출 신청 창을 닫지 못했습니다: %s", e)
            self._dismiss_messageboxes(seconds=2)
            self._own.discard(req_handle)
            return []

        # 거부 팝업이 더 없는 걸 확인하고 제목부터 채운다.
        self._dismiss_messageboxes(seconds=REJECT_POPUP_WAIT)

        self._set_text(req, ID_REQ_SUBJECT, subject, "제목")
        self._select_combo(req, ID_REQ_PREPOST, PREPOST_INDEX, "사전/사후")
        self._set_text(req, ID_REQ_REASON, reason, "사유")

        self._click(req, ID_REQ_SUBMIT, "신청")

        # 신청은 첨부 파일을 올리므로 시간이 걸린다. "신청하시겠습니까?" 와
        # 완료 알림을 처리하면서 창이 닫힐 때까지 기다린다.
        end = time.time() + SUBMIT_TIMEOUT
        tries: dict[int, int] = {}
        while self._alive(req):
            self._guard()
            if time.time() > end:
                raise RuntimeError(
                    f"신청이 {SUBMIT_TIMEOUT:.0f}초 안에 끝나지 않았습니다."
                )
            if not self._handle_popups(tries):
                time.sleep(0.5)
        self._own.discard(req_handle)
        log.info("신청 완료 (%d개)", len(attached))
        return attached

    # ---- 2) 최신 건 열어서 다운로드 ----------------------------------------
    @staticmethod
    def _mtime(p: Path):
        try:
            return p.stat().st_mtime
        except OSError:
            return None

    def download_latest(self, dest: Path, files: list[Path]):
        self.main.set_focus()
        self._click(self.main, ID_MAIN_SEARCH, "검색")
        time.sleep(1.5)

        lv = self.main.child_window(control_id=ID_MAIN_LIST,
                                    class_name="SysListView32").wrapper_object()
        if lv.item_count() == 0:
            raise RuntimeError("신청 목록이 비어 있습니다.")
        lv.get_item(0).click_input(double=True)  # 맨 위 = 최신
        log.info("최신 신청건 열기")

        detail = self._wait(title_re=DETAIL_TITLE_RE, timeout=20)
        detail_handle = detail.wrapper_object().handle
        self._own.add(detail_handle)

        # 원본 폴더에 저장하면 파일이 이미 있으므로 존재 여부로는 확인할 수
        # 없다. 내려받기 전 수정 시각을 기억해 두고 바뀌었는지로 판단한다.
        before = {f.name: self._mtime(dest / f.name) for f in files}

        self._click(detail, ID_DETAIL_DOWNLOAD, "파일다운")
        tries = self._download_pump(dest)

        # 복호화에 시간이 걸리고, 그 사이 "다운로드가 완료되었습니다" 알림이
        # 뜨면 확인을 누른다.
        end = time.time() + DOWNLOAD_TIMEOUT
        while True:
            self._guard()
            self._handle_popups(tries)
            missing = [n for n, t in before.items() if self._mtime(dest / n) == t]
            if not missing or time.time() > end:
                break
            time.sleep(1.0)
        if missing:
            raise RuntimeError(
                f"{len(missing)}개가 원본 파일명으로 저장되지 않았습니다: {missing[:3]}"
            )
        # 남은 완료 알림을 닫고 상세 창도 닫는다. 남겨두면 다음 배치가
        # 이 창을 최신 건으로 착각한다.
        self._dismiss_messageboxes(seconds=2)
        try:
            detail.close()
            time.sleep(0.5)
        except Exception as e:  # noqa: BLE001
            log.warning("문서반출 창을 닫지 못했습니다: %s", e)
        self._own.discard(detail_handle)

        log.info("다운로드 %d개 저장: %s", len(files), dest)
        return len(files)

    def _download_pump(self, dest: Path, idle_timeout=15.0):
        """폴더 선택 / 저장 대화상자와 그 사이 팝업을 더 안 뜰 때까지 처리한다.

        눌러본 팝업 기록을 돌려준다. 다운로드가 끝날 때까지 계속 쓰인다.
        """
        dest.mkdir(parents=True, exist_ok=True)
        tries: dict[int, int] = {}
        saved = 0
        last = time.time()
        while time.time() - last < idle_timeout:
            self._guard()
            handled = self._handle_popups(tries)
            if not handled:
                found = self._find_file_dialog()
                if found:
                    title, dlg = found
                    edits, _ = self._scan(dlg)
                    edit = self._filename_edit(edits)
                    original = ((edit.window_text() if edit else "") or "").strip()
                    if original:
                        # "다른 이름으로 저장": 미리 채워진 원본 파일명을 유지한다.
                        target = str(dest / Path(original).name)
                        what = Path(original).name
                    else:
                        # "다운로드 할 폴더를 선택해 주세요": 폴더만 지정하면
                        # 프로그램이 원본 파일명 그대로 저장한다.
                        target = str(dest)
                        what = f"{title} -> {dest}"
                    if not self._fill_file_dialog(dlg, target):
                        raise RuntimeError(f"경로가 너무 깁니다: {target}")
                    saved += 1
                    log.info("저장 %d: %s", saved, what)
                    handled = True
            if handled:
                last = time.time()
            else:
                time.sleep(0.4)
        return tries

    # ---- 전체 실행 -------------------------------------------------------
    def run(self, files: list[Path], dest: Path | None, subject: str, reason: str):
        """dest 가 None 이면 각 파일을 원본 폴더에 저장한다."""
        batches = make_batches(files, dest)
        done = 0
        skipped = 0
        for n, (batch, folder) in enumerate(batches, 1):
            log.info("=== 배치 %d/%d (%d개 -> %s) ===", n, len(batches), len(batch), folder)
            attached = self.request_batch(batch, subject, reason)
            skipped += len(batch) - len(attached)
            if not attached:
                continue
            done += self.download_latest(folder, attached)
        log.info("끝. 요청 %d개 / 저장 %d개 / 건너뜀 %d개", len(files), done, skipped)
        return done


def watch_stop_key(stop: threading.Event, finished: threading.Event, poll=0.1):
    """ESC 를 누르면 중지 플래그를 세운다.

    자동화가 마우스와 포커스를 가져가므로 GUI 창의 키 바인딩은 안 먹는다.
    포커스와 무관하게 눌림을 읽는 GetAsyncKeyState 로 확인한다.
    """
    import ctypes

    try:
        user32 = ctypes.windll.user32
        user32.GetAsyncKeyState.restype = ctypes.c_short
    except AttributeError:  # 윈도우가 아니면 ESC 감시 없이 중지 버튼만
        return
    VK_ESCAPE = 0x1B
    while not stop.is_set() and not finished.is_set():
        if user32.GetAsyncKeyState(VK_ESCAPE) & 0x8000:
            stop.set()
            return
        time.sleep(poll)


def dump_windows(out_path: Path | None = None):
    """지금 떠 있는 ESAgent 창 구조를 출력/저장 (새 팝업 구조 파악용)."""
    import contextlib
    import io

    from pywinauto import Application

    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        app = Application(backend="win32").connect(title=MAIN_TITLE, timeout=10)
        for w in app.windows(visible_only=True, top_level_only=True):
            print(f"\n===== {w.window_text()!r} ({w.friendly_class_name()}) =====")
            try:
                w.print_control_identifiers(depth=6)
            except Exception as e:  # noqa: BLE001
                print("  실패:", e)

    text = buf.getvalue()
    if out_path is None:
        print(text)
        return None
    out_path.write_text(text, encoding="utf-8")
    return out_path


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    # tkinterdnd2 가 있으면 드래그앤드롭이 되는 root 를 쓴다. 없으면 버튼만.
    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        root, dnd = TkinterDnD.Tk(), True
    except Exception:  # noqa: BLE001
        root, dnd, DND_FILES = tk.Tk(), False, None

    root.title("MarkAny 복호화 자동화")
    root.geometry("720x580")

    paths: list[str] = []
    msgs: queue.Queue[str] = queue.Queue()

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="복호화할 파일 / 폴더"
                       + ("  (여기로 끌어다 놓으세요)" if dnd else "")).pack(anchor="w")
    lb = tk.Listbox(frm, height=8)
    lb.pack(fill="both", expand=True)

    def add_path(p: str):
        p = p.strip()
        if not p or p in paths:
            return
        paths.append(p)
        lb.insert("end", p + ("  (폴더)" if Path(p).is_dir() else ""))

    if dnd:
        # 윈도우는 공백이 든 경로를 중괄호로 감싸 보낸다. splitlist 가 풀어준다.
        lb.drop_target_register(DND_FILES)
        lb.dnd_bind("<<Drop>>",
                    lambda e: [add_path(p) for p in root.tk.splitlist(e.data)])

    row = ttk.Frame(frm)
    row.pack(fill="x", pady=4)

    def add_files():
        for f in filedialog.askopenfilenames():
            add_path(f)

    def add_folder():
        d = filedialog.askdirectory()
        if d:
            add_path(d)

    def clear():
        paths.clear()
        lb.delete(0, "end")

    ttk.Button(row, text="파일 추가", command=add_files).pack(side="left")
    ttk.Button(row, text="폴더 추가", command=add_folder).pack(side="left", padx=4)
    ttk.Button(row, text="비우기", command=clear).pack(side="left")

    opts = ttk.Frame(frm)
    opts.pack(fill="x", pady=6)

    ttk.Label(opts, text="저장 위치").grid(row=0, column=0, sticky="w")
    mode_var = tk.StringVar(value="fixed")
    modes = ttk.Frame(opts)
    modes.grid(row=0, column=1, sticky="w", padx=4)
    ttk.Radiobutton(modes, text="지정 폴더", variable=mode_var,
                    value="fixed").pack(side="left")
    ttk.Radiobutton(modes, text="원본 폴더 (원본을 덮어씁니다)", variable=mode_var,
                    value="source").pack(side="left", padx=8)

    ttk.Label(opts, text="저장 폴더").grid(row=1, column=0, sticky="w")
    dest_var = tk.StringVar(value=str(Path.home() / "Downloads"))
    dest_entry = ttk.Entry(opts, textvariable=dest_var, width=60)
    dest_entry.grid(row=1, column=1, padx=4)
    browse_btn = ttk.Button(
        opts, text="찾아보기",
        command=lambda: dest_var.set(filedialog.askdirectory() or dest_var.get()))
    browse_btn.grid(row=1, column=2)

    def sync_mode(*_):
        state = "disabled" if mode_var.get() == "source" else "normal"
        dest_entry.configure(state=state)
        browse_btn.configure(state=state)

    mode_var.trace_add("write", sync_mode)

    ttk.Label(opts, text="제목 / 사유").grid(row=2, column=0, sticky="w", pady=4)
    subject_var = tk.StringVar(value="복호화A")
    ttk.Entry(opts, textvariable=subject_var, width=60).grid(row=2, column=1, padx=4, sticky="w")

    precheck_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(opts, text="이미 복호화된 파일 미리 제외 (헤더 검사)",
                    variable=precheck_var).grid(row=3, column=1, sticky="w", pady=2)

    logbox = tk.Text(frm, height=14, state="disabled")
    logbox.pack(fill="both", expand=True, pady=6)

    btns = ttk.Frame(frm)
    btns.pack(fill="x")

    def save_dump():
        """콤보 드롭다운이나 모르는 팝업이 떴을 때 구조를 파일로 남긴다."""
        out = Path(dest_var.get() or Path.home()) / "markany_dump.txt"
        try:
            out.parent.mkdir(parents=True, exist_ok=True)
            dump_windows(out)
            messagebox.showinfo("저장됨", f"창 구조를 저장했습니다.\n{out}")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("실패", str(e))

    stop = threading.Event()
    finished = threading.Event()

    ttk.Button(btns, text="창 구조 저장", command=save_dump).pack(side="left")
    ttk.Label(btns, text="  진행 중 ESC 를 누르면 중지").pack(side="left")
    start_btn = ttk.Button(btns, text="시작")
    start_btn.pack(side="right")
    stop_btn = ttk.Button(btns, text="중지", state="disabled")
    stop_btn.pack(side="right", padx=4)

    class GuiHandler(logging.Handler):
        def emit(self, record):
            msgs.put(self.format(record))

    def pump():
        while not msgs.empty():
            logbox.configure(state="normal")
            logbox.insert("end", msgs.get() + "\n")
            logbox.see("end")
            logbox.configure(state="disabled")
        root.after(200, pump)

    def worker(files, dest, subject):
        try:
            MarkAny(stop).run(files, dest, subject, subject)
            msgs.put("완료되었습니다.")
        except Exception as e:  # noqa: BLE001
            log.error("중단: %s", e)
        finally:
            finished.set()
            root.after(0, lambda: (start_btn.configure(state="normal"),
                                   stop_btn.configure(state="disabled")))

    def request_stop():
        stop.set()
        msgs.put("중지 요청됨. 진행 중인 단계가 끝나면 멈춥니다.")

    def start():
        files, archives = split_archives(collect_files(paths))
        decrypted: list[Path] = []
        if precheck_var.get():
            files, decrypted = split_decrypted(files)
        if not files:
            messagebox.showwarning(
                "확인",
                "복호화할 파일이 없습니다. "
                f"(압축파일 {len(archives)}개, 이미 복호화됨 {len(decrypted)}개)"
                if archives or decrypted else "파일이나 폴더를 먼저 추가하세요.")
            return

        to_source = mode_var.get() == "source"
        if to_source:
            if not messagebox.askyesno(
                    "원본 폴더에 저장",
                    f"복호화한 {len(files)}개 파일을 원본 폴더에 같은 이름으로 저장합니다.\n"
                    "원본 파일은 덮어써지며 되돌릴 수 없습니다.\n\n계속할까요?"):
                return
            dest = None
        else:
            if not dest_var.get().strip():
                messagebox.showwarning("확인", "저장 폴더를 지정하세요.")
                return
            dest = Path(dest_var.get())
            dest.mkdir(parents=True, exist_ok=True)

        log_dir = Path(dest_var.get().strip() or Path.home())
        log_dir.mkdir(parents=True, exist_ok=True)
        setup_logging(log_dir, GuiHandler())
        for f in archives:
            log.warning("압축파일 제외: %s", f.name)
        for f in decrypted:
            log.warning("이미 복호화됨 제외 (헤더 확인): %s", f.name)
        batches = make_batches(files, dest)
        log.info("대상 %d개 파일, %d배치, 저장 위치: %s (중지: ESC)",
                 len(files), len(batches), "원본 폴더" if to_source else dest)
        stop.clear()
        finished.clear()
        start_btn.configure(state="disabled")
        stop_btn.configure(state="normal")
        threading.Thread(target=watch_stop_key, args=(stop, finished),
                         daemon=True).start()
        threading.Thread(target=worker, args=(files, dest, subject_var.get()),
                         daemon=True).start()

    start_btn.configure(command=start)
    stop_btn.configure(command=request_stop)
    pump()
    root.mainloop()


def setup_logging(dest: Path, extra: logging.Handler | None = None):
    log.handlers.clear()
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S")
    fh = logging.FileHandler(dest / f"markany_{time.strftime('%Y%m%d_%H%M%S')}.log",
                             encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    log.addHandler(sh)
    if extra:
        extra.setFormatter(fmt)
        log.addHandler(extra)


def selftest():
    import tempfile

    assert list(chunks(list(range(32)), 15)) == [
        list(range(15)), list(range(15, 30)), [30, 31]
    ]
    assert list(chunks([], 15)) == []

    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        (root / "sub").mkdir()
        (root / "a.pdf").write_text("x")
        (root / "sub" / "b.hwp").write_text("x")
        (root / "~$tmp.hwp").write_text("x")
        got = [f.name for f in collect_files([root])]
        assert got == ["a.pdf", "b.hwp"], got
        # 폴더와 그 안의 파일을 같이 넣어도 중복되지 않는다
        got = collect_files([root, root / "a.pdf"])
        assert len(got) == 2, got

        # 첨부는 원본 폴더별로 묶여야 대화상자에서 한 번에 선택할 수 있다
        groups = group_by_folder(collect_files([root]))
        assert list(groups) == [root, root / "sub"], list(groups)
        assert [f.name for f in groups[root]] == ["a.pdf"]
        assert [f.name for f in groups[root / "sub"]] == ["b.hwp"]

    # 헤더로 복호화 여부 판별. 원본 시그니처가 보이면 이미 풀린 파일이다.
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        cases = {
            "풀린.hwp": _OLE2 + b"rest",
            "잠긴.hwp": b"MADRM\x00\x01\x02\x03",
            "풀린.pdf": b"%PDF-1.7\n",
            "잠긴.pdf": b"\x00\x01\x02\x03%PDF",   # 헤더가 앞에 없으면 암호화
            "풀린.docx": _ZIP + b"rest",
            "메모.txt": "아무거나".encode(),       # 시그니처 없는 확장자
        }
        for name, data in cases.items():
            (root / name).write_bytes(data)
        todo, done = split_decrypted([root / n for n in cases])
        assert [f.name for f in done] == ["풀린.hwp", "풀린.pdf", "풀린.docx"], done
        # 판단 못 하는 건 첨부해서 MarkAny 가 결정하게 둔다
        assert [f.name for f in todo] == ["잠긴.hwp", "잠긴.pdf", "메모.txt"], todo
        assert not looks_decrypted(root / "없는파일.hwp")  # 못 읽으면 False

    # 압축파일은 MarkAny 가 거부하므로 첨부 대상에서 미리 뺀다.
    mixed = [Path("D:/원본/문서.hwp"), Path("D:/원본/모음.ZIP"),
             Path("D:/원본/자료.tar.gz"), Path("D:/원본/보고서.pdf"),
             Path("D:/원본/백업.alz")]
    keep, archives = split_archives(mixed)
    assert [f.name for f in keep] == ["문서.hwp", "보고서.pdf"], keep
    assert [f.name for f in archives] == ["모음.ZIP", "자료.tar.gz", "백업.alz"], archives
    assert len(keep) + len(archives) == len(mixed)  # 빠지는 파일 없음

    # 첨부 목록 판독. 거부된 파일을 건너뛰려면 실제로 붙은 이름을 알아야 한다.
    class _LV:
        def __init__(self, texts):
            self._texts = texts

        def item_count(self):
            return len(self._texts)

        def get_item(self, i):
            text = self._texts[i]
            return type("It", (), {"text": lambda self, t=text: t})()

    lv = _LV([r"D:\원본\가 나.hwp", "다.PDF", "   ", None])
    names = MarkAny._attached_names(lv)
    assert names == {"가 나.hwp", "다.pdf"}, names  # 경로는 이름만, 대소문자 무시

    wanted = [Path("D:/원본/가 나.hwp"), Path("D:/원본/다.pdf"),
              Path("D:/원본/압축.zip")]
    ok = [f for f in wanted if f.name.lower() in names]
    assert [f.name for f in ok] == ["가 나.hwp", "다.pdf"], ok  # 거부된 zip 은 빠진다

    # 저장 위치. 지정 폴더면 15개씩, 원본 폴더면 폴더 단위로도 쪼갠다
    # (다운로드는 배치마다 폴더를 한 번만 고를 수 있다).
    a = [Path(f"D:/원본/a{i}.hwp") for i in range(20)]
    b = [Path(f"D:/다른/b{i}.hwp") for i in range(3)]

    fixed = make_batches(a + b, Path("C:/저장"))
    assert [len(p) for p, _ in fixed] == [15, 8], fixed
    assert all(d == Path("C:/저장") for _, d in fixed)

    src = make_batches(a + b, None)
    assert [(len(p), d) for p, d in src] == [
        (15, Path("D:/원본")), (5, Path("D:/원본")), (3, Path("D:/다른"))], src
    # 원본 폴더 모드는 배치 안 파일이 모두 그 폴더 소속이어야 한다
    assert all(f.parent == d for p, d in src for f in p)
    # 어느 쪽이든 파일이 빠지거나 순서가 섞이지 않는다
    for batches in (fixed, src):
        assert [f for p, _ in batches for f in p] == a + b

    # 열기 대화상자에는 Edit 이 여러 개다. 주소 표시줄이 아니라
    # 파일명 칸을 골라야 한다 (예전에 여기서 탐지가 통째로 실패했다).
    class _Ctrl:
        def __init__(self, cls, cid, visible=True):
            self._cls, self._cid, self._vis = cls, cid, visible

        def class_name(self):
            return self._cls

        def control_id(self):
            return self._cid

        def is_visible(self):
            return self._vis

    class _Win:
        def __init__(self, children):
            self._children = children

        def descendants(self):
            return self._children

    # 메시지박스의 "확인"은 Static 라벨이 아니라 진짜 Button 이어야 한다.
    # Static 을 눌러봐야 창이 안 닫혀서 같은 창을 계속 두드리게 된다.
    class _Named(_Ctrl):
        def __init__(self, cls, text):
            super().__init__(cls, 0)
            self._text = text

        def window_text(self):
            return self._text

        def friendly_class_name(self):
            return self._cls

    label = _Named("Static", "확인")
    button = _Named("Button", "확인")
    assert MarkAny._find_button(_Win([label, button]), OK_TEXTS) is button
    assert MarkAny._find_button(_Win([button, label]), OK_TEXTS) is button
    # 오너 드로우 버튼은 pywinauto 가 CheckBox 로 부른다. Button 다음 순위.
    owner_drawn = _Named("CheckBox", "확인")
    assert MarkAny._find_button(_Win([label, owner_drawn]), OK_TEXTS) is owner_drawn
    assert MarkAny._find_button(_Win([owner_drawn, button]), OK_TEXTS) is button
    assert MarkAny._find_button(_Win([label]), OK_TEXTS) is label  # 최후 수단
    assert MarkAny._find_button(_Win([_Named("Button", "취소")]), OK_TEXTS) is None
    assert MarkAny._find_button(_Win([_Named("Button", "&예")]), OK_TEXTS) is not None

    address = _Ctrl("Edit", 41477)
    filename = _Ctrl("Edit", 1001)
    dlg = _Win([address, filename, _Ctrl("Button", 1), _Ctrl("Button", 2)])

    edits, buttons = MarkAny._scan(dlg)
    assert len(edits) == 2 and buttons == {1, 2}, (edits, buttons)
    assert MarkAny._filename_edit(edits) is filename
    assert MarkAny._button(dlg, 2) is dlg.descendants()[3]
    assert MarkAny._button(dlg, 99) is None
    # 주소 표시줄이 숨어 있어도 결과는 같다
    assert MarkAny._filename_edit([_Ctrl("Edit", 41477, visible=False), filename]) is filename

    # 파일명 칸 길이 제한. 넘기면 이름이 잘려 "파일 이름이 올바르지 않습니다" 가 뜬다.
    long_names = [Path(f"D:/원본/{'가' * 40}_{i}.hwp") for i in range(8)]
    parts = pack_names(long_names, limit=250)
    assert sum(len(p) for p in parts) == 8, parts
    assert [f for p in parts for f in p] == long_names  # 순서와 누락 없음
    assert all(len(quoted_names(p)) <= 250 for p in parts), [len(quoted_names(p)) for p in parts]
    assert len(parts) > 1, "제한을 넘겼는데 나뉘지 않았다"
    # 한 개면 따옴표 없이, 여러 개면 따옴표로 감싼다
    assert quoted_names(long_names[:1]) == long_names[0].name
    assert quoted_names(long_names[:2]).startswith('"')
    # 제한보다 긴 이름 하나는 혼자 담긴다 (자를 방법이 없다)
    huge = [Path("D:/원본/" + "나" * 300 + ".hwp")]
    assert pack_names(huge, limit=250) == [huge]

    print("selftest: ok")


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else ""
    if arg == "--selftest":
        selftest()
    elif arg == "--dump":
        logging.basicConfig(level=logging.INFO)
        dump_windows()
    else:
        gui()
