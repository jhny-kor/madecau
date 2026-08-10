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
from pathlib import Path

BATCH_SIZE = 15  # 반출 신청 팝업의 "(개수 : 15 / 사이즈 : 무제한)"

# captures/*.json 에서 읽은 컨트롤 ID
MAIN_TITLE = "MADRMAgent"
ID_MAIN_REQUEST = 1014      # 반출 신청 버튼
ID_MAIN_SEARCH = 1010       # 검색 버튼
ID_MAIN_LIST = 1002         # 신청 목록 SysListView32

REQ_TITLE = "반출 신청"
ID_REQ_SUBJECT = 1036       # 제목 Edit
ID_REQ_PREPOST = 1031       # 사전/사후 콤보
ID_REQ_REASON = 1034        # 사유 Edit
ID_REQ_ATTACH = 1048        # 파일첨부 버튼
ID_REQ_FILELIST = 1047      # 첨부 목록 SysListView32
ID_REQ_SUBMIT = 1037        # 신청 버튼

DETAIL_TITLE_RE = r"문서반출.*"
ID_DETAIL_DOWNLOAD = 1057   # 파일다운 버튼

# 열기/저장 대화상자 제목. 구조로도 판별하므로 목록에 없어도 동작한다.
FILE_DIALOG_TITLES = ("열기", "Open", "다른 이름으로 저장", "Save",
                      "파일 선택", "찾아보기", "Browse")
APPLY_ALL_TEXT = "이하 동일"  # "이하 동일 파일에 적용" 체크박스
OK_TEXTS = ("확인", "예", "&예", "OK", "&Yes")

# 자동 중단 키워드 (README: 승인/OTP/관리자 화면은 건드리지 않는다)
ABORT_KEYWORDS = ("OTP", "관리자 권한", "본인 인증", "인증서")

# 사전/사후 콤보 드롭다운 구조를 못 찾을 때의 키보드 대체 입력.
# 드롭다운을 연 상태로 Capture-MarkAnyUi.ps1 을 돌려 구조를 확인하는 편이 안전하다.
COMBO_FALLBACK_KEYS: str | None = None  # 예: "{DOWN}{ENTER}"

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


def chunks(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def group_by_folder(files: list[Path]) -> dict[Path, list[Path]]:
    """파일선택 대화상자는 폴더 단위로 다중 선택하므로 원본 폴더별로 묶는다."""
    out: dict[Path, list[Path]] = {}
    for f in files:
        out.setdefault(f.parent, []).append(f)
    return out


# --------------------------------------------------------------------------
# 자동화
# --------------------------------------------------------------------------
class NeedsCapture(RuntimeError):
    """UI 구조를 몰라서 진행할 수 없음. 캡처를 더 떠야 한다."""


class Aborted(RuntimeError):
    """승인/OTP 등 사람이 처리해야 하는 화면이 떠서 중단."""


class MarkAny:
    def __init__(self):
        from pywinauto import Application  # Windows 전용이라 지연 import

        self.app = Application(backend="win32").connect(title=MAIN_TITLE, timeout=15)
        self.main = self.app.window(title=MAIN_TITLE)
        self.main.wait("visible ready", timeout=15)

    # ---- 공용 헬퍼 -------------------------------------------------------
    def _win(self, **kw):
        return self.app.window(**kw)

    def _wait(self, timeout=20, **kw):
        w = self._win(**kw)
        w.wait("visible ready", timeout=timeout)
        return w

    def _click(self, parent, control_id, name=""):
        btn = parent.child_window(control_id=control_id)
        btn.wait("visible enabled", timeout=15)
        btn.click_input()
        log.info("클릭: %s (id=%s)", name or control_id, control_id)
        time.sleep(0.3)

    def _set_text(self, parent, control_id, text, name=""):
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
        for title, _ in self._visible_dialogs():
            for kw in ABORT_KEYWORDS:
                if kw in title:
                    raise Aborted(f"사람이 처리해야 하는 화면입니다: {title!r}")

    def _dismiss_messageboxes(self, known_titles, seconds=3.0):
        """신청 확인/완료 같은 작은 메시지박스를 확인 버튼으로 닫는다."""
        end = time.time() + seconds
        while time.time() < end:
            self._guard()
            for title, w in self._visible_dialogs():
                if any(k in title for k in known_titles):
                    continue
                btn = self._find_button(w, OK_TEXTS)
                if btn is not None:
                    log.info("메시지박스 닫기: %r", title)
                    btn.click_input()
                    time.sleep(0.5)
                    end = time.time() + seconds
            time.sleep(0.3)

    @staticmethod
    def _find_button(win, texts):
        for c in win.descendants():
            try:
                t = c.window_text()
                cls = c.friendly_class_name()
            except Exception:
                continue
            if t and any(t.strip() == x or t.strip() == x.replace("&", "") for x in texts):
                if cls in ("Button", "Pane", "CheckBox", "Static"):
                    return c
        return None

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
    def _select_combo(self, parent, control_id, value, name=""):
        """커스텀 드로우 콤보(class=Button)에서 항목을 고른다."""
        from pywinauto import Desktop

        before = {w.handle for _, w in self._visible_dialogs()}
        self._click(parent, control_id, name)
        time.sleep(0.5)

        candidates = []
        for w in Desktop(backend="win32").windows(class_name="ComboLBox", visible_only=True):
            candidates.append(w)
        for _, w in self._visible_dialogs():
            if w.handle not in before:
                candidates.append(w)
        candidates.append(parent)

        for cand in candidates:
            try:
                item = self._find_by_text(cand, value)
            except Exception:
                continue
            if item is not None:
                item.click_input()
                log.info("선택: %s = %s", name or control_id, value)
                time.sleep(0.3)
                return

        if COMBO_FALLBACK_KEYS:
            from pywinauto.keyboard import send_keys
            log.warning("드롭다운을 못 찾아 키보드 대체 입력 사용: %s", COMBO_FALLBACK_KEYS)
            send_keys(COMBO_FALLBACK_KEYS)
            return

        raise NeedsCapture(
            f"{name or control_id} 드롭다운에서 {value!r} 항목을 찾지 못했습니다.\n"
            "드롭다운을 연 상태로 Capture-MarkAnyUi.ps1 을 실행해 JSON을 남겨주세요."
        )

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

    def _fill_file_dialog(self, dlg, text: str):
        edits, _ = self._scan(dlg)
        edit = self._filename_edit(edits)
        btn = self._button(dlg, 1)  # 열기(O) / 저장(S)
        if edit is None or btn is None:
            raise NeedsCapture("파일 대화상자에서 파일명 칸이나 확인 버튼을 찾지 못했습니다.")
        edit.set_edit_text(text)
        time.sleep(0.2)
        btn.click_input()
        time.sleep(0.6)

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

        파일명 칸에 폴더 경로를 넣고 열면 그 폴더로 이동하고, 이어서 따옴표로
        묶은 파일명들을 넣으면 한 번에 선택된다. 첨부된 개수를 돌려준다.
        """
        before = lv.item_count()
        self._click(req, ID_REQ_ATTACH, "파일첨부")
        _, dlg = self._wait_file_dialog()

        self._fill_file_dialog(dlg, str(folder))          # 폴더로 이동
        names = " ".join(f'"{f.name}"' for f in files)    # 그 폴더 안에서 선택
        self._fill_file_dialog(dlg, names)

        if not self._wait_dialog_closed(dlg):
            # 다중 선택을 막는 대화상자면 "파일을 찾을 수 없습니다" 류가 뜬다.
            self._dismiss_messageboxes({MAIN_TITLE, REQ_TITLE}, seconds=2)
            cancel = self._button(dlg, 2)
            if cancel is not None:
                cancel.click_input()
                time.sleep(0.5)
            return 0
        return lv.item_count() - before

    def _attach_all(self, req, files: list[Path]):
        lv = req.child_window(control_id=ID_REQ_FILELIST,
                              class_name="SysListView32").wrapper_object()
        for folder, group in group_by_folder(files).items():
            added = self._attach_from_folder(req, lv, folder, group)
            if added == len(group):
                log.info("첨부 %d개: %s", added, folder)
                continue
            if added:
                raise RuntimeError(
                    f"{folder} 에서 {len(group)}개 중 {added}개만 첨부되었습니다."
                )
            # ponytail: 다중 선택을 안 받는 대화상자면 한 개씩. 느리지만 확실하다.
            log.warning("다중 선택이 안 되어 한 개씩 첨부합니다: %s", folder)
            for f in group:
                if self._attach_from_folder(req, lv, folder, [f]) != 1:
                    raise RuntimeError(f"첨부 실패: {f}")
                log.info("첨부: %s", f.name)

        got = lv.item_count()
        if got != len(files):
            raise RuntimeError(f"첨부 개수 불일치: 기대 {len(files)}, 실제 {got}")

    # ---- 1) 반출 신청 -----------------------------------------------------
    def request_batch(self, files: list[Path], subject: str, reason: str):
        self.main.set_focus()
        self._click(self.main, ID_MAIN_REQUEST, "반출 신청")
        req = self._wait(title=REQ_TITLE, timeout=20)

        self._attach_all(req, files)

        self._set_text(req, ID_REQ_SUBJECT, subject, "제목")
        self._select_combo(req, ID_REQ_PREPOST, "사후", "사전/사후")
        self._set_text(req, ID_REQ_REASON, reason, "사유")

        self._click(req, ID_REQ_SUBMIT, "신청")
        self._dismiss_messageboxes({MAIN_TITLE, REQ_TITLE}, seconds=4)
        req.wait_not("visible", timeout=30)
        log.info("신청 완료 (%d개)", len(files))

    # ---- 2) 최신 건 열어서 다운로드 ----------------------------------------
    def download_latest(self, dest: Path, expect: int):
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
        self._click(detail, ID_DETAIL_DOWNLOAD, "파일다운")
        saved = self._download_pump(dest, expect)
        log.info("다운로드 %d개 저장: %s", saved, dest)
        return saved

    def _download_pump(self, dest: Path, expect: int, idle_timeout=12.0):
        """저장 대화상자 / '이하 동일 파일에 적용' 팝업이 더 안 뜰 때까지 처리."""
        dest.mkdir(parents=True, exist_ok=True)
        saved = 0
        last = time.time()
        while time.time() - last < idle_timeout and saved < expect:
            self._guard()
            handled = False
            for title, w in self._visible_dialogs():
                # 덮어쓰기 확인이 저장 대화상자보다 먼저
                yes = self._find_button(w, ("예", "&예", "&Yes"))
                if yes is not None and "확인" in title:
                    yes.click_input()
                    handled = True
                    break
                apply_all = self._find_by_text(w, APPLY_ALL_TEXT)
                if apply_all is not None:
                    log.info("'이하 동일 파일에 적용' 체크 후 확인")
                    apply_all.click_input()
                    time.sleep(0.2)
                    ok = self._find_button(w, OK_TEXTS)
                    if ok is None:
                        raise NeedsCapture(f"{title!r} 팝업에서 확인 버튼을 못 찾았습니다.")
                    ok.click_input()
                    handled = True
                    break
            if not handled:
                found = self._find_file_dialog()
                if found:
                    _, dlg = found
                    edits, _ = self._scan(dlg)
                    edit = self._filename_edit(edits)
                    name = Path((edit.window_text() if edit else "") or
                                f"file_{saved + 1}").name
                    self._fill_file_dialog(dlg, str(dest / name))
                    saved += 1
                    log.info("저장 %d: %s", saved, name)
                    handled = True
            if handled:
                last = time.time()
            else:
                time.sleep(0.4)
        return saved

    # ---- 전체 실행 -------------------------------------------------------
    def run(self, files: list[Path], dest: Path, subject: str, reason: str):
        total = len(files)
        done = 0
        for n, batch in enumerate(chunks(files, BATCH_SIZE), 1):
            log.info("=== 배치 %d (%d개) ===", n, len(batch))
            self.request_batch(batch, subject, reason)
            done += self.download_latest(dest, expect=len(batch))
        log.info("끝. 요청 %d개 / 저장 %d개", total, done)
        return done


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

    root = tk.Tk()
    root.title("MarkAny 복호화 자동화")
    root.geometry("720x560")

    paths: list[str] = []
    msgs: queue.Queue[str] = queue.Queue()

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="복호화할 파일 / 폴더").pack(anchor="w")
    lb = tk.Listbox(frm, height=8)
    lb.pack(fill="both", expand=True)

    row = ttk.Frame(frm)
    row.pack(fill="x", pady=4)

    def add_files():
        for f in filedialog.askopenfilenames():
            paths.append(f)
            lb.insert("end", f)

    def add_folder():
        d = filedialog.askdirectory()
        if d:
            paths.append(d)
            lb.insert("end", d + "  (폴더)")

    def clear():
        paths.clear()
        lb.delete(0, "end")

    ttk.Button(row, text="파일 추가", command=add_files).pack(side="left")
    ttk.Button(row, text="폴더 추가", command=add_folder).pack(side="left", padx=4)
    ttk.Button(row, text="비우기", command=clear).pack(side="left")

    opts = ttk.Frame(frm)
    opts.pack(fill="x", pady=6)
    ttk.Label(opts, text="저장 폴더").grid(row=0, column=0, sticky="w")
    dest_var = tk.StringVar(value=str(Path.home() / "Downloads"))
    ttk.Entry(opts, textvariable=dest_var, width=60).grid(row=0, column=1, padx=4)
    ttk.Button(opts, text="찾아보기",
               command=lambda: dest_var.set(filedialog.askdirectory() or dest_var.get())
               ).grid(row=0, column=2)

    ttk.Label(opts, text="제목 / 사유").grid(row=1, column=0, sticky="w", pady=4)
    subject_var = tk.StringVar(value="복호화A")
    ttk.Entry(opts, textvariable=subject_var, width=60).grid(row=1, column=1, padx=4, sticky="w")

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

    ttk.Button(btns, text="창 구조 저장", command=save_dump).pack(side="left")
    start_btn = ttk.Button(btns, text="시작")
    start_btn.pack(side="right")

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
            MarkAny().run(files, dest, subject, subject)
            msgs.put("완료되었습니다.")
        except Exception as e:  # noqa: BLE001
            log.error("중단: %s", e)
        finally:
            root.after(0, lambda: start_btn.configure(state="normal"))

    def start():
        files = collect_files(paths)
        dest = Path(dest_var.get())
        if not files:
            messagebox.showwarning("확인", "파일이나 폴더를 먼저 추가하세요.")
            return
        if not dest_var.get().strip():
            messagebox.showwarning("확인", "저장 폴더를 지정하세요.")
            return
        dest.mkdir(parents=True, exist_ok=True)
        setup_logging(dest, GuiHandler())
        log.info("대상 %d개 파일, %d배치", len(files), -(-len(files) // BATCH_SIZE))
        start_btn.configure(state="disabled")
        threading.Thread(target=worker, args=(files, dest, subject_var.get()),
                         daemon=True).start()

    start_btn.configure(command=start)
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
