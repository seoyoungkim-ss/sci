"""Shared xlwings connection helper for xlwings_loader.py and
segment_loader.py.

Background: xlwings_loader.py's ONE master survey file opens fine via
plain COM automation (no manual pre-open needed). The per-department
"문항별 결과" detail files (up to 44 of them) instead hit Excel's
generic "읽기 전용이거나 손상되었거나 암호화되어 있습니다" error under
the exact same kind of automated open — but open fine by double-
clicking. Manually pre-opening each file works around it (attaching to
an already-open book needs no new Open call at all) but doesn't scale
to 44 files kept open at once, and the goal is a fully automated run.

Since the two file sets behave differently under the identical COM
call, the most likely cause isn't DRM refusing automation outright
(that would break the master file too) but a per-file interactive
prompt that silently blocks/fails under automation — most commonly
Excel's own "읽기 전용으로 여시겠습니까?" (open read-only?) dialog,
which a real user clicking through by double-click never notices, but
a script has no way to answer. open_book_directly() below passes
ignore_read_only_recommended=True (and notify=False, update_links=False)
specifically to suppress that prompt at the COM layer, per Excel's own
Workbooks.Open parameters — xlwings exposes these directly as kwargs
on books.open(). It also retries a couple of times with a short delay,
in case a DRM hook simply needs a moment after Excel launches before
Open succeeds.

open_or_attach() tries that direct automated open FIRST; only if it
still fails does it fall back to looking for an already-open copy (in
case the file was opened manually) rather than giving up immediately —
this keeps a fully automated run working when the direct open
succeeds, while still letting the manual-open workaround apply on
whichever files it doesn't.

If the failure is specifically HRESULT 0x800AC472 ("Excel cannot
access the file because a dialog box is open" — pywin32 reports it as
pywintypes.com_error with -2146777998 as its first element, same
number), ignore_read_only_recommended alone didn't cover it: that
flag only suppresses Excel's OWN built-in read-only prompt, and this
error means SOME dialog is genuinely up and blocking every COM call —
most likely a DRM/security add-in's own popup, which no Workbooks.Open
parameter can suppress since Excel doesn't know it exists.
_dismiss_blocking_dialogs() finds it and answers it programmatically
(Enter, i.e. its default button) rather than just sleeping and
retrying the identical call into the same still-open dialog. This is a
heuristic: it can't know what the dialog actually says, only that
whatever button is already focused is what a user tabbing/enter-ing
through it would trigger — closest to what happens when a person
double-clicks the file and clicks through without reading it, but the
title being dismissed is always printed so this is auditable if it
guesses wrong on some other, unrelated dialog.
"""
import time
from pathlib import Path


def _find_open_book(path):
    import xlwings as xw

    target_path = Path(path).resolve()
    for app in xw.apps:
        for book in app.books:
            try:
                book_path = Path(book.fullname).resolve()
            except Exception:
                continue
            if book_path == target_path or book.name == target_path.name:
                return app, book
    return None, None


BLOCKING_DIALOG_HRESULT = -2146777998  # 0x800AC472: "a dialog box is open"


def _is_blocking_dialog_error(err):
    return getattr(err, "args", None) and err.args[0] == BLOCKING_DIALOG_HRESULT


def _dismiss_blocking_dialogs(pid, timeout=5.0):
    """Best-effort: finds any visible top-level window belonging to
    Excel's process (pid) that is a standard Windows dialog box (class
    "#32770" — the main Excel window itself is a different class) and
    sends Enter to accept its default button. Returns True if it found
    and dismissed something. Needs pywin32 (already an xlwings
    dependency on Windows, so this should always be available there);
    silently no-ops if it isn't, since this whole thing is only ever a
    secondary attempt after the plain retries above."""
    try:
        import win32con
        import win32gui
        import win32process
    except ImportError:
        return False

    dismissed = False

    def enum_handler(hwnd, _):
        nonlocal dismissed
        if dismissed or not win32gui.IsWindowVisible(hwnd):
            return
        try:
            _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            return
        if win_pid != pid:
            return
        if win32gui.GetClassName(hwnd) != "#32770":  # standard dialog box class
            return
        title = win32gui.GetWindowText(hwnd)
        print(f"  (Excel 대화상자 감지: '{title}' — 기본 버튼으로 응답 시도)")
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        dismissed = True

    deadline = time.time() + timeout
    while time.time() < deadline and not dismissed:
        win32gui.EnumWindows(enum_handler, None)
        if not dismissed:
            time.sleep(0.3)
    return dismissed


def open_or_attach(path, retries=3, retry_delay=2.0):
    """Returns (app, book, owns_app). owns_app is True only when this
    call opened a brand-new Excel instance itself (safe to app.quit()
    when done) — an ATTACHED book belongs to whatever Excel window the
    user already had open, so callers must never close() or quit() it,
    only read from it.

    Order per attempt: (1) try a direct automated open with the
    read-only prompt suppressed, (2) on the specific "dialog box is
    open" error, try to find and dismiss it before the next retry,
    (3) after all retries, look for an already-open copy of this exact
    file instead of giving up, (4) only then raise, with full context."""
    import xlwings as xw

    app = xw.App(visible=True)
    last_err = None
    for attempt in range(1, retries + 2):
        try:
            book = app.books.open(
                str(path),
                update_links=False,
                ignore_read_only_recommended=True,
                notify=False,
            )
            return app, book, True
        except Exception as err:
            last_err = err
            if attempt <= retries:
                if _is_blocking_dialog_error(err):
                    _dismiss_blocking_dialogs(app.pid)
                else:
                    time.sleep(retry_delay)

    # Direct open never succeeded — before giving up, check whether this
    # exact file happens to already be open somewhere (manually opened
    # as a workaround) and attach to that instead of failing outright.
    found_app, found_book = _find_open_book(path)
    if found_book is not None:
        app.quit()  # the fresh instance we opened above never got a book into it
        print(f"  (자동 열기는 실패했지만 이미 열려 있는 Excel 창에 연결: '{found_book.name}')")
        return found_app, found_book, False

    app.quit()
    raise RuntimeError(
        f"Excel에서 '{path}' 파일을 열지 못했습니다 ({retries + 1}회 시도): {last_err}\n"
        "이미 더블클릭으로 파일을 열어둔 상태로 스크립트를 다시 실행해보세요 — 그러면 "
        "새로 열지 않고 이미 열려 있는 창에 연결을 시도합니다. 그래도 안 되면 DRM "
        "플러그인이 이 파일에 대해 COM 자동화 자체를 막고 있는 것일 수 있습니다."
    ) from last_err


def close_or_detach(app, book, owns_app):
    """Mirrors open_or_attach: only closes/quits what this script
    itself opened. An attached book/app is left exactly as the user
    had it."""
    if not owns_app:
        return
    try:
        book.close()
    finally:
        app.quit()
