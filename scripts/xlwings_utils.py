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


def open_or_attach(path, retries=2, retry_delay=2.0):
    """Returns (app, book, owns_app). owns_app is True only when this
    call opened a brand-new Excel instance itself (safe to app.quit()
    when done) — an ATTACHED book belongs to whatever Excel window the
    user already had open, so callers must never close() or quit() it,
    only read from it.

    Order: (1) try a direct automated open with the read-only prompt
    suppressed, retrying a couple of times, (2) if that still fails,
    look for an already-open copy of this exact file instead of giving
    up, (3) only then raise, with both attempts' context."""
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
