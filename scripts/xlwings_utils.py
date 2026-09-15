"""Shared xlwings connection helper for xlwings_loader.py and
segment_loader.py.

If double-clicking the file in Explorer opens it fine but a script
driving Excel via COM automation gets Excel's generic "읽기
전용이거나 손상되었거나 암호화되어 있습니다" error, the DRM/IRM plugin
is very likely hooking Excel's own interactive "open via double-click"
path specifically, and either isn't invoked at all — or fails — when a
script opens the same file via Workbooks.Open over COM, even with the
Excel window visible.

The workaround: open the file normally yourself first (double-click,
already confirmed to work), leave that Excel window open, and let the
script ATTACH to that already-open, already-decrypted workbook instead
of asking Excel to open a fresh copy itself. xlwings can already see
every running Excel instance and its open books, so this only needs a
name/path match — no new Open call, so no DRM re-decryption involved.
"""
from pathlib import Path


def open_or_attach(path):
    """Returns (app, book, owns_app). owns_app is True only when this
    call opened a brand-new Excel instance itself (safe to app.quit()
    when done) — an ATTACHED book belongs to whatever Excel window the
    user already had open, so callers must never close() or quit() it,
    only read from it."""
    import xlwings as xw

    target_path = Path(path).resolve()
    for app in xw.apps:
        for book in app.books:
            try:
                book_path = Path(book.fullname).resolve()
            except Exception:
                continue
            if book_path == target_path or book.name == target_path.name:
                print(f"  (이미 열려 있는 Excel 창에 연결: '{book.name}' — 새로 열지 않음)")
                return app, book, False

    # Not already open anywhere — open a new instance ourselves.
    # visible=True: some DRM plugins also fail on a fresh COM-driven
    # Open even with the window visible, but this is still worth trying
    # before giving up — visible at least matches a normal open as
    # closely as automation allows.
    app = xw.App(visible=True)
    try:
        book = app.books.open(str(path))
    except Exception as err:
        app.quit()
        raise RuntimeError(
            f"Excel에서 '{path}' 파일을 열지 못했습니다: {err}\n"
            "이미 더블클릭으로 파일을 열어둔 상태로 스크립트를 다시 실행해보세요 — "
            "그러면 새로 열지 않고 이미 열려 있는 창에 연결을 시도합니다. "
            "그래도 안 되면 DRM 플러그인이 COM 자동화 자체를 막고 있는 것일 수 있습니다."
        ) from err
    return app, book, True


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
