"""Guardrails for simplified receiving actions in the Ftrace editor."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_receiving_explains_auto_save_and_keeps_recovery_under_advanced():
    template = (ROOT / "templates/kitchen/ftrace.html").read_text()
    assert "下載時會自動儲存，不必再按一次儲存" in template
    assert "只儲存，暫不下載" in template
    assert "自動儲存並下載收貨 XLS" in template
    assert "進階操作：找回備份／重新帶入資料" in template
    assert 'id="save-draft"' in template
    assert 'id="restore-browser-draft"' in template
    assert 'id="migrate-drafts"' in template
    assert 'id="reset-draft"' in template
    assert "已匯出」；這不代表已經上傳" in template


def test_download_saves_shared_record_before_starting_file_download():
    script = (ROOT / "static/ftrace.js").read_text()
    assert "const file=await response.blob();await saveShared('exported');" in script
    assert script.index("await saveShared('exported')") < script.index("download(file,")
    assert "if (settings.kind === 'receiving' && !window.confirm(" in script
    assert "目前畫面上尚未儲存的修改會被取代" in script
    assert "本頁未儲存的修改與本日瀏覽器備份將被捨棄" in script
