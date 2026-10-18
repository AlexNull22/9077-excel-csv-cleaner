"""Verify the desktop worker, status updates, and report action without a network."""
import importlib.util
from pathlib import Path
import sys
import tempfile
import time

root = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("gui9077", root / "9077_桌面工具.py")
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)

with tempfile.TemporaryDirectory(prefix="9077_gui_") as folder:
    app = module.CleanerApp()
    app.withdraw()
    app.output_path.set(str(Path(folder) / "result"))
    app.start_run()
    deadline = time.monotonic() + 15
    try:
        while app.metric_values["processed"].get() == "--" and time.monotonic() < deadline:
            app.update()
            time.sleep(0.05)
        counts = {key: value.get() for key, value in app.metric_values.items()}
        assert counts == {"processed": "24", "valid": "13", "invalid": "5", "duplicate": "3", "filtered": "3"}, app.status_text.get()
        assert len(app.table.get_children()) == 24
        assert app.report.is_file()
        assert app.report_button.instate(["!disabled"])
        assert app.run_button.instate(["!disabled"])
        print("Desktop integration: PASS. 24 audit rows; 13 kept, 5 invalid, 3 duplicates, 3 filtered.")
    finally:
        app.destroy()
