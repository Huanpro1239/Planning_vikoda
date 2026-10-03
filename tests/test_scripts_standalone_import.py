import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class ScriptsStandaloneImportTests(unittest.TestCase):
    """CI chạy `python scripts/<x>.py` nên sys.path[0] là scripts/, không phải repo root."""

    def test_every_script_imports_without_repo_root_on_sys_path(self):
        scripts = sorted((ROOT / "scripts").rglob("*.py"))
        self.assertTrue(scripts)
        with tempfile.TemporaryDirectory() as cwd:
            for script in scripts:
                with self.subTest(script=script.relative_to(ROOT).as_posix()):
                    result = subprocess.run(
                        [
                            sys.executable,
                            "-c",
                            "import runpy, sys; "
                            "runpy.run_path(sys.argv[1], run_name='not_main')",
                            str(script),
                        ],
                        cwd=cwd,
                        capture_output=True,
                        text=True,
                        timeout=120,
                    )
                    self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == "__main__":
    unittest.main()
