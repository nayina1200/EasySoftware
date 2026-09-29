from __future__ import annotations

import os
import shutil
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
POWERSHELL = shutil.which("powershell.exe") or shutil.which("powershell") or shutil.which("pwsh")


@unittest.skipUnless(os.name == "nt" and POWERSHELL, "Windows PowerShell is required for CDP policy checks")
class OrderDownloadGuiPolicyTests(unittest.TestCase):
    def test_mocked_cdp_policy_fails_closed(self):
        completed = subprocess.run(
            [
                POWERSHELL,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(ROOT / "scripts" / "order_material_downloader_gui.ps1"),
                "-SelfTest",
            ],
            capture_output=True,
            text=False,
            timeout=30,
            check=False,
        )
        output = (completed.stdout + completed.stderr).decode("utf-8", errors="replace")
        self.assertEqual(completed.returncode, 0, output)
        self.assertIn("Order download policy self-test passed.", output)


if __name__ == "__main__":
    unittest.main()
