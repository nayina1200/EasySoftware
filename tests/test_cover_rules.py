from __future__ import annotations

import base64
import json
import shutil
import subprocess
import unittest
from pathlib import Path

import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import cover_rules


class SharedCoverRuleTests(unittest.TestCase):
    def test_python_classifier_matches_versioned_golden_cases(self):
        cases = cover_rules.load_rules()["golden_cases"]
        for case in cases:
            with self.subTest(text=case["text"]):
                actual = cover_rules.classify_cover_text(case["text"], case.get("title"))
                self.assertEqual(
                    {key: actual[key] for key in ("normalized", "kind", "reason")},
                    {key: case[key] for key in ("normalized", "kind", "reason")},
                )

    @unittest.skipUnless(shutil.which("powershell.exe") or shutil.which("pwsh"), "PowerShell is not installed")
    def test_powershell_classifier_matches_same_golden_cases(self):
        shell = shutil.which("powershell.exe") or shutil.which("pwsh")
        rules_script = (ROOT / "scripts" / "cover_rules.ps1").as_posix().replace("'", "''")
        rules_json = (ROOT / "scripts" / "cover_rules.json").as_posix().replace("'", "''")
        command = f"""
. '{rules_script}'
$cases = Get-Content -LiteralPath '{rules_json}' -Raw -Encoding UTF8 | ConvertFrom-Json
$result = @($cases.golden_cases | ForEach-Object {{
    Get-CoverClassification -Text ([string]$_.text) -Title ([string]$_.title)
}})
$json = ConvertTo-Json -InputObject $result -Depth 4 -Compress
[Console]::WriteLine('COVERJSON:' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json)))
"""
        encoded = base64.b64encode(command.encode("utf-16le")).decode("ascii")
        completed = subprocess.run(
            [shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        output = next((line.split(":", 1)[1] for line in reversed(completed.stdout.splitlines()) if line.startswith("COVERJSON:")), "")
        self.assertTrue(output, completed.stdout)
        actual = json.loads(base64.b64decode(output).decode("utf-8"))
        expected = [
            {key: case[key] for key in ("normalized", "kind", "reason")}
            for case in cover_rules.load_rules()["golden_cases"]
        ]
        self.assertEqual(actual, expected)

    def test_sentence_and_directory_text_are_not_cover_labels(self):
        self.assertFalse(cover_rules.is_manual_label("这是用户手册的介绍。"))
        self.assertFalse(cover_rules.is_manual_label("目录"))
        self.assertFalse(cover_rules.is_manual_label("项目文档"))


if __name__ == "__main__":
    unittest.main()
