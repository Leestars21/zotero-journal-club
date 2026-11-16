import os
import subprocess
import tempfile
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts' / 'render_powerpoint.ps1'
POWERSHELL = Path(r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe')


@unittest.skipUnless(os.name == 'nt' and POWERSHELL.is_file(), 'Windows PowerShell guard test')
class OfficeGuardTests(unittest.TestCase):
    def test_existing_output_is_preserved_before_office_is_started(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source, output = root / 'source.pptx', root / 'existing.pdf'
            source.write_bytes(b'not opened: preflight must run first')
            output.write_bytes(b'existing user output')
            result = subprocess.run([str(POWERSHELL), '-NoProfile', '-NonInteractive', '-File', str(SCRIPT),
                                     str(source), str(output)], capture_output=True, timeout=15)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn(b'Refusing to overwrite', result.stderr)
            self.assertEqual(output.read_bytes(), b'existing user output')
            self.assertEqual(source.read_bytes(), b'not opened: preflight must run first')


if __name__ == '__main__':
    unittest.main()
