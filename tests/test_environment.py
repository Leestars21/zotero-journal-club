import importlib.util
import tempfile
import unittest
from pathlib import Path

path = Path(__file__).resolve().parents[1] / 'zotero-journal-club' / 'scripts' / 'inspect_environment.py'
spec = importlib.util.spec_from_file_location('inspect_environment', path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class EnvironmentTests(unittest.TestCase):
    def test_explicit_directory_is_a_candidate_and_database_is_not_opened(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / 'zotero.sqlite').write_bytes(b'not a database')
            (root / 'storage').mkdir()
            result = module.inspect(data_dir=root)
            self.assertEqual(result['zotero_data_candidates'][0]['origin'], 'explicit')
            self.assertTrue(result['zotero_data_candidates'][0]['database_present'])
            self.assertTrue(result['zotero_data_candidates'][0]['storage_present'])

    def test_profile_custom_directory(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            profile = root / 'profile'
            profile.mkdir()
            (root / 'profiles.ini').write_text('[Profile0]\nIsRelative=1\nPath=profile\n', encoding='utf-8')
            (profile / 'prefs.js').write_text('user_pref("extensions.zotero.dataDir", "D:/Other/Zotero");', encoding='utf-8')
            self.assertEqual(module.profile_data_dirs(root), [Path('D:/Other/Zotero')])

    def test_version_from_explicit_app_directory(self):
        with tempfile.TemporaryDirectory() as folder:
            app = Path(folder) / 'app'
            app.mkdir()
            (app / 'application.ini').write_text('[App]\nName=Zotero\nVersion=9.1\n', encoding='utf-8')
            self.assertEqual(module.inspect(data_dir=folder, app_dir=folder)['zotero_version'], '9.1')

if __name__ == '__main__':
    unittest.main()
