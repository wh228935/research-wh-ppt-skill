"""Read-only Windows dependency discovery. Never installs anything."""
import argparse
import importlib.util
import json
import platform
import shutil
import subprocess
import sys
from pathlib import Path


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--powershell', help='Explicit PowerShell 7 executable selected by host')
    a = p.parse_args()
    ps = a.powershell or shutil.which('pwsh')
    result = {'platform': platform.platform(), 'python': sys.executable,
              'python_version': platform.python_version(), 'core_standard_library': True,
              'office_registration_is_not_actual_validation': True,
              'optional_pdf_modules': {x: bool(importlib.util.find_spec(x)) for x in ['pymupdf', 'fitz', 'pypdf']},
              'powershell': ps, 'office': {}, 'font_microsoft_yahei': None, 'missing': []}
    if sys.version_info < (3, 10):
        result['missing'].append({'tool': 'Python 3.10+', 'purpose': 'core helpers', 'action': 'ask user before installation'})
    if ps:
        try:
            r = subprocess.run([ps, '-NoProfile', '-Command', '$PSVersionTable.PSVersion.ToString()'],
                               capture_output=True, text=True, timeout=15)
            result['powershell_version'] = r.stdout.strip()
            if r.returncode or int(r.stdout.strip().split('.')[0]) < 7:
                result['missing'].append({'tool': 'PowerShell 7', 'purpose': 'COM automation', 'action': 'ask user'})
        except (OSError, ValueError, subprocess.TimeoutExpired) as e:
            result['powershell_error'] = str(e)
    else:
        result['missing'].append({'tool': 'PowerShell 7', 'purpose': 'COM automation', 'action': 'locate approved executable; do not use 5.1'})
    if sys.platform == 'win32':
        import winreg
        for label, prog in [('PowerPoint','PowerPoint.Application'), ('WPS','KWPP.Application'), ('Word','Word.Application')]:
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog + '\\CLSID') as key:
                    result['office'][label] = {'registered': True, 'clsid': winreg.QueryValue(key, None)}
            except OSError:
                result['office'][label] = {'registered': False}
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts') as k:
                fonts = [winreg.EnumValue(k, i)[0] for i in range(winreg.QueryInfoKey(k)[1])]
            result['font_microsoft_yahei'] = any('yahei' in x.lower() or '微软雅黑' in x for x in fonts)
        except OSError:
            pass
    else:
        result['missing'].append({'tool': 'Windows', 'purpose': 'this package targets Windows only', 'action': 'do not claim unsupported platforms'})
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
