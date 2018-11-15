import sys
import os
import os.path
import stat
import tempfile
import shutil
import subprocess

_isPy3k = sys.version_info[0] >= 3

if _isPy3k:
    import winreg
else:
    import _winreg as winreg

_is64bitOS = os.environ.get('PROCESSOR_ARCHITECTURE') == 'AMD64' or os.environ.get(
    'PROCESSOR_ARCHITEW6432') == 'AMD64'


def EnumPythonInstallations():

    result = []

    def ProcessKey(key, arch):
        index = 0
        while True:
            subkeyname = winreg.EnumKey(key, index)
            with winreg.OpenKey(key, os.path.join(subkeyname, 'InstallPath'), 0, winreg.KEY_READ) as subkey:
                path = winreg.QueryValueEx(subkey, '')[0]
                result.append((arch, subkeyname, path))
            index += 1

    def ProcessHKLM(view):
        if view == winreg.KEY_WOW64_32KEY:
            arch = '32 bit'
        elif view == winreg.KEY_WOW64_64KEY:
            arch = '64 bit'
        else:
            raise AssertionError
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, os.path.join('Software', 'Python', 'PythonCore'), 0, winreg.KEY_READ | view) as key:
                ProcessKey(key, arch)
        except WindowsError:
            pass

    def ProcessHKCU():
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, os.path.join('Software', 'Python', 'PythonCore'), 0, winreg.KEY_READ) as key:
                ProcessKey(key, '64 bit')
        except WindowsError:
            pass

    ProcessHKLM(winreg.KEY_WOW64_32KEY)
    if _is64bitOS:
        ProcessHKLM(winreg.KEY_WOW64_64KEY)
    ProcessHKCU()
    return result

cwd = os.getcwd()
os.chdir(cwd + '\\DATA')
tmp = tempfile.mkdtemp()
try:
    for fileName in 'OrcFxAPI.py', 'OrcFxAPIConfig.py', 'setup.py':
        shutil.copy2(fileName, tmp)
        # make writeable otherwise later upgrades will fail
        os.chmod(os.path.join(tmp, fileName), stat.S_IWRITE)
    os.chdir(tmp)
    for arch, subkeyname, path in EnumPythonInstallations():
        if subkeyname[-3:] == '-32':
            arch = '32 bit'
            ver = subkeyname[:-3]
        else:
            ver = subkeyname
        major, minor = map(int, ver.split('.'))
        if major < 2:
            continue
        if major == 2 and minor < 6:
            continue
        print('Installing to {0} Python version {1} located at {2}'.format(arch, ver, path))
        subprocess.check_call([os.path.join(path, 'python.exe'), 'setup.py', 'install', '--force'])
        print('Done.\n')
finally:
    os.chdir(cwd)
    shutil.rmtree(tmp)
