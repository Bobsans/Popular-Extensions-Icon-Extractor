import sys
import os
import platform
import shutil
import tempfile
import zipfile

try:
    from urllib.request import urlretrieve
except ImportError:
    from urllib import urlretrieve

# TODO:
#   Add output directory option
#   better error and exception handling


def choose_icon(extension, icon_number, resource_dir):
    """Choose an icon using its resource number."""
    if os.listdir(resource_dir):
        if icon_number >= 0:
            icon_number += 1
        elif icon_number < 0:
            icon_number = -icon_number

        aux = None
        found = False
        i = 0
        for f in os.listdir(resource_dir):
            file_name = os.path.join(resource_dir, f)
            if '_%s.' % icon_number in f:
                found = True
                break
            elif i == icon_number - 1:
                aux = file_name
            i += 1

        if not found and aux:
            file_name = aux

        shutil.copy(file_name, os.path.join('icons', extension.lower() + '.ico'))


def icon_extract(resource, resource_dir):
    """Extract icons from resource"""
    path = os.path.dirname(os.path.abspath(__file__))
    return os.system('%s /save "%s" "%s" -icons' % (os.path.join(path, 'iconsext.exe'), resource, resource_dir))


def get_temp_directory():
    """Get temporary directory name"""
    return os.path.join(tempfile.gettempdir(), 'iconsextractor')


def create_temp_directory():
    """Create a temporary directory and return its path. Delete any existing directories"""
    tmp_dir = get_temp_directory()
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)
    return tmp_dir


def download_resource(url, file_to_extract):
    path = os.path.dirname(os.path.abspath(__file__))
    urlretrieve(url, os.path.join(path, os.path.basename(url)))
    with zipfile.ZipFile(os.path.join(path, os.path.basename(url))) as zip_file:
        for name in zip_file.namelist():
            if name.lower() == file_to_extract:
                zip_file.extract(name, path)
                break
    os.remove(os.path.join(path, os.path.basename(url)))


def main():
    if platform.system().lower() != 'windows':
        sys.stderr.write('Windows is required.')
        exit(1)

    if os.path.exists('icons') and os.listdir('icons'):
        sys.stderr.write("'icons' directory exists and is not empty. Please remove it first.\n")
        exit(1)

    # Create icons directory
    if not os.path.exists('icons'):
        os.mkdir('icons')

    script_path = os.path.dirname(os.path.abspath(__file__))

    if not os.path.exists(os.path.join(script_path, 'FileTypesMan.exe')):
        print('FileTypesMan.exe not found. Trying to download it.')
        download_resource('http://www.nirsoft.net/utils/filetypesman.zip', 'filetypesman.exe')
    if not os.path.exists(os.path.join(script_path, 'iconsext.exe')):
        print('iconsext.exe not found. Trying to download it.')
        download_resource('http://www.nirsoft.net/utils/iconsext.zip', 'iconsext.exe')

    lookup_paths = os.environ.get('PATH').split(';')
    tmp_dir = create_temp_directory()
    list_file = os.path.join(tmp_dir, 'list.txt')
    code = os.system('%s /stab %s' % (os.path.join(script_path, 'filetypesman.exe'), list_file))
    if code != 0:
        sys.exit(code)

    with open(list_file, encoding='utf-16') as lf:
        ext_list = [row.split('\t') for row in lf]

    for item in ext_list:
        if item[0] and item[0][0] == '.':
            extension = item[0][1:]
            icon_path = item[10].replace('\"', '').replace('&quot;', '') # Remove quotes

            # Example cases (with qoutes removed):
            # 1)
            # 2) %1
            # 3) C:\Python27\DLLs\py.ico
            # 4) C:\Windows\Installer\{AC76BA86-7AD7-FFFF-7B44-AB0000000001}\PDFFile_8.ico,0
            # 5) C:\Windows\System32\zipfldr.dll
            # 6) C:\Windows\System32\icardres.dll,-4112
            # 7) C:\PROGRA~1\COMMON~1\MICROS~1\OFFICE14\MSORES.DLL,-560
            # 8) imageres.dll,-67
            # 9) cryptui.dll,-3425
            # 10) %SystemRoot%\System32\shell32.dll,2
            # 11) %SystemRoot%\System32\GfxUIEx.exe, 2
            # 12) %SystemRoot%\system32\wmploc.dll,-730
            # 13) C:\Program Files (x86)\qBittorrent\qbittorrent.exe,1
            # 14) &quot;%ProgramFiles%\Windows Journal\Journal.exe&quot;,2

            if icon_path and len(icon_path) > 2:
                icon_path2 = ''
                resource_number = 0

                # Handle cases with an icon number
                if ',' in icon_path:
                    icon_path, resource_number = icon_path.split(',')
                    if icon_path == '':
                        continue
                    resource_number = int(resource_number)

                # Evaluate environment variables and workaround specific case for %ProgramFiles%
                if '%ProgramFiles%' in icon_path:
                    icon_path2 = os.path.expandvars(icon_path.replace('%ProgramFiles%', '%ProgramW6432%'))
                icon_path = os.path.expandvars(icon_path)

                # Handle cases 3 and 4
                if icon_path.lower().endswith('.ico'):
                    shutil.copy(icon_path, os.path.join('icons', extension.lower() + '.ico'))
                    continue

                # Try to extract the icon
                resource_dir = os.path.join(tmp_dir, os.path.basename(icon_path))
                if not os.path.exists(resource_dir):
                    os.mkdir(resource_dir)
                    icon_extract(icon_path, resource_dir)
                    if icon_path2:
                        icon_extract(icon_path2, resource_dir)

                # If the directory is still empty and the path of the resource was relative
                if not os.listdir(resource_dir) and not os.path.isabs(icon_path):
                    # Try again with absolute paths
                    for path_prefix in lookup_paths:
                        icon_extract(os.path.join(path_prefix, icon_path), resource_dir)
                        if os.listdir(resource_dir):
                            continue

                choose_icon(extension, resource_number, resource_dir)

    # Delete temporary directory
    shutil.rmtree(get_temp_directory())

if __name__ == '__main__':
    main()
