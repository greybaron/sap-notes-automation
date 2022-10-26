import sys
import os
from pathlib import Path
import shutil
import requests

def check_browser_install():
    if sys.platform == "win32":
        chromepath = Path.home().joinpath("sap-automation-chromium/chrome.exe")
    elif sys.platform == "darwin":
        chromepath = Path.home().joinpath("sap-automation-chromium/Chromium.app/Contents/MacOS/Chromium")
    else:
        raise NotImplementedError(f"not running windows or macos, platform {sys.platform} not supported")

    if not chromepath.exists():
        download_chromium(('win' if sys.platform == 'win32' else 'mac'))
    
    return chromepath


def download_chromium(platform):
    print(f"downloading Chromium for {'Windows' if platform == 'win' else 'macOS'}")

    path = Path.home().joinpath("sap-automation-chromium/")

    try:
        shutil.rmtree(path)
    except FileNotFoundError:
        pass

    os.makedirs(path)


    zip_path = path.joinpath("chromium.zip")

    zip_response = requests.get(f"https://github.com/greybaron/sap-notes-automation/raw/main/chromium/chromium-{platform}.zip")

    # this is dumb as the zip will be written to ram first and not streamed in chunks to storage
    zip_path.write_bytes(zip_response.content)


    print("Download finished")

    
    if platform == "win":
        shutil.unpack_archive(zip_path, path)

    # shutil.unpack screws up permissions and Gatekeeper probably fucks shit up as well,
    # so using system unzip here
    if platform == "mac":
        os.system(f'unzip "{zip_path}" -d "{path}"')
    
    os.remove(zip_path)
    

def getListOfFiles(path):
    # create a list of all files, including those in subdirectories
    # recursively calls itself when encountering a folder, and keeps looping until it hits no more subfolders in any given directory
    nodesAtLevel = os.listdir(path)

    files = []
    # Iterate over all the entries
    for subnode in nodesAtLevel:
        newPath = os.path.join(path, subnode)
        # If entry is a directory then get the list of files in this directory 
        if os.path.isdir(newPath):
            files += getListOfFiles(newPath)
        else:
            files.append(Path(newPath))
    return files   
