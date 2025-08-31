import os
import logging
import shutil
import json
import sys
import subprocess
import re
import tempfile
import ctypes
from ctypes import windll, wintypes, byref, POINTER
import comtypes
import comtypes.client
import urllib.request

def inet_read(url):
    """Python equivalent of AutoIt's InetRead() function using WinINet API."""
    try:
        logging.info(f"Using InetRead equivalent to download: {url}")

        # Load wininet.dll
        wininet = windll.wininet

        # Define function prototypes
        InternetOpenW = wininet.InternetOpenW
        InternetOpenUrlW = wininet.InternetOpenUrlW
        InternetReadFile = wininet.InternetReadFile
        InternetCloseHandle = wininet.InternetCloseHandle

        # Set up function argument types
        InternetOpenW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD]
        InternetOpenW.restype = wintypes.HANDLE

        InternetOpenUrlW.argtypes = [wintypes.HANDLE, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
                                     wintypes.DWORD, wintypes.DWORD]
        InternetOpenUrlW.restype = wintypes.HANDLE

        InternetReadFile.argtypes = [wintypes.HANDLE, wintypes.LPVOID, wintypes.DWORD, POINTER(wintypes.DWORD)]
        InternetReadFile.restype = wintypes.BOOL

        InternetCloseHandle.argtypes = [wintypes.HANDLE]
        InternetCloseHandle.restype = wintypes.BOOL

        # Constants
        INTERNET_OPEN_TYPE_PRECONFIG = 0
        INTERNET_FLAG_RELOAD = 0x80000000
        INTERNET_FLAG_NO_CACHE_WRITE = 0x04000000

        # Open internet connection
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        h_internet = InternetOpenW(user_agent, INTERNET_OPEN_TYPE_PRECONFIG, None, None, 0)

        if not h_internet:
            logging.error("Failed to initialize WinINet")
            return None

        try:
            # Open URL
            h_url = InternetOpenUrlW(
                h_internet,
                url,
                None,
                0,
                INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE,
                0
            )

            if not h_url:
                logging.error(f"Failed to open URL: {url}")
                return None

            try:
                # Read data in chunks
                data = b""
                buffer_size = 8192
                buffer = (ctypes.c_ubyte * buffer_size)()
                bytes_read = wintypes.DWORD(0)

                while True:
                    if not InternetReadFile(h_url, buffer, buffer_size, byref(bytes_read)):
                        break

                    if bytes_read.value == 0:
                        break

                    data += bytes(buffer[:bytes_read.value])

                logging.info(f"Successfully downloaded {len(data)} bytes using InetRead equivalent")
                return data

            finally:
                InternetCloseHandle(h_url)
        finally:
            InternetCloseHandle(h_internet)

    except Exception as e:
        logging.error(f"Error in InetRead equivalent: {e}")
        return None

def download_file_inet_read(url, save_path):
    """Download a file using the InetRead equivalent and save it."""
    try:
        data = inet_read(url)
        if data is None:
            return False

        with open(save_path, 'wb') as f:
            f.write(data)

        logging.info(f"File saved to: {save_path}")
        return True

    except Exception as e:
        logging.error(f"Error saving downloaded file: {e}")
        return False

def extract_zip_with_shell(zip_path, extract_path):
    """Extract a zip file using Shell.Application."""
    try:
        logging.info(f"Extracting {zip_path} to {extract_path} using Shell.Application")

        # Make sure extraction directory exists
        if not os.path.exists(extract_path):
            os.makedirs(extract_path)

        # Create Shell Application COM object
        shell = comtypes.client.CreateObject("Shell.Application")

        # Get zip items
        zip_folder = shell.NameSpace(zip_path)
        if not zip_folder:
            logging.error(f"Could not access ZIP file: {zip_path}")
            return False

        zip_items = zip_folder.Items()

        # Extract files
        # Flag 20 = 16 (no progress dialog) + 4 (do not display a user interface if an error occurs)
        dest_folder = shell.NameSpace(extract_path)
        dest_folder.CopyHere(zip_items, 20)

        # Wait for extraction to complete. Check if the expected file exist after extraction
        extracted_driver = os.path.join(extract_path, "msedgedriver.exe")
        if not os.path.exists(extracted_driver):
            # Try to find the driver in the extracted files (in case of subfolders)
            found = False
            for root, files in os.walk(extract_path):
                for file in files:
                    if file.lower() == "msedgedriver.exe":
                        found = True
                        break
                if found:
                    break
            if not found:
                logging.warning("msedgedriver.exe not found after extraction.")
                return False

        logging.info("ZIP extraction completed using Shell.Application")
        return True

    except Exception as e:
        logging.error(f"Error extracting ZIP with Shell.Application: {e}")
        return False

def download_file_with_wininet(url, save_path):
    """Download a file using WinINet API."""
    try:
        logging.info(f"Attempting to download {url} using WinINet API")

        # Try using the built-in Windows downloader first
        try:
            # Define the required constants for URLDownloadToFile
            BINDF_GETNEWESTVERSION = 0x00000010

            # Define the function prototype
            urlmon = ctypes.windll.urlmon
            URLDownloadToFile = urlmon.URLDownloadToFileW
            URLDownloadToFile.argtypes = [
                wintypes.LPVOID, wintypes.LPCWSTR, wintypes.LPCWSTR,
                wintypes.DWORD, wintypes.LPVOID
            ]
            URLDownloadToFile.restype = wintypes.HRESULT

            # Download the file
            result = URLDownloadToFile(
                None, url, save_path, BINDF_GETNEWESTVERSION, None
            )

            if result == 0:  # S_OK
                logging.info(f"Successfully downloaded to {save_path} using WinINet")
                return True
            else:
                logging.warning(f"WinINet download failed with code {result}, trying alternative")
        except Exception as e:
            logging.warning(f"WinINet API error: {e}, trying alternative")

        # Fallback to urllib if WinINet fails
        logging.info(f"Attempting to download with urllib")

        # Set up a custom opener with IE headers
        opener = urllib.request.build_opener()
        opener.addheaders = [
            ('User-Agent',
             'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36 Edg/91.0.864.59'),
            ('Accept', 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'),
            ('Accept-Language', 'en-US,en;q=0.5'),
            ('Connection', 'keep-alive'),
        ]
        urllib.request.install_opener(opener)

        # Create a request with custom timeout and flags
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=30) as response, open(save_path, 'wb') as out_file:
            shutil.copyfileobj(response, out_file)

        logging.info(f"Successfully downloaded to {save_path} using urllib")
        return True

    except Exception as e:
        logging.error(f"Error downloading file with WinINet and urllib: {e}")
        return False

def get_edge_version():
    """Get the installed Edge version using various methods."""
    try:
        logging.info("Detecting Edge browser version...")
        edge_version = ""

        # Method 1: Using wmic command
        try:
            edge_paths = [
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            ]

            for edge_path in edge_paths:
                if os.path.exists(edge_path):
                    # Fix: avoid backslash in f-string expression
                    wmic_path = edge_path.replace('\\', '\\\\').replace('"', '\\"')
                    wmic_query = f'name="{wmic_path}"'
                    result = subprocess.run(
                        ['wmic', 'datafile', 'where', wmic_query, 'get', 'Version', '/value'],
                        capture_output=True, text=True)
                    version_match = re.search(r'Version=(.+)', result.stdout)
                    if version_match:
                        edge_version = version_match.group(1).strip()
                        logging.info(f"Edge version detected (wmic): {edge_version}")
                        return edge_version
        except Exception as e:
            logging.warning(f"Could not get Edge version using wmic: {e}")

        # Method 2: Using reg query
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKEY_CURRENT_USER\\Software\\Microsoft\\Edge\\BLBeacon', '/v', 'version'],
                capture_output=True, text=True
            )
            version_match = re.search(r'version\s+REG_SZ\s+(.+)', result.stdout)
            if version_match:
                edge_version = version_match.group(1).strip()
                logging.info(f"Edge version detected (registry): {edge_version}")
                return edge_version
        except Exception as e:
            logging.warning(f"Could not get Edge version using registry: {e}")

        # Method 3: Try checking version file
        try:
            edge_dir = r"C:\Program Files (x86)\Microsoft\Edge\Application"
            if not os.path.exists(edge_dir):
                edge_dir = r"C:\Program Files\Microsoft\Edge\Application"

            if os.path.exists(edge_dir):
                version_file = os.path.join(edge_dir, "product_versions.json")
                if os.path.exists(version_file):
                    with open(version_file, 'r') as f:
                        data = json.load(f)
                        edge_version = data.get('product', {}).get('version', '')
                        if edge_version:
                            logging.info(f"Edge version detected (version file): {edge_version}")
                            return edge_version
        except Exception as e:
            logging.warning(f"Could not get Edge version from version file: {e}")

        # If all methods fail, use a default version
        if not edge_version:
            logging.warning("Could not detect Edge version, using default latest version")
            try:
                # Try to get the latest version number
                temp_file = os.path.join(tempfile.gettempdir(), "edge_version.txt")
                if download_file_with_wininet("https://msedgedriver.azureedge.net/LATEST_STABLE", temp_file):
                    with open(temp_file, 'r') as f:
                        edge_version = f.read().strip()
                    os.remove(temp_file)
                    logging.info(f"Using latest stable Edge driver version: {edge_version}")
                    return edge_version
            except Exception as e:
                logging.warning(f"Could not get latest version: {e}")

            # Hardcoded fallback
            edge_version = "136.0.3240.64"  # Update this based on known version that works
            logging.info(f"Using fallback Edge version: {edge_version}")

        return edge_version

    except Exception as e:
        logging.error(f"Error detecting Edge version: {e}")
        # Fallback to a default version that's likely to work on most systems
        edge_version = "136.0.3240.64"  # Update this based on the version in the error message
        logging.info(f"Using fallback Edge version due to error: {edge_version}")
        return edge_version

def get_edge_driver_path():
    """Get or download Edge driver."""
    try:
        # Always use the directory of the script, not the PyInstaller temp dir
        base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        # Path where we'll store the Edge driver
        driver_path = os.path.join(base_path, "msedgedriver.exe")

        # Check if the driver already exists in expected locations
        possible_driver_locations = [
            # Check in the current directory
            driver_path,
            # Check in the executable's directory (if different from base_path)
            os.path.join(os.path.dirname(sys.executable), "msedgedriver.exe")
        ]

        for path in possible_driver_locations:
            if os.path.exists(path):
                logging.info(f"Edge driver found at: {path}")
                return path

        # If we get here, we need to download the driver
        logging.info("Edge driver not found in expected locations.")

        # Get the Edge version
        edge_version = get_edge_version()
        if not edge_version:
            logging.error("Could not determine Edge version.")
            return None

        # Create a temporary directory for downloading
        temp_dir = tempfile.mkdtemp()
        try:
            # Download URL
            download_url = f"https://msedgedriver.azureedge.net/{edge_version}/edgedriver_win64.zip"
            zip_path = os.path.join(temp_dir, "edgedriver.zip")

            # Download using InetRead equivalent
            logging.info(f"Downloading Edge driver from: {download_url}")
            if not download_file_inet_read(download_url, zip_path):
                logging.error("Failed to download Edge driver using InetRead method.")
                return None

            # Extract using Shell.Application
            logging.info("Extracting Edge driver using Shell.Application.")
            if not extract_zip_with_shell(zip_path, base_path):
                logging.error("Failed to extract Edge driver using Shell.Application.")
                return None

            # Clean up zip file
            try:
                os.remove(zip_path)
            except:
                pass

            # Check if the driver was extracted successfully
            if os.path.exists(driver_path):
                logging.info(f"Edge driver downloaded and extracted to: {driver_path}")
                return driver_path
            else:
                # Look for the driver in extracted files
                for root, files in os.walk(base_path):
                    for file in files:
                        if file.lower() == "msedgedriver.exe":
                            found_path = os.path.join(root, file)
                            if os.path.dirname(found_path) != base_path:
                                shutil.move(found_path, driver_path)
                                logging.info(f"Moved Edge driver to: {driver_path}")
                            return driver_path if os.path.exists(driver_path) else found_path

                logging.error("Edge driver not found in extracted files.")
                return None

        except Exception as e:
            logging.error(f"Error downloading or extracting Edge driver: {e}")
            return None
        finally:
            # Clean up the temporary directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

    except Exception as e:
        logging.error(f"Error setting up Edge driver: {e}")
        return None


def initialize_logging(folder_path):
    """Initialize logging to file"""
    log_file = os.path.join(folder_path, LOG_FILE_NAME)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    logging.info(f"Logging initialized. Log file: {log_file}")

LOG_FILE_NAME = "inet_read.log"

if __name__ == "__main__":
    # Example usage: download and extract Edge driver
    folder = os.path.dirname(os.path.abspath(__file__))
    initialize_logging(folder)
    driver_path = get_edge_driver_path()
    if driver_path:
        print(f"Edge driver is ready at: {driver_path}")
    else:
        print("Failed to set up Edge driver.")