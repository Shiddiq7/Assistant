import speech_recognition as sr
import pyttsx3
import os
import subprocess
import time
import sys
import json
import winreg
from pathlib import Path
import glob
import ctypes
import win32com.client
import win32gui
import win32con
import pythoncom
from threading import Thread
from concurrent.futures import ThreadPoolExecutor
import pickle
from datetime import datetime, timedelta


class AppCache:
    def __init__(self, cache_file='app_cache.txt', cache_duration=timedelta(hours=24)):
        self.cache_file = cache_file
        self.cache_duration = cache_duration
        self.app_cache = {}
        self.load_cache()

    def load_cache(self):
        """Memuat cache aplikasi dari file"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'rb') as f:
                    cached_data = pickle.load(f)
                    if datetime.now() - cached_data['timestamp'] < self.cache_duration:
                        self.app_cache = cached_data['apps']
        except Exception as e:
            print(f"Terjadi kesalahan saat memuat cache: {e}")
            self.app_cache = {}

    def save_cache(self):
        """Menyimpan cache aplikasi ke file"""
        try:
            cache_data = {
                'timestamp': datetime.now(),
                'apps': self.app_cache
            }
            with open(self.cache_file, 'wb') as f:
                pickle.dump(cache_data, f)
        except Exception as e:
            print(f"Terjadi kesalahan saat menyimpan cache: {e}")

    def get_app(self, app_name):
        """Mengambil aplikasi dari cache"""
        return self.app_cache.get(app_name.lower())

    def add_app(self, app_name, app_info):
        """Menambahkan aplikasi ke cache"""
        self.app_cache[app_name.lower()] = app_info
        self.save_cache()

class AppLauncher:
    def __init__(self):
        self.shell = None
        self.store_apps = {}
        self.app_cache = AppCache()
        self.common_locations = self._get_common_locations()
        self.init_com_objects()
        self.refresh_store_apps()

    def _get_common_locations(self):
        """Get list of common app locations"""
        locations = set()
        system_paths = [
            os.environ.get("ProgramFiles", ""),
            os.environ.get("ProgramFiles(x86)", ""),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs"),
            os.path.join(os.environ.get("APPDATA", ""), "Microsoft\\Windows\\Start Menu\\Programs"),
            os.path.join(os.environ.get("ProgramData", ""), "Microsoft\\Windows\\Start Menu\\Programs"),
            "C:\\Windows\\System32"
        ]
        return [path for path in system_paths if path and os.path.exists(path)]

    def init_com_objects(self):
        """Initialize COM objects"""
        try:
            pythoncom.CoInitialize()
            self.shell = win32com.client.Dispatch("WScript.Shell")
        except Exception as e:
            print(f"Error initializing COM objects: {e}")

    def refresh_store_apps(self):
        """Refresh Store apps list using PowerShell"""
        powershell_cmd = """
        Get-AppxPackage | Select-Object -Property PackageFamilyName, InstallLocation | ConvertTo-Json
        """
        try:
            result = subprocess.run(
                ["powershell", "-Command", powershell_cmd],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.stdout.strip():
                apps = json.loads(result.stdout)
                if isinstance(apps, list):
                    self.store_apps = {
                        app['PackageFamilyName'].lower(): app
                        for app in apps if app['PackageFamilyName']
                    }
        except Exception as e:
            print(f"Error refreshing store apps: {e}")

    def find_app_quick(self, app_name):
        """Quick app search in common locations"""
        app_name_lower = app_name.lower()
        
        # Check cache first
        cached_app = self.app_cache.get_app(app_name_lower)
        if cached_app:
            return cached_app

        # Check Store apps
        if app_name_lower in self.store_apps:
            app_info = {
                'name': app_name,
                'type': 'store',
                'info': self.store_apps[app_name_lower]
            }
            self.app_cache.add_app(app_name_lower, app_info)
            return app_info

        # Quick search in common locations
        for location in self.common_locations:
            # Check direct exe
            exe_path = os.path.join(location, f"{app_name}.exe")
            if os.path.exists(exe_path):
                app_info = {
                    'name': app_name,
                    'type': 'desktop',
                    'path': exe_path
                }
                self.app_cache.add_app(app_name_lower, app_info)
                return app_info

            # Check direct shortcut
            lnk_path = os.path.join(location, f"{app_name}.lnk")
            if os.path.exists(lnk_path):
                app_info = {
                    'name': app_name,
                    'type': 'shortcut',
                    'path': lnk_path
                }
                self.app_cache.add_app(app_name_lower, app_info)
                return app_info

        return None

    def find_app_deep(self, app_name):
        """Deep app search with limited depth"""
        app_name_lower = app_name.lower()
        max_depth = 2  # Limit search depth

        def search_directory(directory, current_depth):
            if current_depth > max_depth:
                return None

            try:
                for item in os.listdir(directory):
                    if current_depth > max_depth:
                        break

                    full_path = os.path.join(directory, item)
                    item_lower = item.lower()

                    # Check if file matches
                    if item_lower in [f"{app_name_lower}.exe", f"{app_name_lower}.lnk"]:
                        return {
                            'name': app_name,
                            'type': 'shortcut' if item_lower.endswith('.lnk') else 'desktop',
                            'path': full_path
                        }

                    # Recurse into directories
                    if os.path.isdir(full_path):
                        result = search_directory(full_path, current_depth + 1)
                        if result:
                            return result
            except Exception:
                pass
            return None

        # Search in common locations with limited depth
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = [
                executor.submit(search_directory, location, 0)
                for location in self.common_locations
            ]
            for future in futures:
                try:
                    result = future.result(timeout=2)
                    if result:
                        self.app_cache.add_app(app_name_lower, result)
                        return result
                except Exception:
                    continue

        return None

    def launch_app(self, app_info):
        """Launch application based on type"""
        try:
            if app_info['type'] == 'store':
                return self.launch_store_app(app_info['info'])
            elif app_info['type'] in ['desktop', 'shortcut']:
                return self.launch_desktop_app(app_info['path'])
        except Exception as e:
            print(f"Error launching app: {e}")
        return False

    def launch_store_app(self, app_info):
        """Launch Store app"""
        try:
            cmd = f"start shell:AppsFolder\\{app_info['PackageFamilyName']}!App"
            subprocess.run(cmd, shell=True, check=True)
            return True
        except Exception:
            return False

    def launch_desktop_app(self, path):
        """Launch desktop application"""
        try:
            if os.path.exists(path):
                os.startfile(path)
                return True
        except Exception:
            pass
        return False

def type_text(text, delay=0.03):
    for char in text:
        sys.stdout.write(char)
        sys.stdout.flush()
        time.sleep(delay)
    print()

def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Silakan bicara...")
        audio = recognizer.listen(source)
        try:
            text = recognizer.recognize_google(audio, language='id-ID')
            print(f"Anda berkata: {text}")
            return text
        except sr.UnknownValueError:
            print("Maaf, saya tidak mengerti apa yang Anda katakan.")
        except sr.RequestError:
            print("Maaf, ada masalah dengan layanan pengenalan suara.")
    return ""

def speak(text):
    print(text)
    engine = pyttsx3.init()
    engine.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0')
    engine.say(text)
    engine.runAndWait()

def execute_command(command):
    command = command.lower()
    
    if "buka" in command:
        app_name = command.replace("buka ", "").strip()
        
        # Initialize launcher
        launcher = AppLauncher()
        
        # Quick search first
        app_info = launcher.find_app_quick(app_name)
        
        # Deep search if quick search fails
        if not app_info:
            speak(f"Mencari {app_name}, mohon tunggu sebentar...")
            app_info = launcher.find_app_deep(app_name)
        
        if app_info and launcher.launch_app(app_info):
            speak(f"Membuka {app_name}")
        else:
            speak(f"Maaf, aplikasi {app_name} tidak ditemukan")
            
    elif "tutup" in command:
        app_name = command.replace("tutup ", "").strip()
        try:
            os.system(f"taskkill /F /IM {app_name}.exe")
            speak(f"Menutup {app_name}")
        except Exception:
            speak(f"Maaf, tidak dapat menutup {app_name}")
            
    elif "cari" in command:
        try:
            # Parse the search command more flexibly
            parts = command.split()
            search_terms = []
            browser = "chrome"  # default browser
            site = None
            
            # Skip the "cari" word
            i = 1
            while i < len(parts):
                if parts[i] == "di":
                    if i + 1 < len(parts):
                        # Check if it's a browser or a site
                        if parts[i + 1] in ["chrome", "firefox", "edge", "opera"]:
                            browser = parts[i + 1]
                        else:
                            site = parts[i + 1]
                        i += 2
                        continue
                search_terms.append(parts[i])
                i += 1
            
            search_query = " ".join(search_terms)
            
            # Define browser configurations
            browser_configs = {
                'chrome': {
                    'paths': [
                        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
                    ],
                    'search_urls': {
                        'google': 'https://www.google.com/search?q={}',
                        'youtube': 'https://www.youtube.com/results?search_query={}',
                        'github': 'https://github.com/search?q={}',
                        'stackoverflow': 'https://stackoverflow.com/search?q={}',
                        'maps': 'https://www.google.com/maps/search/{}',
                        'images': 'https://www.google.com/search?q={}&tbm=isch'
                    }
                },
                'firefox': {
                    'paths': [
                        r'C:\Program Files\Mozilla Firefox\firefox.exe',
                        r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe'
                    ],
                    'search_urls': {
                        'google': 'https://www.google.com/search?q={}',
                        'youtube': 'https://www.youtube.com/results?search_query={}',
                        'github': 'https://github.com/search?q={}',
                        'stackoverflow': 'https://stackoverflow.com/search?q={}',
                        'maps': 'https://www.google.com/maps/search/{}',
                        'images': 'https://www.google.com/search?q={}&tbm=isch'
                    }
                },
                'edge': {
                    'paths': [
                        r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
                        r'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
                    ],
                    'search_urls': {
                        'google': 'https://www.google.com/search?q={}',
                        'youtube': 'https://www.youtube.com/results?search_query={}',
                        'github': 'https://github.com/search?q={}',
                        'stackoverflow': 'https://stackoverflow.com/search?q={}',
                        'maps': 'https://www.google.com/maps/search/{}',
                        'images': 'https://www.google.com/search?q={}&tbm=isch'
                    }
                },
                'opera': {
                    'paths': [
                        r'C:\Users\shiddiq\AppData\Local\Programs\Opera GX\opera.exe'
                    ],
                    'search_urls': {
                        'google': 'https://www.google.com/search?q={}',
                        'youtube': 'https://www.youtube.com/results?search_query={}',
                        'github': 'https://github.com/search?q={}',
                        'stackoverflow': 'https://stackoverflow.com/search?q={}',
                        'maps': 'https://www.google.com/maps/search/{}',
                        'images': 'https://www.google.com/search?q={}&tbm=isch'
                    }
                }
            }
            
            # Find browser executable
            browser_path = None
            for path in browser_configs[browser]['paths']:
                if os.path.exists(path):
                    browser_path = path
                    break
            
            if not browser_path:
                speak(f"Maaf, browser {browser} tidak ditemukan")
                return
            
            # Construct search URL based on site
            search_url = None
            if site:
                site_lower = site.lower()
                if site_lower in browser_configs[browser]['search_urls']:
                    search_url = browser_configs[browser]['search_urls'][site_lower].format(search_query)
                else:
                    # For other sites, search on the site directly if it's a valid domain
                    if '.' in site:  # Basic domain validation
                        search_url = f"https://{site}/search?q={search_query}"
                    else:
                        search_url = browser_configs[browser]['search_urls']['google'].format(f"site:{site} {search_query}")
            else:
                # Default to Google search
                search_url = browser_configs[browser]['search_urls']['google'].format(search_query)
            
            # Launch browser with search
            subprocess.Popen([browser_path, search_url])
            
            # Provide feedback
            if site:
                speak(f"Mencari '{search_query}' di {site} menggunakan {browser}")
            else:
                speak(f"Mencari '{search_query}' menggunakan {browser}")
                
        except Exception as e:
            speak(f"Maaf, terjadi kesalahan saat melakukan pencarian: {str(e)}")
            
    else:
        try:
            output = subprocess.getoutput(command)
            print(output)
            speak("Perintah berhasil dijalankan")
        except:
            speak("Maaf, saya tidak dapat menjalankan perintah tersebut")

def main():
    speak("Halo! Saya adalah asisten virtual Anda. Apa yang bisa saya bantu hari ini?")
    while True:
        command = listen()
        if "berhenti" in command.lower():
            speak("Sampai jumpa!")
            break
        elif command:
            execute_command(command)

if __name__ == "__main__":
    main()