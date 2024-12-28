import speech_recognition as sr
import pyttsx3
import os
import subprocess
import time
import sys

def type_text(text, delay=0.03):
    for char in text:
        sys.stdout.write(char)
        sys.stdout.flush()
        time.sleep(delay)
    print()

def listen():
    # Inisialisasi recognizer
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        type_text("Silakan bicara...")
        audio = recognizer.listen(source)
        try:
            # Menggunakan Google Web Speech API
            text = recognizer.recognize_google(audio, language='id-ID')
            type_text(f"Anda berkata: {text}")
            return text
        except sr.UnknownValueError:
            type_text("Maaf, saya tidak mengerti apa yang Anda katakan.")
        except sr.RequestError:
            type_text("Maaf, ada masalah dengan layanan pengenalan suara.")
    return ""

def speak(text):
    # Inisialisasi engine text-to-speech
    type_text(text)
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def execute_command(command):
    command = command.lower()
    
    # Menangani perintah untuk membuka aplikasi
    if "buka" in command:
        app_name = command.replace("buka ", "").strip().lower()
        try:
            if os.name == 'nt':  # Windows
                # Pemetaan nama aplikasi ke path khusus
                app_mappings = {
                    'spotify': r'C:\Users\%USERNAME%\AppData\Roaming\Spotify\Spotify.exe',
                    'chrome': r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                    'firefox': r'C:\Program Files\Mozilla Firefox\firefox.exe',
                    'word': r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
                    'excel': r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE',
                    'powerpoint': r'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE',
                    'notepad': r'C:\Windows\System32\notepad.exe',
                    'paint': r'C:\Windows\System32\mspaint.exe',
                    'calculator': r'C:\Windows\System32\calc.exe',
                    'vlc': r'C:\Program Files\VideoLAN\VLC\vlc.exe',
                    'winamp': r'C:\Program Files (x86)\Winamp\winamp.exe',
                    'photoshop': r'C:\Program Files\Adobe\Adobe Photoshop CC 2020\Photoshop.exe',
                    'visual studio code': r'C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe',
                    'steam': r'C:\Program Files (x86)\Steam\Steam.exe',
                    'discord': r'C:\Users\%USERNAME%\AppData\Local\Discord\app-1.0.9003\Discord.exe'
                }
                
                # Cek apakah ada mapping khusus untuk aplikasi
                if app_name in app_mappings:
                    path = os.path.expandvars(app_mappings[app_name])
                    if os.path.exists(path):
                        os.startfile(path)
                        speak(f"Membuka {app_name}")
                        return
                
                # Cek di berbagai lokasi program Windows
                program_paths = [
                    os.path.join(os.environ["ProgramFiles"], f"{app_name}.exe"),
                    os.path.join(os.environ["ProgramFiles"], f"{app_name}\\{app_name}.exe"),
                    os.path.join(os.environ["ProgramFiles(x86)"], f"{app_name}.exe"), 
                    os.path.join(os.environ["ProgramFiles(x86)"], f"{app_name}\\{app_name}.exe"),
                    os.path.join(os.environ["APPDATA"], f"{app_name}.exe"),
                    os.path.join(os.environ["LOCALAPPDATA"], f"{app_name}.exe"),
                    os.path.join(os.environ["LOCALAPPDATA"], f"{app_name}\\{app_name}.exe"),
                    os.path.join("C:\\Windows\\System32", f"{app_name}.exe")
                ]
                
                # Cari file executable
                found = False
                for path in program_paths:
                    if os.path.exists(path):
                        os.startfile(path)
                        speak(f"Membuka {app_name}")
                        found = True
                        break
                
                if not found:
                    # Coba cari di Start Menu
                    start_menu = os.path.join(os.environ["ProgramData"], "Microsoft\\Windows\\Start Menu\\Programs")
                    user_start_menu = os.path.join(os.environ["APPDATA"], "Microsoft\\Windows\\Start Menu\\Programs")
                    
                    for menu_path in [start_menu, user_start_menu]:
                        for root, dirs, files in os.walk(menu_path):
                            for file in files:
                                if app_name in file.lower() and file.endswith(".lnk"):
                                    os.startfile(os.path.join(root, file))
                                    speak(f"Membuka {app_name}")
                                    found = True
                                    break
                            if found:
                                break
                            
                    if not found:
                        try:
                            # Coba buka menggunakan nama aplikasi langsung
                            os.system(f"start {app_name}")
                            speak(f"Membuka {app_name}")
                        except Exception as e:
                            if "The system cannot find the file" in str(e):
                                speak(f"Maaf, aplikasi {app_name} tidak ditemukan di sistem")
                            else:
                                speak(f"Terjadi kesalahan saat membuka {app_name}")
                
            else:  # Linux/Mac
                # Pemetaan nama aplikasi ke path khusus untuk Linux/Mac
                app_mappings = {
                    'spotify': '/snap/bin/spotify',  # Untuk Linux dengan Snap
                    # Tambahkan aplikasi lain di sini
                }
                
                # Cek apakah ada mapping khusus untuk aplikasi
                if app_name in app_mappings:
                    path = app_mappings[app_name]
                    if os.path.exists(path):
                        subprocess.Popen([path])
                        speak(f"Membuka {app_name}")
                        return
                
                # Cek di lokasi umum aplikasi Linux/Mac
                common_paths = [
                    f"/usr/bin/{app_name}",
                    f"/usr/local/bin/{app_name}",
                    f"/snap/bin/{app_name}",
                    f"/opt/{app_name}",
                    f"/opt/{app_name}/bin/{app_name}",
                    f"/Applications/{app_name}.app",  # untuk Mac
                    f"/usr/share/applications/{app_name}.desktop"
                ]
                
                # Cari executable
                found = False
                for path in common_paths:
                    if os.path.exists(path):
                        if path.endswith('.desktop'):
                            subprocess.Popen(['xdg-open', path])
                        else:
                            subprocess.Popen([path])
                        speak(f"Membuka {app_name}")
                        found = True
                        break
                
                if not found:
                    # Coba cari menggunakan which
                    try:
                        app_path = subprocess.getoutput(f"which {app_name}")
                        if app_path and os.path.exists(app_path):
                            subprocess.Popen([app_path])
                            speak(f"Membuka {app_name}")
                            found = True
                    except:
                        pass
                        
                    if not found:
                        try:
                            # Coba buka langsung
                            subprocess.Popen([app_name])
                            speak(f"Membuka {app_name}")
                        except Exception as e:
                            if "No such file or directory" in str(e):
                                speak(f"Maaf, aplikasi {app_name} tidak ditemukan di sistem")
                            else:
                                speak(f"Terjadi kesalahan saat membuka {app_name}")
                
        except Exception as e:
            type_text(f"Error: {str(e)}")
            if "The system cannot find the file" in str(e):
                speak(f"Maaf, aplikasi {app_name} tidak ditemukan di sistem")
            else:
                speak(f"Maaf, tidak dapat membuka {app_name}")

            
    #  Menangani perintah untuk menutup aplikasi
    elif "tutup" in command:
        app_name = command.replace("tutup ", "").strip().lower()
        try:
            if os.name == 'nt':  # Windows
                os.system(f"taskkill /F /IM {app_name}.exe")
                speak(f"Menutup {app_name}")
            else:  # Linux/Mac
                os.system(f"pkill -f {app_name}")
                speak(f"Menutup {app_name}")
        except Exception as e:
            type_text(f"Error: {str(e)}")
            if "The system cannot find the file" in str(e):
                speak(f"Maaf, aplikasi {app_name} tidak ditemukan di sistem")
            else:
                speak(f"Maaf, tidak dapat menutup {app_name}")


    
    # Menangani perintah pencarian web
    elif "cari" in command:
        # Memisahkan perintah browser dan query pencarian
        parts = command.replace("cari ", "").strip().split(" di ")
        search_query = parts[0]
        browser = parts[1].lower() if len(parts) > 1 else "chrome"  # Default ke Chrome jika browser tidak disebutkan
        
        try:
            if os.name == 'nt':  # Windows
                browser_paths = {
                    'chrome': r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                    'firefox': r'C:\Program Files\Mozilla Firefox\firefox.exe',
                    'opera': r'C:\Users\shiddiq\AppData\Local\Programs\Opera GX\opera.exe',
                    'edge': r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
                }
                
                # Cek path alternatif untuk Opera
                if browser == 'opera' and not os.path.exists(browser_paths['opera']):
                    alt_paths = [
                        r'C:\Users\shiddiq\AppData\Local\Programs\Opera GX\opera.exe',
                        r'C:\Program Files\Opera GX\opera.exe',
                        r'C:\Program Files (x86)\Opera GX\opera.exe'
                    ]
                    for path in alt_paths:
                        if os.path.exists(path):
                            browser_paths['opera'] = path
                            break
                
                if browser in browser_paths and os.path.exists(browser_paths[browser]):
                    subprocess.Popen([browser_paths[browser], f'https://www.google.com/search?q={search_query}'])
                    speak(f"Mencari {search_query} di {browser}")
                else:
                    speak(f"Maaf, browser {browser} tidak ditemukan")
            
            else:  # Linux/Mac
                browser_commands = {
                    'chrome': 'google-chrome',
                    'firefox': 'firefox',
                    'opera': 'opera',
                    'safari': 'safari'
                }
                
                if browser in browser_commands and os.system(f"which {browser_commands[browser]}") == 0:
                    os.system(f"{browser_commands[browser]} 'https://www.google.com/search?q={search_query}'")
                    speak(f"Mencari {search_query} di {browser}")
                else:
                    os.system(f"xdg-open 'https://www.google.com/search?q={search_query}'")
                    speak(f"Mencari {search_query} menggunakan browser default")
                    
        except Exception as e:
            speak(f"Maaf, terjadi kesalahan saat membuka {browser}")
    
    else:
        try:
            # Mencoba menjalankan perintah langsung
            output = subprocess.getoutput(command)
            type_text(output)
            speak("Perintah berhasil dijalankan")
        except:
            speak("Maaf, saya tidak dapat menjalankan perintah tersebut")
            

def main():
    type_text = "Halo! Saya adalah asisten virtual Anda. Apa yang bisa saya bantu hari ini?"
    speak(type_text)
    while True:
        command = listen()
        if "berhenti" in command.lower():
            speak("Sampai jumpa!")
            break
        elif command:
            execute_command(command)

if __name__ == "__main__":
    main()
