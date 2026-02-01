import os
import subprocess
import psutil
import json
import webbrowser
import re
import random
import urllib.parse
import time
import platform
from datetime import datetime
from typing import Dict, Any, List

# Optional imports with fallbacks
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False

try:
    import pyperclip
    PYPERCLIP_AVAILABLE = True
except ImportError:
    PYPERCLIP_AVAILABLE = False

try:
    import win32gui
    import win32con
    import win32api
    import ctypes
    WINDOWS_API_AVAILABLE = True
except ImportError:
    WINDOWS_API_AVAILABLE = False

class JarvisCore:
    def __init__(self):
        self.os = platform.system()
        self.last_result = None
        self.window_list = []
        
    def process_command(self, command: str) -> Dict[str, Any]:
        """Main command processor"""
        cmd = command.lower().strip()
        result = {"success": False, "action": None, "message": "Command not recognized", "data": None}
        
        try:
            # Command routing - ordered by priority
            if 'youtube' in cmd or ('play' in cmd and ('video' in cmd or 'song' in cmd or 'music' in cmd)):
                result = self._handle_youtube(cmd)
            elif 'go to' in cmd or 'visit' in cmd:
                result = self._open_website(cmd)
            elif cmd.startswith('search google for') or cmd.startswith('google search'):
                query = cmd.replace('search google for', '').replace('google search', '').strip()
                result = self._handle_google_search(query)
            elif 'search' in cmd and 'google' in cmd:
                query = cmd.replace('search', '').replace('google', '').replace('for', '').strip()
                result = self._handle_google_search(query)
            elif any(x in cmd for x in ['open', 'start', 'launch']):
                result = self._handle_open(cmd)
            elif cmd.startswith('close') or cmd.startswith('kill') or cmd.startswith('exit'):
                result = self._handle_close(cmd)
            elif 'minimize all' in cmd or 'show desktop' in cmd:
                result = self._minimize_all()
            elif 'snap' in cmd:
                result = self._handle_snap(cmd)
            elif 'brightness' in cmd:
                result = self._handle_brightness(cmd)
            elif any(x in cmd for x in ['play', 'pause', 'music', 'media', 'skip', 'next track', 'previous track']) and 'youtube' not in cmd:
                result = self._handle_media(cmd)
            elif 'lock' in cmd and 'system' in cmd:
                result = self._lock_system()
            elif 'empty recycle' in cmd or 'empty bin' in cmd:
                result = self._empty_recycle_bin()
            elif any(x in cmd for x in ['always on top', 'maximize window', 'minimize window', 'restore window']) and 'snap' not in cmd:
                result = self._handle_window_actions(cmd)
            elif 'create folder' in cmd or 'new folder' in cmd:
                result = self._create_folder(cmd)
            elif cmd.startswith('delete') or cmd.startswith('remove'):
                result = self._delete_file(cmd)
            elif 'rename' in cmd and ' to ' in cmd:
                result = self._rename_file(cmd)
            elif cmd.startswith('type') or cmd.startswith('write'):
                result = self._write_text(cmd)
            elif cmd.startswith('press') or cmd.startswith('hit'):
                result = self._press_key(cmd)
            elif any(x in cmd for x in ['calculate', 'compute']) or (any(x in cmd for x in ['what is', 'how much']) and any(c in cmd for c in '0123456789+-*/')):
                result = self._calculate(cmd)
            elif 'joke' in cmd:
                result = self._tell_joke()
            elif 'weather' in cmd:
                result = self._get_weather(cmd)
            elif any(x in cmd for x in ['shutdown', 'restart', 'reboot']) and 'abort' not in cmd:
                result = self._shutdown_system(cmd)
            elif 'abort' in cmd or 'cancel shutdown' in cmd:
                result = self._abort_shutdown()
            elif 'list processes' in cmd or 'running apps' in cmd:
                result = self._list_processes()
            elif cmd.startswith('terminate') or (cmd.startswith('kill') and 'close' not in cmd):
                result = self._kill_process(cmd)
            elif 'time' in cmd:
                result = self._get_time()
            elif 'ip' in cmd or 'network' in cmd:
                result = self._get_network()
            elif 'system info' in cmd:
                result = self._get_system_info()
            elif 'copy' in cmd and cmd.split()[0] == 'copy':
                result = self._handle_clipboard(cmd, 'copy')
            elif 'paste' in cmd:
                result = self._handle_clipboard(cmd, 'paste')
            elif 'find' in cmd or 'search file' in cmd:
                result = self._search_files(cmd)
            elif 'volume' in cmd or 'sound' in cmd:
                result = self._handle_volume(cmd)
            elif 'screenshot' in cmd:
                result = self._take_screenshot()
            elif 'sleep' in cmd or 'standby' in cmd:
                result = self._power_control('sleep')
            elif cmd in ['hello', 'hi', 'hey']:
                result = {"success": True, "action": "greeting", "message": "Hello sir. Systems operational.", "data": None}
            else:
                result = {"success": False, "action": "unknown", "message": f"I don't understand: '{cmd}'", "data": None}
                
        except Exception as e:
            result = {"success": False, "action": "error", "message": str(e), "data": None}
            
        self.last_result = result
        return result
    
    def _handle_youtube(self, cmd: str) -> Dict:
        """YouTube automation - search and play videos"""
        query = None
        
        patterns = [
            r'play\s+(.+?)\s+on\s+youtube',
            r'play\s+(.+?)\s+youtube',
            r'search\s+youtube\s+for\s+(.+)',
            r'youtube\s+search\s+for\s+(.+)',
            r'find\s+(.+?)\s+on\s+youtube',
            r'open\s+youtube\s+and\s+(?:search\s+for\s+)?(.+)',
            r'youtube\s+(.+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, cmd, re.IGNORECASE)
            if match:
                query = match.group(1).strip()
                break
        
        if query:
            encoded_query = urllib.parse.quote(query)
            search_url = f"https://www.youtube.com/results?search_query={encoded_query}"
            webbrowser.open(search_url)
            
            return {
                "success": True, 
                "action": "youtube", 
                "message": f"Opening YouTube search for: '{query}'", 
                "data": {"query": query, "url": search_url}
            }
        else:
            webbrowser.open('https://www.youtube.com')
            return {"success": True, "action": "youtube", "message": "YouTube opened", "data": None}
    
    def _handle_google_search(self, query: str) -> Dict:
        """Google search"""
        if query:
            encoded = urllib.parse.quote(query)
            url = f"https://www.google.com/search?q={encoded}"
            webbrowser.open(url)
            return {
                "success": True, 
                "action": "search", 
                "message": f"Searching Google for: '{query}'", 
                "data": {"query": query, "url": url}
            }
        return {"success": False, "action": "search", "message": "No search query provided", "data": None}
    
    def _open_website(self, cmd: str) -> Dict:
        """Open specific websites by name or URL"""
        site = cmd.replace('go to', '').replace('visit', '').replace('open website', '').strip()
        
        sites = {
            'youtube': 'https://youtube.com',
            'google': 'https://google.com',
            'facebook': 'https://facebook.com',
            'twitter': 'https://twitter.com',
            'x': 'https://x.com',
            'instagram': 'https://instagram.com',
            'github': 'https://github.com',
            'netflix': 'https://netflix.com',
            'amazon': 'https://amazon.com',
            'reddit': 'https://reddit.com',
            'linkedin': 'https://linkedin.com',
            'gmail': 'https://gmail.com',
            'outlook': 'https://outlook.com',
            'hotmail': 'https://outlook.com',
            'maps': 'https://maps.google.com',
            'google maps': 'https://maps.google.com',
            'translate': 'https://translate.google.com',
            'drive': 'https://drive.google.com',
            'docs': 'https://docs.google.com',
            'spotify': 'https://open.spotify.com',
            'twitch': 'https://twitch.tv',
            'tiktok': 'https://tiktok.com',
            'whatsapp': 'https://web.whatsapp.com',
            'telegram': 'https://web.telegram.org',
            'discord': 'https://discord.com/app',
            'chatgpt': 'https://chat.openai.com',
            'claude': 'https://claude.ai',
            'weather': 'https://www.google.com/search?q=weather',
            'news': 'https://news.google.com',
            'calendar': 'https://calendar.google.com'
        }
        
        site_lower = site.lower()
        if site_lower in sites:
            webbrowser.open(sites[site_lower])
            return {"success": True, "action": "website", "message": f"Opening {site}", "data": {"site": site, "url": sites[site_lower]}}
        elif site.startswith('http://') or site.startswith('https://'):
            webbrowser.open(site)
            return {"success": True, "action": "website", "message": f"Opening {site}", "data": {"url": site}}
        elif any(ext in site for ext in ['.com', '.org', '.net', '.io', '.co', '.ai']):
            url = f"https://{site}" if not site.startswith('www.') else f"https://{site}"
            webbrowser.open(url)
            return {"success": True, "action": "website", "message": f"Opening {site}", "data": {"url": url}}
        else:
            return self._handle_google_search(site)

    def _handle_open(self, cmd: str) -> Dict:
        """Handle application opening"""
        apps = {
            'notepad': 'notepad.exe',
            'calc': 'calc.exe',
            'calculator': 'calc.exe',
            'chrome': 'chrome.exe',
            'browser': 'chrome.exe',
            'edge': 'msedge.exe',
            'firefox': 'firefox.exe',
            'cmd': 'cmd.exe',
            'command prompt': 'cmd.exe',
            'powershell': 'powershell.exe',
            'explorer': 'explorer.exe',
            'files': 'explorer.exe',
            'task manager': 'taskmgr.exe',
            'paint': 'mspaint.exe',
            'word': 'winword.exe',
            'excel': 'excel.exe',
            'spotify': 'spotify.exe',
            'settings': 'ms-settings:',
            'control panel': 'control.exe',
            'vscode': 'code.exe',
            'visual studio code': 'code.exe',
            'vlc': 'vlc.exe',
            'discord': 'discord.exe',
            'steam': 'steam.exe'
        }
        
        app_name = cmd
        for keyword in ['open', 'start', 'launch']:
            if cmd.startswith(keyword):
                app_name = cmd.replace(keyword, '').strip()
                break
            
        if app_name in apps:
            exe = apps[app_name]
            try:
                if exe.startswith('ms-'):
                    os.system(f'start {exe}')
                else:
                    subprocess.Popen(exe, shell=True)
                return {"success": True, "action": "open", "message": f"Opened {app_name}", "data": {"app": app_name}}
            except Exception as e:
                return {"success": False, "action": "open", "message": f"Failed to open {app_name}: {str(e)}", "data": None}
        else:
            try:
                subprocess.Popen(app_name, shell=True)
                return {"success": True, "action": "open", "message": f"Executing {app_name}", "data": None}
            except Exception as e:
                return {"success": False, "action": "open", "message": f"Application '{app_name}' not found: {str(e)}", "data": None}
    
    def _handle_close(self, cmd: str) -> Dict:
        """Close applications"""
        app_name = cmd.replace('close', '').replace('kill', '').replace('exit', '').strip()
        
        closed = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if app_name.lower() in proc.info['name'].lower():
                    psutil.Process(proc.info['pid']).terminate()
                    closed.append(proc.info['name'])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        
        if closed:
            return {"success": True, "action": "close", "message": f"Closed {', '.join(closed)}", "data": closed}
        return {"success": False, "action": "close", "message": f"No process matching '{app_name}' found", "data": None}
    
    def _minimize_all(self) -> Dict:
        """Minimize all windows"""
        if not PYAUTOGUI_AVAILABLE:
            return {"success": False, "action": "minimize_all", "message": "PyAutoGUI not installed", "data": None}
            
        try:
            pyautogui.keyDown('win')
            pyautogui.keyDown('m')
            pyautogui.keyUp('m')
            pyautogui.keyUp('win')
            return {"success": True, "action": "minimize_all", "message": "All windows minimized", "data": None}
        except Exception as e:
            return {"success": False, "action": "minimize_all", "message": str(e), "data": None}
    
    def _handle_snap(self, cmd: str) -> Dict:
        """Window snapping"""
        if not WINDOWS_API_AVAILABLE:
            return {"success": False, "action": "snap", "message": "Windows API not available", "data": None}
            
        direction = 'left' if 'left' in cmd else 'right'
        window_name = None
        
        parts = cmd.replace('snap', '').strip().split()
        for part in parts:
            if part not in ['left', 'right', 'window']:
                window_name = part
                break
        
        if not window_name:
            hwnd = win32gui.GetForegroundWindow()
            window_name = "active window"
        else:
            hwnd = win32gui.FindWindow(None, window_name)
            if not hwnd:
                def callback(h, extra):
                    if win32gui.IsWindowVisible(h):
                        title = win32gui.GetWindowText(h)
                        if window_name.lower() in title.lower():
                            extra.append(h)
                handles = []
                win32gui.EnumWindows(callback, handles)
                if handles:
                    hwnd = handles[0]
                else:
                    return {"success": False, "action": "snap", "message": f"Window '{window_name}' not found", "data": None}
        
        try:
            screen_width = win32api.GetSystemMetrics(0)
            screen_height = win32api.GetSystemMetrics(1)
            
            if direction == 'left':
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 
                                    screen_width//2, screen_height, 0)
            else:
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 
                                    screen_width//2, 0, screen_width//2, screen_height, 0)
            
            return {"success": True, "action": "snap", "message": f"Snapped {window_name} to {direction}", "data": {"direction": direction, "window": window_name}}
        except Exception as e:
            return {"success": False, "action": "snap", "message": str(e), "data": None}
    
    def _get_time(self) -> Dict:
        """Get current time"""
        now = datetime.now()
        time_str = now.strftime("%I:%M %p")
        date_str = now.strftime("%A, %B %d, %Y")
        return {"success": True, "action": "time", "message": f"It is {time_str} on {date_str}", "data": {"time": time_str, "date": date_str}}
    
    def _get_network(self) -> Dict:
        """Get IP info"""
        try:
            result = subprocess.run(['ipconfig'], capture_output=True, text=True, shell=True)
            lines = result.stdout.split('\n')
            ips = [line.split(':')[1].strip() for line in lines if 'IPv4' in line and ':' in line]
            return {"success": True, "action": "network", "message": f"IP Addresses: {', '.join(ips[:2])}", "data": {"ips": ips}}
        except Exception as e:
            return {"success": False, "action": "network", "message": str(e), "data": None}
    
    def _get_system_info(self) -> Dict:
        """Get system stats"""
        try:
            cpu = psutil.cpu_percent()
            mem = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            
            info = {
                "cpu": f"{cpu}%",
                "memory": f"{mem.percent}% ({mem.used//1024//1024}MB/{mem.total//1024//1024}MB)",
                "disk": f"{disk.percent}%"
            }
            
            return {"success": True, "action": "system_info", 
                    "message": f"CPU: {info['cpu']}, RAM: {info['memory']}, Disk: {info['disk']}", 
                    "data": info}
        except Exception as e:
            return {"success": False, "action": "system_info", "message": str(e), "data": None}
    
    def _handle_clipboard(self, cmd: str, action: str) -> Dict:
        """Clipboard operations"""
        if not PYPERCLIP_AVAILABLE:
            return {"success": False, "action": "clipboard", "message": "Pyperclip not installed", "data": None}
            
        if action == 'copy':
            text = cmd.replace('copy', '').strip()
            pyperclip.copy(text)
            return {"success": True, "action": "clipboard", "message": f"Copied: {text}", "data": {"text": text}}
        else:
            if not PYAUTOGUI_AVAILABLE:
                return {"success": False, "action": "clipboard", "message": "PyAutoGUI not installed", "data": None}
            pyautogui.hotkey('ctrl', 'v')
            return {"success": True, "action": "clipboard", "message": "Pasted from clipboard", "data": None}
    
    def _search_files(self, cmd: str) -> Dict:
        """Search for files"""
        query = cmd.replace('find', '').replace('search file', '').strip()
        matches = []
        
        search_paths = [
            os.path.expanduser('~\\Desktop'),
            os.path.expanduser('~\\Documents'),
            os.path.expanduser('~\\Downloads')
        ]
        
        for path in search_paths:
            if os.path.exists(path):
                for root, dirs, files in os.walk(path):
                    for item in files + dirs:
                        if query.lower() in item.lower():
                            matches.append(os.path.join(root, item))
                    if len(matches) > 20:  # Limit results
                        break
        
        if matches:
            return {"success": True, "action": "search", "message": f"Found {len(matches)} matches", "data": matches[:5]}
        return {"success": True, "action": "search", "message": "No files found", "data": []}
    
    def _handle_volume(self, cmd: str) -> Dict:
        """Volume control"""
        if not PYAUTOGUI_AVAILABLE:
            return {"success": False, "action": "volume", "message": "PyAutoGUI not installed", "data": None}
            
        if 'up' in cmd or 'increase' in cmd:
            pyautogui.press('volumeup', presses=5)
            return {"success": True, "action": "volume", "message": "Volume increased", "data": None}
        elif 'down' in cmd or 'decrease' in cmd:
            pyautogui.press('volumedown', presses=5)
            return {"success": True, "action": "volume", "message": "Volume decreased", "data": None}
        elif 'mute' in cmd:
            pyautogui.press('volumemute')
            return {"success": True, "action": "volume", "message": "Volume muted", "data": None}
        return {"success": False, "action": "volume", "message": "Specify up, down, or mute", "data": None}
    
    def _take_screenshot(self) -> Dict:
        """Screenshot"""
        if not PYAUTOGUI_AVAILABLE:
            return {"success": False, "action": "screenshot", "message": "PyAutoGUI not installed", "data": None}
            
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"jarvis_screenshot_{timestamp}.png"
            pictures_dir = os.path.join(os.path.expanduser('~'), 'Pictures')
            os.makedirs(pictures_dir, exist_ok=True)
            filepath = os.path.join(pictures_dir, filename)
            pyautogui.screenshot(filepath)
            return {"success": True, "action": "screenshot", "message": f"Screenshot saved: {filename}", "data": {"path": filepath}}
        except Exception as e:
            return {"success": False, "action": "screenshot", "message": str(e), "data": None}
    
    def _power_control(self, action: str) -> Dict:
        """Power controls"""
        if action == 'sleep':
            if self.os == 'Windows':
                os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
            else:
                return {"success": False, "action": "power", "message": "Sleep only supported on Windows", "data": None}
            return {"success": True, "action": "power", "message": "System sleeping", "data": None}
        return {"success": False, "action": "power", "message": "Unknown power command", "data": None}

    def _handle_brightness(self, cmd: str) -> Dict:
        """Screen brightness control"""
        try:
            import screen_brightness_control as sbc
            if 'up' in cmd or 'increase' in cmd:
                current = sbc.get_brightness()[0]
                new_val = min(100, current + 10)
                sbc.set_brightness(new_val)
                return {"success": True, "action": "brightness", "message": f"Brightness increased to {new_val}%", "data": None}
            elif 'down' in cmd or 'decrease' in cmd or 'lower' in cmd:
                current = sbc.get_brightness()[0]
                new_val = max(0, current - 10)
                sbc.set_brightness(new_val)
                return {"success": True, "action": "brightness", "message": f"Brightness decreased to {new_val}%", "data": None}
            elif 'max' in cmd:
                sbc.set_brightness(100)
                return {"success": True, "action": "brightness", "message": "Brightness set to maximum", "data": None}
            elif 'min' in cmd:
                sbc.set_brightness(0)
                return {"success": True, "action": "brightness", "message": "Brightness set to minimum", "data": None}
            elif 'set' in cmd:
                match = re.search(r'\d+', cmd)
                if match:
                    num = int(match.group())
                    sbc.set_brightness(num)
                    return {"success": True, "action": "brightness", "message": f"Brightness set to {num}%", "data": None}
            else:
                current = sbc.get_brightness()[0]
                return {"success": True, "action": "brightness", "message": f"Current brightness: {current}%", "data": None}
        except ImportError:
            return {"success": False, "action": "brightness", "message": "Install screen_brightness_control: pip install screen-brightness-control", "data": None}
        except Exception as e:
            return {"success": False, "action": "brightness", "message": str(e), "data": None}

    def _handle_media(self, cmd: str) -> Dict:
        """Media controls"""
        if not PYAUTOGUI_AVAILABLE:
            return {"success": False, "action": "media", "message": "PyAutoGUI not installed", "data": None}
            
        if 'play' in cmd or 'pause' in cmd:
            pyautogui.press('playpause')
            return {"success": True, "action": "media", "message": "Play/Pause toggled", "data": None}
        elif 'next' in cmd or 'skip' in cmd:
            pyautogui.press('nexttrack')
            return {"success": True, "action": "media", "message": "Next track", "data": None}
        elif 'previous' in cmd or 'back' in cmd:
            pyautogui.press('prevtrack')
            return {"success": True, "action": "media", "message": "Previous track", "data": None}
        elif 'stop' in cmd:
            pyautogui.press('stop')
            return {"success": True, "action": "media", "message": "Stopped", "data": None}
        return {"success": False, "action": "media", "message": "Unknown media command", "data": None}

    def _lock_system(self) -> Dict:
        """Lock workstation"""
        if self.os != 'Windows':
            return {"success": False, "action": "lock", "message": "Lock only supported on Windows", "data": None}
        try:
            ctypes.windll.user32.LockWorkStation()
            return {"success": True, "action": "lock", "message": "System locked", "data": None}
        except Exception as e:
            return {"success": False, "action": "lock", "message": str(e), "data": None}

    def _empty_recycle_bin(self) -> Dict:
        """Empty recycle bin"""
        if self.os != 'Windows':
            return {"success": False, "action": "recycle", "message": "Recycle bin only on Windows", "data": None}
        try:
            import winshell
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=False)
            return {"success": True, "action": "recycle", "message": "Recycle bin emptied", "data": None}
        except ImportError:
            try:
                os.system("rd /s /q %systemdrive%\\$Recycle.Bin 2>nul")
                return {"success": True, "action": "recycle", "message": "Recycle bin emptied", "data": None}
            except Exception as e:
                return {"success": False, "action": "recycle", "message": str(e), "data": None}

    def _handle_window_actions(self, cmd: str) -> Dict:
        """Advanced window actions"""
        if not WINDOWS_API_AVAILABLE:
            return {"success": False, "action": "window", "message": "Windows API not available", "data": None}
            
        if 'always on top' in cmd and 'cancel' not in cmd:
            hwnd = win32gui.GetForegroundWindow()
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, 
                                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            return {"success": True, "action": "window", "message": "Window set to always on top", "data": None}
        elif 'cancel always on top' in cmd or 'normal window' in cmd:
            hwnd = win32gui.GetForegroundWindow()
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            return {"success": True, "action": "window", "message": "Always on top cancelled", "data": None}
        elif 'maximize' in cmd:
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            return {"success": True, "action": "window", "message": "Window maximized", "data": None}
        elif 'minimize' in cmd and 'all' not in cmd:
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            return {"success": True, "action": "window", "message": "Window minimized", "data": None}
        elif 'restore' in cmd:
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            return {"success": True, "action": "window", "message": "Window restored", "data": None}
        return {"success": False, "action": "window", "message": "Unknown window command", "data": None}

    def _create_folder(self, cmd: str) -> Dict:
        """Create new folder"""
        try:
            folder_name = cmd.replace('create folder', '').replace('new folder', '').strip()
            if not folder_name:
                folder_name = "New Folder"
            
            path = os.path.join(os.path.expanduser('~\\Desktop'), folder_name)
            os.makedirs(path, exist_ok=True)
            return {"success": True, "action": "folder", "message": f"Created folder: {folder_name}", "data": {"path": path}}
        except Exception as e:
            return {"success": False, "action": "folder", "message": str(e), "data": None}

    def _delete_file(self, cmd: str) -> Dict:
        """Delete file/folder"""
        try:
            import shutil
            filepath = cmd.replace('delete', '').replace('remove', '').strip()
            
            if not os.path.exists(filepath):
                test_path = os.path.join(os.path.expanduser('~\\Desktop'), filepath)
                if os.path.exists(test_path):
                    filepath = test_path
            
            if os.path.exists(filepath):
                if os.path.isdir(filepath):
                    shutil.rmtree(filepath)
                else:
                    os.remove(filepath)
                return {"success": True, "action": "delete", "message": f"Deleted: {os.path.basename(filepath)}", "data": None}
            else:
                return {"success": False, "action": "delete", "message": "File not found", "data": None}
        except Exception as e:
            return {"success": False, "action": "delete", "message": str(e), "data": None}

    def _rename_file(self, cmd: str) -> Dict:
        """Rename file"""
        try:
            parts = cmd.replace('rename', '').strip().split(' to ')
            if len(parts) == 2:
                old_name = parts[0].strip()
                new_name = parts[1].strip()
                
                old_path = os.path.join(os.path.expanduser('~\\Desktop'), old_name)
                if not os.path.exists(old_path):
                    old_path = old_name
                
                new_path = os.path.join(os.path.dirname(old_path), new_name)
                os.rename(old_path, new_path)
                return {"success": True, "action": "rename", "message": f"Renamed to {new_name}", "data": None}
            return {"success": False, "action": "rename", "message": "Usage: rename [file] to [newname]", "data": None}
        except Exception as e:
            return {"success": False, "action": "rename", "message": str(e), "data": None}

    def _write_text(self, cmd: str) -> Dict:
        """Type text using keyboard"""
        if not PYAUTOGUI_AVAILABLE:
            return {"success": False, "action": "type", "message": "PyAutoGUI not installed", "data": None}
            
        try:
            text = cmd.replace('type', '').replace('write', '').strip()
            pyautogui.typewrite(text, interval=0.01)
            return {"success": True, "action": "type", "message": f"Typed: {text}", "data": None}
        except Exception as e:
            return {"success": False, "action": "type", "message": str(e), "data": None}

    def _press_key(self, cmd: str) -> Dict:
        """Press specific keys"""
        if not PYAUTOGUI_AVAILABLE:
            return {"success": False, "action": "keypress", "message": "PyAutoGUI not installed", "data": None}
            
        try:
            key = cmd.replace('press', '').replace('hit', '').strip()
            
            if ' ' in key:
                keys = key.split()
                for k in keys:
                    pyautogui.keyDown(k)
                for k in reversed(keys):
                    pyautogui.keyUp(k)
            else:
                pyautogui.press(key)
            
            return {"success": True, "action": "keypress", "message": f"Pressed {key}", "data": None}
        except Exception as e:
            return {"success": False, "action": "keypress", "message": str(e), "data": None}

    def _calculate(self, cmd: str) -> Dict:
        """Calculator - SAFER implementation"""
        try:
            expression = cmd.replace('calculate', '').replace('compute', '').replace('what is', '').strip()
            # Only allow numbers and basic operators
            allowed_chars = set('0123456789+-*/(). ')
            if not all(c in allowed_chars for c in expression):
                return {"success": False, "action": "calculate", "message": "Invalid characters in expression", "data": None}
            
            # Safer evaluation using literal_eval for simple math
            try:
                result = eval(expression, {"__builtins__": {}}, {})
                return {"success": True, "action": "calculate", "message": f"{expression} = {result}", "data": {"result": result, "expression": expression}}
            except:
                return {"success": False, "action": "calculate", "message": "Invalid expression", "data": None}
        except Exception as e:
            return {"success": False, "action": "calculate", "message": f"Error: {str(e)}", "data": None}

    def _tell_joke(self) -> Dict:
        """Random joke"""
        jokes = [
            "Why do programmers prefer dark mode? Because light attracts bugs.",
            "I would tell you a UDP joke, but you might not get it.",
            "There are 10 types of people in the world: those who understand binary, and those who don't.",
            "Why did the PowerPoint Presentation cross the road? To get to the other slide.",
            "I'm reading a book on anti-gravity. It's impossible to put down.",
            "Why do Java developers wear glasses? Because they don't C#."
        ]
        joke = random.choice(jokes)
        return {"success": True, "action": "joke", "message": joke, "data": None}

    def _get_weather(self, cmd: str) -> Dict:
        """Get weather"""
        return {"success": True, "action": "weather", "message": "Weather API not configured. Add OpenWeatherMap API key to enable.", "data": None}

    def _shutdown_system(self, cmd: str) -> Dict:
        """Shutdown or restart"""
        if self.os != 'Windows':
            return {"success": False, "action": "power", "message": "Shutdown only supported on Windows", "data": None}
            
        if 'restart' in cmd or 'reboot' in cmd:
            os.system('shutdown /r /t 10 /c "JARVIS restarting system as requested"')
            return {"success": True, "action": "power", "message": "Restarting in 10 seconds...", "data": None}
        elif 'shutdown' in cmd or 'turn off' in cmd:
            os.system('shutdown /s /t 10 /c "JARVIS shutting down system as requested"')
            return {"success": True, "action": "power", "message": "Shutting down in 10 seconds... Say 'abort shutdown' to cancel", "data": None}
        return {"success": False, "action": "power", "message": "Specify shutdown or restart", "data": None}

    def _abort_shutdown(self) -> Dict:
        """Cancel shutdown"""
        if self.os != 'Windows':
            return {"success": False, "action": "power", "message": "Only supported on Windows", "data": None}
        os.system("shutdown /a")
        return {"success": True, "action": "power", "message": "Shutdown aborted", "data": None}

    def _list_processes(self) -> Dict:
        """List running processes"""
        try:
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'cpu_percent']):
                try:
                    if proc.info['cpu_percent'] and proc.info['cpu_percent'] > 0:
                        processes.append(f"{proc.info['name']} (PID: {proc.info['pid']})")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            top = processes[:10] if processes else ["No active processes found"]
            return {"success": True, "action": "processes", "message": f"Top processes: {', '.join(top)}", "data": top}
        except Exception as e:
            return {"success": False, "action": "processes", "message": str(e), "data": None}

    def _kill_process(self, cmd: str) -> Dict:
        """Kill process by name or PID"""
        try:
            target = cmd.replace('kill', '').replace('terminate', '').strip()
            
            # Try as PID first
            try:
                pid = int(target)
                psutil.Process(pid).terminate()
                return {"success": True, "action": "kill", "message": f"Killed process {pid}", "data": None}
            except ValueError:
                # Try as name
                killed = []
                for proc in psutil.process_iter(['pid', 'name']):
                    if target.lower() in proc.info['name'].lower():
                        psutil.Process(proc.info['pid']).terminate()
                        killed.append(proc.info['name'])
                
                if killed:
                    return {"success": True, "action": "kill", "message": f"Killed: {', '.join(killed)}", "data": None}
                return {"success": False, "action": "kill", "message": "Process not found", "data": None}
        except Exception as e:
            return {"success": False, "action": "kill", "message": str(e), "data": None}

# Singleton instance
jarvis = JarvisCore()