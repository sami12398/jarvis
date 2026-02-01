import os
import subprocess
import psutil
import pyautogui
import pyperclip
import win32gui
import win32con
import win32api
import ctypes
import json
import webbrowser
import re
import random
import urllib.parse
import time
from datetime import datetime
from typing import Dict, Any, Optional

# Constants
SC_MINIMIZEALL = 0xF032  # Windows minimize all constant

class JarvisCore:
    def __init__(self):
        self.os = 'Windows'
        self.last_result = None
        self.window_list = []
        
    def process_command(self, command: str) -> Dict[str, Any]:
        """Main command processor"""
        cmd = command.lower().strip()
        result = {"success": False, "action": None, "message": "Command not recognized", "data": None}
        
        try:
            # Command routing - ordered by priority (most specific first)
            
            # Browser & YouTube (high priority)
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
            
            # Power commands (MUST be before open/start check since 'restart' contains 'start')
            elif any(x in cmd for x in ['shutdown', 'restart', 'reboot', 'turn off']) and 'abort' not in cmd:
                result = self._shutdown_system(cmd)
            elif 'abort' in cmd or 'cancel shutdown' in cmd:
                result = self._shutdown_system(cmd)
            
            # System & Window Management
            elif any(x in cmd for x in ['open', 'start', 'launch']) and 'restart' not in cmd:
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
            
            # File Operations
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
            
            # Utilities
            elif any(x in cmd for x in ['calculate', 'compute']) or (any(x in cmd for x in ['what is', 'how much']) and any(c in cmd for c in '0123456789+-*/')):
                result = self._calculate(cmd)
            elif 'joke' in cmd:
                result = self._tell_joke()
            elif 'weather' in cmd:
                result = self._get_weather(cmd)
            elif 'list processes' in cmd or 'running apps' in cmd:
                result = self._list_processes()
            elif cmd.startswith('terminate') or (cmd.startswith('kill') and 'kill' in cmd and 'close' not in cmd):
                result = self._kill_process(cmd)
            
            # Info & Clipboard
            elif 'time' in cmd:
                result = self._get_time()
            elif 'ip' in cmd or 'network' in cmd:
                result = self._get_network()
            elif 'system info' in cmd:
                result = self._get_system_info()
            elif 'copy' in cmd and 'copy' == cmd.split()[0]:
                result = self._handle_clipboard(cmd, 'copy')
            elif 'paste' in cmd:
                result = self._handle_clipboard(cmd, 'paste')
            elif 'find' in cmd or 'search file' in cmd:
                result = self._search_files(cmd)
            
            # Audio & Power
            elif 'volume' in cmd or 'sound' in cmd:
                result = self._handle_volume(cmd)
            elif 'screenshot' in cmd:
                result = self._take_screenshot()
            elif 'sleep' in cmd or 'standby' in cmd:
                result = self._power_control('sleep')
            
            # Greetings
            elif cmd in ['hello', 'hi', 'hey']:
                result = {"success": True, "action": "greeting", "message": "Hello sir. Systems operational.", "data": None}
            else:
                result = {"success": False, "action": "unknown", "message": f"I don't understand: '{cmd}'", "data": None}
                
        except Exception as e:
            result = {"success": False, "action": "error", "message": str(e), "data": None}
            
        self.last_result = result
        return result
    
    # ==================== BROWSER & YOUTUBE METHODS ====================
    
    def _handle_youtube(self, cmd: str) -> Dict:
        """YouTube automation - search and play videos"""
        query = None
        
        # Extract search query from various patterns
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
            
            # Determine if user wants to play or just search
            if 'play' in cmd or 'watch' in cmd:
                # Direct search URL (first result will be playable)
                search_url = f"https://www.youtube.com/results?search_query={encoded_query}"
                webbrowser.open(search_url)
                
                return {
                    "success": True, 
                    "action": "youtube", 
                    "message": f"Opening YouTube search for: '{query}'. Click the first video to play.", 
                    "data": {"query": query, "url": search_url, "action": "search_play"}
                }
            else:
                # Just search
                search_url = f"https://www.youtube.com/results?search_query={encoded_query}"
                webbrowser.open(search_url)
                return {
                    "success": True, 
                    "action": "youtube", 
                    "message": f"YouTube search: '{query}'", 
                    "data": {"query": query, "url": search_url}
                }
        else:
            # Just open YouTube homepage
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
        
        # Common website shortcuts
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
        
        # Check if it's a known site
        site_lower = site.lower()
        if site_lower in sites:
            webbrowser.open(sites[site_lower])
            return {"success": True, "action": "website", "message": f"Opening {site}", "data": {"site": site, "url": sites[site_lower]}}
        
        # Check if it's a direct URL
        elif site.startswith('http://') or site.startswith('https://'):
            webbrowser.open(site)
            return {"success": True, "action": "website", "message": f"Opening {site}", "data": {"url": site}}
        
        # Check if it looks like a domain (contains .com, .org, etc.)
        elif any(ext in site for ext in ['.com', '.org', '.net', '.io', '.co', '.ai']):
            url = f"https://{site}" if not site.startswith('www.') else f"https://{site}"
            webbrowser.open(url)
            return {"success": True, "action": "website", "message": f"Opening {site}", "data": {"url": url}}
        
        # Otherwise, search for it
        else:
            return self._handle_google_search(site)

    # ==================== EXISTING METHODS (unchanged) ====================
    
    def _handle_open(self, cmd: str) -> Dict:
        """Handle application opening"""
        apps = {
            # Text Editors & Utilities
            'notepad': 'notepad.exe',
            'calc': 'calc.exe',
            'calculator': 'calc.exe',
            'paint': 'mspaint.exe',
            'snipping tool': 'snippingtool.exe',
            'snip': 'snippingtool.exe',
            
            # Browsers
            'chrome': 'chrome.exe',
            'google chrome': 'chrome.exe',
            'browser': 'chrome.exe',
            'edge': 'msedge.exe',
            'microsoft edge': 'msedge.exe',
            'firefox': 'firefox.exe',
            'brave': 'brave.exe',
            'opera': 'opera.exe',
            
            # System Tools
            'cmd': 'cmd.exe',
            'command prompt': 'cmd.exe',
            'terminal': 'wt.exe',
            'windows terminal': 'wt.exe',
            'powershell': 'powershell.exe',
            'task manager': 'taskmgr.exe',
            'taskmgr': 'taskmgr.exe',
            'settings': 'ms-settings:',
            'control panel': 'control.exe',
            'control': 'control.exe',
            'device manager': 'devmgmt.msc',
            'disk management': 'diskmgmt.msc',
            'services': 'services.msc',
            'registry': 'regedit.exe',
            'regedit': 'regedit.exe',
            
            # File Explorer - Multiple variations for better matching
            'explorer': 'explorer.exe',
            'file explorer': 'explorer.exe',
            'files': 'explorer.exe',
            'windows explorer': 'explorer.exe',
            'my computer': 'explorer.exe',
            'this pc': 'explorer.exe',
            'folder': 'explorer.exe',
            'folders': 'explorer.exe',
            
            # Microsoft Office
            'word': 'winword.exe',
            'microsoft word': 'winword.exe',
            'excel': 'excel.exe',
            'microsoft excel': 'excel.exe',
            'powerpoint': 'powerpnt.exe',
            'outlook': 'outlook.exe',
            'onenote': 'onenote.exe',
            
            # Code Editors & Development
            'vscode': 'code.exe',
            'vs code': 'code.exe',
            'visual studio code': 'code.exe',
            'visual studio': 'devenv.exe',
            'sublime': 'sublime_text.exe',
            'sublime text': 'sublime_text.exe',
            'notepad++': 'notepad++.exe',
            'notepad plus plus': 'notepad++.exe',
            
            # Communication Apps
            'discord': 'discord.exe',
            'telegram': 'telegram.exe',
            'telegram desktop': 'telegram.exe',
            'slack': 'slack.exe',
            'zoom': 'zoom.exe',
            'teams': 'ms-teams.exe',
            'microsoft teams': 'ms-teams.exe',
            'skype': 'skype.exe',
            
            # Media & Entertainment
            'spotify': 'spotify.exe',
            'vlc': 'vlc.exe',
            'vlc player': 'vlc.exe',
            'media player': 'wmplayer.exe',
            'windows media player': 'wmplayer.exe',
            'movies': 'mswindowsvideo:',
            'photos': 'ms-photos:',
            'groove': 'mswindowsmusic:',
            'itunes': 'itunes.exe',
            'obs': 'obs64.exe',
            'obs studio': 'obs64.exe',
            
            # Productivity
            'notion': 'notion.exe',
            'obsidian': 'obsidian.exe',
            
            # Cloud Storage
            'onedrive': 'onedrive.exe',
            'dropbox': 'dropbox.exe',
            
            # Graphics & Design
            'photoshop': 'photoshop.exe',
            'figma': 'figma.exe',
            'blender': 'blender.exe',
            'gimp': 'gimp.exe',
            
            # File Compression
            'winrar': 'winrar.exe',
            '7zip': '7zFM.exe',
            '7-zip': '7zFM.exe',
            
            # Remote Desktop
            'remote desktop': 'mstsc.exe',
            'anydesk': 'anydesk.exe',
            'teamviewer': 'teamviewer.exe'
        }
        
        # Extract app name
        app_name = cmd
        for keyword in ['open', 'start', 'launch']:
            if cmd.startswith(keyword):
                app_name = cmd.replace(keyword, '').strip()
                break
            
        # Don't intercept browser commands if they're for web
        if app_name in ['chrome', 'browser', 'edge', 'firefox'] and ('search' in cmd or 'youtube' in cmd or 'go to' in cmd):
            # Let browser commands fall through to web search if they contain search terms
            pass
            
        # Find match
        if app_name in apps:
            exe = apps[app_name]
            
            # Handle different launch types
            if exe.startswith('ms-') or exe.startswith('mswindows'):
                # Windows Store app / protocol
                os.system(f'start "" "{exe}"')
                return {"success": True, "action": "open", "message": f"Opened {app_name}", "data": {"app": app_name}}
            elif exe.endswith('.msc'):
                # Microsoft Management Console snap-in
                os.system(f'start mmc {exe}')
                return {"success": True, "action": "open", "message": f"Opened {app_name}", "data": {"app": app_name}}
            else:
                # Try as executable name in PATH
                subprocess.Popen(exe, shell=True)
                return {"success": True, "action": "open", "message": f"Opened {app_name}", "data": {"app": app_name}}
        else:
            # Try direct execution
            try:
                subprocess.Popen(app_name, shell=True)
                return {"success": True, "action": "open", "message": f"Executing {app_name}", "data": None}
            except:
                return {"success": False, "action": "open", "message": f"Application '{app_name}' not found", "data": None}
    
    def _handle_close(self, cmd: str) -> Dict:
        """Close applications"""
        app_name = cmd.replace('close', '').replace('kill', '').replace('exit', '').strip()
        
        closed = []
        for proc in psutil.process_iter(['pid', 'name']):
            if app_name.lower() in proc.info['name'].lower():
                try:
                    psutil.Process(proc.info['pid']).terminate()
                    closed.append(proc.info['name'])
                except:
                    pass
        
        if closed:
            return {"success": True, "action": "close", "message": f"Closed {', '.join(closed)}", "data": closed}
        return {"success": False, "action": "close", "message": f"No process matching '{app_name}' found", "data": None}
    
    def _minimize_all(self) -> Dict:
        """Minimize all windows using Win+M keyboard shortcut (more reliable)"""
        try:
            pyautogui.keyDown('win')
            pyautogui.keyDown('m')
            pyautogui.keyUp('m')
            pyautogui.keyUp('win')
            return {"success": True, "action": "minimize_all", "message": "All windows minimized", "data": None}
        except Exception as e:
            # Fallback to Windows API method if pyautogui fails
            try:
                win32gui.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, 
                                   SC_MINIMIZEALL, 0)
                return {"success": True, "action": "minimize_all", "message": "All windows minimized", "data": None}
            except:
                return {"success": False, "action": "minimize_all", "message": str(e), "data": None}
    
    def _handle_snap(self, cmd: str) -> Dict:
        """Window snapping"""
        direction = 'left' if 'left' in cmd else 'right'
        window_name = None
        
        # Extract window name
        parts = cmd.replace('snap', '').strip().split()
        for part in parts:
            if part not in ['left', 'right', 'window']:
                window_name = part
                break
        
        # Get window handle
        if not window_name:
            hwnd = win32gui.GetForegroundWindow()
            window_name = "active window"
        else:
            hwnd = win32gui.FindWindow(None, window_name)
            if not hwnd:
                # Try partial match
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
        
        # Perform snap
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)
        
        if direction == 'left':
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 
                                screen_width//2, screen_height, 0)
        else:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 
                                screen_width//2, 0, screen_width//2, screen_height, 0)
        
        return {"success": True, "action": "snap", "message": f"Snapped {window_name} to {direction}", "data": {"direction": direction, "window": window_name}}
    
    def _get_time(self) -> Dict:
        """Get current time"""
        now = datetime.now()
        time_str = now.strftime("%I:%M %p")
        date_str = now.strftime("%A, %B %d, %Y")
        return {"success": True, "action": "time", "message": f"It is {time_str} on {date_str}", "data": {"time": time_str, "date": date_str}}
    
    def _get_network(self) -> Dict:
        """Get IP info"""
        try:
            result = subprocess.run(['ipconfig'], capture_output=True, text=True)
            lines = result.stdout.split('\n')
            ips = [line.split(':')[1].strip() for line in lines if 'IPv4' in line and ':' in line]
            return {"success": True, "action": "network", "message": f"IP Addresses: {', '.join(ips[:2])}", "data": {"ips": ips}}
        except Exception as e:
            return {"success": False, "action": "network", "message": str(e), "data": None}
    
    def _get_system_info(self) -> Dict:
        """Get system stats"""
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
    
    def _handle_clipboard(self, cmd: str, action: str) -> Dict:
        """Clipboard operations"""
        if action == 'copy':
            text = cmd.replace('copy', '').strip()
            pyperclip.copy(text)
            return {"success": True, "action": "clipboard", "message": f"Copied: {text}", "data": {"text": text}}
        else:
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
                for item in os.listdir(path):
                    if query.lower() in item.lower():
                        matches.append(os.path.join(path, item))
        
        if matches:
            return {"success": True, "action": "search", "message": f"Found {len(matches)} matches: {', '.join([os.path.basename(m) for m in matches[:3]])}", "data": matches[:5]}
        return {"success": True, "action": "search", "message": "No files found", "data": []}
    
    def _handle_volume(self, cmd: str) -> Dict:
        """Volume control using keyboard shortcuts"""
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
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"jarvis_screenshot_{timestamp}.png"
            filepath = os.path.join(os.path.expanduser('~\\Pictures'), filename)
            pyautogui.screenshot(filepath)
            return {"success": True, "action": "screenshot", "message": f"Screenshot saved: {filename}", "data": {"path": filepath}}
        except Exception as e:
            return {"success": False, "action": "screenshot", "message": str(e), "data": None}
    
    def _power_control(self, action: str) -> Dict:
        """Power controls (sleep, shutdown)"""
        if action == 'sleep':
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
            return {"success": True, "action": "power", "message": "System sleeping", "data": None}
        return {"success": False, "action": "power", "message": "Unknown power command", "data": None}

    # ==================== NEW METHODS ====================
    
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
                num = int(re.search(r'\d+', cmd).group())
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
        try:
            ctypes.windll.user32.LockWorkStation()
            return {"success": True, "action": "lock", "message": "System locked", "data": None}
        except Exception as e:
            return {"success": False, "action": "lock", "message": str(e), "data": None}

    def _empty_recycle_bin(self) -> Dict:
        """Empty recycle bin"""
        try:
            import winshell
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=False)
            return {"success": True, "action": "recycle", "message": "Recycle bin emptied", "data": None}
        except:
            os.system("rd /s /q %systemdrive%\\$Recycle.Bin 2>nul")
            return {"success": True, "action": "recycle", "message": "Recycle bin emptied", "data": None}

    def _handle_window_actions(self, cmd: str) -> Dict:
        """Advanced window actions"""
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
            return {"success": True, "action": "(folder", "message": f"Created folder: {folder_name}", "data": {"path": path}}
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
        try:
            text = cmd.replace('type', '').replace('write', '').strip()
            pyautogui.typewrite(text, interval=0.01)
            return {"success": True, "action": "type", "message": f"Typed: {text}", "data": None}
        except Exception as e:
            return {"success": False, "action": "type", "message": str(e), "data": None}

    def _press_key(self, cmd: str) -> Dict:
        """Press specific keys"""
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
        """Calculator"""
        try:
            expression = cmd.replace('calculate', '').replace('compute', '').replace('what is', '').strip()
            expression = re.sub(r'[^0-9+\-*/(). ]', '', expression)
            result = eval(expression)
            return {"success": True, "action": "calculate", "message": f"{expression} = {result}", "data": {"result": result, "expression": expression}}
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
        try:
            if 'abort' in cmd or 'cancel' in cmd:
                result = subprocess.run(['shutdown', '/a'], capture_output=True, text=True)
                return {"success": True, "action": "power", "message": "Shutdown aborted", "data": None}
            elif 'restart' in cmd or 'reboot' in cmd:
                result = subprocess.run(['shutdown', '/r', '/t', '10', '/c', 'JARVIS restarting system'], capture_output=True, text=True)
                if result.returncode == 0:
                    return {"success": True, "action": "power", "message": "Restarting in 10 seconds... Say 'abort shutdown' to cancel", "data": None}
                else:
                    return {"success": False, "action": "power", "message": f"Restart failed: {result.stderr}", "data": None}
            elif 'shutdown' in cmd or 'turn off' in cmd:
                result = subprocess.run(['shutdown', '/s', '/t', '10', '/c', 'JARVIS shutting down system'], capture_output=True, text=True)
                if result.returncode == 0:
                    return {"success": True, "action": "power", "message": "Shutting down in 10 seconds... Say 'abort shutdown' to cancel", "data": None}
                else:
                    return {"success": False, "action": "power", "message": f"Shutdown failed: {result.stderr}", "data": None}
            return {"success": False, "action": "power", "message": "Specify shutdown, restart, or abort", "data": None}
        except Exception as e:
            return {"success": False, "action": "power", "message": f"Power command failed: {str(e)}", "data": None}

    def _list_processes(self) -> Dict:
        """List running processes"""
        try:
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'cpu_percent']):
                try:
                    if proc.info['cpu_percent'] and proc.info['cpu_percent'] > 0:
                        processes.append(f"{proc.info['name']} (PID: {proc.info['pid']})")
                except:
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