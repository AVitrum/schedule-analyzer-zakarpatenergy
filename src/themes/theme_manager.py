import sys
import subprocess


class ModernTheme:
    
    @staticmethod
    def detect_system_theme():
        try:
            if sys.platform == "darwin":  # macOS
                try:
                    result = subprocess.run(
                        ['defaults', 'read', '-g', 'AppleInterfaceStyle'],
                        capture_output=True,
                        text=True,
                        timeout=1
                    )
                    return result.returncode == 0 and 'Dark' in result.stdout
                except (subprocess.TimeoutExpired, FileNotFoundError):
                    pass

            elif sys.platform == "win32":  # Windows
                import winreg
                registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
                key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                winreg.CloseKey(key)
                return value == 0

            elif sys.platform.startswith("linux"):  # Linux
                try:
                    result = subprocess.run(
                        ['gsettings', 'get', 'org.gnome.desktop.interface', 'gtk-theme'],
                        capture_output=True,
                        text=True,
                        timeout=1
                    )
                    if result.returncode == 0:
                        theme = result.stdout.lower()
                        return 'dark' in theme
                except (subprocess.TimeoutExpired, FileNotFoundError):
                    pass
        except Exception:
            pass

        return False
    
    @staticmethod
    def get_theme_colors(dark_mode=False):
        if dark_mode:
            return {
                'bg': '#1e1e1e',
                'fg': '#e0e0e0',
                'secondary_bg': '#2d2d2d',
                'input_bg': '#3c3c3c',
                'input_fg': '#ffffff',
                'button_bg': '#0e639c',
                'button_hover': '#1177bb',
                'button_fg': '#ffffff',
                'accent': '#0e639c',
                'border': '#3c3c3c',
                'text_secondary': '#a0a0a0',
                'success': '#4ec9b0',
                'error': '#f48771',
                'warning': '#ce9178'
            }
        else:
            return {
                'bg': '#f5f5f5',
                'fg': '#2c2c2c',
                'secondary_bg': '#ffffff',
                'input_bg': '#ffffff',
                'input_fg': '#000000',
                'button_bg': '#0078d4',
                'button_hover': '#106ebe',
                'button_fg': '#ffffff',
                'accent': '#0078d4',
                'border': '#d0d0d0',
                'text_secondary': '#707070',
                'success': '#107c10',
                'error': '#d13438',
                'warning': '#ca5010'
            }
