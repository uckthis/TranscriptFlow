import subprocess
import json
import threading
import time
from PyQt6.QtCore import QObject, pyqtSignal

class USBManager:
    @staticmethod
    def list_hid_devices():
        """Uses PowerShell to list HID class devices"""
        try:
            cmd = "Get-PnpDevice -Class HIDClass | Select-Object FriendlyName, InstanceId, Status | ConvertTo-Json"
            result = subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                devices = json.loads(result.stdout)
                # handle single vs list return from PS JSON
                if isinstance(devices, dict):
                    devices = [devices]
                return devices
            return []
        except Exception as e:
            print(f"Hardware Error: {e}")
            return []

class FootPedalManager(QObject):
    pedalPressed = pyqtSignal(int)
    pedalReleased = pyqtSignal(int)
    
    def __init__(self, config=None):
        super().__init__()
        self.config = config or {}
        self.running = False
        self.thread = None
        self.device = None
        self.last_state = {} # button_id -> state

    def start(self, device_id=None):
        if self.running: self.stop()
        
        target_id = device_id or self.config.get('hardware', {}).get('pedal_id')
        if not target_id: return False

        # Note: Actual HID reading requires 'hidapi' library.
        # This implementation provides the loop and signal structure.
        self.running = True
        self.thread = threading.Thread(target=self._monitor_loop, args=(target_id,), daemon=True)
        self.thread.start()
        return True

    def stop(self):
        self.running = False
        if self.thread:
            self.thread.join(timeout=1.0)
            self.thread = None

    def _monitor_loop(self, device_id):
        """Background thread to poll or wait for HID events"""
        # Placeholder for real HID library integration
        # In a real scenario, we'd use hid.device() here
        print(f"Monitoring Foot Pedal: {device_id}")
        
        while self.running:
            # Poll every 50ms for low latency
            time.sleep(0.05)
            # Logic here would read from hidapi
            # if data: handle_data(data)
            pass
