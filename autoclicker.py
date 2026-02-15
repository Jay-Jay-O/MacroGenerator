import ctypes
import time

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

class AutoClicker:
    # Mouse Event Constants
    MOUSEEVENTF_LEFTDOWN = 0x0002
    MOUSEEVENTF_LEFTUP = 0x0004
    MOUSEEVENTF_RIGHTDOWN = 0x0008
    MOUSEEVENTF_RIGHTUP = 0x0010
    MOUSEEVENTF_MIDDLEDOWN = 0x0020
    MOUSEEVENTF_MIDDLEUP = 0x0040
    MOUSEEVENTF_WHEEL = 0x0800

    # Key Event Constants
    KEYEVENTF_KEYUP = 0x0002

    # Button Mappings
    BUTTON_LEFT = 1
    BUTTON_RIGHT = 2
    BUTTON_MIDDLE = 3

    def __init__(self):
        self.delay_ms = 300.0
        self.x = 0
        self.y = 0
        # Access Windows User32 API
        self.user32 = ctypes.windll.user32

    def getCursorPos(self):
        """Returns the current (x, y) tuple of the mouse cursor."""
        pt = POINT()
        self.user32.GetCursorPos(ctypes.byref(pt))
        return pt.x, pt.y

    def isKeyPressed(self, vk_code):
        """Returns True if the key is currently pressed."""
        # GetAsyncKeyState returns header bit set if pressed
        return (self.user32.GetAsyncKeyState(vk_code) & 0x8000) != 0

    def getKeyName(self, vk_code):
        """Returns the human-readable name of the key."""
        # Fix for some special keys not mapping nicely by default
        special_keys = {
            0x08: "BACKSPACE", 0x09: "TAB", 0x0D: "ENTER", 0x10: "SHIFT",
            0x11: "CTRL", 0x12: "ALT", 0x13: "PAUSE", 0x14: "CAPS LOCK",
            0x1B: "ESC", 0x20: "SPACE", 0x21: "PAGE UP", 0x22: "PAGE DOWN",
            0x23: "END", 0x24: "HOME", 0x25: "LEFT", 0x26: "UP",
            0x27: "RIGHT", 0x28: "DOWN", 0x2C: "PRINT SCREEN", 0x2D: "INSERT",
            0x2E: "DELETE", 0x5B: "LWIN", 0x5C: "RWIN", 0x5D: "APPS",
            0xA0: "LSHIFT", 0xA1: "RSHIFT", 0xA2: "LCTRL", 0xA3: "RCTRL",
            0xA4: "LALT", 0xA5: "RALT"
        }
        
        if vk_code in special_keys:
            return special_keys[vk_code]
            
        # Try to map to character
        # MapVirtualKeyW: uCode, uMapType (2 = MAPVK_VK_TO_CHAR)
        scan_code = self.user32.MapVirtualKeyW(vk_code, 2)
        if scan_code > 0:
            char = chr(scan_code)
            # Filter non-printable
            if char.isprintable():
                return char.upper()
        
        return f"Unknown ({vk_code})"

    def setDelay(self, ms):
        """Sets the delay after clicking in milliseconds."""
        self.delay_ms = float(ms)

    def mouseMove(self, x, y):
        """Sets the target X and Y coordinates."""
        self.x = int(x)
        self.y = int(y)

    def clickMouse(self, button, action_delay_ms, next_action_delay_ms):
        """
        Moves to stored coordinates, clicks the specified button,
        waits 250ms, releases, and waits for the configured delay.
        """
        # Move actual cursor to self.x, self.y
        self.user32.SetCursorPos(self.x, self.y)

        down_event = 0
        up_event = 0

        if button == self.BUTTON_LEFT:
            down_event = self.MOUSEEVENTF_LEFTDOWN
            up_event = self.MOUSEEVENTF_LEFTUP
        elif button == self.BUTTON_RIGHT:
            down_event = self.MOUSEEVENTF_RIGHTDOWN
            up_event = self.MOUSEEVENTF_RIGHTUP
        elif button == self.BUTTON_MIDDLE:
            down_event = self.MOUSEEVENTF_MIDDLEDOWN
            up_event = self.MOUSEEVENTF_MIDDLEUP
        
        if down_event != 0:
            # Press
            self.user32.mouse_event(down_event, 0, 0, 0, 0)
            time.sleep(action_delay_ms / 1000.0)
            # Release
            self.user32.mouse_event(up_event, 0, 0, 0, 0)
            # Wait for the configured delay
            time.sleep(next_action_delay_ms / 1000.0)

    def keyPress(self, vk_code, delay_ms):
        """
        Simulates a key press and release.
        vk_code: Virtual Key Code (e.g., 0x41 for 'A', 0x0D for Enter)
        """
        self.user32.keybd_event(vk_code, 0, 0, 0)
        time.sleep(delay_ms / 1000.0)
        self.user32.keybd_event(vk_code, 0, self.KEYEVENTF_KEYUP, 0)

    def keyDown(self, vk_code):
        """Presses a key down."""
        self.user32.keybd_event(vk_code, 0, 0, 0)

    def keyUp(self, vk_code):
        """Releases a key."""
        self.user32.keybd_event(vk_code, 0, self.KEYEVENTF_KEYUP, 0)
    
    def keyHold(self, vk_code, duration_ms):
        """Presses a key, waits, and releases it."""
        self.keyDown(vk_code)
        time.sleep(duration_ms / 1000.0)
        self.keyUp(vk_code)

    def mouseDown(self, button):
        """Presses a mouse button down."""
        if button == self.BUTTON_LEFT:
            self.user32.mouse_event(self.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        elif button == self.BUTTON_RIGHT:
            self.user32.mouse_event(self.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
        elif button == self.BUTTON_MIDDLE:
            self.user32.mouse_event(self.MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)

    def mouseUp(self, button):
        """Releases a mouse button."""
        if button == self.BUTTON_LEFT:
            self.user32.mouse_event(self.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        elif button == self.BUTTON_RIGHT:
            self.user32.mouse_event(self.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
        elif button == self.BUTTON_MIDDLE:
            self.user32.mouse_event(self.MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)

    def mouseHold(self, button, duration_ms):
        """Presses a mouse button, waits, and releases it."""
        self.mouseDown(button)
        time.sleep(duration_ms / 1000.0)
        self.mouseUp(button)
        
    def mouseDrag(self, x1, y1, x2, y2, button, duration_ms):
        """Drag from (x1, y1) to (x2, y2)."""
        self.mouseMove(x1, y1)
        # FORCE move to start position
        self.user32.SetCursorPos(x1, y1)
        
        # Initial press and hold
        self.mouseDown(button)
        time.sleep(0.1) 
        
        # Move to end point
        self.mouseMove(x2, y2)
        # Apply actual cursor movement
        self.user32.SetCursorPos(x2, y2)
        
        # Hold at end point
        if duration_ms > 0:
            time.sleep(duration_ms / 1000.0)
            
        self.mouseUp(button)
        
    def shortcut(self, vk_list, duration_ms):
        """Presses multiple keys, holds, then releases in reverse."""
        for vk in vk_list:
            self.keyDown(vk)
        
        if duration_ms > 0:
            time.sleep(duration_ms / 1000.0)
            
        for vk in reversed(vk_list):
            self.keyUp(vk)
    
    def mouseScroll(self, delta):
        """
        Scrolls the mouse wheel.
        Positive delta scrolls up, negative scrolls down.
        Standard delta is ±120 per notch.
        """
        self.user32.mouse_event(self.MOUSEEVENTF_WHEEL, 0, 0, delta, 0)
