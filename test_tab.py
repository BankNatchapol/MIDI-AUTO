from Quartz import CGEventCreateKeyboardEvent, CGEventPost, kCGHIDEventTap
import time

def keycode(k):  # example for '2'
    return 19  # macOS keycode for '2'

def tap_key(code, down_time=0.02):
    CGEventPost(kCGHIDEventTap, CGEventCreateKeyboardEvent(None, code, True))
    time.sleep(down_time)
    CGEventPost(kCGHIDEventTap, CGEventCreateKeyboardEvent(None, code, False))

time.sleep(5)
tap_key(keycode('2'))
tap_key(keycode('2'))