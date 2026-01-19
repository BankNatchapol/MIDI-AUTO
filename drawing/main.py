import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import threading
import time

from Quartz.CoreGraphics import (
    CGEventCreateMouseEvent, CGEventPost,
    kCGHIDEventTap, kCGEventLeftMouseDown, kCGEventLeftMouseUp
)
import pyautogui

STOP_FLAG = False
DRAW_THREAD = None
screen_w, screen_h = pyautogui.size()

def quartz_click(x, y):
    """
    Low-level click for macOS (Quartz) so clicks register in games/editors.
    """
    ev_down = CGEventCreateMouseEvent(None, kCGEventLeftMouseDown, (x, y), 0)
    CGEventPost(kCGHIDEventTap, ev_down)
    time.sleep(0.01)
    ev_up = CGEventCreateMouseEvent(None, kCGEventLeftMouseUp, (x, y), 0)
    CGEventPost(kCGHIDEventTap, ev_up)

def pixelate_and_threshold(img, res, thresh):
    """
    Resize the image to (res x res) and threshold to B/W.
    """
    small = img.resize((res, res), Image.BILINEAR)
    gray = small.convert("L")
    bw = gray.point(lambda x: 0 if x < thresh else 255, "1")
    return bw.convert("RGB")

class PixelArtGridDrawer:
    def __init__(self, master):
        self.master = master
        master.title("Pixel Art Drawer — Real Grid")

        self.original_img = None
        self.preview_img = None
        self.tk_preview = None

        # grid corners & cell size
        self.tl_x = None
        self.tl_y = None
        self.br_x = None
        self.br_y = None
        self.cell_w = None
        self.cell_h = None

        self.center_x = None
        self.center_y = None

        # load
        tk.Button(master, text="Load Image", command=self.load_image).pack(pady=4)

        # preview
        self.canvas = tk.Canvas(master, width=360, height=360, bg="white")
        self.canvas.pack()

        self.status_lbl = tk.Label(master, text="Load an image to begin")
        self.status_lbl.pack()

        # parameters
        param_frame = tk.Frame(master)
        param_frame.pack(pady=6)

        tk.Label(param_frame, text="Pixel Resolution:").grid(row=0, column=0, sticky="e")
        self.res_entry = tk.Entry(param_frame, width=6)
        self.res_entry.grid(row=0, column=1)
        self.res_entry.insert(0, "30")

        tk.Label(param_frame, text="Threshold:").grid(row=1, column=0, sticky="e")
        self.thresh_entry = tk.Entry(param_frame, width=6)
        self.thresh_entry.grid(row=1, column=1)
        self.thresh_entry.insert(0, "128")

        tk.Label(param_frame, text="Delay (ms):").grid(row=2, column=0, sticky="e")
        self.delay_entry = tk.Entry(param_frame, width=6)
        self.delay_entry.grid(row=2, column=1)
        self.delay_entry.insert(0, "0.0")

        tk.Button(param_frame, text="Preview", command=self.update_preview).grid(row=0, column=2, rowspan=3, padx=8)

        # grid corners
        self.btn_tl = tk.Button(master, text="Click Top-Left Grid Corner", command=self.set_tl, state="disabled")
        self.btn_tl.pack(pady=2)

        self.btn_br = tk.Button(master, text="Click Bottom-Right Grid Corner", command=self.set_br, state="disabled")
        self.btn_br.pack(pady=2)

        # center selection
        self.btn_center = tk.Button(master, text="Set Art Center (move mouse → ENTER)",
                                    command=self.capture_center, state="disabled")
        self.btn_center.pack(pady=2)

        # draw & stop
        self.draw_btn = tk.Button(master, text="Start Drawing", command=self.start_drawing, state="disabled")
        self.draw_btn.pack(pady=5)

        self.stop_btn = tk.Button(master, text="Stop Drawing", command=self.stop_drawing, state="disabled")
        self.stop_btn.pack(pady=5)

        master.bind("<Return>", self.on_enter)

    def load_image(self):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg"), ("All", "*.*")])
        if not path:
            return
        try:
            self.original_img = Image.open(path).convert("RGB")
        except Exception as e:
            messagebox.showerror("Error", f"Cannot load image: {e}")
            return
        self.status_lbl.config(text="Image loaded — enter params and press Preview.")
        self.preview_img = None
        self.btn_tl.config(state="disabled")
        self.btn_br.config(state="disabled")
        self.btn_center.config(state="disabled")
        self.draw_btn.config(state="disabled")

    def update_preview(self):
        if self.original_img is None:
            return

        try:
            res = int(self.res_entry.get())
            thresh = int(self.thresh_entry.get())
        except ValueError:
            messagebox.showerror("Invalid input", "Resolution & threshold must be integers!")
            return

        # pixelate & threshold
        self.preview_img = pixelate_and_threshold(self.original_img, res, thresh)
        disp = self.preview_img.resize((360, 360), Image.NEAREST)
        self.tk_preview = ImageTk.PhotoImage(disp)
        self.canvas.create_image(0,0,anchor="nw", image=self.tk_preview)

        self.status_lbl.config(text="Preview ready — define grid corners.")
        self.btn_tl.config(state="normal")
        self.btn_br.config(state="normal")
        self.btn_center.config(state="disabled")
        self.draw_btn.config(state="disabled")

    def set_tl(self):
        self.status_lbl.config(text="Move mouse to **top-left grid cell**, then press ENTER")
        self.stage = "tl"

    def set_br(self):
        self.status_lbl.config(text="Move mouse to **bottom-right grid cell**, then press ENTER")
        self.stage = "br"

    def capture_center(self):
        self.status_lbl.config(text="Move mouse to art center, then press ENTER")
        self.stage = "center"

    def on_enter(self, event):
        pos = pyautogui.position()

        if not hasattr(self, "stage") or self.preview_img is None:
            return

        if self.stage == "tl":
            self.tl_x, self.tl_y = pos.x, pos.y
            self.status_lbl.config(text=f"Top-Left grid corner set at ({self.tl_x},{self.tl_y}) — set bottom-right.")
            self.stage = None

        elif self.stage == "br":
            self.br_x, self.br_y = pos.x, pos.y
            self.status_lbl.config(text=f"Bottom-Right grid corner set at ({self.br_x},{self.br_y}).")
            self.stage = None

            # calculate grid spacing
            dx = abs(self.br_x - self.tl_x)
            dy = abs(self.br_y - self.tl_y)
            try:
                res = int(self.res_entry.get())
                # (res-1) spans number of cell intervals
                self.cell_w = dx / (res - 1)
                self.cell_h = dy / (res - 1)
                self.status_lbl.config(
                    text=f"Cell size calculated: {self.cell_w:.2f} x {self.cell_h:.2f} — now set center."
                )
                self.btn_center.config(state="normal")
            except Exception:
                messagebox.showerror("Error", "Invalid resolution for grid calculation.")

        elif self.stage == "center":
            self.center_x, self.center_y = pos.x, pos.y
            self.status_lbl.config(
                text=f"Art center set at ({self.center_x},{self.center_y}) — ready to draw."
            )
            self.stage = None
            self.draw_btn.config(state="normal")

    def start_drawing(self):
        global STOP_FLAG, DRAW_THREAD
        STOP_FLAG = False
        self.stop_btn.config(state="normal")
        self.draw_btn.config(state="disabled")
        self.status_lbl.config(text="Drawing…")

        DRAW_THREAD = threading.Thread(target=self.draw_loop)
        DRAW_THREAD.start()

    def stop_drawing(self):
        global STOP_FLAG
        STOP_FLAG = True
        self.status_lbl.config(text="Stopping…")

    def draw_loop(self):
        global STOP_FLAG

        try:
            res = int(self.res_entry.get())
            delay_s = float(self.delay_entry.get()) / 1000.0
        except ValueError:
            messagebox.showerror("Error", "Invalid numeric input for drawing.")
            return

        pixels = self.preview_img.load()

        # origin = top-left based on center
        origin_x = self.center_x - self.cell_w * (res - 1)/2
        origin_y = self.center_y - self.cell_h * (res - 1)/2

        time.sleep(0.1)
        count = 0
        for y in range(res):
            if STOP_FLAG:
                break
            for x in range(res):
                if STOP_FLAG:
                    break
                if pixels[x, y] == (0, 0, 0):
                    tx = origin_x + x*self.cell_w
                    ty = origin_y + y*self.cell_h
                    tx, ty = int(tx), int(ty)
                    if 0 <= tx < screen_w and 0 <= ty < screen_h:
                        quartz_click(tx, ty)
                        count += 1
                        if delay_s > 0:
                            time.sleep(delay_s)

        self.status_lbl.config(
            text=f"Finished {count} clicks." if not STOP_FLAG else "Drawing stopped."
        )
        self.stop_btn.config(state="disabled")
        self.draw_btn.config(state="normal")

print("[APP] Real Canvas Grid Pixel Drawer")
root = tk.Tk()
app = PixelArtGridDrawer(root)
root.mainloop()
