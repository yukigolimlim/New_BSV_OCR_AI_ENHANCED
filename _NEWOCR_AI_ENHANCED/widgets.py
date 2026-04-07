"""
widgets.py — DocExtract Pro
============================
Reusable custom Tkinter widgets.
  - GradientCanvas  : canvas that fills itself with a vertical colour gradient
  - Spinner         : animated arc spinner shown during processing
"""
import math
import tkinter as tk

from utils import (
    hex_blend,
    NAVY, NAVY_MID, NAVY_LIGHT, NAVY_PALE, NAVY_GHOST,
    BORDER_LIGHT, WHITE, CARD_WHITE,
)


class GradientCanvas(tk.Canvas):
    """
    A tk.Canvas that paints a smooth vertical gradient between two colours.
    The gradient redraws automatically on resize.
    """

    def __init__(self, parent, c1: str, c2: str, steps: int = 60, **kw):
        super().__init__(parent, highlightthickness=0, bd=0, **kw)
        self.c1    = c1
        self.c2    = c2
        self.steps = steps
        self.bind("<Configure>", self._draw)

    def _draw(self, e=None):
        self.delete("g")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 2 or h < 2:
            return
        sh = h / self.steps
        for i in range(self.steps):
            c = hex_blend(self.c1, self.c2, i / self.steps)
            self.create_rectangle(
                0, int(i * sh), w, int((i + 1) * sh) + 1,
                fill=c, outline="", tags="g"
            )
        self.lower("g")


class Spinner(tk.Canvas):
    """
    Animated circular spinner.
    Call .start() to begin spinning and .stop() to hide it.
    """

    def __init__(self, parent, size: int = 96, bg: str = CARD_WHITE, **kw):
        super().__init__(
            parent, width=size, height=size,
            bg=bg, highlightthickness=0, **kw
        )
        self.size  = size
        self.angle = 0
        self._job  = None

    def _draw(self):
        self.delete("all")
        cx = cy = self.size / 2
        r  = cx - 12

        # Background ring
        self.create_oval(
            cx - r, cy - r, cx + r, cy + r,
            outline=BORDER_LIGHT, width=8
        )
        # Main arc
        self.create_arc(
            cx - r, cy - r, cx + r, cy + r,
            start=self.angle % 360, extent=300,
            outline=NAVY, width=8, style="arc"
        )
        # Trailing ghost arc
        self.create_arc(
            cx - r, cy - r, cx + r, cy + r,
            start=(self.angle + 300) % 360, extent=60,
            outline=NAVY_GHOST, width=8, style="arc"
        )
        # Dot at the leading edge
        rad = math.radians(self.angle % 360)
        dx  = cx + r * math.cos(rad)
        dy  = cy - r * math.sin(rad)
        self.create_oval(dx - 7, dy - 7, dx + 7, dy + 7,
                         fill=NAVY_MID, outline=NAVY_PALE, width=2)
        self.create_oval(dx - 3, dy - 3, dx + 3, dy + 3,
                         fill=WHITE, outline="")

    def start(self):
        self._spin()

    def stop(self):
        if self._job:
            self.after_cancel(self._job)
            self._job = None
        self.delete("all")

    def _spin(self):
        self.angle += 5
        self._draw()
        self._job = self.after(14, self._spin)