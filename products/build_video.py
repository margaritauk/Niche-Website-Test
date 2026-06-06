#!/usr/bin/env python3
"""
Build a short slideshow MP4 per product from its 5 generated images.

Etsy accepts a 5-15s video (no audio needed); this makes a ~12s, 1080x1080 H.264 clip
that gently pans/fades through the 5 product images. Run build_mockups.py first.

Usage:
  python3 build_video.py            # all products
  python3 build_video.py str driver # specific variants
"""
import os, sys
import numpy as np
from PIL import Image
import imageio.v2 as imageio
from build_tracker import VARIANTS

OUT_SIZE = 1080
FPS = 30
SECONDS_PER_IMAGE = 2.4
FADE = 0.4  # seconds of crossfade between slides
IMG_ORDER = ["01-hero", "02-profit", "03-steps", "04-included", "05-summary"]


def load(path):
    im = Image.open(path).convert("RGB").resize((OUT_SIZE, OUT_SIZE), Image.LANCZOS)
    return np.asarray(im, dtype=np.float32)


def build(cfg):
    base = os.path.join(os.path.dirname(os.path.abspath(__file__)), cfg["dir"])
    idir = os.path.join(base, "images")
    frames_imgs = [load(os.path.join(idir, f"{n}.png")) for n in IMG_ORDER]
    hold = int(SECONDS_PER_IMAGE * FPS)
    fade = int(FADE * FPS)
    out_path = os.path.join(base, "preview-video.mp4")
    w = imageio.get_writer(out_path, fps=FPS, codec="libx264", quality=8,
                           macro_block_size=8, ffmpeg_log_level="error")
    for i, fr in enumerate(frames_imgs):
        # hold the slide
        for _ in range(hold):
            w.append_data(fr.astype(np.uint8))
        # crossfade into the next slide
        if i < len(frames_imgs) - 1:
            nxt = frames_imgs[i + 1]
            for f in range(fade):
                a = (f + 1) / fade
                w.append_data((fr * (1 - a) + nxt * a).astype(np.uint8))
    w.close()
    dur = (hold * len(frames_imgs) + fade * (len(frames_imgs) - 1)) / FPS
    print(f"wrote {os.path.relpath(out_path)}  (~{dur:.0f}s)")


if __name__ == "__main__":
    keys = sys.argv[1:] or list(VARIANTS)
    for k in keys:
        build(VARIANTS[k])
