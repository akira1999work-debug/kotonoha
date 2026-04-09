"""マイクアイコン (.ico + .png) を生成する
Cairo で glow_preview と同じ紫フラットデザインを描画
"""
import math
import ctypes
import numpy as np
from pathlib import Path
from PIL import Image
import cairo


# 色パレット (voice_input.py の glow_preview と統一)
PURPLE_MAIN = (140, 80, 220)
PURPLE_DEEP = (88, 30, 180)
RED_MAIN = (235, 55, 70)


def cairo_surface_to_pil(surface: cairo.ImageSurface) -> Image.Image:
    w = surface.get_width()
    h = surface.get_height()
    data = bytes(surface.get_data())
    arr = np.frombuffer(data, dtype=np.uint8).reshape(h, w, 4)
    rgba = arr[:, :, [2, 1, 0, 3]].copy()  # BGRA -> RGBA
    alpha = rgba[:, :, 3:4].astype(np.float32)
    with np.errstate(divide="ignore", invalid="ignore"):
        rgba[:, :, :3] = np.where(
            alpha > 0,
            np.clip(rgba[:, :, :3].astype(np.float32) * 255.0 / alpha, 0, 255),
            0,
        ).astype(np.uint8)
    return Image.fromarray(rgba, "RGBA")


def draw_mic(size: int, mode: str = "active") -> Image.Image:
    """
    mode: 'active' (紫) | 'recording' (赤 + グロー) | 'off' (灰色)
    """
    surface = cairo.ImageSurface(cairo.FORMAT_ARGB32, size, size)
    ctx = cairo.Context(surface)
    ctx.set_antialias(cairo.ANTIALIAS_BEST)

    cx = size / 2
    cy = size / 2

    if mode == "recording":
        mic_rgb = RED_MAIN
        # 録音中はグロー付き
        for r_mult, alpha in [(0.95, 0.25), (0.85, 0.4), (0.75, 0.55)]:
            halo_r = (size * 0.48) * r_mult
            ctx.arc(cx, cy, halo_r, 0, 2 * math.pi)
            ctx.set_source_rgba(
                mic_rgb[0] / 255, mic_rgb[1] / 255, mic_rgb[2] / 255, alpha,
            )
            ctx.fill()
        mic_r_size = size * 0.36
    elif mode == "off":
        mic_rgb = (130, 130, 130)
        mic_r_size = size * 0.42
    else:  # active
        mic_rgb = PURPLE_MAIN
        mic_r_size = size * 0.42

    # ボタン本体 (円、フラット)
    ctx.arc(cx, cy, mic_r_size, 0, 2 * math.pi)
    ctx.set_source_rgba(
        mic_rgb[0] / 255, mic_rgb[1] / 255, mic_rgb[2] / 255, 1.0,
    )
    ctx.fill()

    # マイクアイコン (白カプセル + スタンド + ベース)
    ctx.set_source_rgba(1.0, 1.0, 1.0, 1.0)

    # カプセル (角丸長方形)
    cap_w = size * 0.10
    cap_h = size * 0.18
    cap_y_offset = -size * 0.06  # やや上寄せ

    # 角丸の path
    cap_x0 = cx - cap_w
    cap_y0 = cy + cap_y_offset - cap_h
    cap_x1 = cx + cap_w
    cap_y1 = cy + cap_y_offset + cap_h * 0.4  # 下側は少し短く
    cap_r = cap_w

    ctx.new_sub_path()
    ctx.arc(cap_x1 - cap_r, cap_y0 + cap_r, cap_r, -math.pi / 2, 0)
    ctx.arc(cap_x1 - cap_r, cap_y1 - cap_r, cap_r, 0, math.pi / 2)
    ctx.arc(cap_x0 + cap_r, cap_y1 - cap_r, cap_r, math.pi / 2, math.pi)
    ctx.arc(cap_x0 + cap_r, cap_y0 + cap_r, cap_r, math.pi, 3 * math.pi / 2)
    ctx.close_path()
    ctx.fill()

    # スタンド (縦線)
    ctx.set_line_width(max(1, size * 0.033))
    ctx.set_line_cap(cairo.LINE_CAP_ROUND)
    stand_y0 = cy + cap_y_offset + cap_h * 0.5
    stand_y1 = cy + cap_y_offset + cap_h * 0.85
    ctx.move_to(cx, stand_y0)
    ctx.line_to(cx, stand_y1)
    ctx.stroke()

    # ベース (横線)
    base_half = size * 0.10
    ctx.move_to(cx - base_half, stand_y1)
    ctx.line_to(cx + base_half, stand_y1)
    ctx.stroke()

    return cairo_surface_to_pil(surface)


def main():
    out_dir = Path(__file__).parent
    ico_sizes = [16, 24, 32, 48, 64, 128, 256]

    # .ico は 256x256 高品質ベースから PIL に downscale させる
    big = draw_mic(256, "active")
    ico_path = out_dir / "mic.ico"
    big.save(
        ico_path, format="ICO",
        sizes=[(sz, sz) for sz in ico_sizes],
    )
    print(f"Saved: {ico_path}")

    # トレイ用 PNG 3バリエーション
    for mode, name in [
        ("active", "mic.png"),
        ("recording", "mic_rec.png"),
        ("off", "mic_off.png"),
    ]:
        p = out_dir / name
        draw_mic(128, mode).save(p, format="PNG")
        print(f"Saved: {p} ({mode})")


if __name__ == "__main__":
    main()
