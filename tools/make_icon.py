"""MarkAnyAuto.ico 를 굽는다.  실행:  python3 tools/make_icon.py

MarkAny 의 각진 아스테리스크 마크를 변형해 가운데를 열쇠구멍으로 뚫었다.
16/32px 는 구멍이 메워지므로 구멍을 키운 별도 도안에서 뽑는다.
"""
import math
from pathlib import Path

from PIL import Image, ImageDraw

N = 1024
NAVY = (16, 68, 143, 255)          # MarkAny 마크에서 뽑은 네이비
CLEAR = (0, 0, 0, 0)
TILT = -14                         # 원본처럼 살짝 기울인다
BLADES = [(0.485, 0.055), (0.455, -0.05), (0.50, 0.05),
          (0.46, -0.055), (0.49, 0.045), (0.445, -0.05)]  # 길이, 끝 사선


def _blade(cx, cy, ang, length, skew, half):
    a = math.radians(ang)
    dx, dy = math.cos(a), math.sin(a)
    px, py = -dy, dx
    L, W, S = length * N, half * N, skew * N
    return [(cx + px * W, cy + py * W),
            (cx + dx * L + px * W + dx * S, cy + dy * L + py * W + dy * S),
            (cx + dx * L - px * W - dx * S, cy + dy * L - py * W - dy * S),
            (cx - px * W, cy - py * W)]


def draw(half: float, hole: float) -> Image.Image:
    im = Image.new("RGBA", (N, N), CLEAR)
    d = ImageDraw.Draw(im)
    for i, (length, skew) in enumerate(BLADES):
        d.polygon(_blade(N / 2, N / 2, TILT + i * 60, length, skew, half), fill=NAVY)
    cx, cy, r = N / 2, N * 0.455, N * hole
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=CLEAR)
    d.polygon([(cx - r * 0.62, cy + r * 0.35), (cx + r * 0.62, cy + r * 0.35),
               (cx + r * 1.05, cy + r * 2.45), (cx - r * 1.05, cy + r * 2.45)], fill=CLEAR)
    return im


def main():
    out = Path(__file__).resolve().parent.parent / "MarkAnyAuto.ico"
    big, small = draw(0.145, 0.118), draw(0.175, 0.150)
    frames = [big.resize((s, s), Image.LANCZOS) for s in (256, 48)]
    frames += [small.resize((s, s), Image.LANCZOS) for s in (32, 16)]
    frames[0].save(out, format="ICO", sizes=[(f.width, f.width) for f in frames],
                   append_images=frames[1:])
    print(f"{out}  {', '.join(str(f.width) for f in frames)}")


if __name__ == "__main__":
    main()
