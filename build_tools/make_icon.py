"""生成程序图标 assets/app.ico（蓝白风：蓝色圆角底 + 白色文档 + 对齐线）。"""

from __future__ import annotations

import os

from PIL import Image, ImageDraw

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(BASE, "assets", "app.ico")
PNG = os.path.join(BASE, "assets", "app_256.png")


def draw(size: int = 256) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size / 256.0

    def r(*v):
        return [x * s for x in v]

    # 蓝色圆角背景（上下渐变用两层近似）
    d.rounded_rectangle(r(8, 8, 248, 248), radius=44 * s, fill=(37, 99, 235, 255))
    d.rounded_rectangle(r(8, 8, 248, 150), radius=44 * s, fill=(59, 116, 240, 255))
    d.rounded_rectangle(r(8, 60, 248, 248), radius=44 * s, fill=(37, 99, 235, 255))

    # 白色纸张
    d.rounded_rectangle(r(58, 40, 198, 216), radius=14 * s, fill=(255, 255, 255, 255))
    # 折角
    d.polygon([tuple(r(166, 40))[0:2], tuple(r(198, 72))[0:2], tuple(r(166, 72))[0:2]],
              fill=(219, 231, 254, 255))

    # 文字线条（模拟排版：标题粗、正文细、逐行对齐）
    blue = (37, 99, 235, 255)
    grey = (148, 163, 184, 255)
    d.rounded_rectangle(r(78, 74, 150, 88), radius=6 * s, fill=blue)
    for i, (x1, x2, y) in enumerate([(78, 178, 104), (78, 168, 122),
                                     (78, 178, 140), (78, 148, 158)]):
        d.rounded_rectangle(r(x1, y, x2, y + 9), radius=4 * s, fill=grey)
    # 底部对勾块（代表"优化完成"）
    d.rounded_rectangle(r(118, 180, 198, 232), radius=16 * s, fill=(16, 185, 129, 255))
    d.line([tuple(r(134, 206)), tuple(r(152, 222)), tuple(r(184, 190))],
           fill=(255, 255, 255, 255), width=int(11 * s), joint="curve")
    return img


def main() -> None:
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    img = draw(256)
    img.save(PNG)
    img.save(OUT, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64),
                         (128, 128), (256, 256)])
    print("saved", OUT, PNG)


if __name__ == "__main__":
    main()
