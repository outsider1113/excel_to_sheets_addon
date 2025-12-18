from pathlib import Path
from PIL import Image

# === CONFIG ===
BASE_LOGO_PATH = Path("CMX METALS - NEW LOGO.png")  # change if your file is named differently
OUTPUT_DIR = Path("output_icons")
SIZES = [16, 32, 80, 128, 256]          # add/remove sizes as needed


def make_square(img: Image.Image, fill_color=(0, 0, 0, 0)) -> Image.Image:
    """
    Pad the image to a square (keeping aspect ratio) using transparent background.
    This avoids distortion when resizing non-square logos.
    """
    w, h = img.size
    if w == h:
        return img

    size = max(w, h)
    new_img = Image.new("RGBA", (size, size), fill_color)
    offset = ((size - w) // 2, (size - h) // 2)
    new_img.paste(img, offset)
    return new_img


def main():
    if not BASE_LOGO_PATH.exists():
        raise FileNotFoundError(f"Base logo not found: {BASE_LOGO_PATH}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Open base logo
    img = Image.open(BASE_LOGO_PATH).convert("RGBA")
    img_sq = make_square(img)

    for size in SIZES:
        resized = img_sq.resize((size, size), Image.LANCZOS)
        out_path = OUTPUT_DIR / f"icon-{size}.png"
        resized.save(out_path, format="PNG")
        print(f"Saved {out_path}")

    print("Done!")


if __name__ == "__main__":
    main()
