#!/usr/bin/env python3
"""
Eye Reflection Analyzer
Analyzes what is reflected in the eyes of a person in a portrait photo.
"""

import sys
import os
import argparse
import base64
import json

try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

try:
    from PIL import Image, ImageEnhance, ImageFilter
    import io
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


# Haar cascade paths (standard OpenCV install)
CASCADE_FACE = cv2.data.haarcascades + "haarcascade_frontalface_default.xml" if HAS_CV2 else ""
CASCADE_EYE  = cv2.data.haarcascades + "haarcascade_eye.xml"                  if HAS_CV2 else ""


def check_dependencies():
    missing = []
    if not HAS_CV2:
        missing.append("opencv-python")
    if not HAS_PIL:
        missing.append("Pillow")
    if not HAS_ANTHROPIC:
        missing.append("anthropic")
    if missing:
        print(f"[!] Missing dependencies: {', '.join(missing)}")
        print(f"    Install with: pip install {' '.join(missing)}")
        sys.exit(1)


def load_image(path: str) -> np.ndarray:
    img = cv2.imread(path)
    if img is None:
        print(f"[!] Cannot load image: {path}")
        sys.exit(1)
    return img


def detect_eyes(img: np.ndarray) -> list[dict]:
    """Return list of eye crops as dicts with 'region' and 'crop' keys."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.equalizeHist(gray)

    face_cascade = cv2.CascadeClassifier(CASCADE_FACE)
    eye_cascade  = cv2.CascadeClassifier(CASCADE_EYE)

    faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(60, 60))
    if len(faces) == 0:
        print("[!] No face detected. Trying full-image eye search.")
        eyes_raw = eye_cascade.detectMultiScale(gray, scaleFactor=1.05, minNeighbors=4, minSize=(20, 20))
        regions = [{"face": None, "eye": e} for e in eyes_raw]
    else:
        regions = []
        for (fx, fy, fw, fh) in faces:
            roi_gray = gray[fy:fy+fh, fx:fx+fw]
            eyes = eye_cascade.detectMultiScale(roi_gray, scaleFactor=1.05, minNeighbors=4, minSize=(15, 15))
            for e in eyes:
                ex, ey, ew, eh = e
                regions.append({
                    "face": (fx, fy, fw, fh),
                    "eye":  (fx + ex, fy + ey, ew, eh),
                })

    results = []
    for r in regions:
        x, y, w, h = r["eye"]
        # Add padding around the eye
        pad = int(max(w, h) * 0.5)
        x1 = max(0, x - pad)
        y1 = max(0, y - pad)
        x2 = min(img.shape[1], x + w + pad)
        y2 = min(img.shape[0], y + h + pad)
        crop = img[y1:y2, x1:x2]
        results.append({"region": (x1, y1, x2 - x1, y2 - y1), "crop": crop})

    return results


def enhance_eye_crop(crop_bgr: np.ndarray, scale: int = 8) -> Image.Image:
    """Upscale and sharpen the eye crop to make reflections more visible."""
    h, w = crop_bgr.shape[:2]
    large = cv2.resize(crop_bgr, (w * scale, h * scale), interpolation=cv2.INTER_CUBIC)

    # Increase local contrast with CLAHE
    lab = cv2.cvtColor(large, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    l = clahe.apply(l)
    large = cv2.cvtColor(cv2.merge([l, a, b]), cv2.COLOR_LAB2BGR)

    # Convert to PIL for final sharpening
    pil_img = Image.fromarray(cv2.cvtColor(large, cv2.COLOR_BGR2RGB))
    pil_img = ImageEnhance.Sharpness(pil_img).enhance(2.5)
    pil_img = ImageEnhance.Contrast(pil_img).enhance(1.4)
    return pil_img


def image_to_base64(pil_img: Image.Image) -> str:
    buf = io.BytesIO()
    pil_img.save(buf, format="JPEG", quality=95)
    return base64.standard_b64encode(buf.getvalue()).decode()


def analyze_reflection_with_claude(eye_images: list[Image.Image], original_b64: str) -> str:
    """Send eye crops + original photo to Claude and ask about reflections."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return "[!] ANTHROPIC_API_KEY not set. Set it to enable AI analysis."

    client = anthropic.Anthropic(api_key=api_key)

    content = [
        {
            "type": "text",
            "text": (
                "I am analyzing a portrait photo to determine what is reflected "
                "in the subject's eyes. I will provide you with:\n"
                "1. The full original portrait\n"
                "2. Cropped and enhanced close-ups of each detected eye\n\n"
                "Please examine the eye crops carefully and describe:\n"
                "- What objects, shapes, colors, or scenes are visible as "
                "  reflections in the eyes (corneal reflections / catchlights)\n"
                "- Any light sources reflected (windows, lamps, flash, sky)\n"
                "- Any environmental details you can infer from the reflections\n"
                "- How confident you are in your assessment\n\n"
                "Be as specific as possible about shapes and colors you observe."
            ),
        },
        {
            "type": "text",
            "text": "Full original portrait:",
        },
        {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": original_b64,
            },
        },
    ]

    for i, eye_img in enumerate(eye_images, 1):
        content.append({"type": "text", "text": f"Enhanced eye crop #{i}:"})
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": image_to_base64(eye_img),
            },
        })

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=1024,
        messages=[{"role": "user", "content": content}],
    )
    return response.content[0].text


def save_debug_crops(eye_data: list[dict], output_dir: str):
    os.makedirs(output_dir, exist_ok=True)
    for i, ed in enumerate(eye_data, 1):
        enhanced = enhance_eye_crop(ed["crop"])
        path = os.path.join(output_dir, f"eye_{i}_enhanced.jpg")
        enhanced.save(path, quality=95)
        print(f"[+] Saved enhanced eye #{i} -> {path}")


def prompt_image_path() -> str:
    """Ask the user for the portrait image path, validating it exists."""
    while True:
        path = input("\nPercorso della foto da analizzare: ").strip()
        if not path:
            print("[!] Il percorso non può essere vuoto.")
            continue
        # Remove accidental quotes pasted from file managers
        path = path.strip("'\"")
        if not os.path.isfile(path):
            print(f"[!] File non trovato: {path}")
            continue
        ext = os.path.splitext(path)[1].lower()
        if ext not in {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"}:
            print(f"[!] Formato non supportato ({ext}). Usa JPG, PNG, BMP, TIFF o WebP.")
            continue
        return path


def main():
    parser = argparse.ArgumentParser(
        description="Analyze reflections in a person's eyes from a portrait photo."
    )
    parser.add_argument(
        "image",
        nargs="?",
        default=None,
        help="Path to the portrait image (JPG/PNG). If omitted, you will be prompted.",
    )
    parser.add_argument(
        "--save-crops",
        metavar="DIR",
        default=None,
        help="Directory to save enhanced eye crops for manual inspection",
    )
    parser.add_argument(
        "--no-ai",
        action="store_true",
        help="Skip Claude AI analysis (only detect and save crops)",
    )
    args = parser.parse_args()

    check_dependencies()

    print("=" * 60)
    print("  EYE REFLECTION ANALYZER")
    print("=" * 60)

    # Ask for path interactively if not supplied as argument
    image_path = args.image if args.image else prompt_image_path()

    print(f"\n[*] Loading image: {image_path}")
    img = load_image(image_path)
    print(f"[*] Image size: {img.shape[1]}x{img.shape[0]} px")

    print("[*] Detecting eyes...")
    eye_data = detect_eyes(img)
    if not eye_data:
        print("[!] No eyes detected in the image.")
        sys.exit(1)

    print(f"[+] Found {len(eye_data)} eye region(s)")

    if args.save_crops:
        save_debug_crops(eye_data, args.save_crops)

    if args.no_ai:
        print("[*] --no-ai flag set, skipping Claude analysis.")
        return

    print("[*] Enhancing eye crops for analysis...")
    enhanced_eyes = [enhance_eye_crop(ed["crop"]) for ed in eye_data]

    # Prepare original image as base64 (resize if very large)
    pil_orig = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    max_side = 1600
    if max(pil_orig.size) > max_side:
        pil_orig.thumbnail((max_side, max_side), Image.LANCZOS)
    original_b64 = image_to_base64(pil_orig)

    print("[*] Sending to Claude for reflection analysis...")
    result = analyze_reflection_with_claude(enhanced_eyes, original_b64)

    print("\n" + "=" * 60)
    print("EYE REFLECTION ANALYSIS")
    print("=" * 60)
    print(result)
    print("=" * 60)


if __name__ == "__main__":
    main()
