import io
import os
import zipfile
from dataclasses import dataclass
from typing import List, Optional, Tuple

import cv2
import fitz  # PyMuPDF
import numpy as np
import streamlit as st
from PIL import Image, ImageOps


@dataclass
class CandidateImage:
    source_name: str
    image: Image.Image
    face_boxes: List[Tuple[int, int, int, int]]
    score: float


def load_face_cascade() -> cv2.CascadeClassifier:
    cascade_path = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
    cascade = cv2.CascadeClassifier(cascade_path)
    if cascade.empty():
        raise RuntimeError("Could not load OpenCV face cascade.")
    return cascade


def detect_faces(pil_img: Image.Image, cascade: cv2.CascadeClassifier) -> List[Tuple[int, int, int, int]]:
    rgb = np.array(pil_img.convert("RGB"))
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    faces = cascade.detectMultiScale(
        gray,
        scaleFactor=1.1,
        minNeighbors=5,
        minSize=(40, 40),
    )
    return [(int(x), int(y), int(w), int(h)) for (x, y, w, h) in faces]


def score_candidate(img: Image.Image, faces: List[Tuple[int, int, int, int]]) -> float:
    w, h = img.size
    pixel_count = w * h
    if pixel_count <= 0:
        return -1e9

    resolution_score = min(pixel_count / (700 * 700), 4.0)
    if not faces:
        return resolution_score * 0.1

    face_areas = [fw * fh for (_, _, fw, fh) in faces]
    largest_face_area = max(face_areas)
    face_ratio = largest_face_area / pixel_count
    face_count_penalty = abs(len(faces) - 1) * 0.35
    ratio_center_bonus = 1.0 - min(abs(face_ratio - 0.12) / 0.12, 1.0)

    return 3.0 + (face_ratio * 12.0) + ratio_center_bonus + (resolution_score * 0.2) - face_count_penalty


def square_crop_around_face(img: Image.Image, face_box: Optional[Tuple[int, int, int, int]]) -> Image.Image:
    w, h = img.size
    if w <= 0 or h <= 0:
        return img

    if face_box is None:
        side = min(w, h)
        left = (w - side) // 2
        top = (h - side) // 2
        return img.crop((left, top, left + side, top + side))

    x, y, fw, fh = face_box
    cx = x + (fw / 2.0)
    cy = y + (fh / 2.0)

    # Larger multiplier gives breathing room so the crop does not feel too tight.
    target_side = int(max(fw, fh) * 2.6)
    side = max(target_side, int(max(fw, fh) * 2.1))
    side = min(side, min(w, h))
    side = max(side, 128)

    left = int(round(cx - side / 2))
    top = int(round(cy - side / 2))

    # Clamp while preserving square dimensions.
    left = max(0, min(left, w - side))
    top = max(0, min(top, h - side))
    right = left + side
    bottom = top + side

    return img.crop((left, top, right, bottom))


def extract_images_from_docx(file_bytes: bytes, filename: str) -> List[Tuple[str, Image.Image]]:
    out: List[Tuple[str, Image.Image]] = []
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        for name in zf.namelist():
            if not name.startswith("word/media/"):
                continue
            raw = zf.read(name)
            try:
                img = Image.open(io.BytesIO(raw))
                img.load()
                out.append((f"{filename}:{name}", ImageOps.exif_transpose(img).convert("RGB")))
            except Exception:
                continue
    return out


def extract_images_from_pdf(file_bytes: bytes, filename: str) -> List[Tuple[str, Image.Image]]:
    out: List[Tuple[str, Image.Image]] = []
    seen_xrefs = set()

    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        for page_index in range(len(doc)):
            page = doc[page_index]
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)
                try:
                    extracted = doc.extract_image(xref)
                    raw = extracted.get("image")
                    if not raw:
                        continue
                    img = Image.open(io.BytesIO(raw))
                    img.load()
                    out.append((f"{filename}:page{page_index + 1}:xref{xref}", ImageOps.exif_transpose(img).convert("RGB")))
                except Exception:
                    continue
    return out


def gather_candidates(uploaded_file, cascade: cv2.CascadeClassifier) -> List[CandidateImage]:
    name = uploaded_file.name
    data = uploaded_file.read()
    ext = os.path.splitext(name)[1].lower()

    if ext == ".docx":
        extracted = extract_images_from_docx(data, name)
    elif ext == ".pdf":
        extracted = extract_images_from_pdf(data, name)
    else:
        return []

    candidates: List[CandidateImage] = []
    for src_name, pil_img in extracted:
        w, h = pil_img.size
        if w < 80 or h < 80:
            continue
        faces = detect_faces(pil_img, cascade)
        score = score_candidate(pil_img, faces)
        candidates.append(CandidateImage(source_name=src_name, image=pil_img, face_boxes=faces, score=score))
    return candidates


def pick_best_profile(candidates: List[CandidateImage]) -> Optional[CandidateImage]:
    if not candidates:
        return None
    return max(candidates, key=lambda c: c.score)


def image_to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def build_zip(items: List[Tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in items:
            zf.writestr(name, data)
    return buf.getvalue()


def main() -> None:
    st.set_page_config(page_title="CV Profile Picture Extractor", page_icon="CV", layout="wide")
    st.title("CV Profile Picture Extractor")
    st.write("Upload CV files (`.pdf` or `.docx`), and this tool extracts and square-crops likely profile photos.")

    uploaded_files = st.file_uploader(
        "Drop files here or click to upload",
        type=["pdf", "docx"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Add one or more CV files to start.")
        return

    run = st.button("Extract Profile Pictures", type="primary")
    if not run:
        return

    cascade = load_face_cascade()
    results = []
    zip_items: List[Tuple[str, bytes]] = []

    with st.spinner("Extracting images and finding profile photos..."):
        for up in uploaded_files:
            candidates = gather_candidates(up, cascade)
            best = pick_best_profile(candidates)
            if best is None:
                results.append((up.name, None, "No usable images found in this file."))
                continue

            largest_face = max(best.face_boxes, key=lambda b: b[2] * b[3]) if best.face_boxes else None
            cropped = square_crop_around_face(best.image, largest_face)
            png_bytes = image_to_png_bytes(cropped)
            out_name = os.path.splitext(up.name)[0] + "_profile.png"
            zip_items.append((out_name, png_bytes))
            results.append((up.name, (cropped, png_bytes, out_name), f"Picked image: {best.source_name}"))

    st.subheader("Results")
    for idx, (source_name, payload, note) in enumerate(results):
        st.markdown(f"**{source_name}**")
        st.caption(note)
        if payload is None:
            continue

        cropped, png_bytes, out_name = payload
        col1, col2 = st.columns([1, 2])
        with col1:
            st.image(cropped, caption="Square crop", use_container_width=True)
        with col2:
            st.download_button(
                label=f"Download {out_name}",
                data=png_bytes,
                file_name=out_name,
                mime="image/png",
                key=f"dl_{idx}_{source_name}",
            )

    if zip_items:
        zip_bytes = build_zip(zip_items)
        st.download_button(
            label="Download All as ZIP",
            data=zip_bytes,
            file_name="cv_profile_pictures.zip",
            mime="application/zip",
        )


if __name__ == "__main__":
    main()
