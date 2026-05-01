import os
import re
import zipfile
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont


# -----------------------------
# CONFIGURATION
# -----------------------------
DEFAULT_WIDTH = 340
DEFAULT_HEIGHT = 340
TITLE_TEXT = "FICHE TECHNIQUE"

BG_COLOR = (239, 239, 239)
BORDER_COLOR = (20, 20, 20)
HEADER_COLOR = (20, 20, 20)
TEXT_COLOR = (25, 25, 25)
LINE_COLOR = (195, 195, 195)


# -----------------------------
# FONCTIONS
# -----------------------------
def safe_filename(value):
    value = str(value).strip()

    # Corrige les EAN lus comme 1234567890123.0
    if value.endswith(".0"):
        value = value[:-2]

    value = re.sub(r'[\\/*?:"<>|]+', "_", value)
    value = re.sub(r"\s+", "_", value)

    return value or "sans_ean"


def load_font(size, bold=False):
    if bold:
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "arialbd.ttf",
            "Arial Bold.ttf",
        ]
    else:
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "arial.ttf",
            "Arial.ttf",
        ]

    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass

    return ImageFont.load_default()


def text_width(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0]


def text_height(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[3] - bbox[1]


def truncate_text(draw, text, font, max_width):
    text = str(text).strip()

    if text_width(draw, text, font) <= max_width:
        return text

    ellipsis = "..."
    while text and text_width(draw, text + ellipsis, font) > max_width:
        text = text[:-1]

    return text + ellipsis


def fit_font_single_line(draw, text, max_width, max_height, preferred_size, min_size):
    text = "" if pd.isna(text) else str(text).strip()

    for size in range(preferred_size, min_size - 1, -1):
        font = load_font(size, bold=False)

        if (
            text_width(draw, text, font) <= max_width
            and text_height(draw, text, font) <= max_height
        ):
            return font, text

    font = load_font(min_size, bold=False)
    text = truncate_text(draw, text, font, max_width)

    return font, text


def create_base_image(width, height):
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    outer_margin = max(5, int(width * 0.025))
    border_width = max(2, int(width * 0.006))
    radius = max(12, int(width * 0.045))

    # Fond principal
    draw.rounded_rectangle(
        [
            (outer_margin, outer_margin),
            (width - outer_margin, height - outer_margin),
        ],
        radius=radius,
        fill=BG_COLOR,
        outline=BORDER_COLOR,
        width=border_width,
    )

    # Header noir
    header_h = int(height * 0.125)

    draw.rectangle(
        [
            (outer_margin + border_width, outer_margin + border_width),
            (width - outer_margin - border_width, outer_margin + header_h),
        ],
        fill=HEADER_COLOR,
    )

    # Titre
    title_font = load_font(int(height * 0.045), bold=True)
    title_w = text_width(draw, TITLE_TEXT, title_font)
    title_h = text_height(draw, TITLE_TEXT, title_font)

    title_x = (width - title_w) // 2
    title_y = outer_margin + (header_h - title_h) // 2 - 2

    draw.text(
        (title_x, title_y),
        TITLE_TEXT,
        fill=(255, 255, 255),
        font=title_font,
    )

    # Lignes horizontales
    body_left = int(width * 0.065)
    body_right = int(width * 0.935)
    body_top = int(height * 0.18)
    body_bottom = int(height * 0.93)

    row_count = 10
    row_h = (body_bottom - body_top) / row_count

    for i in range(row_count + 1):
        y = int(body_top + i * row_h)
        draw.line(
            (body_left, y, body_right, y),
            fill=LINE_COLOR,
            width=max(1, int(width * 0.002)),
        )

    return img


def draw_lines_on_image(img, values):
    draw = ImageDraw.Draw(img)
    width, height = img.size

    text_left = int(width * 0.065)
    text_right = int(width * 0.935)

    body_top = int(height * 0.18)
    body_bottom = int(height * 0.93)

    row_count = 10
    row_h = (body_bottom - body_top) / row_count

    max_text_width = text_right - text_left

    preferred_size = int(height * 0.055)
    min_size = int(height * 0.033)

    for idx in range(10):
        text = values[idx] if idx < len(values) else ""
        text = "" if pd.isna(text) else str(text).strip()

        if not text:
            continue

        row_top = int(body_top + idx * row_h)
        row_bottom = int(body_top + (idx + 1) * row_h)
        available_h = row_bottom - row_top

        font, final_text = fit_font_single_line(
            draw=draw,
            text=text,
            max_width=max_text_width,
            max_height=available_h - 2,
            preferred_size=preferred_size,
            min_size=min_size,
        )

        h = text_height(draw, final_text, font)
        y = row_top + max(0, (available_h - h) // 2) - 1

        draw.text(
            (text_left, y),
            final_text,
            fill=TEXT_COLOR,
            font=font,
        )

    return img


def create_excel_template():
    columns = ["ean"] + [f"L{i}" for i in range(1, 11)]

    example_data = [
        {
            "ean": "1234567890123",
            "L1": "Volume : 300 L",
            "L2": "Classe : A++",
            "L3": "Couleur : Inox",
            "L4": "Bruit : 39 dB",
            "L5": "No Frost",
            "L6": "Multi Air Flow",
            "L7": "Dim. : 186x60x64 cm",
            "L8": "Garantie : 2 ans",
            "L9": "",
            "L10": "",
        },
        {
            "ean": "",
            "L1": "",
            "L2": "",
            "L3": "",
            "L4": "",
            "L5": "",
            "L6": "",
            "L7": "",
            "L8": "",
            "L9": "",
            "L10": "",
        },
    ]

    df_template = pd.DataFrame(example_data, columns=columns)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Produits")

        worksheet = writer.sheets["Produits"]

        # Largeur des colonnes
        worksheet.column_dimensions["A"].width = 18

        for col in range(2, 12):
            col_letter = worksheet.cell(row=1, column=col).column_letter
            worksheet.column_dimensions[col_letter].width = 28

    output.seek(0)
    return output


def read_excel_file(uploaded_file):
    df = pd.read_excel(uploaded_file, dtype=str, keep_default_na=False)

    # Nettoyer les noms des colonnes
    df.columns = [str(c).strip() for c in df.columns]

    lower_cols = {c.lower(): c for c in df.columns}

    if "ean" not in lower_cols:
        raise ValueError("La colonne 'ean' est obligatoire.")

    df = df.rename(columns={lower_cols["ean"]: "ean"})

    # Renommer L1...L10 si nécessaire
    lower_cols = {c.lower(): c for c in df.columns}

    for i in range(1, 11):
        col_lower = f"l{i}"

        if col_lower in lower_cols:
            df = df.rename(columns={lower_cols[col_lower]: f"L{i}"})

    # Ajouter colonnes manquantes
    for i in range(1, 11):
        col = f"L{i}"

        if col not in df.columns:
            df[col] = ""

    df = df[["ean"] + [f"L{i}" for i in range(1, 11)]]

    # Supprimer les lignes sans EAN
    df["ean"] = df["ean"].astype(str).str.strip()
    df = df[df["ean"] != ""]

    return df


def generate_images_and_zip(df, width, height):
    temp_dir = tempfile.mkdtemp()
    images_dir = os.path.join(temp_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    generated_files = []
    used_names = {}

    for _, row in df.iterrows():
        ean = safe_filename(row["ean"])

        if not ean:
            continue

        # Évite d’écraser si le même EAN existe plusieurs fois
        base_name = ean

        if base_name in used_names:
            used_names[base_name] += 1
            ean = f"{base_name}_{used_names[base_name]}"
        else:
            used_names[base_name] = 0

        values = [row.get(f"L{i}", "") for i in range(1, 11)]

        img = create_base_image(width, height)
        img = draw_lines_on_image(img, values)

        image_path = os.path.join(images_dir, f"{ean}.png")
        img.save(image_path)

        generated_files.append(image_path)

    zip_path = os.path.join(temp_dir, "fiches_techniques.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file_path in generated_files:
            zipf.write(file_path, arcname=os.path.basename(file_path))

    return zip_path, generated_files


# -----------------------------
# INTERFACE STREAMLIT
# -----------------------------
st.set_page_config(
    page_title="Générateur Fiches Techniques",
    page_icon="🏷️",
    layout="centered",
)

st.title("🏷️ Générateur de Fiches Techniques ESL")

st.write(
    "Importez un fichier Excel contenant les colonnes : "
    "`ean`, `L1`, `L2`, ..., `L10`."
)

# Dimension fixe
width = DEFAULT_WIDTH
height = DEFAULT_HEIGHT

st.info("Dimension des images générées : 340 × 340 px")

# Bouton modèle Excel
template_excel = create_excel_template()

st.download_button(
    label="📥 Télécharger le modèle Excel à remplir",
    data=template_excel,
    file_name="modele_fiches_techniques.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.divider()

uploaded_file = st.file_uploader(
    "Importer le fichier Excel rempli",
    type=["xlsx", "xls"],
)

if uploaded_file is not None:
    try:
        df = read_excel_file(uploaded_file)

        st.success("Fichier Excel chargé avec succès.")

        st.write("Aperçu du fichier :")
        st.dataframe(df.head(10), use_container_width=True)

        st.write(f"Nombre d’articles détectés : **{len(df)}**")

        if st.button("⚙️ Générer les images"):
            zip_path, generated_files = generate_images_and_zip(df, width, height)

            if len(generated_files) == 0:
                st.warning("Aucune image générée. Vérifiez que la colonne EAN est remplie.")
            else:
                st.success(f"{len(generated_files)} image(s) générée(s).")

                st.write("Aperçu des premières images :")

                for file_path in generated_files[:3]:
                    st.image(
                        file_path,
                        caption=os.path.basename(file_path),
                        width=250,
                    )

                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📥 Télécharger le ZIP",
                        data=f,
                        file_name="fiches_techniques.zip",
                        mime="application/zip",
                    )

    except Exception as e:
        st.error(f"Erreur : {e}")