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

# Dossier où se trouve app.py (pour retrouver template.png + DejaVuSans.ttf)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PATH = os.path.join(BASE_DIR, "template.png")


# -----------------------------
# OUTILS
# -----------------------------
def safe_filename(value):
    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    value = re.sub(r'[\\/*?:"<>|]+', "_", value)
    value = re.sub(r"\s+", "_", value)

    return value or "sans_ean"


def load_font(size, bold=False):
    """
    Charge une police Unicode (DejaVuSans) capable d'afficher
    é, è, à, ç, ×, °, etc.

    Cherche d'abord dans le dossier du projet (à côté de app.py),
    puis dans les emplacements système. Lève une erreur explicite
    si rien n'est trouvé, pour éviter de retomber silencieusement
    sur la police bitmap par défaut (qui afficherait des carrés □).
    """
    if bold:
        candidates = [
            os.path.join(BASE_DIR, "DejaVuSans-Bold.ttf"),
            os.path.join(BASE_DIR, "fonts", "DejaVuSans-Bold.ttf"),
            # fallback : police normale si la bold n'est pas dispo
            os.path.join(BASE_DIR, "DejaVuSans.ttf"),
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
        ]
    else:
        candidates = [
            os.path.join(BASE_DIR, "DejaVuSans.ttf"),
            os.path.join(BASE_DIR, "fonts", "DejaVuSans.ttf"),
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "C:/Windows/Fonts/arial.ttf",
        ]

    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue

    raise RuntimeError(
        "Police Unicode introuvable. Ajoutez 'DejaVuSans.ttf' "
        "à la racine du projet (à côté de app.py)."
    )


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


# -----------------------------
# TEMPLATE IMAGE
# -----------------------------
def create_base_image(width, height):
    """
    Utilise le vrai template image au lieu de le redessiner.
    """
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(
            f"Le fichier template '{TEMPLATE_PATH}' est introuvable. "
            "Ajoutez votre image vide dans le projet."
        )

    img = Image.open(TEMPLATE_PATH).convert("RGB")
    img = img.resize((width, height))
    return img


def draw_lines_on_image(img, values):
    """
    Écriture ajustée pour votre template 340x340.
    Police plus grande + caractères spéciaux OK.
    """
    draw = ImageDraw.Draw(img)
    width, height = img.size

    # Coordonnées adaptées à votre template
    text_left = 22
    text_right = 318

    # Zone texte sous le header noir
    body_top = 56
    body_bottom = 305

    row_count = 10
    row_h = (body_bottom - body_top) / row_count
    max_text_width = text_right - text_left

    # Texte plus grand
    preferred_size = 18
    min_size = 13

    for idx in range(10):
        text = values[idx] if idx < len(values) else ""
        text = "" if pd.isna(text) else str(text).strip()

        if not text:
            continue

        # Nettoyage caractères
        text = (
            text.replace("x", "×")
            .replace(" X ", " × ")
            .replace("  ", " ")
        )

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
        y = row_top + max(0, (available_h - h) // 2) - 2

        draw.text(
            (text_left, y),
            final_text,
            fill=(20, 20, 20),
            font=font,
        )

    return img


# -----------------------------
# EXCEL TEMPLATE
# -----------------------------
def create_excel_template():
    columns = ["ean"] + [f"L{i}" for i in range(1, 11)]

    example_data = [
        {
            "ean": "1234567890123",
            "L1": "VOLUME NET : 486L",
            "L2": "VOLUME REFRIG. : 348L",
            "L3": "VOLUME CONGEL. : 138L",
            "L4": "CLASSE ÉNERG. : A++",
            "L5": "CONSO. : 231Kwh/an",
            "L6": "POUV. CONGEL. : 16kg/24h",
            "L7": "COULEUR : INOX",
            "L8": "NIVEAU SONORE : 35 db",
            "L9": "DIMENSIONS : 201/75/68",
            "L10": "GARANTIE : 2 ANS",
        }
    ]

    df_template = pd.DataFrame(example_data, columns=columns)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Produits")
        worksheet = writer.sheets["Produits"]

        worksheet.column_dimensions["A"].width = 20
        for col in range(2, 12):
            col_letter = worksheet.cell(row=1, column=col).column_letter
            worksheet.column_dimensions[col_letter].width = 28

    output.seek(0)
    return output


def read_excel_file(uploaded_file):
    """
    Lit un fichier Excel (.xlsx ou .xls).
    Choisit explicitement le moteur en fonction de l'extension :
        - .xlsx -> openpyxl
        - .xls  -> xlrd (version < 2.0 requise dans requirements.txt)
    """
    filename = getattr(uploaded_file, "name", "").lower()

    if filename.endswith(".xls"):
        engine = "xlrd"
    else:
        engine = "openpyxl"

    df = pd.read_excel(
        uploaded_file,
        dtype=str,
        keep_default_na=False,
        engine=engine,
    )

    df.columns = [str(c).strip() for c in df.columns]
    lower_cols = {c.lower(): c for c in df.columns}

    if "ean" not in lower_cols:
        raise ValueError("La colonne 'ean' est obligatoire.")

    df = df.rename(columns={lower_cols["ean"]: "ean"})

    lower_cols = {c.lower(): c for c in df.columns}

    for i in range(1, 11):
        col_lower = f"l{i}"
        if col_lower in lower_cols:
            df = df.rename(columns={lower_cols[col_lower]: f"L{i}"})

    for i in range(1, 11):
        col = f"L{i}"
        if col not in df.columns:
            df[col] = ""

    df = df[["ean"] + [f"L{i}" for i in range(1, 11)]]

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
        img.save(image_path, format="PNG")

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
    page_title="Générateur Fiches Techniques ESL",
    page_icon="🏷️",
    layout="centered",
)

st.title("🏷️ Générateur de Fiches Techniques ESL")
st.write("Importez un fichier Excel contenant les colonnes : `ean`, `L1`, `L2`, ..., `L10`.")

width = DEFAULT_WIDTH
height = DEFAULT_HEIGHT

st.info("Dimension des images générées : 340 × 340 px")
st.info("Le rendu dépend du fichier template.png utilisé comme modèle.")

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

        st.write(f"Nombre d'articles détectés : **{len(df)}**")

        if st.button("⚙️ Générer les images"):
            zip_path, generated_files = generate_images_and_zip(df, width, height)

            if len(generated_files) == 0:
                st.warning("Aucune image générée. Vérifiez que la colonne EAN est remplie.")
            else:
                st.success(f"{len(generated_files)} image(s) générée(s).")

                st.write("Aperçu des premières images :")
                for file_path in generated_files[:3]:
                    st.image(file_path, caption=os.path.basename(file_path), width=250)

                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📥 Télécharger le ZIP",
                        data=f,
                        file_name="fiches_techniques.zip",
                        mime="application/zip",
                    )

    except Exception as e:
        st.error(f"Erreur : {e}")
