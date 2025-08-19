import os, io, json, base64
from io import BytesIO

import streamlit as st
import pandas as pd
from dotenv import load_dotenv
from openai import OpenAI
from PIL import Image
from pdf2image import convert_from_bytes

# --------- Config de base ----------
st.set_page_config(page_title="Analyseur de Factures avec IA", layout="centered")

# Charger la clé API depuis .env
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY", "")
if not API_KEY:
    st.error("❌ Clé OpenAI absente. Ouvre le fichier `.env` et ajoute OPENAI_API_KEY=sk-proj-... puis relance.")
    st.stop()

client = OpenAI(api_key=API_KEY)

# --------- Helpers ----------
def pdf_to_image_bytes(pdf_bytes: bytes, dpi: int = 250) -> bytes:
    """Prend un PDF (bytes), retourne la 1ère page en JPEG (bytes)."""
    pages = convert_from_bytes(pdf_bytes, dpi=dpi)
    if not pages:
        raise RuntimeError("PDF illisible (Poppler requis ?)")
    buf = BytesIO()
    pages[0].save(buf, format="JPEG", quality=90)
    return buf.getvalue()

def image_file_to_bytes(file) -> bytes:
    """Prend un fichier image streamlit (jpg/png), retourne bytes JPEG."""
    img = Image.open(file).convert("RGB")
    buf = BytesIO()
    img.save(buf, format="JPEG", quality=90)
    return buf.getvalue()

def to_data_url(img_bytes: bytes, mime="image/jpeg") -> str:
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"

def call_openai_vision(image_data_url: str) -> dict:
    """Appelle le modèle vision pour extraire un JSON structuré de la facture."""
    messages = [
        {
            "role": "system",
            "content": "Tu es un expert en lecture de factures. Réponds uniquement en JSON valide."
        },
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": (
                        "Extrait ces champs depuis l'image de facture et renvoie UNIQUEMENT un JSON valide :\n"
                        "{\n"
                        '  "fournisseur": "",\n'
                        '  "siren_ou_siret": "",\n'
                        '  "tva_intra": "",\n'
                        '  "client_nom": "",\n'
                        '  "client_adresse": "",\n'
                        '  "numero_facture": "",\n'
                        '  "date_facture": "",\n'
                        '  "date_echeance": "",\n'
                        '  "total_ht": "",\n'
                        '  "tva": "",\n'
                        '  "total_ttc": "",\n'
                        '  "devise": "",\n'
                        '  "lignes": [\n'
                        '    {"description":"", "quantite":"", "unite":"", "prix_unitaire_ht":"", "tva_pourcent":"", "total_ligne_ttc":""}\n'
                        "  ]\n"
                        "}\n"
                        "Formats tolérés pour les dates : AAAA-MM-JJ ou JJ/MM/AAAA."
                    )
                },
                {
                    "type": "image_url",
                    "image_url": {"url": image_data_url}
                }
            ]
        }
    ]

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=messages,
        temperature=0
    )
    content = resp.choices[0].message.content.strip()

    # Essayer de parser en JSON directement
    try:
        return json.loads(content)
    except Exception:
        # Si le modèle entoure le JSON de ```json ... ```
        import re
        m = re.search(r"\{.*\}", content, flags=re.S)
        if m:
            return json.loads(m.group(0))
        raise ValueError(f"Réponse non JSON: {content[:300]}")

def json_to_excel_bytes(data: dict) -> bytes:
    """Crée un Excel avec 2 onglets (Facture + Lignes) et renvoie bytes."""
    # Onglet 1: Métadonnées
    meta = {
        "fournisseur": data.get("fournisseur", ""),
        "siren_ou_siret": data.get("siren_ou_siret", ""),
        "tva_intra": data.get("tva_intra", ""),
        "client_nom": data.get("client_nom", ""),
        "client_adresse": data.get("client_adresse", ""),
        "numero_facture": data.get("numero_facture", ""),
        "date_facture": data.get("date_facture", ""),
        "date_echeance": data.get("date_echeance", ""),
        "total_ht": data.get("total_ht", ""),
        "tva": data.get("tva", ""),
        "total_ttc": data.get("total_ttc", ""),
        "devise": data.get("devise", ""),
    }
    df_meta = pd.DataFrame([meta])

    # Onglet 2: Lignes
    lignes = data.get("lignes", [])
    if not isinstance(lignes, list):
        lignes = []
    df_lignes = pd.DataFrame(lignes)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        df_meta.to_excel(xw, index=False, sheet_name="Facture")
        if not df_lignes.empty:
            df_lignes.to_excel(xw, index=False, sheet_name="Lignes")
        else:
            # au moins une feuille Lignes vide
            pd.DataFrame(columns=["description","quantite","unite","prix_unitaire_ht","tva_pourcent","total_ligne_ttc"]).to_excel(
                xw, index=False, sheet_name="Lignes"
            )
    out.seek(0)
    return out.getvalue()

# --------- UI ----------
st.title("📑 Analyseur de Factures avec IA")
st.caption("Choisis une facture (PDF ou image) et récupère un Excel prêt pour la comptabilité.")

f = st.file_uploader("Dépose un fichier ici", type=["pdf","jpg","jpeg","png"])

col1, col2 = st.columns(2)
with col1:
    lancer = st.button("🔍 Analyser")
with col2:
    st.write("")

if lancer:
    if not f:
        st.warning("Ajoute d'abord un fichier.")
        st.stop()

    try:
        # Prépare l'image (JPEG) pour l'IA
        if f.type == "application/pdf":
            img_bytes = pdf_to_image_bytes(f.read(), dpi=250)
        else:
            img_bytes = image_file_to_bytes(f)

        # Aperçu
        st.image(img_bytes, caption="Aperçu (1ère page si PDF)", use_container_width=True)

        # Appel IA
        data_url = to_data_url(img_bytes, mime="image/jpeg")
        with st.spinner("Analyse IA en cours…"):
            data = call_openai_vision(data_url)

        # Affichage JSON
        st.subheader("🧾 Résultat (JSON)")
        st.code(json.dumps(data, ensure_ascii=False, indent=2), language="json")

        # Excel
        excel_bytes = json_to_excel_bytes(data)
        st.download_button(
            label="💾 Télécharger en Excel",
            data=excel_bytes,
            file_name=f"facture_{data.get('numero_facture','extrait')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Récap rapide
        st.subheader("✅ Récapitulatif")
        recap_cols = ["fournisseur","numero_facture","date_facture","date_echeance","total_ht","tva","total_ttc","devise"]
        recap = {k: data.get(k,"") for k in recap_cols}
        st.table(pd.DataFrame([recap]))
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
        st.info("Vérifie : clé API (.env), Poppler installé (pour PDF), et que l'image est lisible.")