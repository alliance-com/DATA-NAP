"""
NAP SARL — Téléchargement des bons en WORD et en PDF
À placer à la RACINE du projet (à côté de app.py).

Le PDF est créé à partir du fichier Word, à la demande, puis conservé à côté du Word
(ex. bons_paiement/Bon_Paiement_CHQ-0400-2026.docx  ->  bons_paiement/Bon_Paiement_CHQ-0400-2026.pdf).
Si le Word est régénéré, le bouton propose de refaire le PDF.

Conversion (premier moteur disponible) :
  1) Microsoft Word  -> installer :  pip install docx2pdf        (Word doit être installé sur l'ordinateur)
  2) LibreOffice     -> gratuit :    https://www.libreoffice.org  (aucun réglage, détecté automatiquement)
"""
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

VERSION_TELECHARGEMENT = 1
MIME_WORD = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_PDF = "application/pdf"
DELAI_CONVERSION = 180   # secondes


# ==================== Conversion ====================
def chemin_pdf(chemin_docx):
    return os.path.splitext(chemin_docx)[0] + ".pdf"


def pdf_a_jour(chemin_docx):
    """True si le PDF existe et n'est pas plus ancien que le Word."""
    pdf = chemin_pdf(chemin_docx)
    return os.path.exists(pdf) and os.path.getmtime(pdf) >= os.path.getmtime(chemin_docx) - 2


def _libreoffice():
    for nom in ("soffice", "libreoffice", "soffice.exe"):
        chemin = shutil.which(nom)
        if chemin:
            return chemin
    for chemin in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                   r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                   "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                   "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice"):
        if os.path.exists(chemin):
            return chemin
    return None


def _word_disponible():
    return sys.platform in ("win32", "darwin") and importlib.util.find_spec("docx2pdf") is not None


def moteur_pdf():
    """Nom du moteur de conversion disponible, ou None."""
    if _word_disponible():
        return "Microsoft Word (docx2pdf)"
    if _libreoffice():
        return "LibreOffice"
    return None


def _convertir_word(docx, pdf):
    from docx2pdf import convert
    try:                                   # Windows : Word est piloté par COM, dans un fil d'exécution
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    convert(os.path.abspath(docx), os.path.abspath(pdf))


def _convertir_libreoffice(docx, pdf):
    exe = _libreoffice()
    with tempfile.TemporaryDirectory() as tmp:
        sortie = os.path.join(tmp, "sortie")
        os.makedirs(sortie)
        profil = Path(os.path.join(tmp, "profil")).as_uri()   # profil isolé : évite les conflits si LibreOffice est ouvert
        cmd = [exe, f"-env:UserInstallation={profil}", "--headless", "--norestore", "--convert-to", "pdf",
               "--outdir", sortie, os.path.abspath(docx)]
        options = {"capture_output": True, "text": True, "timeout": DELAI_CONVERSION}
        if sys.platform == "win32":
            options["creationflags"] = 0x08000000   # CREATE_NO_WINDOW : pas de fenêtre noire
        resultat = subprocess.run(cmd, **options)
        produit = os.path.join(sortie, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
        if not os.path.exists(produit):
            raise RuntimeError((resultat.stderr or resultat.stdout or "aucun PDF produit").strip()[:300])
        shutil.copyfile(produit, pdf)


def convertir_en_pdf(chemin_docx, forcer=False):
    """Crée le PDF d'un bon Word. -> (ok, chemin du PDF  ou  message d'erreur)"""
    if not os.path.exists(chemin_docx):
        return False, "Fichier Word introuvable."
    pdf = chemin_pdf(chemin_docx)
    if not forcer and pdf_a_jour(chemin_docx):
        return True, pdf
    debut = time.time() - 1
    erreurs = []
    moteurs = []
    if _word_disponible():
        moteurs.append(("Microsoft Word", _convertir_word))
    if _libreoffice():
        moteurs.append(("LibreOffice", _convertir_libreoffice))
    for nom, fonction in moteurs:
        try:
            fonction(chemin_docx, pdf)
            if os.path.exists(pdf) and os.path.getmtime(pdf) >= debut:
                return True, pdf
            erreurs.append(f"{nom} : PDF non créé")
        except Exception as e:
            erreurs.append(f"{nom} : {e}")
    if erreurs:
        return False, "Conversion PDF impossible — " + " | ".join(erreurs)
    return False, ("Aucun convertisseur PDF trouvé. Installez LibreOffice (gratuit : libreoffice.org) "
                   "ou, si Microsoft Word est installé, lancez : pip install docx2pdf — puis relancez l'application.")


# ==================== Affichage Streamlit ====================
def afficher_moteur():
    """Petit rappel sur le moteur PDF (ou comment en installer un)."""
    import streamlit as st
    moteur = moteur_pdf()
    if moteur:
        st.caption(f"🖨️ Conversion PDF : {moteur}")
    else:
        st.warning("🖨️ Pour télécharger en PDF, installez **LibreOffice** (gratuit, libreoffice.org) ou, si Microsoft "
                   "Word est installé, lancez `pip install docx2pdf`. Puis relancez l'application. "
                   "Le téléchargement Word reste disponible.")


def afficher_erreur_pdf():
    import streamlit as st
    message = st.session_state.pop("_erreur_pdf", None)
    if message:
        st.error(message)


def afficher_boutons(chemin_docx, cle, conteneur_word, conteneur_pdf, avec_creation=True):
    """Deux boutons : ⬇️ Word et ⬇️ PDF (ou « Créer PDF » s'il n'existe pas encore)."""
    import streamlit as st
    nom = os.path.basename(chemin_docx)
    with conteneur_word:
        with open(chemin_docx, "rb") as f:
            st.download_button("⬇️ Word", data=f.read(), file_name=nom, mime=MIME_WORD, key=f"w_{cle}",
                               width="stretch")
    with conteneur_pdf:
        pdf = chemin_pdf(chemin_docx)
        if pdf_a_jour(chemin_docx):
            with open(pdf, "rb") as f:
                st.download_button("⬇️ PDF", data=f.read(), file_name=os.path.basename(pdf), mime=MIME_PDF,
                                   key=f"p_{cle}", width="stretch")
        elif not avec_creation:
            st.button("PDF indisponible", key=f"x_{cle}", disabled=True, width="stretch")
        else:
            libelle = "🔄 Refaire le PDF" if os.path.exists(pdf) else "📄 Créer le PDF"
            if st.button(libelle, key=f"c_{cle}", width="stretch"):
                with st.spinner("Conversion en PDF…"):
                    ok, resultat = convertir_en_pdf(chemin_docx, forcer=True)
                if not ok:
                    st.session_state["_erreur_pdf"] = resultat
                st.rerun()


def afficher_apres_generation(chemin_docx, cle):
    """Juste après la création d'un bon : crée le PDF tout de suite et propose Word + PDF."""
    import streamlit as st
    with st.spinner("Création du PDF…"):
        ok, resultat = convertir_en_pdf(chemin_docx)
    c1, c2, _ = st.columns([1, 1, 3])
    afficher_boutons(chemin_docx, cle, c1, c2, avec_creation=False)
    if not ok:
        st.warning(resultat)


def afficher_tout_creer(fichiers, cle):
    """Bouton « Créer tous les PDF manquants »."""
    import streamlit as st
    manquants = [f for f in fichiers if not pdf_a_jour(f)]
    if not manquants or not moteur_pdf():
        return
    if st.button(f"📄 Créer les {len(manquants)} PDF manquant(s)", key=f"tout_{cle}"):
        barre = st.progress(0.0, text="Conversion en PDF…")
        echecs, dernier_message = 0, ""
        for i, f in enumerate(manquants, start=1):
            ok, resultat = convertir_en_pdf(f)
            if not ok:
                echecs, dernier_message = echecs + 1, resultat
            barre.progress(i / len(manquants), text=f"Conversion en PDF… {i}/{len(manquants)}")
        if echecs:
            st.session_state["_erreur_pdf"] = f"{echecs} PDF non créé(s). {dernier_message}"
        st.rerun()
