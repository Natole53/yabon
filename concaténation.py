import pandas as pd
import os
from produits import produits


# CONFIGURATION
FICHIER_SORTIE = "Annuel.ods"
DOSSIER_SOURCES = "hebdo_2025"
SEMAINES = range(1, 52)

# VARIABLES GLOBALES
batch_pre = None
recettes_inconnues = 0
recettes_sans_ph_t0 = 0


# FONCTIONS PRODUITS
def get_format(produits, nom_produit):
    return produits[nom_produit]['format']

def get_code_produit(produits, nom_produit):
    return produits[nom_produit]['code produit']

def get_code_base(produits, nom_produit):
    return produits[nom_produit]['code base']

# CONSTRUCTION LIGNE SORTANTE
def construire_ligne_sortante(ligne_source, colonnes_sortie):
    global batch_pre
    global recettes_inconnues
    global recettes_sans_ph_t0

    nouvelle_ligne = {}

    recette = ligne_source.iloc[0]
    date_prod = ligne_source.iloc[1]
    batch = ligne_source.iloc[4]
    tournee = ligne_source.iloc[5]
    ph_t0 = ligne_source.iloc[11]
    ph_37_8j = ligne_source.iloc[24]
    ph_55_8j = ligne_source.iloc[28]

    # Nettoyage recette
    if isinstance(recette, str):
        recette = recette.strip()

    if recette not in produits:
        recettes_inconnues += 1
        return None

    # Format date
    if not pd.isna(date_prod):
        date_prod = pd.to_datetime(date_prod).strftime("%d/%m/%Y")

    # Complétion batch
    if pd.isna(batch):
        batch = batch_pre
    batch_pre = batch

    # Vérification pH T0
    if pd.isna(ph_t0):
        recettes_sans_ph_t0 += 1
        return None

    format_produit = get_format(produits, recette)
    code_base = get_code_base(produits, recette)
    code_produit = get_code_produit(produits, recette)

    for colonne in colonnes_sortie:
        if colonne == "Recette":
            nouvelle_ligne[colonne] = recette
        elif colonne == "Format":
            nouvelle_ligne[colonne] = format_produit
        elif colonne == "Code Produit":
            nouvelle_ligne[colonne] = code_produit
        elif colonne == "Code Base":
            nouvelle_ligne[colonne] = code_base
        elif colonne == "Date de Prod":
            nouvelle_ligne[colonne] = date_prod
        elif colonne == "Batch":
            nouvelle_ligne[colonne] = batch
        elif colonne == "Tournée":
            nouvelle_ligne[colonne] = tournee
        elif colonne == "pH T0":
            nouvelle_ligne[colonne] = ph_t0
        elif colonne == "pH 37°C 8j":
            nouvelle_ligne[colonne] = ph_37_8j
        elif colonne == "pH 55°C 8j":
            nouvelle_ligne[colonne] = ph_55_8j
        else:
            nouvelle_ligne[colonne] = None

    return nouvelle_ligne

# VÉRIFICATION FICHIER SORTIE
if not os.path.exists(FICHIER_SORTIE):
    raise FileNotFoundError("Le fichier Annuel.ods doit déjà exister.")

df_existant = pd.read_excel(
    FICHIER_SORTIE,
    engine="odf",
    header=2
)

# TRAITEMENT MULTI-FICHIERS
nouvelles_lignes = []

for semaine in SEMAINES:

    semaine_str = f"{semaine:02d}"  # Force 2 chiffres

    nom_fichier = f"TABLEAU 2025 ENCOURS - 24H - INCUBATION S{semaine_str}.xlsm"
    nom_feuille = f"S{semaine_str}"
    chemin_fichier = os.path.join(DOSSIER_SOURCES, nom_fichier)

    if not os.path.exists(chemin_fichier):
        print(f"⚠ {nom_fichier} introuvable → ignoré.")
        continue

    print(f"📂 Traitement {nom_fichier}")

    df_source = pd.read_excel(
        chemin_fichier,
        sheet_name=nom_feuille,
        skiprows=4
    )

    lignes_vides_consecutives = 0

    for _, ligne_source in df_source.iterrows():

        recette = ligne_source.iloc[0]

        # Détection lignes vides
        if pd.isna(recette) or (isinstance(recette, str) and recette.strip() == ""):
            lignes_vides_consecutives += 1
        else:
            lignes_vides_consecutives = 0

        # Arrêt après 2 lignes vides
        if lignes_vides_consecutives >= 2:
            print("⛔ Deux lignes vides → arrêt feuille")
            break

        ligne_sortante = construire_ligne_sortante(
            ligne_source,
            df_existant.columns
        )

        if ligne_sortante is not None:
            nouvelles_lignes.append(ligne_sortante)

# CONCATÉNATION FINALE
if nouvelles_lignes:

    df_nouveau = pd.DataFrame(nouvelles_lignes)

    # Forcer les mêmes colonnes et le même ordre
    df_nouveau = df_nouveau.reindex(columns=df_existant.columns)

    # Supprimer colonnes totalement vides
    df_nouveau = df_nouveau.dropna(axis=1, how="all")

    df_final = pd.concat(
        [df_existant, df_nouveau],
        ignore_index=True
    )

    df_final.to_excel(
        FICHIER_SORTIE,
        index=False,
        engine="odf",
        startrow=2
    )

    print("✅ Mise à jour terminée.")

else:
    print("⚠ Aucune nouvelle ligne ajoutée.")


print("\n📊 Résumé :")
print(f" - Recettes inconnues : {recettes_inconnues}")
print(f" - Recettes sans pH T0 : {recettes_sans_ph_t0}")
print(f" - Total ignorées : {recettes_inconnues + recettes_sans_ph_t0}")