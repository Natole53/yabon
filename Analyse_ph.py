import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np

FICHIER_ENTRANT = "Annuel.ods"
FICHIER_SORTIE = "Resultats_analyse.xlsx"

# VÉRIFICATION
if not os.path.exists(FICHIER_ENTRANT):
    raise FileNotFoundError("Le fichier Annuel.ods doit déjà exister.")

# ===== LECTURE DES TABLEAUX =====
df_data = pd.read_excel(
    FICHIER_ENTRANT,
    engine="odf",
    header=4
)


# FONCTION : ETUDE D'UNE LIGNE DU TABLEAU SOURCE 
def etudier_ligne_source(
    ligne_source,
    res_format_37,
    res_code_produit_37,
    res_batch_37,
    res_format_55,
    res_code_produit_55,
    res_batch_55):

    format = ligne_source.iloc[0]
    code_produit = ligne_source.iloc[2]
    batch = ligne_source.iloc[5]

    valid_37_8j = ligne_source.iloc[11]
    valid_55_8j = ligne_source.iloc[13]

    def maj_dico(dico, cle, valid):
        if cle not in dico:
            dico[cle] = {"c": 0, "nv": 0,"total" : 0}
        if valid == "c":
            dico[cle]["c"] += 1
        elif valid == "nv":
            dico[cle]["nv"] += 1
        dico[cle]["total"] += 1

    # Fonction pour nettoyer la valeur
    def clean_valid(val):
        if pd.isna(val):
            return None
        val = str(val).strip().lower()
        if val in ["c", "nv"]:
            return val
        return None

    # TRAITEMENT 37°C
    valid_37_8j = clean_valid(valid_37_8j)
    if valid_37_8j:
        maj_dico(res_format_37, format, valid_37_8j)
        maj_dico(res_code_produit_37, code_produit, valid_37_8j)
        maj_dico(res_batch_37, batch, valid_37_8j)

    # TRAITEMENT 55°C
    valid_55_8j = clean_valid(valid_55_8j)
    if valid_55_8j:
        maj_dico(res_format_55, format, valid_55_8j)
        maj_dico(res_code_produit_55, code_produit, valid_55_8j)
        maj_dico(res_batch_55, batch, valid_55_8j)

# FONCTION : CONSTRUCTION D'UNE LIGNE POUR LE TABLEAU DE SORTIE
res_format_37 = {}
res_code_produit_37 = {}
res_batch_37 = {}
res_format_55 = {}
res_code_produit_55 = {}
res_batch_55 = {}

for _, ligne_source in df_data.iterrows():
    etudier_ligne_source(
        ligne_source,
        res_format_37,
        res_code_produit_37,
        res_batch_37,
        res_format_55,
        res_code_produit_55,
        res_batch_55
    )

# FONCTION : TRI DICO PAR NOMBRE DE TOTAL
def sort(dico):
    return sorted(dico.items(), key=lambda x: x[1]["total"], reverse=True)

# FONCTION : CONVERSION DICO EN DATAFRAME
def dico_to_dataframe(dico):

    data = []
    for cle, valeurs in dico.items():
        data.append({
            "Categorie": cle,
            "Conformes": valeurs["c"],
            "Non conformes": valeurs["nv"],
            "Total": valeurs["total"]
        })

    df = pd.DataFrame(data)
    df = df.sort_values("Total", ascending=False)

    return df

df_format_37 = dico_to_dataframe(res_format_37)
df_code_37 = dico_to_dataframe(res_code_produit_37)
df_batch_37 = dico_to_dataframe(res_batch_37)

df_format_55 = dico_to_dataframe(res_format_55)
df_code_55 = dico_to_dataframe(res_code_produit_55)
df_batch_55 = dico_to_dataframe(res_batch_55)

with pd.ExcelWriter(FICHIER_SORTIE) as writer:

    df_format_37.to_excel(writer, sheet_name="Format 37C", index=False)
    df_code_37.to_excel(writer, sheet_name="Code produit 37C", index=False)
    df_batch_37.to_excel(writer, sheet_name="Batch 37C", index=False)

    df_format_55.to_excel(writer, sheet_name="Format 55C", index=False)
    df_code_55.to_excel(writer, sheet_name="Code produit 55C", index=False)
    df_batch_55.to_excel(writer, sheet_name="Batch 55C", index=False)

print("Fichier Excel généré :", FICHIER_SORTIE)

# FONCTION GRAPHIQUE PARETO
def graphique_pareto(dico, titre):

    # On garde uniquement les non-valides
    nv_data = {k: v["nv"] for k, v in dico.items() if v["nv"] > 0}

    if not nv_data:
        print(f"Aucune non-validité pour : {titre}")
        return

    # Tri décroissant
    nv_data = dict(sorted(nv_data.items(), key=lambda x: x[1], reverse=True))

    categories = list(nv_data.keys())
    valeurs = list(nv_data.values())

    # Calcul cumul %
    cumul = np.cumsum(valeurs)
    cumul_pct = cumul / cumul[-1] * 100

    fig, ax1 = plt.subplots(figsize=(10, 6))

    bars = ax1.bar(categories, valeurs, color="skyblue")
    ax1.set_ylabel("Nombre de non-valides")
    ax1.set_xlabel("Catégories")
    ax1.set_title(titre)
    ax1.tick_params(axis='x', rotation=45)

    for bar in bars:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2,
                 height,
                 f"{int(height)}",
                 ha='center',
                 va='bottom')

    ax2 = ax1.twinx()
    ax2.plot(categories, cumul_pct, color="red", marker='o')
    ax2.set_ylabel("Pourcentage cumulatif (%)")
    ax2.set_ylim(0, 110)
    ax2.axhline(80, color="green", linestyle="--")

    plt.tight_layout()
    plt.show()

graphique_pareto(res_code_produit_37, "Pareto des non-valides - Code produit 37°C")
graphique_pareto(res_code_produit_55, "Pareto des non-valides - Code produit 55°C")

graphique_pareto(res_format_37, "Pareto des non-valides - Format 37°C")
graphique_pareto(res_format_55, "Pareto des non-valides - Format 55°C")

graphique_pareto(res_batch_37, "Pareto des non-valides - Batch 37°C")
graphique_pareto(res_batch_55, "Pareto des non-valides - Batch 55°C")