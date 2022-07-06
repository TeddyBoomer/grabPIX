"""PIX grabber
v0.0.6

Ce script permet d'aggréger et résumer les résultats PIX de plusieurs classes:

1. on exporte les CSV depuis orga.pix.fr dans un même dossier
2. on y place ce script, on l'exécute

Sortie:
-------

1. Un diagramme en boîte et moustache des répartitions des nombres de PIX des
participants qui ont validé la certification: PIX.png

2. Un fichier tableur
  * onglet valides: les résumés statistiques des notes pour les validés, par classe
  * onglet repartition: nombre de validés/rejetés par classe

Prérequis Python:
-----------------

On utilise essentiellement les modules numpy, pandas, matplotlib, seaborn
donc à installer au préalable:
(linux)
> pip install numpy pandas matplotlib seaborn
(win)
> py -3 -m pip install numpy pandas matplotlib seaborn

GNU GPL v3.0

"""

import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import seaborn as sns
from functools import reduce
# from mpl_toolkits.axes_grid1.axes_divider import make_axes_area_auto_adjustable

data = []
# renseigner les groupes de classes ici
groups = { "3e": ["3PM"],
           "tgen": ["TGEN1", "TGEN2", "TGEN3"],
           "techno": ["TST2S1", "TST2S2", "TSTMG1", "TSTMG2"],
           "tpro": ["TASSP", "TEPC1", "TEPC2", "TGA", "TMCV", "TOL"],
           "bts": ["SIO2", "SIO2A", "NDRC2"]}

ginv = dict(reduce(lambda a,b: a+b, [[[e, k] for e in v] for k,v in groups.items()]))

identite = ["Prénom", "Nom", "Date de naissance"]
comps =["1.1", "1.2", "1.3", "2.1", "2.2", "2.3", "2.4", "3.1", "3.2", "3.3",
        "3.4", "4.1", "4.2", "4.3", "5.1", "5.2"]
date_certif = "Date de passage de la certification"
fields = identite + ["Statut", "Nombre de Pix"] + comps + [date_certif]
# 0 et - mis à NA pour le décompte des compétences
na_comps =dict([[e, ["-", "0"]] for e in comps]) 
# liste des classes pour régler la hauteur des images
classes = []

for e in os.listdir("."):
    if e.endswith(".csv"):
        d = pd.read_csv(e, sep=";", usecols=fields,
                        na_values=na_comps, parse_dates=[len(fields)-1])
        classe = e.split("_")[2].split(".")[0] # vilaine capture classe
        classes.append(classe)
        d.insert(2, "classe", classe)
        d.insert(3, "groupe", ginv[classe])
        data.append(d)

pd.set_option("display.precision", 0)
# regroupement en DataFrame
df = pd.concat(data, ignore_index=True).convert_dtypes()
#nb de compétences par élève
nbcomps = df[comps].count(axis=1)
df.insert(23, "Nombre de compétences", nbcomps)

#########################################
# stratégie de purge des doublons élève #
#########################################
# id unique: prénom nom date de naissance
pndn = df[["Prénom", "Nom", "Date de naissance"]].aggregate(''.join, axis=1)
df.insert(25, "id", pndn)
# purge des doublons les plus anciens
df.sort_values(by=date_certif, ascending=False, inplace=True)
df.drop_duplicates(subset="id", keep="first", inplace=True)
# pour une purge sur le nb de pix: trier sur "Nombre de Pix"
# df.sort_values(by="Nombre de Pix", ascending=False, inplace=True)
# df.drop_duplicates(subset="id", keep="first", inplace=True)

# réorganisation correcte des données
df.sort_values(by=["groupe", "classe"], inplace=True)

##############
# graphiques #
##############
sns.set_theme(style="whitegrid")
for val in ["Nombre de Pix", "Nombre de compétences"]:
    F = plt.figure(figsize=(7, 1.5 + 0.5*len(classes)))
    ax = F.add_axes(rect=[0.15, 0.1, 0.60, 0.85], adjustable='datalim')
    _ = sns.boxplot(x=val, y="classe", data=df, palette="Set3",
                     hue="groupe", whis=(0,100), ax=ax)
    ax.legend(loc='upper left', bbox_to_anchor=(1.02, 0.8))
    # whis: réglage moustaches vers Q0=min, Q4=max
    #plt.xticks(rotation=90)
    #make_axes_area_auto_adjustable(ax)
    # plt.show()
    plt.savefig(f"PIX-{val.replace(' ', '_')}.png")
    del F

##################
# export tableur #
##################
df.set_index(["groupe", "classe"], inplace=True)
dg = df[["Nombre de Pix", "Nombre de compétences"]]

with pd.ExcelWriter("statsPIX.xlsx", engine='xlsxwriter') as writer: # xlsx
    # résumé des niveaux par groupe et classe
    dg.groupby(["groupe", "classe"]).describe().to_excel(writer,
                                                         float_format="%.2f",
                                                         sheet_name="valides")

    # résumé des niveaux par groupe
    dg.groupby("groupe").describe().to_excel(writer,
                                             float_format="%.2f",
                                             sheet_name="valides",
                                             startrow=len(classes)+4)

    # résumé global des niveaux 
    dg.describe().to_excel(writer,
                           float_format="%.2f",
                           sheet_name="valides",
                           startrow=len(classes)+len(groups)+2*4)
    
    # résumé des validés/rejetés
    ALL = df["Statut"].value_counts()
    ALL.to_excel(writer, sheet_name="repartition")
    a = df.groupby(["groupe", "classe"])["Statut"].value_counts()
    dh = pd.concat([a[:,:,"Validée"], a[:,:,"Rejetée"]], axis=1,
                   keys=["Validée", "Rejetée"]).replace(np.nan, 0).convert_dtypes()
    dh.to_excel(writer, sheet_name="repartition", startrow=4)
