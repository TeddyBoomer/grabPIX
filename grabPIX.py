"""PIX grabber
v0.0.4

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
comps =["1.1", "1.2", "1.3", "2.1", "2.2", "2.3", "2.4", "3.1", "3.2", "3.3",
        "3.4", "4.1", "4.2", "4.3", "5.1", "5.2"]
# 0 et - mis à NA pour le décompte des compétences
na_comps =dict([[e, ["-", "0"]] for e in comps]) 
# liste des classes pour régler la hauteur des images
classes = []

for e in os.listdir("."):
    if e.endswith(".csv"):
        d = pd.read_csv(e, sep=";", usecols=["Statut", "Nombre de Pix"] + comps,
                        na_values=na_comps)
        classe = e.split("_")[2].split(".")[0] # vilaine capture classe
        classes.append(classe)
        d.insert(2, "classe", classe)
        d.insert(3, "groupe", ginv[classe])
        data.append(d)

pd.set_option("display.precision", 0)
# regroupement en DataFrame
df = pd.concat(data, ignore_index=True).convert_dtypes()
df.sort_values(by=["groupe", "classe"], inplace=True)
#nb de compétences par élève
nbcomps = df[comps].count(axis=1)
df.insert(20, "Nombre de compétences", nbcomps)


##############
# graphiques #
##############
sns.set_theme(style="whitegrid")
for val in ["Nombre de Pix", "Nombre de compétences"]:
    F = plt.figure(figsize=(7, 1.5 + 0.5*len(classes)))
    ax = F.add_axes(rect=[0.15, 0.1, 0.80, 0.85], adjustable='datalim')
    _ = sns.boxplot(x=val, y="classe", data=df, palette="Set3",
                     hue="groupe", whis=(0,100), ax=ax)
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
                                                         sheet_name="valides")
    # résumé des validés/rejetés
    a = df.groupby(["groupe", "classe"])["Statut"].value_counts()
    dh = pd.concat([a[:,:,"Validée"], a[:,:,"Rejetée"]], axis=1,
                   keys=["Validée", "Rejetée"]).replace(np.nan, 0).convert_dtypes()
    dh.to_excel(writer, sheet_name="repartition")
