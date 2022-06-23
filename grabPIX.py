"""PIX grabber

Ce script permet d'aggréger et résumé les résultats PIX de plusieurs classes:

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

On utilise essentiellement les modules pandas, matplotlib, seaborn
donc à installer au préalable:
(linux)
> pip install pandas matplotlib seaborn
(win)
> py -3 -m pip install pandas matplotlib seaborn

GNU GPL v3.0

"""

import pandas as pd
import os
import matplotlib.pyplot as plt
import seaborn as sns
from functools import reduce
from mpl_toolkits.axes_grid1.axes_divider import make_axes_area_auto_adjustable


data = []
# renseigner les groupes de classes ici
groups = { "3e": ["3PM"], "tgen": ["TGEN1", "TGEN2", "TGEN3"],
           "techno": ["TST2S1", "TST2S2", "TSTMG1", "TSTMG2"],
           "tpro": ["TASSP", "TEPC1", "TEPC2", "TGA", "TMCV", "TOL"],
           "bts": ["SIO2", "SIO2A", "NDRC2"]}

ginv = dict(reduce(lambda a,b: a+b, [[[e, k] for e in v] for k,v in groups.items()]))

for e in os.listdir("."):
    if e.endswith(".csv"):
        d = pd.read_csv(e, sep=";", usecols=["Statut", "Nombre de Pix"])
        classe = e.split("_")[2].split(".")[0] # vilaine capture classe
        d.insert(2, "classe", classe)
        d.insert(3, "groupe", ginv[classe])
        data.append(d)

pd.set_option("display.precision", 2)
# regroupement en DataFrame
df = pd.concat(data, ignore_index=True)
df.sort_values(by=["groupe", "classe"], inplace=True)

sns.set_theme(style="whitegrid")
plt.figure(figsize=(9,9))
ax = sns.boxplot(x="Nombre de Pix", y="classe", data=df, palette="Set3",
                 hue="groupe", whis=(0,100))
# whis: réglage moustaches vers Q0=min, Q4=max
#plt.xticks(rotation=90)
make_axes_area_auto_adjustable(ax)
# plt.show()
plt.savefig("PIX.png")

# export tableur
df.set_index(["groupe", "classe"], inplace=True)
with pd.ExcelWriter("statsPIX.xlsx", engine='xlsxwriter') as writer: # xlsx
    # résumé des niveaux par groupe et classe
    df.groupby(["groupe", "classe"]).describe().to_excel(writer,
                                                         sheet_name="valides")
    # résumé des validés/rejetés
    df.groupby("classe")["Statut"].value_counts().to_excel(writer,
                                                           sheet_name="repartition")


