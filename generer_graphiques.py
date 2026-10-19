"""Génère les graphiques du sujet 1 (pauvreté en France) dans le dossier graphiques/."""
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

OUT = Path(__file__).parent / "graphiques"
OUT.mkdir(exist_ok=True)

# Palette (rôles)
SURFACE = "#fcfcfb"
TEXT = "#0b0b0b"
TEXT_2 = "#52514e"
MUTED = "#8a8984"
GRID = "#e6e5e1"
BLUE = "#2a78d6"
BLUE_LIGHT = "#cde2fb"
ORANGE = "#eb6834"
AQUA = "#1baf7a"
NEUTRAL = "#c9c8c2"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 15,
    "text.color": TEXT,
    "axes.edgecolor": GRID,
    "axes.labelcolor": TEXT_2,
    "xtick.color": TEXT_2,
    "ytick.color": TEXT_2,
    "figure.facecolor": SURFACE,
    "axes.facecolor": SURFACE,
    "savefig.facecolor": SURFACE,
})

W, H = 16, 9  # 16:9 pour les diapos

SRC_ECO = "Paul Crauchet, « La face cachée du modèle Macron révélée par l'Insee », L'Essentiel de l'Éco, 08/07/2025"
SRC_IREF = "Élodie Messéant, « Hausse du taux de pauvreté : le socialisme ne fonctionne pas », IREF, 26/08/2025"
SRC_OFCE = "Pierre Madec, « Pauvreté monétaire en France : les causes d'un décrochage inquiétant », OFCE, 27/03/2026"


def fr(x, dec=1):
    """Format français : virgule décimale, espace pour les milliers."""
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")


def titre(fig, t, sous_titre=None):
    fig.text(0.05, 0.93, t, fontsize=28, fontweight="bold", color=TEXT, va="top")
    if sous_titre:
        fig.text(0.05, 0.855, sous_titre, fontsize=17, color=TEXT_2, va="top")


def source(fig, s):
    fig.text(0.05, 0.035, "Source : " + s, fontsize=12, color=MUTED)


def axes_sobres(ax, garder_bas=True):
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_visible(garder_bas)
    ax.tick_params(length=0)


def save(fig, nom):
    fig.savefig(OUT / nom, dpi=120)
    plt.close(fig)
    print("OK", nom)


def boite(ax, x, y, w, h, texte, fond, bord=None, couleur_texte=TEXT, taille=16, gras=False):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=0.02",
                                fc=fond, ec=bord or fond, lw=2))
    ax.text(x + w / 2, y + h / 2, texte, ha="center", va="center", fontsize=taille,
            color=couleur_texte, fontweight="bold" if gras else "normal")


def fleche(ax, x1, y1, x2, y2):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", color=TEXT_2, lw=2, mutation_scale=22))


def schema(fig):
    ax = fig.add_axes([0.03, 0.08, 0.94, 0.74])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    return ax


# ---------------------------------------------------------------------------
# 1. Accroche : évolution du taux de pauvreté
# ---------------------------------------------------------------------------
def g1_taux():
    annees = [2017, 2022, 2023]
    taux = [13.7, 14.5, 15.4]

    fig = plt.figure(figsize=(W, H))
    titre(fig, "Le taux de pauvreté au plus haut depuis 1996",
          "Part de la population sous le seuil de pauvreté (60 % du niveau de vie médian), en %")
    ax = fig.add_axes([0.07, 0.13, 0.86, 0.66])
    ax.plot(annees, taux, color=BLUE, lw=2.5, zorder=2)
    ax.scatter(annees, taux, s=120, color=BLUE, edgecolor=SURFACE, linewidth=2, zorder=3)
    for a, t in zip(annees, taux):
        ax.annotate(fr(t) + " %", (a, t), xytext=(0, 16), textcoords="offset points",
                    ha="center", fontsize=17, color=TEXT_2, fontweight="bold" if a == 2023 else "normal")
    ax.annotate("+ 650 000 pauvres\nen un an", xy=(2022.5, 14.9), xytext=(2020.6, 11.4),
                fontsize=16, color=TEXT, ha="center",
                arrowprops=dict(arrowstyle="-", color=MUTED, lw=1.5))
    ax.text(2023, 11.6, "9,8 millions\nde personnes", ha="center", fontsize=20,
            fontweight="bold", color=TEXT)
    ax.set_ylim(0, 18)
    ax.set_xlim(2016.3, 2023.8)
    ax.set_xticks(annees)
    ax.set_yticks([0, 5, 10, 15])
    ax.yaxis.grid(True, color=GRID, lw=1)
    ax.set_axisbelow(True)
    axes_sobres(ax)
    source(fig, "Insee (juillet 2025) ; L'Essentiel de l'Éco (08/07/2025). Seules les années citées sont représentées.")
    save(fig, "1_taux_pauvrete.png")


# ---------------------------------------------------------------------------
# 2. Constat : qui est pauvre ?
# ---------------------------------------------------------------------------
def g2_qui():
    items = [
        ("Inactifs (hors retraités)", 37.3),
        ("Chômeurs", 36.1),
        ("Familles monoparentales", 34.3),
        ("Micro-entrepreneurs", 19.2),
        ("Ensemble de la population", 15.4),
    ][::-1]
    fig = plt.figure(figsize=(W, H))
    titre(fig, "Qui est pauvre en France ?",
          "Taux de pauvreté par catégorie en 2023, en %")
    ax = fig.add_axes([0.32, 0.12, 0.60, 0.68])
    labels = [i[0] for i in items]
    vals = [i[1] for i in items]
    couleurs = [NEUTRAL if l.startswith("Ensemble") else BLUE for l in labels]
    ax.barh(labels, vals, color=couleurs, height=0.6)
    for y, v in enumerate(vals):
        ax.text(v + 0.6, y, fr(v) + " %", va="center", fontsize=18, fontweight="bold")
    ax.axvline(15.4, color=TEXT_2, lw=1.5, ls="--", zorder=0)
    ax.set_xlim(0, 45)
    ax.set_xticks([])
    ax.tick_params(axis="y", labelsize=17, labelcolor=TEXT)
    axes_sobres(ax, garder_bas=False)
    source(fig, "Insee (juillet 2025), chiffres cités par L'Essentiel de l'Éco (08/07/2025)")
    save(fig, "2_qui_est_pauvre.png")


# ---------------------------------------------------------------------------
# 3a. Point de vue 1 : les inégalités se creusent (Gini)
# ---------------------------------------------------------------------------
def g3_gini():
    fig = plt.figure(figsize=(W, H))
    titre(fig, "Gauche : « les inégalités se creusent depuis 2017 »",
          "Indice de Gini (0 = égalité parfaite, 1 = une personne a tout)")
    ax = fig.add_axes([0.07, 0.24, 0.86, 0.50])
    ax.set_xlim(0.25, 0.32)
    ax.set_ylim(-1, 1)
    ax.hlines(0, 0.25, 0.32, color=GRID, lw=6)
    ax.annotate("", xy=(0.2965, 0), xytext=(0.2815, 0),
                arrowprops=dict(arrowstyle="-|>", color=ORANGE, lw=3, mutation_scale=30))
    for x, lab, coul in [(0.281, "2017\n0,281", BLUE), (0.297, "2023\n0,297", ORANGE)]:
        ax.scatter([x], [0], s=500, color=coul, edgecolor=SURFACE, linewidth=3, zorder=3)
        ax.text(x, 0.25, lab, ha="center", va="bottom", fontsize=22, fontweight="bold")
    ax.set_xticks([0.25, 0.26, 0.27, 0.28, 0.29, 0.30, 0.31, 0.32])
    ax.set_xticklabels([fr(t, 2) for t in ax.get_xticks()], fontsize=14)
    ax.set_yticks([])
    axes_sobres(ax, garder_bas=False)
    fig.text(0.07, 0.13, "Plus l'indice est élevé, plus la société est inégalitaire. "
             "L'axe est zoomé : il ne part pas de 0.", fontsize=16, color=TEXT_2)
    source(fig, SRC_ECO)
    save(fig, "3a_gini.png")


# ---------------------------------------------------------------------------
# 3b. Point de vue 1 : le chiffre officiel sous-estime la pauvreté
# ---------------------------------------------------------------------------
def g3_sous_estime():
    fig = plt.figure(figsize=(W, H))
    titre(fig, "Gauche : « le chiffre officiel sous-estime la pauvreté »",
          "Nombre de personnes pauvres, en millions")
    ax = fig.add_axes([0.07, 0.22, 0.86, 0.42])
    ax.barh([0], [9.8], color=BLUE, height=0.5)
    ax.barh([0], [2.1], left=[9.8 + 0.06], color=ORANGE, height=0.5)
    ax.text(4.9, 0, "9,8 millions comptés par l'Insee", ha="center", va="center",
            fontsize=19, color="#ffffff", fontweight="bold")
    ax.text(9.8 + 0.06 + 1.05, 0, "+ 2,1", ha="center", va="center",
            fontsize=19, color="#ffffff", fontweight="bold")
    ax.text(11.95, 0.42, "≈ 12 millions", ha="right", va="bottom", fontsize=24, fontweight="bold")
    ax.set_xlim(0, 12.2)
    ax.set_ylim(-0.6, 0.9)
    ax.set_xticks([0, 2, 4, 6, 8, 10, 12])
    ax.set_yticks([])
    axes_sobres(ax)
    fig.text(0.07, 0.11, "Non comptés : sans-abri, résidents de foyers, étudiants précaires, outre-mer",
             fontsize=17, color=TEXT_2)
    source(fig, SRC_ECO)
    save(fig, "3b_pauvrete_sous_estimee.png")


# ---------------------------------------------------------------------------
# 4a. Point de vue 2 : la pauvreté relative (exemple fictif)
# ---------------------------------------------------------------------------
def g4_relative():
    fig = plt.figure(figsize=(W, H))
    titre(fig, "Libéral : « la pauvreté relative est trompeuse »",
          "Exemple fictif : un ménage modeste dans deux pays imaginaires, revenu mensuel en €")
    ax = fig.add_axes([0.10, 0.14, 0.80, 0.62])
    pays = ["Pays A\npauvre mais égalitaire", "Pays B\nriche mais inégalitaire"]
    revenu = [500, 1200]
    seuil = [420, 1500]
    ax.bar([0, 1], revenu, color=[BLUE, BLUE], width=0.45)
    for x, (r, s) in enumerate(zip(revenu, seuil)):
        ax.hlines(s, x - 0.33, x + 0.33, color=ORANGE, lw=4)
        ax.text(x + 0.35, s, f"seuil : {fr(s, 0)} €", va="center", fontsize=16, color=TEXT)
        ax.text(x, r / 2, f"{fr(r, 0)} €", ha="center", va="center", fontsize=22,
                color="#ffffff", fontweight="bold")
    ax.text(0, 640, "PAS pauvre", ha="center", fontsize=20, fontweight="bold", color=TEXT)
    ax.text(1, 1640, "« PAUVRE »", ha="center", fontsize=20, fontweight="bold", color=TEXT)
    ax.set_xticks([0, 1])
    ax.set_xticklabels(pays, fontsize=17, color=TEXT)
    ax.set_xlim(-0.5, 1.75)
    ax.set_ylim(0, 1900)
    ax.set_yticks([])
    axes_sobres(ax)
    fig.text(0.05, 0.79, "Le ménage de B gagne 2,4 fois plus que celui de A… mais c'est lui qui est compté pauvre.",
             fontsize=16, color=TEXT_2)
    source(fig, "chiffres inventés pour illustrer l'argument de l'IREF (26/08/2025). Seuil = 60 % du revenu médian du pays.")
    save(fig, "4a_pauvrete_relative.png")


# ---------------------------------------------------------------------------
# 4b. Point de vue 2 : la France redistribue déjà énormément
# ---------------------------------------------------------------------------
def g4_redistribution():
    fig = plt.figure(figsize=(W, H))
    titre(fig, "Libéral : « la France redistribue déjà énormément »",
          "Combien de fois plus gagnent les plus aisés que les plus modestes ?")
    ax = fig.add_axes([0.20, 0.16, 0.60, 0.62])
    labels = ["Avant impôts\net aides", "Après impôts\net aides"]
    vals = [18, 3]
    ax.bar(labels, vals, color=[NEUTRAL, BLUE], width=0.5)
    for x, v in enumerate(vals):
        ax.text(x, v + 0.4, f"× {v}", ha="center", fontsize=28, fontweight="bold")
    ax.set_ylim(0, 21)
    ax.set_yticks([])
    ax.tick_params(axis="x", labelsize=18, labelcolor=TEXT)
    axes_sobres(ax)
    fig.text(0.60, 0.62, "Prélèvements obligatoires\nFrance : 45,6 % du PIB\nUE : ≈ 40 %",
             fontsize=18, color=TEXT_2, va="top")
    source(fig, SRC_IREF)
    save(fig, "4b_redistribution.png")


# ---------------------------------------------------------------------------
# 5. Point de vue 3 : OFCE, quatre causes
# ---------------------------------------------------------------------------
def g5_ofce():
    fig = plt.figure(figsize=(W, H))
    titre(fig, "Milieu (OFCE) : quatre causes à séparer",
          "D'où vient la hausse du taux de pauvreté ?")
    ax = schema(fig)
    effets = [
        ("1. Le seuil", "monte avec le\nrevenu médian"),
        ("2. La démographie", "ex. plus de familles\nmonoparentales"),
        ("3. Les aides sociales", "sont-elles\nassez efficaces ?"),
        ("4. Les revenus\ndu travail", "emploi, salaires,\nprécarité"),
    ]
    for i, (t, d) in enumerate(effets):
        x = 0.02 + i * 0.245
        boite(ax, x, 0.62, 0.22, 0.30, "", "#ffffff", bord=GRID)
        ax.text(x + 0.11, 0.84, t, ha="center", va="center", fontsize=17, fontweight="bold")
        ax.text(x + 0.11, 0.71, d, ha="center", va="center", fontsize=14, color=TEXT_2)
        fleche(ax, x + 0.11, 0.62, 0.5, 0.42)
    boite(ax, 0.32, 0.22, 0.36, 0.20, "Taux de pauvreté : 15,4 %", AQUA,
          couleur_texte="#ffffff", taille=22, gras=True)
    ax.text(0.5, 0.08, "+ regarder le « reste à vivre » : ce qui reste après loyer, énergie et dépenses obligatoires",
            ha="center", fontsize=16, color=TEXT_2, style="italic")
    source(fig, SRC_OFCE)
    save(fig, "5_schema_ofce.png")


# ---------------------------------------------------------------------------
# 6. Synthèse
# ---------------------------------------------------------------------------
def g6_synthese():
    fig = plt.figure(figsize=(W, H))
    titre(fig, "Trois lectures d'un même chiffre", "15,4 % de pauvres : qu'en conclure ?")
    ax = schema(fig)
    colonnes = [
        ("Gauche\n(L'Essentiel de l'Éco)", BLUE,
         ["Des choix politiques\ndepuis 2017", "Aides coupées ou\nmal revalorisées",
          "Protéger davantage\nles plus modestes", "Le chiffre sous-estime\n(≈ 12 millions)"]),
        ("Milieu\n(OFCE)", AQUA,
         ["Plusieurs causes\nmêlées", "Efficacité\nà évaluer",
          "Analyser avant d'agir,\n« reste à vivre »", "Compléter le revenu\npar le reste à vivre"]),
        ("Libéral\n(IREF)", ORANGE,
         ["Trop d'État", "Créent dépendance,\ndécouragent le travail",
          "Liberté économique,\nemploi, investissement", "Le chiffre exagère :\nil mesure les inégalités"]),
    ]
    lignes = ["Diagnostic", "Les aides", "Solution", "La mesure"]
    wl, wc = 0.16, 0.27
    for j, l in enumerate(lignes):
        ax.text(0.01, 0.62 - j * 0.19 + 0.08, l, fontsize=17, fontweight="bold", color=TEXT_2, va="center")
    for i, (nom, coul, cells) in enumerate(colonnes):
        x = wl + i * (wc + 0.01)
        boite(ax, x, 0.82, wc, 0.16, nom, coul, couleur_texte="#ffffff", taille=16, gras=True)
        for j, c in enumerate(cells):
            boite(ax, x, 0.62 - j * 0.19, wc, 0.17, c, "#ffffff", bord=GRID, taille=15)
    source(fig, "synthèse des trois articles étudiés")
    save(fig, "6_synthese.png")


if __name__ == "__main__":
    for f in OUT.glob("*.png"):
        f.unlink()
    g1_taux()
    g2_qui()
    g3_gini()
    g3_sous_estime()
    g4_relative()
    g4_redistribution()
    g5_ofce()
    g6_synthese()
