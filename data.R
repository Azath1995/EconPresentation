# Tableau de bord macroéconomique — France (Shiny)

Application R Shiny qui suit la conjoncture française avec des données publiques
téléchargées depuis [DBnomics](https://db.nomics.world) (Eurostat, BCE, INSEE).

## Installation

```r
install.packages(c("shiny", "bslib", "plotly", "httr", "jsonlite"))
```

`bslib` doit être en version 0.6 ou plus récente.

## Lancement

```r
shiny::runApp("tableau_macro_france")
```

Au premier lancement, les séries sont téléchargées puis gardées 12 h dans le
dossier `cache/`. Sans connexion, l'application réutilise la dernière copie.

Mode démonstration hors ligne (données **fictives**, pour tester l'interface) :

```r
Sys.setenv(MACRO_DEMO = "1"); shiny::runApp("tableau_macro_france")
```

## Contenu

| Onglet | Ce qu'on y trouve |
|---|---|
| Synthèse | 6 indicateurs de une avec code couleur et libellé (PIB, chômage, inflation, solde public, dette, écart OAT-Bund) et 4 graphiques |
| Activité | Croissance, PIB en volume, contributions à la croissance, production industrielle, climat des affaires, confiance par secteur |
| Ménages | Consommation, ventes au détail, confiance des consommateurs, taux d'épargne et d'investissement |
| Emploi | Chômage (total et jeunes), emploi, taux d'emploi, salaires comparés à l'inflation, défaillances |
| Prix | Inflation globale et sous-jacente, par composante, prix de production |
| Finances publiques | Solde, dette, charge d'intérêts, dépenses et recettes |
| Taux et crédit | Taux à 10 ans, taux de dépôt BCE, écart France-Allemagne, taux et croissance des crédits |
| Extérieur | Solde courant, biens et services, exportations et importations |
| Comparaison | France, Allemagne, Italie, Espagne, zone euro sur un même graphique, plus un tableau des dernières valeurs |
| Explorateur | Ajout de n'importe quelle série DBnomics par son identifiant |
| Diagnostic | Liste des séries qui répondent et de l'identifiant retenu |

## Graphiques modulables

Chaque carte a un bouton **Régler** qui permet de :

- choisir les séries affichées, parmi tout le catalogue ;
- appliquer une transformation : niveau, variation sur la période, variation
  sur un an, base 100 au début de la fenêtre (les taux varient en points, les
  niveaux en %) ;
- passer en lignes, barres ou aires ;
- lisser par moyenne mobile (3, 6 ou 12 périodes) ;
- tracer la moyenne historique en pointillés ;
- exporter les données affichées en CSV.

La période se règle pour tous les graphiques dans la barre latérale.

## Ajouter ou corriger une série

Tout est dans `R/catalogue.R`. Chaque indicateur liste des identifiants
DBnomics candidats, essayés dans l'ordre : si Eurostat change d'année de base
(par exemple `I15` vers `I21`), la candidate suivante prend le relais.
Après une modification, lancez l'onglet **Diagnostic** pour vérifier.

## Limites

- Les indices PMI (S&P Global) et les notations souveraines ne sont pas en
  accès libre et ne figurent donc pas dans l'application.
- Les seuils de couleur de la synthèse (2 % d'inflation, 3 % de déficit,
  60 % de dette...) sont des repères simples, modifiables dans `app.R` (`KPIS`).
