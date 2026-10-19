# Présentation d'économie — Conjoncture et revue de presse

*15-min economics presentation: French macroeconomic dashboard (magic square, euro-area comparison) and press review of three debates (poverty, Zucman wealth tax, funded pensions). Script-generated PowerPoint, verified sources.*

Présentation de 15 minutes en trois parties : tableau de bord de la conjoncture française, trois sujets d'actualité (pauvreté, taxe Zucman, retraites et capitalisation), puis l'impact de l'actualité sur les indicateurs.

## Fichiers

| Fichier | Contenu |
|---|---|
| `Presentation_Economie.pptx` | Le PowerPoint final (29 diapos), généré par `generer_powerpoint.js` |
| `generer_powerpoint.js` | Script qui construit le PowerPoint (pptxgenjs) |
| `pitch_presentation.md` | Texte de l'oral complet et minuté (français) |
| `pitch_presentation_EN.md` | Le même texte en anglais, pour s'entraîner |
| `sujet1_pauvrete.md`, `sujet2_zucman.md`, `sujet3_retraites.md` | Textes détaillés de chaque sujet, avec les sources |
| `app.R`, `data.R` | Fichiers de l'application Shiny de tableau de bord (à vérifier : noms possiblement inversés) |
| `generer_graphiques.py`, `graphiques/` | Anciens graphiques PNG, qui ne servent plus au PowerPoint (`4a_pauvrete_relative.png` contient un exemple fictif) |

## Régénérer le PowerPoint

Prérequis : Node.js.

```bash
npm install
node generer_powerpoint.js
```

Fermer `Presentation_Economie.pptx` dans PowerPoint avant de lancer le script, sinon le fichier est verrouillé. Pour écrire ailleurs, définir la variable d'environnement `PPTX_OUT` avec un autre chemin.

Le script applique le thème de couleurs avec `apply_theme.js` (skill pptx de Claude Code). Le chemin est défini en haut de `generer_powerpoint.js`.

## À mettre à jour avant la présentation

- Les chiffres de marché de la diapo 2 : taux à 10 ans et écart avec l'Allemagne.
- Les publications qui sortent entre-temps : inflation et PIB d'octobre (Insee, 30 octobre), confiance des ménages (27 octobre), décision de Moody's (23 octobre).
