# Markdown to DOCX Converter

Ce script Python permet de convertir facilement des fichiers écrits en **Markdown** en documents Word au format `.docx`.  
Il est particulièrement adapté pour transformer des documents Markdown structurés (comme ceux générés par Claude.ai) en fichiers Word modifiables et bien formatés.

---

## Fonctionnalités principales

- Conversion des titres Markdown (`#`, `##`, `###`, `####`) en styles de titres Word hiérarchiques (H1, H2, H3, H4)  
- Gestion du formatage de texte :  
  - **Gras** (`**texte**`)  
  - *Italique* (`*texte*`)  
  - ***Gras + Italique*** (`***texte***`)  
  - `Code inline` (`\`code\``)  
- Support des listes à puces (`-`, `*`) et numérotées (`1.`, `2.`, ...)  
- Prise en charge des cases à cocher (✅, □)  
- Gestion des citations (`> citation`)  
- Conversion des tableaux Markdown en tableaux Word  
- Conservation des liens cliquables Markdown `[texte](URL)`  
- Possibilité d'adapter la mise en forme selon la structure de ton playbook Markdown (ex: niveaux de titre pour chapitres, sous-sections, détails)

---

## Contexte d'utilisation

Ce convertisseur est idéal si tu as des documents Markdown produits par des outils comme Claude.ai ou autres éditeurs Markdown, avec la hiérarchie et conventions suivantes :

| Markdown | Usage typique dans le document |
| -------- | ------------------------------ |
| `#`      | Titre principal (H1) - Titre du document uniquement |
| `##`     | Chapitres principaux (H2) - Ex : « Qu’est-ce que l’AEO » |
| `###`    | Sous-sections (H3) - Ex : « Phase 1 : Audit et analyse » |
| `####`   | Détails spécifiques (H4) - Ex : « Jour 1-2 : Diagnostic initial » |

### Exemples de formatage texte

- `**texte en gras**` → Mise en évidence importante  
- `*texte en italique*` → Nuance ou terme technique  
- `***texte gras et italique***` → Super important  
- `` `code ou terme technique` `` → Élément technique précis  

### Listes et autres éléments

- Puces : `-`, `*`  
- Listes numérotées : `1.`, `2.`  
- Checklists : `□` (case vide), `✅` (case cochée)  
- Citations : `> citation importante`  
- Tableaux au format Markdown  

---

## Installation

1. Clone ce repository ou télécharge le script Python.  
2. Assure-toi d'avoir Python 3.x installé.  
3. Installe les dépendances nécessaires (exemple avec `python-docx` et `markdown` si utilisés) :  

```bash
pip install python-docx markdown
