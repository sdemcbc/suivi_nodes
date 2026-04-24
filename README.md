# Site Down Dashboard

Dashboard web pour visualiser les coupures de sites télécom depuis un fichier Excel.

## Structure du projet

```
appli site down/
├── index.html       ← Page principale
├── style.css        ← Styles (thème dark industriel)
├── app.js           ← Logique JS (import, filtres, tri, export)
└── README.md        ← Ce fichier
```

> SheetJS est chargé depuis CDN (cdnjs.cloudflare.com). Aucune installation nécessaire.

## Utilisation

1. Ouvrir `index.html` dans un navigateur (Chrome, Firefox, Edge)
2. Importer un fichier `.xlsx`, `.xls` ou `.csv`
3. Filtrer par Location, Site ID, Site Name, et/ou Duration
4. Cliquer sur les en-têtes pour trier
5. Exporter les résultats filtrés en CSV

## Colonnes Excel attendues

Le fichier Excel doit contenir ces colonnes (noms flexibles) :

| Donnée      | Noms acceptés                                        |
|-------------|------------------------------------------------------|
| Location    | location, loc, ville, city, region, site location    |
| Technologie | technology, technologie, tech, type, network type    |
| Site ID     | site id, siteid, id, site_id, site number, code site |
| Site Name   | site name, sitename, nom site, site_name, name       |
| Duration    | duration, durée, outage duration, down time, coupure |

## Fonctionnalités

- Import drag & drop ou sélection de fichier
- 5 filtres dynamiques (Location, Site ID, Site Name, Duration min/max + slider)
- Tri par colonne (clic sur l'en-tête)
- KPI en temps réel (total sites, sites filtrés, durée totale, locations uniques)
- Code couleur sur la durée : vert < 1h · orange 1–3h · rouge > 3h
- Export CSV des résultats filtrés
- Responsive mobile
