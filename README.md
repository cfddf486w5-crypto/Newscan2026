# Newscan 2026

Application PWA HTML pour :
- importer un fichier Excel (`.xlsx`, `.xls`, `.csv`) avec colonnes `code barre`, `produit`, `location`, `qty` (ordre libre),
- scanner des articles via scanner USB (mode clavier),
- constituer une liste de demande de commande,
- exporter la demande en CSV,
- conserver les données en `localStorage` (aucun compte, aucun mot de passe).

## Lancer en local

```bash
python3 -m http.server 4173
```

Puis ouvrir `http://localhost:4173`.
