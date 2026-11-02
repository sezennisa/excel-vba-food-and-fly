# Food & Fly – Outil Excel VBA connecté à des API

🇫🇷 [Français](#-français) · 🇬🇧 [English](#-english)

---

## 🇫🇷 Français

Projet réalisé dans le cadre du **Master MBFA – parcours SIEF** (Université de Montpellier, 2025–2026), en binôme.

### Objectif
Construire, entièrement en **Excel VBA**, un outil qui interroge des API web pour comparer des produits alimentaires du monde entier, puis trouver le pays où un produit est le moins cher.

### Fonctionnement
1. L'utilisateur saisit un mot-clé dans la feuille **Recherche** (par exemple « chocolat »).
2. Le mot est **traduit automatiquement en 6 langues** (anglais, espagnol, allemand, italien, portugais, néerlandais) pour élargir la recherche.
3. Pour chaque traduction, la macro interroge l'**API OpenFoodFacts** page par page (jusqu'à 500 produits), filtre les résultats hors sujet et **supprime les doublons** grâce au code-barres.
4. Les résultats sont écrits dans la feuille **Resultats** : nom, marque, quantité, pays, **Nutri-Score** et **Eco-Score** (mis en couleur), code-barres.
5. Les prix sont récupérés pour chaque produit via l'**API OpenPrices**.
6. Une seconde macro identifie le **produit le moins cher** et le pays associé, puis propose d'ouvrir **Google Flights** vers ce pays.

### Compétences mises en œuvre
- Requêtes **HTTP** vers des API REST et lecture de réponses **JSON** (parseur écrit à la main en VBA)
- Pagination, gestion des erreurs, pause entre les appels pour respecter les serveurs
- Nettoyage des données, dédoublonnage, mise en forme conditionnelle
- Code commenté et organisé en fonctions réutilisables

### Contenu du dépôt
| Fichier | Description |
|---|---|
| `Food_and_Fly.xlsm` | Classeur Excel avec les macros et un exemple de résultats |
| `Module1.bas` | Code VBA exporté, lisible directement sur GitHub |

### Utilisation
Ouvrir `Food_and_Fly.xlsm` dans Excel (Windows), activer les macros, saisir un mot en **D5** de la feuille *Recherche*, puis lancer `RechercherProduits`. La macro `TrouverVolMoinsCher` se lance ensuite depuis les résultats.

---

## 🇬🇧 English

Project completed as part of the **Master's in Money, Banking, Finance & Insurance – Economic and Financial Information Systems** (University of Montpellier, 2025–2026), in a pair with **Anton**.

### Goal
Build a tool entirely in **Excel VBA** that calls web APIs to compare food products from around the world, then find the country where a product is cheapest.

### How it works
1. The user types a keyword in the **Recherche** sheet (for example "chocolat").
2. The keyword is **automatically translated into 6 languages** (English, Spanish, German, Italian, Portuguese, Dutch) to widen the search.
3. For each translation, the macro queries the **OpenFoodFacts API** page by page (up to 500 products), filters out irrelevant results and **removes duplicates** using the barcode.
4. Results are written to the **Resultats** sheet: name, brand, quantity, countries, **Nutri-Score** and **Eco-Score** (colour-coded), barcode.
5. Prices are retrieved for each product from the **OpenPrices API**.
6. A second macro finds the **cheapest product** and its country, then offers to open **Google Flights** to that country.

### Skills
- **HTTP** requests to REST APIs and **JSON** parsing (hand-written parser in VBA)
- Pagination, error handling, throttling between calls
- Data cleaning, de-duplication, conditional formatting
- Commented code organised into reusable functions

### Repository contents
| File | Description |
|---|---|
| `Food_and_Fly.xlsm` | Excel workbook with the macros and sample results |
| `Module1.bas` | Exported VBA code, readable directly on GitHub |

### Usage
Open `Food_and_Fly.xlsm` in Excel (Windows), enable macros, type a keyword in cell **D5** of the *Recherche* sheet, then run `RechercherProduits`. Then run `TrouverVolMoinsCher` from the results.

---

**Authors:** Sezen Nisa Dikmen & Anton
**Data sources:** [OpenFoodFacts](https://world.openfoodfacts.org) · [Open Prices](https://prices.openfoodfacts.org)
