# Exemples de Rapports - GUEDE OIL SARL

Ce dossier contient des exemples de rapports générés par le système.

## 📊 Rapports disponibles

### 1. rapport_complet.xlsx
**Contenu** : Rapport complet de toutes les ventes (tous mois confondus)
- 413 transactions
- Ventes totales : 37,629,961 CFA
- Période : Janvier - Mars 2026
- Destinations : MBACKE1 et MBACKE2

**Commande utilisée** :
```bash
python3 rapport_ventes.py --export exemples_rapports/rapport_complet.xlsx
```

### 2. rapport_janvier.xlsx
**Contenu** : Rapport des ventes du mois de Janvier 2026
- 242 transactions
- Ventes totales : 28,760,977 CFA
- Destinations : MBACKE1 et MBACKE2

**Commande utilisée** :
```bash
python3 rapport_ventes.py --mois 1 --export exemples_rapports/rapport_janvier.xlsx
```

### 3. rapport_mbacke1.xlsx
**Contenu** : Rapport des ventes de GUEDE-OIL-MBACKE1 uniquement
- 363 transactions
- Ventes totales : 34,921,661 CFA
- Période : Janvier - Mars 2026

**Commande utilisée** :
```bash
python3 rapport_ventes.py --destination GUEDE-OIL-MBACKE1 --export exemples_rapports/rapport_mbacke1.xlsx
```

## 📋 Structure des fichiers Excel

Chaque fichier Excel contient 5 feuilles :

1. **Résumé** - Vue d'ensemble et statistiques générales
2. **Par Destination** - Répartition par localisation (MBACKE1/MBACKE2)
3. **Par Produit** - Détails des ventes par produit (GASOIL, SUPER, etc.)
4. **Par Client** - Liste des clients avec transactions et montants
5. **Détails Transactions** - Liste complète de toutes les transactions

## 🔍 Points clés des analyses

### Principaux clients :
1. COMMUNE DE MBACKE GASOIL
2. SAREQ GROUP
3. COMMUNE DE MBACKE SUPER
4. S.CHEIKH MBACKE VOITURE ET CAMION
5. MCG SARL

### Produits les plus vendus :
1. GASOIL - 9,723.9 litres
2. SUPER - 2,852.9 litres
3. ESPESE (espèces/avances)

### Répartition par destination :
- GUEDE-OIL-MBACKE1 : 92.8% des ventes
- GUEDE-OIL-MBACKE2 : 7.2% des ventes

### Types de paiement :
- CREDIT : 71.9%
- BANQREMB : 22.7%
- REMBOURSEMENT : 5.4%

## 🎯 Utilisation

Ces exemples peuvent servir de modèles pour :
- Présentations à la direction
- Analyses mensuelles
- Comparaisons entre destinations
- Suivi des clients importants
- Planification des stocks

## 📝 Notes

- Les montants sont en Francs CFA
- Les dates sont au format JJ/MM/AAAA
- Les rapports peuvent être ouverts avec Microsoft Excel ou LibreOffice Calc

---
GUEDE OIL SARL - Mars 2026
