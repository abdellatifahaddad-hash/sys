# Guide Rapide - Rapport des Ventes

## 🚀 Démarrage rapide

### Installation
```bash
pip3 install openpyxl
```

### Commandes essentielles

#### Rapport complet
```bash
python3 rapport_ventes.py
```

#### Rapport du mois de Mars
```bash
python3 rapport_ventes.py --mois 3
```

#### Rapport avec export Excel
```bash
python3 rapport_ventes.py --export mon_rapport.xlsx
```

#### Rapport Mbacke1 uniquement
```bash
python3 rapport_ventes.py --destination GUEDE-OIL-MBACKE1
```

## 📊 Types de rapports disponibles

| Type | Commande |
|------|----------|
| Rapport annuel | `python3 rapport_ventes.py` |
| Rapport mensuel | `python3 rapport_ventes.py --mois 3` |
| Rapport par destination | `python3 rapport_ventes.py --destination GUEDE-OIL-MBACKE1` |
| Export Excel | `python3 rapport_ventes.py --export rapport.xlsx` |

## 🔢 Numéros de mois

1. Janvier
2. Février
3. Mars
4. Avril
5. Mai
6. Juin
7. Juillet
8. Août
9. Septembre
10. Octobre
11. Novembre
12. Décembre

## 🏢 Destinations disponibles

- `GUEDE-OIL-MBACKE1`
- `GUEDE-OIL-MBACKE2`

## 📈 Ce que contient le rapport

✅ Nombre total de transactions  
✅ Ventes totales en CFA  
✅ Ventes par destination  
✅ Ventes par type de paiement  
✅ Top 10 produits  
✅ Top 10 clients  
✅ Ventes par mois  

## 💾 Format Excel exporté

Le fichier Excel contient 5 feuilles :

1. **Résumé** - Vue d'ensemble
2. **Par Destination** - Détails par localisation
3. **Par Produit** - Détails par produit
4. **Par Client** - Détails par client
5. **Détails Transactions** - Liste complète

## 📞 Support

Pour plus d'aide :
```bash
python3 rapport_ventes.py --help
```

---
GUEDE OIL SARL - Mars 2026
