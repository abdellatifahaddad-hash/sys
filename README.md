# Système de Gestion GUEDE OIL SARL

## 📊 Générateur de Rapport des Ventes

Ce système permet de générer des rapports détaillés des ventes à partir du fichier Excel `MBACKE.xlsm`.

### 🎯 Fonctionnalités

Le générateur de rapport des ventes offre les analyses suivantes :

- **Vue d'ensemble** : Nombre total de transactions et montant des ventes
- **Ventes par destination** : Répartition entre GUEDE-OIL-MBACKE1 et MBACKE2
- **Ventes par type de paiement** : CREDIT, ESPESE, BANQREMB, REMBOURSEMENT
- **Top produits** : Classement des produits les plus vendus (GASOIL, SUPER, etc.)
- **Top clients** : Liste des meilleurs clients avec nombre de transactions
- **Ventes par mois** : Évolution mensuelle des ventes

### 📋 Prérequis

- Python 3.6 ou supérieur
- Bibliothèque openpyxl

### 🔧 Installation

1. Installer Python 3 (si ce n'est pas déjà fait)

2. Installer les dépendances :
```bash
pip3 install openpyxl
```

### 🚀 Utilisation

#### Commandes de base

1. **Générer un rapport complet (tous les mois)** :
```bash
python3 rapport_ventes.py
```

2. **Filtrer par mois** :
```bash
python3 rapport_ventes.py --mois 3
```
(Affiche uniquement les ventes du mois de Mars)

3. **Filtrer par destination** :
```bash
python3 rapport_ventes.py --destination GUEDE-OIL-MBACKE1
```

4. **Combiner plusieurs filtres** :
```bash
python3 rapport_ventes.py --mois 1 --destination GUEDE-OIL-MBACKE2
```

5. **Exporter vers Excel** :
```bash
python3 rapport_ventes.py --export rapport_janvier.xlsx
```

6. **Filtrer et exporter** :
```bash
python3 rapport_ventes.py --mois 3 --export rapport_mars_2026.xlsx
```

#### Options disponibles

| Option | Description | Exemple |
|--------|-------------|---------|
| `--fichier` | Chemin vers le fichier Excel source | `--fichier sys/MBACKE.xlsm` |
| `--mois` | Filtrer par mois (1-12) | `--mois 3` |
| `--destination` | Filtrer par destination | `--destination GUEDE-OIL-MBACKE1` |
| `--export` | Exporter vers un fichier Excel | `--export rapport.xlsx` |

### 📄 Format du rapport Excel

Le fichier Excel exporté contient 5 feuilles :

1. **Résumé** : Vue d'ensemble et informations générales
2. **Par Destination** : Ventes groupées par destination avec pourcentages
3. **Par Produit** : Détails des ventes par produit (quantité et montant)
4. **Par Client** : Liste des clients avec nombre de transactions et montant
5. **Détails Transactions** : Liste complète de toutes les transactions

### 💡 Exemples pratiques

#### Exemple 1 : Rapport mensuel pour la direction
```bash
python3 rapport_ventes.py --mois 3 --export rapport_direction_mars.xlsx
```

#### Exemple 2 : Analyse d'une destination spécifique
```bash
python3 rapport_ventes.py --destination GUEDE-OIL-MBACKE1
```

#### Exemple 3 : Rapport complet exporté
```bash
python3 rapport_ventes.py --export rapport_complet_2026.xlsx
```

### 📊 Exemple de sortie console

```
================================================================================
                     📊 RAPPORT DES VENTES - GUEDE OIL SARL                      
================================================================================

📈 VUE D'ENSEMBLE
--------------------------------------------------------------------------------
Nombre de transactions:            413
Ventes totales:             37,629,961 CFA

🏢 VENTES PAR DESTINATION
--------------------------------------------------------------------------------
GUEDE-OIL-MBACKE1                   34,921,661 CFA  ( 92.8%)
GUEDE-OIL-MBACKE2                    2,708,300 CFA  (  7.2%)

💳 VENTES PAR TYPE DE PAIEMENT
--------------------------------------------------------------------------------
CREDIT                              27,047,563 CFA  ( 71.9%)
BANQREMB                             8,547,698 CFA  ( 22.7%)
REMBOURSEMENT                        2,034,700 CFA  (  5.4%)

🛢️  TOP 10 PRODUITS
--------------------------------------------------------------------------------
ESPESE                         Qté:        0.0  |    11,295,798 CFA
GASOIL                         Qté:    9,723.9  |     6,612,234 CFA
SUPER                          Qté:    2,852.9  |     2,624,700 CFA
...
```

### 🗂️ Structure des données

Le système extrait les données des feuilles Excel suivantes :

- **etatclient** : Transactions de ventes clients (ventes au détail)
- **FACTURES** : Factures fournisseurs et dépenses
- **CLT** : Liste des clients
- **ETATCREDITM** : État des crédits mensuels

### 🔒 Sécurité

- Le script est en lecture seule et ne modifie jamais le fichier Excel source
- Les données sont traitées localement sans connexion externe
- Aucune donnée n'est partagée ou transmise

### ❓ Aide et support

Pour afficher l'aide :
```bash
python3 rapport_ventes.py --help
```

### 📝 Notes

- Les montants sont affichés en Francs CFA
- Les mois sont numérotés de 1 (Janvier) à 12 (Décembre)
- Les destinations disponibles : GUEDE-OIL-MBACKE1 et GUEDE-OIL-MBACKE2
- Le rapport affiche automatiquement le top 10 des produits et clients

### 🐛 Résolution de problèmes

**Problème** : `ModuleNotFoundError: No module named 'openpyxl'`
**Solution** : Installer openpyxl avec `pip3 install openpyxl`

**Problème** : `FileNotFoundError: [Errno 2] No such file or directory: 'sys/MBACKE.xlsm'`
**Solution** : Vérifier que le fichier Excel existe ou utiliser `--fichier` pour spécifier le chemin correct

**Problème** : Le rapport est vide
**Solution** : Vérifier que le fichier Excel contient des données et que les filtres sont corrects

### 🎯 Bonnes pratiques

1. Générer des rapports mensuels pour suivre l'évolution
2. Exporter les rapports importants pour archivage
3. Comparer les performances entre les deux destinations
4. Identifier les clients à fort potentiel
5. Analyser les tendances de ventes par produit

---

**GUEDE OIL SARL** - Système de Gestion  
Mbacké, Sénégal  
Version: 1.0 (Mars 2026)
