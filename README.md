# Modèle de Projection Actuariel

## 📋 Vue d'ensemble

Ce dépôt contient un modèle de projection actuariel pour contrats d'assurance vie, initialement développé en VBA/Excel.

## 📄 Documents Clés

### 📊 [Analyse Complète et Proposition de Migration vers R](./ANALYSE_MIGRATION_R.md)

**Note d'analyse détaillée** présentant :
- ✅ État des lieux du code VBA actuel
- ❌ Identification des 6 catégories de faiblesses majeures
- ✨ Bénéfices de la migration vers R
- 🗺️ Plan de migration détaillé (20 semaines)
- 💰 ROI estimé à 750% sur 3 ans

## 🏗️ Structure Actuelle (VBA)

Le code est organisé en plusieurs modules VBA :

| Fichier | Rôle |
|---------|------|
| `Main.txt` | Module principal, types de données, orchestration |
| `Calculs.txt` | Moteur de calculs actuariels (Euro, UC, Prévoyance) |
| `Initialisation.txt` | Chargement paramètres et données |
| `Fonctions.txt` | Fonctions utilitaires et transformations |
| `Ecriture.txt` | Export des résultats |
| `ExportResultats.txt` | Formatage et export avancé |
| `PRIIPS.txt` | Module spécifique produits PRIIPS |
| `Erreurs.txt` | Gestion basique des erreurs |
| `Suivi.txt` | Fonctions de filtrage et suivi |

## ⚠️ Principales Limitations Identifiées

1. **Architecture monolithique** : Code procédural difficile à maintenir
2. **Performance limitée** : Pas de vectorisation, boucles imbriquées
3. **Pas de tests** : Aucune validation automatique
4. **Dépendance Excel** : Couplage fort avec l'interface Excel
5. **Documentation minimale** : Courbe d'apprentissage très longue
6. **Scalabilité limitée** : Contraintes mémoire Excel (32-bit)

## 🚀 Solution Proposée : Migration vers R

### Bénéfices Clés

| Aspect | Gain |
|--------|------|
| **Performance** | **x30-120** (vectorisation + parallélisation) |
| **Capacité** | **Illimitée** (vs ~50k contrats en VBA) |
| **Tests** | **90%+ couverture** (vs 0% en VBA) |
| **Productivité** | **x4-36** selon les tâches |
| **ROI** | **750% sur 3 ans** |

### Architecture R Cible

```
me/ (Package R)
├── R/                    # Code source modulaire
├── tests/                # Tests unitaires complets
├── vignettes/            # Documentation longue
├── inst/templates/       # Templates R Markdown
└── data/                 # Données de référence
```

## 📅 Planning de Migration

| Phase | Durée | Objectif |
|-------|-------|----------|
| **Phase 1** | 4 semaines | Préparation & environnement R |
| **Phase 2** | 12 semaines | Migration incrémentale des modules |
| **Phase 3** | 4 semaines | Validation & double run VBA/R |
| **Phase 4** | 4 semaines | Optimisation & formation |
| **Total** | **24 semaines** (~6 mois) | |

## 📖 Pour en Savoir Plus

👉 **Consultez l'analyse complète** : [ANALYSE_MIGRATION_R.md](./ANALYSE_MIGRATION_R.md)

Ce document de 900+ lignes détaille :
- Les faiblesses précises du code actuel avec exemples
- L'architecture R recommandée avec exemples de code
- Le plan de migration détaillé semaine par semaine
- L'analyse coût/bénéfice quantifiée
- Les risques et leur mitigation

## 🎯 Prochaines Étapes Recommandées

1. **Validation de l'analyse** par l'équipe technique et métier
2. **Décision go/no-go** sur la migration
3. **Formation R** de l'équipe (2 jours)
4. **Démarrage Phase 1** : Audit détaillé et POC

---

*Dernière mise à jour : 24 novembre 2025*
