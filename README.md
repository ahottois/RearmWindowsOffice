# Script de Réarmement de Licences Windows et Office

Ce script PowerShell permet de réarmer facilement les périodes d'essai des licences Microsoft Windows et Microsoft Office. Il utilise les outils officiels de Microsoft (`slmgr.vbs` pour Windows et `ospprearm.exe` pour Office) pour étendre légitimement les périodes d'évaluation.

## 🚀 Fonctionnalités

- ✅ Réarmement de la licence Windows via `slmgr.vbs /rearm`
- ✅ Réarmement de la licence Office via `ospprearm.exe`
- ✅ Détection automatique des installations d'Office (versions 2010 à 365)
- ✅ Vérification des privilèges administrateur
- ✅ Fermeture automatique des applications Office en cours d'exécution
- ✅ Interface utilisateur simple avec menu interactif
- ✅ Option de redémarrage après le réarmement Windows

## 📋 Prérequis

- Windows 7/8/10/11 ou Windows Server 2008 R2/2012/2016/2019/2022
- PowerShell 3.0 ou supérieur
- Droits d'administrateur sur le système
- Microsoft Office installé (pour la fonction de réarmement Office)

## 📥 Installation

1. Clonez ce dépôt ou téléchargez le fichier `License-Rearm.ps1`
2. Assurez-vous que la politique d'exécution PowerShell permet l'exécution de scripts

```powershell
# Pour vérifier la politique d'exécution actuelle
Get-ExecutionPolicy

# Pour autoriser l'exécution de scripts (exécuter en tant qu'administrateur)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 🎮 Utilisation

1. Ouvrez PowerShell en tant qu'administrateur
2. Naviguez vers le répertoire contenant le script
3. Exécutez le script :

```powershell
.\License-Rearm.ps1
```

4. Choisissez l'option désirée dans le menu :
   - `1` : Réarmer Windows uniquement
   - `2` : Réarmer Office uniquement
   - `3` : Réarmer Windows et Office
   - `Q` : Quitter

## ⚠️ Avertissements

- Le réarmement de Windows nécessite un redémarrage pour prendre effet
- Le nombre de réarmements possibles est limité (généralement à 3 ou 5 fois selon les produits)
- Ce script est conçu pour des scénarios légitimes d'évaluation ou de tests
- N'utilisez pas ce script pour éviter l'achat de licences pour un usage en production
- Les modifications apportées aux systèmes d'activation pourraient rendre ce script obsolète

## 🔄 Limites de réarmement

- **Windows** : Généralement limité à 3 réarmements
- **Office** : Généralement limité à 5 réarmements

## 📜 Utilisations légitimes

Ce script est utile pour :
- Environnements d'apprentissage et de formation
- Tests logiciels et compatibilité d'applications
- Préparation de machines virtuelles pour démonstrations
- Extension temporaire pendant la résolution de problèmes d'activation

## 📋 Notes légales

Ce script utilise uniquement des outils et méthodes légitimes fournis par Microsoft. Il est conçu pour faciliter l'évaluation et les tests, et non pour contourner les licences commerciales. Respectez toujours les conditions de licence de Microsoft.

## 🔧 Dépannage

Si vous rencontrez des erreurs lors de l'exécution :

1. Vérifiez que vous exécutez PowerShell en tant qu'administrateur
2. Assurez-vous que toutes les applications Office sont fermées
3. Vérifiez que les chemins d'installation d'Office sont standards
4. Consultez les logs d'erreur affichés par le script

## 🤝 Contribution

Les contributions sont les bienvenues ! N'hésitez pas à ouvrir une issue ou une pull request pour améliorer ce script.
