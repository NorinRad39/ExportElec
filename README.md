## 📦 Dépendances et Installation

Ce projet utilise plusieurs bibliothèques nécessaires à son fonctionnement :

### 1. SDK TopSolid (NuGet)
Les bibliothèques d'interface (API) de TopSolid sont récupérées via **NuGet**. Assurez-vous d'avoir accès aux flux de paquets configurés dans votre Visual Studio pour restaurer les références.

### 2. OutilsTS (Bibliothèque personnelle)
Ce projet s'appuie sur ma classe utilitaire **`OutilsTs.dll`**. 
Elle est disponible publiquement sur **NuGet**. 
- Vous pouvez l'installer via la console de gestion des paquets :
  `Install-Package OutilsTs`
- Ou via le gestionnaire de solutions NuGet en cherchant "OutilsTs".

### 3. Configuration du projet
- **Cible :** .NET Framework 4.8.1
- **Plateforme :** x64 (Obligatoire pour la compatibilité avec TopSolid)
- **NuGet Restore :** Au premier lancement, faites un clic droit sur la Solution > "Restaurer les packages NuGet".

---

## ⚖️ Licence
Ce projet est partagé sous licence **Creative Commons Attribution-NonCommercial (CC BY-NC 4.0)**.
- **Utilisation gratuite** pour un usage personnel ou interne en entreprise.
- **Revente interdite** : Vous n'êtes pas autorisé à vendre ce logiciel ou une version modifiée de celui-ci.
