# Résolution du crash au démarrage - ExportElec.exe

## 🔴 Problème identifié

L'application `ExportElec.exe` se terminait immédiatement avec le code d'erreur **0xFFFFFFFF (-1)** lors du démarrage, sans afficher aucun message d'erreur explicite.

### Logs de débogage
```
'ExportElec.exe' (CLR v4.0.30319: ExportElec.exe) : Chargé 'OutilsTs.dll'
Les symboles du module 'OutilsTs.dll' n'ont pas été chargés.
Le programme '[17108] ExportElec.exe' s'est arrêté avec le code 4294967295 (0xffffffff).
```

## 🔍 Causes probables

### 1. **Connexion TopSolid échoue**
- Si TopSolid n'est pas démarré ou accessible
- Si l'API TopSolid n'est pas disponible
- Si les DLL TopSolid ne sont pas chargées correctement

### 2. **Exception non capturée dans le constructeur**
Le constructeur `MainWindow()` n'avait aucun bloc try-catch, donc toute exception provoquait un crash immédiat de l'application.

### 3. **Module OutilsTs.dll optimisé**
Le module était compilé en Release (optimisé) sans symboles de débogage, rendant le diagnostic difficile.

## ✅ Solutions implémentées

### 1. **Gestion d'exception dans MainWindow.xaml.cs**

#### Avant :
```csharp
public MainWindow()
{
    InitializeComponent();
    startConnect = new StartConnect();
    startConnect.ConnectionTopsolid();
    LoadSavedPath();
    InitializeForm();
    SelectFile = this.FindName("SelectFile") as Button;
}
```

#### Après :
```csharp
public MainWindow()
{
    try
    {
        InitializeComponent();
        LoadSavedPath();
        SelectFile = this.FindName("SelectFile") as Button;

        // Connexion TopSolid avec gestion d'erreur spécifique
        try
        {
            startConnect = new StartConnect();
            startConnect.ConnectionTopsolid();
        }
        catch (Exception exConnect)
        {
            MessageBox.Show(
                $"Impossible de se connecter à TopSolid.\n\n" +
                $"Erreur: {exConnect.Message}\n\n" +
                $"Assurez-vous que TopSolid est démarré avant de lancer cette application.",
                "Erreur de connexion TopSolid",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            
            AfficherAucunDocumentOuvert();
            return;
        }

        InitializeForm();
    }
    catch (Exception ex)
    {
        // Capturer toute exception critique
        MessageBox.Show(
            $"Erreur critique lors de l'initialisation:\n\n" +
            $"Message: {ex.Message}\n\n" +
            $"Type: {ex.GetType().Name}\n\n" +
            $"StackTrace:\n{ex.StackTrace}",
            "Erreur critique",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        
        AfficherAucunDocumentOuvert();
    }
}
```

### 2. **Gestion globale des exceptions dans App.xaml.cs**

Ajout de gestionnaires d'exceptions au niveau de l'application :

```csharp
public App()
{
    // Capturer les exceptions UI thread
    this.DispatcherUnhandledException += App_DispatcherUnhandledException;
    
    // Capturer les exceptions threads en arrière-plan
    AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
}
```

## 📋 Résultats attendus

Maintenant, au lieu d'un crash silencieux, vous obtiendrez :

1. **Message d'erreur détaillé** avec :
   - Le message d'erreur exact
   - Le type d'exception
   - La pile d'appels (StackTrace)
   
2. **L'application reste ouverte** en mode dégradé plutôt que de crasher

3. **Message spécifique** si TopSolid n'est pas connecté

## 🧪 Tests à effectuer

### Test 1 : TopSolid non démarré
1. Fermer TopSolid complètement
2. Lancer ExportElec.exe
3. **Résultat attendu** : Message "Impossible de se connecter à TopSolid"

### Test 2 : TopSolid démarré sans document
1. Démarrer TopSolid
2. Ne pas ouvrir de document
3. Lancer ExportElec.exe
4. **Résultat attendu** : Application s'ouvre avec message "Aucun document TopSolid ouvert"

### Test 3 : TopSolid avec document
1. Démarrer TopSolid
2. Ouvrir un document avec des électrodes
3. Lancer ExportElec.exe
4. **Résultat attendu** : Application s'ouvre normalement avec la liste des électrodes

## 🔧 Dépannage supplémentaire

Si le problème persiste, vérifiez :

### 1. Dépendances OutilsTs.dll
```powershell
# Vérifier les dépendances de OutilsTs.dll
dumpbin /dependents "bin\Debug\OutilsTs.dll"
```

### 2. Versions .NET Framework
- Projet ExportElec : .NET Framework 4.8.1
- Vérifier que OutilsTs.dll est compatible avec cette version

### 3. Journaux détaillés de Visual Studio
- Activer tous les messages de débogage dans Visual Studio
- Outils → Options → Débogage → Sortie
- Activer tous les messages de chargement de modules

### 4. Compiler OutilsTs.dll en mode Debug
Pour avoir les symboles de débogage et mieux identifier les erreurs :
```
Configuration : Debug au lieu de Release
```

## 📝 Notes importantes

- **Ne jamais laisser un constructeur sans gestion d'exception**
- Toujours vérifier la connectivité aux API externes (TopSolid) avant de les utiliser
- Implémenter une gestion d'exception globale dans App.xaml.cs pour tout projet WPF
- Compiler les bibliothèques internes en mode Debug pendant le développement

## 🎯 Prochaines étapes recommandées

1. **Tester l'application** avec les 3 scénarios ci-dessus
2. **Noter le message d'erreur exact** si le problème persiste
3. **Vérifier les logs** dans la fenêtre de sortie de Visual Studio
4. **Compiler OutilsTs.dll en Debug** pour avoir plus d'informations
5. **Ajouter des logs** dans `StartConnect.ConnectionTopsolid()` pour tracer la connexion

---

**Date de création** : ${new Date().toLocaleDateString('fr-FR')}
**Fichiers modifiés** :
- `MainWindow.xaml.cs` : Ajout gestion d'exception dans constructeur
- `App.xaml.cs` : Ajout gestion d'exception globale
