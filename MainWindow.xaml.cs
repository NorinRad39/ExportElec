using OutilsTs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml.Linq;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cad.Drafting.Automating;
using TopSolid.Cam.NC.Kernel.Automating;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;

namespace ExportElec
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// Permet l'export automatisé d'électrodes depuis TopSolid avec gestion des gaps
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Champs privés

        /// <summary>
        /// Document TopSolid actuellement ouvert
        /// </summary>
        private Document currentDoc;

        /// <summary>
        /// Gestionnaire de connexion à TopSolid
        /// </summary>
        private StartConnect startConnect;

        /// <summary>
        /// Bouton Export
        /// </summary>
        private Button SelectFile;

        #endregion

        #region Constructeur

        /// <summary>
        /// Initialise une nouvelle instance de la fenêtre principale
        /// </summary>
        public MainWindow()
        {
            // Pour capturer plus de détails, ajoutez ceci dans votre code
            
                InitializeComponent();
           
            

            // Connexion à TopSolid
            startConnect = new StartConnect();
            startConnect.ConnectionTopsolid();
            
            // Restaurer le chemin d'export sauvegardé
            LoadSavedPath();
            
            // Initialisation du formulaire avec le document courant
            InitializeForm();

            // Récupérer le bouton Export depuis le XAML
            SelectFile = this.FindName("SelectFile") as Button;
        }

        #endregion

        #region Chargement et initialisation

        /// <summary>
        /// Charge le chemin d'export sauvegardé depuis les paramètres de l'application
        /// </summary>
        private void LoadSavedPath()
        {
            if (!string.IsNullOrEmpty(Properties.Settings.Default.CheminExport))
            {
                chemin.Text = Properties.Settings.Default.CheminExport;
            }
        }

        /// <summary>
        /// Initialise le formulaire en chargeant les données du document TopSolid
        /// et en classifiant les électrodes par gap
        /// </summary>
        private void InitializeForm()
        {
            try
            {
                // Récupérer l'identifiant du document édité sans créer d'instance Document
                DocumentId editedDocId = TSH.Documents.EditedDocument;

                // Vérifier si un document est ouvert AVANT de créer l'instance Document
                if (editedDocId == null || editedDocId.IsEmpty)
                {
                    AfficherAucunDocumentOuvert();
                    return;
                }

                // Maintenant qu'on sait qu'un document est ouvert, créer l'instance Document
                currentDoc = new Document();
                currentDoc.DocId = editedDocId;

                // Document ouvert : activer l'interface et charger les données
                DocumentNameText.Text = currentDoc.DocNomTxt;
                ActiverInterface(true);

                //Etape 1 : Récupération des éléments du document
                List<ElementId> elements = RecupList(currentDoc.DocId);

                if (elements == null || elements.Count == 0)
                {
                    electrodeList.Items.Clear();
                    electrodeList.Items.Add("Aucun élément trouvé dans le document");
                    ActiverBoutonExport(false);
                    return;
                }

                //Etape 2 : Filtrage des opérations
                List<ElementId> operations = ListOperations(elements);

                //Etape 3 : Filtrage des opérations actives
                List<ElementId> operationsActives = OperationsActive(operations);

                //Etape 4 : Recherche des opérations de duplication
                List<ElementId> duplicateOperations = searchOperations(operationsActives, "TopSolid.Kernel.DB.Operations.DuplicateCreation");

                //Etape 5 : Récupération des éléments enfants des opérations de duplication
                List<ElementId> duplicateOperationsChildElementsId = childrenElements(duplicateOperations);

                //Etape 6 : Récupération de l'électrode
                ElementId electrodeId = SearchElectrode(currentDoc.DocId, "Electrode");

                //Etape 7 : Ajout de l'électrode à la liste des shapes
                List<ElementId> ShapesList = AddShapeToList(duplicateOperationsChildElementsId, electrodeId);

                //Etape 8 : Récupération des valeurs des gaps
                gapString(currentDoc, out string gapEb, out string gapDemiFini, out string gapFini);

                //Etape 9 : Tri et nommage des électrodes par volume
                ClassifyAndDisplayElectrodes(ShapesList, gapEb, gapDemiFini, gapFini);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'initialisation du formulaire: {ex.Message}", 
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                AfficherAucunDocumentOuvert();
            }
        }

        /// <summary>
        /// Affiche un message indiquant qu'aucun document n'est ouvert et désactive l'interface
        /// </summary>
        private void AfficherAucunDocumentOuvert()
        {
            DocumentNameText.Text = "Aucun document TopSolid ouvert";
            DocumentNameText.Foreground = System.Windows.Media.Brushes.Red;
            electrodeList.Items.Clear();
            electrodeList.Items.Add("Veuillez ouvrir un document dans TopSolid puis cliquer sur 'Recharger'");
            ActiverInterface(false);
        }

        /// <summary>
        /// Active ou désactive l'interface d'export selon l'état du document
        /// </summary>
        /// <param name="activer">True pour activer, False pour désactiver</param>
        private void ActiverInterface(bool activer)
        {
            // Réinitialiser la couleur du texte si on active
            if (activer)
            {
                DocumentNameText.Foreground = System.Windows.Media.Brushes.Black;
            }
            
            ActiverBoutonExport(activer);
        }

        /// <summary>
        /// Active ou désactive le bouton d'export
        /// </summary>
        /// <param name="activer">True pour activer, False pour désactiver</param>
        private void ActiverBoutonExport(bool activer)
        {
            // Vérifier si le bouton existe dans le XAML (nom probable: SelectFile ou btnExport)
            if (SelectFile != null)
            {
                SelectFile.IsEnabled = activer;
            }
        }

        #endregion

        #region Etape 1 : Récupération des éléments du document

        /// <summary>
        /// Récupère la liste de tous les éléments présents dans le document TopSolid
        /// </summary>
        /// <param name="DocId">Identifiant du document TopSolid</param>
        /// <returns>Liste des identifiants d'éléments ou null en cas d'erreur</returns>
        private List<ElementId> RecupList(DocumentId DocId)
        {
            try
            {
                if (currentDoc?.DocId == null || currentDoc.DocId.IsEmpty)
                {
                    return null;
                }

                List<ElementId> elements = currentDoc.DocElements;
                return elements;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des éléments: {ex.Message}",
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        #endregion

        #region Etape 2 : Filtrage des opérations

        /// <summary>
        /// Filtre la liste des éléments pour ne conserver que les opérations TopSolid
        /// </summary>
        /// <param name="elements">Liste d'éléments à filtrer</param>
        /// <returns>Liste des opérations ou null si aucune opération trouvée</returns>
        private List<ElementId> ListOperations(List<ElementId> elements)
        {
            List<ElementId> operations = new List<ElementId>();

            if (elements != null && elements.Count > 0)
            {
                foreach (var element in elements)
                {
                    // Vérifier si l'élément est une opération
                    if (TSH.Operations.IsOperation(element))
                    {
                        operations.Add(element);
                      }
                }
                return operations;
            }
            else
            {
                return null;
            }
        }

        #endregion

        #region Etape 3 : Filtrage des opérations actives

        /// <summary>
        /// Filtre les opérations pour ne conserver que celles qui sont actives
        /// </summary>
        /// <param name="operations">Liste des opérations à filtrer</param>
        /// <returns>Liste des opérations actives ou null si aucune trouvée</returns>
        private List<ElementId> OperationsActive(List<ElementId> operations)
        {
            List<ElementId> operationsActive = new List<ElementId>();

            if (operations != null && operations.Count > 0)
            {
                foreach (var operation in operations)
                {
                    // Vérifier si l'opération est active
                    if (TSH.Operations.IsActive(operation))
                    {
                        operationsActive.Add(operation);
                    }
                }
                return operationsActive;
            }
            else
            {
                return null;
            }

        }
        #endregion

        #region Etape 4 : Recherche des opérations de duplication

        /// <summary>
        /// Recherche les opérations d'un type spécifique parmi les opérations actives
        /// </summary>
        /// <param name="operationsActive">Liste des opérations actives</param>
        /// <param name="operationName">Nom complet du type d'opération recherché</param>
        /// <returns>Liste des opérations correspondantes ou null si aucune trouvée</returns>
        private List<ElementId> searchOperations(List<ElementId> operationsActive, string operationName)
        {
            List<ElementId> duplicateOperations = new List<ElementId>();

            if (operationsActive != null && operationsActive.Count > 0)
            {
                foreach (var operation in operationsActive)
                {
                    // Récupérer le nom complet de l'élément via l'API TopSolid
                    string operationFullName = TSH.Elements.GetTypeFullName(operation);
                    
                    // Comparer avec le nom d'opération recherché
                    if (operationFullName == operationName)
                    {
                        duplicateOperations.Add(operation);
                    }
                }
                return duplicateOperations;  
            }
            else
            {
                return null;
            }
        }

        #endregion

        #region Etape 5 : Récupération des éléments enfants des opérations de duplication

        /// <summary>
        /// Récupère tous les éléments enfants des opérations spécifiées
        /// </summary>
        /// <param name="Operations">Liste des opérations parentes</param>
        /// <returns>Liste de tous les éléments enfants ou null si aucun</returns>
        private List<ElementId> childrenElements(List<ElementId> Operations)
        {
            List<ElementId> childrenElementsId = new List<ElementId>();

            if (Operations == null)
            {
                return null;
            }

            foreach (var operation in Operations)
            {
                // Récupérer les enfants de l'opération
                List<ElementId> children = TSH.Operations.GetChildren(operation);
                
                if (children != null && children.Count > 0)
                {
                    // Ajouter tous les enfants à la liste
                    childrenElementsId.AddRange(children);
                }
            }
            
            return childrenElementsId;
        }

        #endregion
        
        #region Etape 6 : Recherche de l'électrode

        /// <summary>
        /// Recherche un élément spécifique par son nom dans le document
        /// </summary>
        /// <param name="DocId">Identifiant du document</param>
        /// <param name="electrodeName">Nom de l'électrode à rechercher</param>
        /// <returns>Identifiant de l'électrode trouvée ou ElementId.Empty si non trouvée</returns>
        private ElementId SearchElectrode(DocumentId DocId, string electrodeName)
        {
            ElementId electrodeId = new ElementId();
            try
            {
                if (currentDoc.DocId != null)
                {
                    List<ElementId> elements = currentDoc.DocElements;
                    foreach (var element in elements)
                    {
                        // Récupérer le nom convivial de l'élément
                        string elementName = TSH.Elements.GetFriendlyName(element);
                        if (elementName == electrodeName)
                        {
                            electrodeId = element;
                            break; // Sortir de la boucle une fois l'élément trouvé
                        }
                    }
                    return electrodeId;
                }
                else
                {
                    DocumentNameText.Text = "aucun document ouvert, l'électrode ne peut pas être trouvée";
                    return ElementId.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la recherche de l'électrode: {ex.Message}",
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return ElementId.Empty;
            }
        }

        #endregion

        #region Etape 7 : Ajout de l'électrode à la liste des shapes

        /// <summary>
        /// Ajoute l'électrode principale à la liste des éléments enfants dupliqués
        /// </summary>
        /// <param name="duplicateOperationsChildElementsId">Liste des éléments dupliqués</param>
        /// <param name="electrodeId">Identifiant de l'électrode principale</param>
        /// <returns>Liste complète des shapes ou null en cas d'erreur</returns>
        private List<ElementId> AddShapeToList(List<ElementId> duplicateOperationsChildElementsId, ElementId electrodeId)
        {
            try
            {
                // Vérifier que nous avons au moins un élément ou une électrode
                if ((duplicateOperationsChildElementsId == null || duplicateOperationsChildElementsId.Count == 0) && 
                    (electrodeId == null || electrodeId.IsEmpty))
                {
                    return null;
                }

                // Si la liste est null, la créer
                if (duplicateOperationsChildElementsId == null)
                {
                    duplicateOperationsChildElementsId = new List<ElementId>();
                }

                // Ajouter l'électrode principale si elle existe
                if (electrodeId != null && !electrodeId.IsEmpty)
                {
                    duplicateOperationsChildElementsId.Add(electrodeId);
                }

                return duplicateOperationsChildElementsId.Count > 0 ? duplicateOperationsChildElementsId : null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des shapes: {ex.Message}",
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        #endregion

        #region Etape 8 : Récupération des valeurs des gaps

        /// <summary>
        /// Récupère les valeurs des paramètres de gap (ébauche, demi-fini, fini) depuis le document
        /// </summary>
        /// <param name="Doc">Document TopSolid</param>
        /// <param name="gapEb">Valeur du gap d'ébauche en mm</param>
        /// <param name="gapDemiFini">Valeur du gap demi-fini en mm</param>
        /// <param name="gapFini">Valeur du gap fini en mm</param>
        private void gapString(Document Doc, out string gapEb, out string gapDemiFini, out string gapFini)
        {
            // Initialiser les paramètres out avec des valeurs par défaut
            gapEb = "0";
            gapDemiFini = "0";
            gapFini = "0";

            if (Doc == null || Doc.DocId == null)
            {
                return;
            }

            foreach (var param in Doc.DocParameters)
            {
                string paramName = TSH.Elements.GetFriendlyName(param);

                // Rechercher et extraire les valeurs des paramètres de gap
                if (paramName.StartsWith("00-Gap Eb"))
                {
                    // Récupérer la valeur du paramètre (en mètres) et convertir en mm
                    double value = TSH.Parameters.GetRealValue(param);
                    gapEb = (value * 1000).ToString();
                }
                else if (paramName.StartsWith("01-Gap Demi fini"))
                {
                    double value = TSH.Parameters.GetRealValue(param);
                    gapDemiFini = (value * 1000).ToString();
                }
                else if (paramName.StartsWith("02-Gap Fini"))
                {
                    double value = TSH.Parameters.GetRealValue(param);
                    gapFini = (value * 1000).ToString();
                }
            }
        }

        #endregion

        #region Etape 9 : Tri et nommage des électrodes par volume

        /// <summary>
        /// Classe les électrodes par volume croissant et leur attribue un type (Ebauche, Demi-fini, Fini, Sans GAP)
        /// </summary>
        /// <param name="shapesIds">Liste des identifiants de shapes</param>
        /// <param name="gapEb">Valeur du gap d'ébauche</param>
        /// <param name="gapDemiFini">Valeur du gap demi-fini</param>
        /// <param name="gapFini">Valeur du gap fini</param>
        private void ClassifyAndDisplayElectrodes(List<ElementId> shapesIds, string gapEb, string gapDemiFini, string gapFini)
        {
            // Vider la ListBox
            electrodeList.Items.Clear();

            if (shapesIds == null || shapesIds.Count == 0)
            {
                electrodeList.Items.Add("Aucun élément trouvé");
                return;
            }

            try
            {
                // Créer des objets Element pour tous les shapes
                List<Element> allElements = new List<Element>();
                Element electrodeWithoutGap = null;
                
                foreach (var shapeId in shapesIds)
                {
                    Element element = new Element(shapeId);
                    string friendlyName = TSH.Elements.GetFriendlyName(shapeId);
                    
                    // Traiter l'électrode "Electrode" spécialement (Sans GAP)
                    if (friendlyName == "Electrode")
                    {
                        electrodeWithoutGap = element;
                    }
                    else if (element.IsShape && element.VolumeMm3.HasValue)
                    {
                        // Les autres doivent être des shapes valides avec volume
                        allElements.Add(element);
                    }
                }

                // Trier les électrodes dupliquées par volume croissant
                List<Element> duplicatedElectrodes = allElements.OrderBy(e => e.VolumeMm3).ToList();

                // Classifier selon le nombre d'électrodes dupliquées
                List<ElementItem> items = new List<ElementItem>();

                switch (duplicatedElectrodes.Count)
                {
                    case 0:
                        // Aucune électrode dupliquée
                        break;

                    case 1:
                        // Cas 1 : 1 seule shape → Electrode d'ébauche
                        items.Add(new ElementItem
                        {
                            Name = $"Finition: {gapFini} mm",
                            IsChecked = true,
                            ElementId = duplicatedElectrodes[0].ElementId
                        });
                        break;

                    case 2:
                        // Cas 2 : 2 shapes → Ebauche (petite) + Finition (grosse)
                        items.Add(new ElementItem
                        {
                            Name = $"Ebauche: {gapEb} mm",
                            IsChecked = true,
                            ElementId = duplicatedElectrodes[0].ElementId
                        });
                        items.Add(new ElementItem
                        {
                            Name = $"Finition: {gapFini} mm",
                            IsChecked = true,
                            ElementId = duplicatedElectrodes[1].ElementId
                        });
                        break;

                    default:
                        // Cas 3 : 3+ shapes → Ebauche + Demi finition(s) + Finition
                        // La plus petite → Ebauche
                        items.Add(new ElementItem
                        {
                            Name = $"Ebauche: {gapEb}mm",
                            IsChecked = true,
                            ElementId = duplicatedElectrodes[0].ElementId
                        });

                        // Les moyennes → Demi finition
                        for (int i = 1; i < duplicatedElectrodes.Count - 1; i++)
                        {
                            items.Add(new ElementItem
                            {
                                Name = $"Demi finition: {gapDemiFini} mm",
                                IsChecked = true,
                                ElementId = duplicatedElectrodes[i].ElementId
                            });
                        }

                        // La plus grosse → Finition
                        items.Add(new ElementItem
                        {
                            Name = $"Finition: {gapFini} mm",
                            IsChecked = true,
                            ElementId = duplicatedElectrodes[duplicatedElectrodes.Count - 1].ElementId
                        });
                        break;
                }

                // Ajouter l'électrode "Sans GAP" en dernier si elle existe
                if (electrodeWithoutGap != null)
                {
                    items.Add(new ElementItem
                    {
                        Name = $"Sans GAP",
                        IsChecked = false,
                        ElementId = electrodeWithoutGap.ElementId
                    });
                }

                // Ajouter tous les éléments à la ListBox
                foreach (var item in items)
                {
                    electrodeList.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la classification des électrodes: {ex.Message}",
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion 

        #region Gestion des événements boutons

        #region Bouton Quitter

        /// <summary>
        /// Gère le clic sur le bouton Quitter : sauvegarde les paramètres et ferme l'application
        /// </summary>
        /// <param name="sender">Source de l'événement</param>
        /// <param name="e">Arguments de l'événement</param>
        private void Quit_Click(object sender, RoutedEventArgs e)
        {
            // Sauvegarder les paramètres avant de quitter (par sécurité)
            Properties.Settings.Default.Save();
            
            // Déconnecter tous les hôtes TopSolid
            if (TopSolidHost.IsConnected)
            {
                TopSolidHost.Disconnect();
            }
            if (TopSolidDesignHost.IsConnected)
            {
                TopSolidDesignHost.Disconnect();
            }
            if (TopSolidDraftingHost.IsConnected)
            {
                TopSolidDraftingHost.Disconnect();
            }

            Application.Current.Shutdown();
        }

        #endregion

        #region Bouton Recharger

        /// <summary>
        /// Gère le clic sur le bouton Recharger : recharge le document TopSolid modifié
        /// </summary>
        /// <param name="sender">Source de l'événement</param>
        /// <param name="e">Arguments de l'événement</param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Recharger le document et mettre à jour l'interface
                RechargerDocument();
                
                // Afficher un message de succès seulement si un document est ouvert
                if (currentDoc?.DocId != null && !currentDoc.DocId.IsEmpty)
                {
                    MessageBox.Show("Document rechargé avec succès", "Rechargement", 
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du rechargement du document: {ex.Message}", 
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                AfficherAucunDocumentOuvert();
            }
        }

        /// <summary>
        /// Recharge le document TopSolid courant et met à jour l'interface
        /// </summary>
        private void RechargerDocument()
        {
            // Récupérer l'ID du document édité AVANT de créer l'instance
            DocumentId editedDocId = TSH.Documents.EditedDocument;

            if (editedDocId != null && !editedDocId.IsEmpty)
            {
                // Le document est ouvert, créer l'instance
                currentDoc = new Document();
                currentDoc.DocId = editedDocId;
                
                DocumentNameText.Text = currentDoc.DocNomTxt;
                DocumentNameText.Foreground = System.Windows.Media.Brushes.Black;
                
                // Recharger tous les éléments du formulaire
                InitializeForm();
            }
            else
            {
                AfficherAucunDocumentOuvert();
            }
        }

        #endregion

        #region Bouton Export

        /// <summary>
/// Gère le clic sur le bouton Export : exporte les électrodes sélectionnées
/// </summary>
/// <param name="sender">Source de l'événement</param>
/// <param name="e">Arguments de l'événement</param>
private void SelectFile_Click(object sender, RoutedEventArgs e)
{
    try
    {
        // Vérifier qu'on a un document ouvert
        if (currentDoc?.DocId == null)
        {
            MessageBox.Show("Aucun document ouvert", "Erreur", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Vérifier qu'on a un chemin d'export
        if (string.IsNullOrWhiteSpace(chemin.Text))
        {
            MessageBox.Show("Veuillez sélectionner un dossier de destination", "Erreur", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Récupérer les paramètres AVANT toute modification
        string nomDocu = GetParameterValue(currentDoc, "Nom_docu");
        string designation = GetParameterValue(currentDoc, "Designation");
        string nomElec = GetParameterValue(currentDoc, "Nom elec");

        if (string.IsNullOrEmpty(nomDocu))
        {
            MessageBox.Show("Le paramètre 'Nom_docu' n'a pas été trouvé dans le document", 
                            "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrEmpty(designation))
        {
            MessageBox.Show("Le paramètre 'Designation' n'a pas été trouvé dans le document", 
                            "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrEmpty(nomElec))
        {
            MessageBox.Show("Le paramètre 'Nom elec' n'a pas été trouvé dans le document", 
                            "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Construire le chemin d'export complet avec la structure de dossiers
        string cheminExportFinal = BuildExportPath(chemin.Text, nomDocu, designation);

        if (string.IsNullOrEmpty(cheminExportFinal))
        {
            MessageBox.Show("Impossible de créer la structure de dossiers", 
                            "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        // Créer une liste des noms et états cochés AVANT EnsureIsDirty
        List<(string Name, bool IsChecked)> electrodesToExport = new List<(string, bool)>();
        foreach (ElementItem item in electrodeList.Items)
        {
            electrodesToExport.Add((item.Name, item.IsChecked));
        }

        // IMPORTANT : Rendre le document dirty UNE SEULE FOIS au début
        DocumentId workingDocId = currentDoc.DocId;
        
        if (TSH.Application.StartModification("Préparation export électrodes", false))
        {
            try
            {
                TSH.Documents.EnsureIsDirty(ref workingDocId);
                currentDoc.DocId = workingDocId;
                TSH.Application.EndModification(true, true);
            }
            catch
            {
                TSH.Application.EndModification(false, false);
                MessageBox.Show("Impossible de préparer le document pour l'export", 
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }
        else
        {
            MessageBox.Show("Impossible de démarrer la modification du document", 
                            "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        // RECHARGER les éléments avec le nouveau DocumentId après EnsureIsDirty
        List<ElementId> elements = RecupList(workingDocId);
        List<ElementId> operations = ListOperations(elements);
        List<ElementId> operationsActives = OperationsActive(operations);
        List<ElementId> duplicateOperations = searchOperations(operationsActives, "TopSolid.Kernel.DB.Operations.DuplicateCreation");
        List<ElementId> duplicateOperationsChildElementsId = childrenElements(duplicateOperations);
        ElementId electrodeId = SearchElectrode(workingDocId, "Electrode");
        List<ElementId> ShapesList = AddShapeToList(duplicateOperationsChildElementsId, electrodeId);

        // Recréer les objets Element avec les nouveaux ElementId
        List<Element> allReloadedElements = new List<Element>();
        Element electrodeWithoutGap = null;
        
        if (ShapesList != null)
        {
            foreach (var shapeId in ShapesList)
            {
                Element element = new Element(shapeId);
                string friendlyName = TSH.Elements.GetFriendlyName(shapeId);
                
                // Traiter l'électrode "Electrode" spécialement (Sans GAP)
                if (friendlyName == "Electrode")
                {
                    electrodeWithoutGap = element;
                }
                else if (element.IsShape && element.VolumeMm3.HasValue)
                {
                    // Les autres doivent être des shapes valides avec volume
                    allReloadedElements.Add(element);
                }
            }
        }

        // Trier les électrodes dupliquées par volume croissant
        List<Element> duplicatedElectrodes = allReloadedElements.OrderBy(ele => ele.VolumeMm3).ToList();

        // Recréer la correspondance entre les noms et les nouveaux ElementId
        Dictionary<string, ElementId> electrodeMapping = new Dictionary<string, ElementId>();

        int index = 0;
        foreach (var elec in duplicatedElectrodes)
        {
            if (index < electrodesToExport.Count)
            {
                var originalItem = electrodesToExport[index];
                if (originalItem.Name != "Sans GAP")
                {
                    electrodeMapping[originalItem.Name] = elec.ElementId;
                    index++;
                }
            }
        }

        // Ajouter l'électrode sans GAP
        if (electrodeWithoutGap != null)
        {
            electrodeMapping["Sans GAP"] = electrodeWithoutGap.ElementId;
        }

        int exportCount = 0;
        int errorCount = 0;

        // Parcourir les électrodes à exporter avec les nouveaux ElementId
        foreach (var (name, isChecked) in electrodesToExport)
        {
            if (isChecked && electrodeMapping.ContainsKey(name))
            {
                ElementId electrodeElementId = electrodeMapping[name];
                ElementId representationId = ElementId.Empty;
                
                try
                {
                    // ÉTAPE 1 : Créer et activer la représentation
                    if (TSH.Application.StartModification("Création représentation", false))
                    {
                        try
                        {
                            // Créer une représentation pour isoler l'électrode
                            representationId = TSHD.Representations.CreateRepresentation(workingDocId);
                            
                            if (representationId.IsEmpty)
                            {
                                throw new Exception("La création de la représentation a échoué");
                            }

                            // Ajouter l'électrode cochée dans la représentation
                            TSHD.Representations.AddRepresentationConstituent(representationId, electrodeElementId);

                            // Activer la représentation
                            TSHD.Representations.SetCurrentRepresentation(representationId);

                            // Terminer la modification avec succès
                            TSH.Application.EndModification(true, true);
                        }
                        catch (Exception exRep)
                        {
                            TSH.Application.EndModification(false, false);
                            throw new Exception($"Erreur lors de la création de la représentation: {exRep.Message}", exRep);
                        }
                    }
                    else
                    {
                        throw new Exception("Impossible de démarrer la modification pour créer la représentation");
                    }

                    // ÉTAPE 2 : Exporter (EN DEHORS de tout contexte de modification)
                    try
                    {
                        // Construire le nom du fichier d'export
                        string nomFichier = GenerateElectrodeName(nomElec, name);

                        // Exporter au format STEP
                        OutilsTs.Export.ExportDocId(workingDocId, cheminExportFinal, nomFichier, "step");
                        
                        // Exporter au format Parasolid X_T version 31
                        OutilsTs.Export.ExportDocId(workingDocId, cheminExportFinal, nomFichier, "x_t", new Dictionary<string, string> { { "Version", "31" } });
                    }
                    catch (Exception exExport)
                    {
                        throw new Exception($"Erreur lors de l'export des fichiers: {exExport.Message}", exExport);
                    }

                    // ÉTAPE 3 : Supprimer la représentation
                    if (!representationId.IsEmpty)
                    {
                        if (TSH.Application.StartModification("Suppression représentation", false))
                        {
                            try
                            {
                                // Supprimer la représentation
                                TSH.Elements.Delete(representationId);
                                
                                TSH.Application.EndModification(true, true);
                            }
                            catch
                            {
                                TSH.Application.EndModification(false, false);
                                throw;
                            }
                        }
                    }

                    exportCount++;
                }
                catch (Exception exOuter)
                {
                    errorCount++;
                    MessageBox.Show($"Erreur pour '{name}':\n{exOuter.Message}", 
                                    "Erreur d'export", MessageBoxButton.OK, MessageBoxImage.Error);
                    
                    // Essayer de nettoyer la représentation en cas d'erreur
                    if (!representationId.IsEmpty)
                    {
                        try
                        {
                            if (TSH.Application.StartModification("Nettoyage représentation", false))
                            {
                                try
                                {
                                    TSH.Elements.Delete(representationId);
                                    TSH.Application.EndModification(true, true);
                                }
                                catch
                                {
                                    TSH.Application.EndModification(false, false);
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
        }

        // Message récapitulatif
        if (exportCount > 0)
        {
            string message = $"{exportCount} électrode(s) exportée(s) avec succès dans:\n{cheminExportFinal}";
            if (errorCount > 0)
            {
                message += $"\n{errorCount} erreur(s)";
            }
            MessageBox.Show(message, "Export terminé", MessageBoxButton.OK, 
                            errorCount > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            
            // Ouvrir le dossier d'export après que l'utilisateur a cliqué sur OK
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", cheminExportFinal);
            }
            catch (Exception exExplorer)
            {
                MessageBox.Show($"Impossible d'ouvrir le dossier: {exExplorer.Message}", 
                                "Avertissement", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        else if (errorCount > 0)
        {
            MessageBox.Show("Aucune électrode n'a pu être exportée", "Erreur", 
                            MessageBoxButton.OK, MessageBoxImage.Error);
        }
        else
        {
            MessageBox.Show("Aucune électrode cochée pour l'export", "Information", 
                            MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Erreur générale lors de l'export: {ex.Message}\n\nDétails: {ex.ToString()}", 
                        "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}

        #endregion

        #region Bouton Parcourir

        /// <summary>
        /// Gère le clic sur le bouton Parcourir : ouvre un dialogue de sélection de dossier
        /// </summary>
        /// <param name="sender">Source de l'événement</param>
        /// <param name="e">Arguments de l'événement</param>
        private void parcourir_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Sélectionner un dossier";
                dialog.ShowNewFolderButton = true;
                
                // Définir le dossier initial si un chemin existe déjà
                if (!string.IsNullOrEmpty(chemin.Text) && System.IO.Directory.Exists(chemin.Text))
                {
                    dialog.SelectedPath = chemin.Text;
                }

                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    // Convertir le chemin local en chemin réseau UNC si nécessaire
                    string uncPath = CheminReseau.ConvertToUncPath(dialog.SelectedPath);
                    chemin.Text = uncPath;
                    
                    // Sauvegarder le chemin pour la prochaine utilisation
                    Properties.Settings.Default.CheminExport = uncPath;
                    Properties.Settings.Default.Save();
                }
            }
        }

        #endregion

        #endregion

        #region Méthodes utilitaires

        /// <summary>
        /// Récupère la valeur textuelle d'un paramètre du document par son nom
        /// </summary>
        /// <param name="Doc">Document TopSolid</param>
        /// <param name="parameterName">Nom du paramètre à rechercher</param>
        /// <returns>Valeur du paramètre ou null si non trouvé</returns>
        private string GetParameterValue(Document Doc, string parameterName)
        {
            if (Doc == null || Doc.DocId == null || Doc.DocParameters == null)
            {
                return null;
            }

            foreach (var param in Doc.DocParameters)
            {
                string paramName = TSH.Elements.GetFriendlyName(param);
                
                if (paramName == parameterName)
                {
                    // Récupérer la valeur du paramètre (texte)
                    string value = TSH.Parameters.GetTextValue(param);
                    return value;
                }
            }
            
            return null;
        }

        /// <summary>
        /// Génère le nom de fichier d'export pour une électrode selon son type et gap
        /// </summary>
        /// <param name="nomElec">Nom de base de l'électrode</param>
        /// <param name="itemName">Nom de l'item contenant le type et le gap (ex: "Ebauche: 0.5 mm")</param>
        /// <returns>Nom de fichier formaté (ex: "E101Eb-G05")</returns>
        private string GenerateElectrodeName(string nomElec, string itemName)
        {
            // itemName peut être par exemple : "Ebauche: 0.5 mm", "Demi finition: 0.3 mm", "Finition: 0.1 mm", "Sans GAP"
    
            if (itemName == "Sans GAP")
            {
                return $"{nomElec}-Sans-GAP";
            }

            // Séparer le type et la valeur du gap
            string[] parts = itemName.Split(':');
            if (parts.Length != 2)
            {
                return nomElec; // Fallback
            }

            string type = parts[0].Trim(); // "Ebauche" ou "Demi finition" ou "Finition"
            string gapValue = parts[1].Trim().Replace(" mm", "").Replace(".", "").Replace(",", ""); // "0,5 mm" -> "05"
    
            // Définir l'abréviation du type
            string typeAbbr;
            if (type == "Ebauche")
            {
                typeAbbr = "Eb";
            }
            else if (type == "Demi finition")
            {
                typeAbbr = "DemiFini";
            }
            else if (type == "Finition")
            {
                typeAbbr = "Fini";
            }
            else
            {
                typeAbbr = type.Substring(0, 2); // Fallback: prendre les 2 premières lettres
            }
            
            // Construire le nom : "E101Eb-G05"
            return $"{nomElec}{typeAbbr}-G{gapValue}";
        }

        /// <summary>
        /// Construit le chemin d'export complet avec la structure de dossiers hiérarchique
        /// </summary>
        /// <param name="baseExportPath">Chemin de base de l'export</param>
        /// <param name="nomDocu">Nom du document au format "E01 Ind F test"</param>
        /// <param name="designation">Désignation du projet</param>
        /// <returns>Chemin complet créé ou null en cas d'erreur</returns>
        private string BuildExportPath(string baseExportPath, string nomDocu, string designation)
        {
            try
            {
                // Parser Nom_docu pour extraire les différentes parties
                // Format attendu : "E01 Ind F test" -> numero="E01", indice="Ind F", projet="test"
                
                string[] parts = nomDocu.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                
                if (parts.Length < 4)
                {
                    MessageBox.Show($"Le format du paramètre 'Nom_docu' est invalide.\nValeur: {nomDocu}\nFormat attendu: 'E01 Ind F test'", 
                                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return null;
                }

                // Extraire les parties du nom du document
                string numero = parts[0]; // "E01"
                string indice = $"{parts[1]} {parts[2]}"; // "Ind F"
                string projet = parts[3]; // "test"

                // Construire la structure de dossiers hiérarchique
                // {baseExportPath}\{projet}\{numero} - {designation}\{indice}\Electrode\Electrode parallélisée
                
                string cheminFinal = System.IO.Path.Combine(
                    baseExportPath,
                    projet,
                    $"{numero} - {designation}",
                    indice,
                    "Electrode",
                    "Electrode parallélisée"
                );

                // Créer tous les dossiers nécessaires s'ils n'existent pas
                if (!System.IO.Directory.Exists(cheminFinal))
                {
                    System.IO.Directory.CreateDirectory(cheminFinal);
                }

                return cheminFinal;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la création de la structure de dossiers: {ex.Message}", 
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        #endregion
    }

    #region Classes auxiliaires

    /// <summary>
    /// Représente un élément de la liste des électrodes avec son état de sélection
    /// </summary>
    public class ElementItem
    {
        /// <summary>
        /// Nom affiché de l'électrode (ex: "Ebauche: 0.5 mm")
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Indique si l'électrode est cochée pour l'export
        /// </summary>
        public bool IsChecked { get; set; }

        /// <summary>
        /// Identifiant TopSolid de l'élément
        /// </summary>
        public ElementId ElementId { get; set; }
    }

    #endregion
}

