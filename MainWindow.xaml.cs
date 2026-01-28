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
    /// Logique d'interaction pour App.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private Document currentDoc;
        private StartConnect startConnect;

        public MainWindow()
        {
            InitializeComponent();
            startConnect = new StartConnect();
            startConnect.ConnectionTopsolid();
            InitializeForm();

        }
        private void InitializeForm()
        {
            // Initialisation de currentDoc
            currentDoc = new Document();
            currentDoc.DocId = TSH.Documents.EditedDocument;

            if (currentDoc.DocId != null)
            {
                DocumentNameText.Text = currentDoc.DocNomTxt;
            }
            else
            {
                DocumentNameText.Text = "aucun document ouvert";
            }


            List<ElementId> elements = RecupList(currentDoc.DocId);
            List<ElementId> operations = ListOperations(elements);
            List<ElementId> operationsActives = OperationsActive(operations);
            List<ElementId> duplicateOperations = searchOperations(operationsActives, "TopSolid.Kernel.DB.Operations.DuplicateCreation");
            List<ElementId> duplicateOperationsChildElementsId = childrenElements(duplicateOperations);
            ElementId electrodeId = SearchElectrode(currentDoc.DocId, "Electrode");
            List<ElementId> ShapesList = AddShapeToList(duplicateOperationsChildElementsId, electrodeId);
            List<double> shapeVolumes = RecupShapesVolume(ShapesList);

            PopulateListDoublesBox(shapeVolumes);
        }


        private List<double> RecupShapesVolume(List<ElementId> shapes)
        {
            List<double> shapeVolumes = new List<double>();

            try
            {
                if (shapes != null)
                {
                    foreach (var shape in shapes)
                    {
                        double shapeVolume = TSH.Shapes.GetShapeVolume(shape);
                        shapeVolumes.Add(shapeVolume);
                    }
;
                    return shapeVolumes;
                }
                else
                {
                    DocumentNameText.Text = "aucun document ouvert, la liste des shapes ne peut pas être chargée";
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des shapes: {ex.Message}",
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

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


        private List<ElementId> AddShapeToList(List<ElementId> duplicateOperationsChildElementsId, ElementId electrodeId)
        {
            List<ElementId> shapeList = new List<ElementId>();
            try
            {
                if (duplicateOperationsChildElementsId.Count > 0 && !electrodeId.IsEmpty)
                {
                    duplicateOperationsChildElementsId.Add(electrodeId);
                    return duplicateOperationsChildElementsId;
                }
                else
                {
                    DocumentNameText.Text = "aucun document ouvert, la liste des shapes ne peut pas être chargée";
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des shapes: {ex.Message}",
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        void PopulateListDoublesBox(List<double> shapeVolumes)
        {
            // Vider la ListBox avant de la remplir
            electrodeList.Items.Clear();

            if (shapeVolumes == null || shapeVolumes.Count == 0)
            {
                electrodeList.Items.Add("Aucun élément trouvé");
                return;
            }

            foreach (var shapeVolume in shapeVolumes)
            {
                try
                {
                    // Convertir de m³ en mm³ (1 m³ = 1 000 000 000 mm³)
                    double volumeMm3 = shapeVolume * 1_000_000_000;
                    
                    // Formater le volume pour l'affichage
                    string volumeText = $"Volume: {volumeMm3:F2} mm³";
                    
                    // Créer un objet ElementItem avec le volume
                    var item = new ElementItem
                    {
                        Name = volumeText,
                        IsChecked = false,
                        ElementId = ElementId.Empty // Pas d'ElementId pour un volume
                    };
                    
                    // Ajouter l'élément à la ListBox
                    electrodeList.Items.Add(item);
                }
                catch
                {
                    // En cas d'erreur, afficher le volume brut
                    double volumeMm3 = shapeVolume * 1_000_000_000;
                    electrodeList.Items.Add(new ElementItem 
                    { 
                        Name = $"Volume (erreur): {volumeMm3}",
                        IsChecked = false,
                        ElementId = ElementId.Empty
                    });
                }
            }
        }

        private List<ElementId> RecupList(DocumentId DocId)
        {
            List<ElementId> elements = new List<ElementId>();

            try
            {
                if (currentDoc.DocId != null)
                {
                    elements = currentDoc.DocElements;

                    return elements;
                }
                else
                {
                    DocumentNameText.Text = "aucun document ouvert, la liste des elements ne peut pas être chargée";
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des éléments: {ex.Message}",
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private List<ElementId> ListOperations(List<ElementId> elements)
        {
            List<ElementId> operations = new List<ElementId>();

            if (elements != null || elements.Count > 0)
            {
                foreach (var element in elements)
                {
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


        private List<ElementId> searchOperations (List<ElementId> operationsActive, string operationName)
        {
            List<ElementId> duplicateOperations = new List<ElementId>();

            if (operationsActive != null || operationsActive.Count > 0)
            {
                foreach (var operation in operationsActive)
                {
                    // Récupérer le nom de l'élément via l'API TopSolid
                    string operationFullName = TSH.Elements.GetTypeFullName(operation);
                    
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
        private List<ElementId> OperationsActive(List<ElementId> operations)
        {
            List<ElementId> operationsActive = new List<ElementId>();

            if (operations != null && operations.Count > 0)
            {
                foreach (var operation in operations)
                {
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

        private List<ElementId> childrenElements(List<ElementId> Operations)
        {
            List<ElementId> childrenElementsId = new List<ElementId>();

            if (Operations == null)
            {
                return null;
            }

            foreach (var operation in Operations)
            {
                List<ElementId> children = TSH.Operations.GetChildren(operation);
                
                if (children != null && children.Count > 0)
                {
                    childrenElementsId.AddRange(children); // Ajouter tous les enfants à la liste
                }
            }
            
            return childrenElementsId;
        }

        

        private List<ElementId> ElementsConstituant(List<ElementId> elementIds)
        {
            List<ElementId> setElementIds = new List<ElementId>();
            if (elementIds != null)
            {
                foreach (var elementId in elementIds)
                {
                    List<ElementId> constituents = TSH.Elements.GetConstituents(elementId);
                    if (constituents != null && constituents.Count > 0)
                    {
                        setElementIds.AddRange(constituents);
                    }
                }               
            }
            return setElementIds;
        }

        


        void PopulateListBox(List<ElementId> operationsActive)
        {
            // Vider la ListBox avant de la remplir
            electrodeList.Items.Clear();

            if (operationsActive == null && operationsActive.Count == 0)
            {
                electrodeList.Items.Add("Aucun élément trouvé");
                return;
            }

            foreach (var operationActive in operationsActive)
            {
                try
                {
                    // Récupérer le nom de l'élément via l'API TopSolid
                    string operationName = TSH.Elements.GetTypeFullName(operationActive);

                    // Créer un objet ElementItem au lieu d'ajouter directement la chaîne
                    var item = new ElementItem
                    {
                        Name = operationName,
                        IsChecked = false,
                        ElementId = operationActive
                    };

                    // Ajouter l'élément à la ListBox
                    electrodeList.Items.Add(item);
                }
                catch (Exception ex)
                {
                    // En cas d'erreur, afficher l'ID de l'élément
                    electrodeList.Items.Add(new ElementItem { Name = $"Élément (erreur): {operationActive}", IsChecked = false, ElementId = operationActive });
                }
            }
        }


        private void Quit_Click(object sender, RoutedEventArgs e)
        {
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



        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {

        }



        private void parcourir_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Sélectionner un dossier";
                dialog.ShowNewFolderButton = true;

                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    // Utiliser la méthode de la DLL OutilsTs
                    string uncPath = CheminReseau.ConvertToUncPath(dialog.SelectedPath);
                    chemin.Text = uncPath;
                }
            }
        }


    }

    // Ajouter cette classe dans le fichier MainWindow.xaml.cs
    public class ElementItem
    {
        public string Name { get; set; }
        public bool IsChecked { get; set; }
        public ElementId ElementId { get; set; }
    }
}

