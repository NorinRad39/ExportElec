using System.Collections.Generic;
using System.Linq;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;

namespace OutilsTs
{
    /// <summary>
    /// Classe représentant un élément TopSolid avec ses propriétés étendues.
    /// Encapsule un ElementId et fournit un accès simplifié aux propriétés de l'élément,
    /// y compris la détection automatique des shapes et le calcul de leur volume.
    /// </summary>
    /// <remarks>
    /// Namespace: OutilsTs  
    /// Assembly: OutilsTs (in OutilsTs.dll)
    /// </remarks>
    /// <example>
    /// <code>
    /// // Créer un élément à partir d'un ElementId
    /// ElementId elementId = TSH.Elements.SearchByName(docId, "MonElement");
    /// Element element = new Element(elementId);
    /// 
    /// // Vérifier si c'est un shape et obtenir son volume
    /// if (element.IsShape)
    /// {
    ///     Console.WriteLine($"Volume: {element.VolumeMm3} mm³");
    /// }
    /// </code>
    /// </example>
    public class Element
    {
        #region Champs privés
        private ElementId elementId;
        private string friendlyName;
        private string typeFullName;
        private bool isShape;
        private double? volume; // Nullable car tous les éléments ne sont pas des shapes
        #endregion

        #region Constructeurs
        /// <summary>
        /// Initialise un nouvel élément à partir de son ElementId.
        /// Récupère automatiquement toutes les propriétés de l'élément lors de l'initialisation.
        /// </summary>
        /// <param name="elementId">Identifiant de l'élément TopSolid.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Créer un élément
        /// ElementId id = TSH.Elements.SearchByName(docId, "Electrode");
        /// Element element = new Element(id);
        /// Console.WriteLine($"Nom: {element.FriendlyName}");
        /// </code>
        /// </example>
        public Element(ElementId elementId)
        {
            // Assigne l'identifiant de l'élément
            this.elementId = elementId;
            
            // Initialise automatiquement toutes les propriétés de l'élément
            // (nom, type, volume si c'est un shape)
            InitializeElement();
        }
        #endregion

        #region Propriétés publiques
        /// <summary>
        /// Obtient l'identifiant de l'élément TopSolid.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// ElementId id = element.ElementId;
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="ElementId"/>
        /// Identifiant de l'élément.
        /// </returns>
        public ElementId ElementId 
        { 
            get => elementId; 
        }

        /// <summary>
        /// Obtient le nom convivial de l'élément.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Le nom convivial est le nom affiché dans l'interface TopSolid.
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// string nom = element.FriendlyName; // Ex: "Electrode_1"
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Nom convivial de l'élément, ou "Nom inconnu" en cas d'erreur.
        /// </returns>
        public string FriendlyName 
        { 
            get => friendlyName; 
        }

        /// <summary>
        /// Obtient le nom complet du type de l'élément.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Le nom complet du type inclut le namespace complet de la classe TopSolid.
        /// Exemple: "TopSolid.Kernel.DB.D3.Shapes.Prism"
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// string type = element.TypeFullName;
        /// // Ex: "TopSolid.Kernel.DB.D3.Shapes.Prism"
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="string"/>
        /// Nom complet du type de l'élément, ou "Type inconnu" en cas d'erreur.
        /// </returns>
        public string TypeFullName 
        { 
            get => typeFullName; 
        }

        /// <summary>
        /// Indique si l'élément est un shape (forme 3D).
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Un shape est identifié par son type qui commence par "TopSolid.Kernel.DB.D3.Shapes.".
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// if (element.IsShape)
        /// {
        ///     Console.WriteLine("C'est un shape !");
        /// }
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si l'élément est un shape, sinon <c>false</c>.
        /// </returns>
        public bool IsShape 
        { 
            get => isShape; 
        }

        /// <summary>
        /// Obtient le volume du shape en mètres cubes (m³).
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Cette propriété est <c>null</c> si l'élément n'est pas un shape.
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// if (element.Volume.HasValue)
        /// {
        ///     Console.WriteLine($"Volume: {element.Volume.Value} m³");
        /// }
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="Nullable{Double}"/>
        /// Volume en m³, ou <c>null</c> si l'élément n'est pas un shape.
        /// </returns>
        public double? Volume 
        { 
            get => volume; 
        }

        /// <summary>
        /// Obtient le volume du shape en millimètres cubes (mm³).
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// Cette propriété effectue automatiquement la conversion de m³ en mm³.
        /// Conversion: 1 m³ = 1 000 000 000 mm³.
        /// Cette propriété est <c>null</c> si l'élément n'est pas un shape.
        /// </remarks>
        /// <example>
        /// <code>
        /// Element element = new Element(elementId);
        /// if (element.VolumeMm3.HasValue)
        /// {
        ///     Console.WriteLine($"Volume: {element.VolumeMm3.Value:F2} mm³");
        /// }
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="Nullable{Double}"/>
        /// Volume en mm³, ou <c>null</c> si l'élément n'est pas un shape.
        /// </returns>
        public double? VolumeMm3 
        { 
            // Cast explicite nécessaire pour C# 7.3
            get => volume.HasValue ? (double?)(volume.Value * 1_000_000_000) : null; 
        }
        #endregion

        #region Méthodes privées
        /// <summary>
        /// Initialise les propriétés de l'élément en interrogeant l'API TopSolid.
        /// Cette méthode est appelée automatiquement par le constructeur.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// 
        /// Étapes d'initialisation :
        /// 1. Récupération du nom convivial (FriendlyName)
        /// 2. Récupération du type complet (TypeFullName)
        /// 3. Détection si c'est un shape (analyse du nom du type)
        /// 4. Si c'est un shape, récupération du volume
        /// 
        /// En cas d'erreur sur une propriété, une valeur par défaut est assignée.
        /// </remarks>
        private void InitializeElement()
        {
            // Vérifier si l'ElementId est valide
            // Si vide, on arrête l'initialisation
            if (elementId.IsEmpty) return;

            // --- Étape 1 : Récupération du nom convivial ---
            try
            {
                // Récupérer le nom convivial depuis l'API TopSolid
                friendlyName = TSH.Elements.GetFriendlyName(elementId);
            }
            catch
            {
                // En cas d'erreur, on assigne un nom par défaut
                friendlyName = "Nom inconnu";
            }

            // --- Étape 2 : Récupération du type complet ---
            try
            {
                // Récupérer le nom complet du type depuis l'API TopSolid
                // Ex: "TopSolid.Kernel.DB.D3.Shapes.Prism"
                typeFullName = TSH.Elements.GetTypeFullName(elementId);
            }
            catch
            {
                // En cas d'erreur, on assigne un type par défaut
                typeFullName = "Type inconnu";
            }

            // --- Étape 3 & 4 : Détection shape et récupération du volume ---
            try
            {
                // Vérifier si c'est un shape en analysant le nom du type
                // Un shape a un type qui commence par "TopSolid.Kernel.DB.D3.Shapes."
                isShape = !string.IsNullOrEmpty(typeFullName) && 
                          typeFullName.StartsWith("TopSolid.Kernel.DB.D3.Shapes.");

                // Si c'est un shape, récupérer son volume en m³
                if (isShape)
                {
                    volume = TSH.Shapes.GetShapeVolume(elementId);
                }
                else
                {
                    // Si ce n'est pas un shape, le volume est null
                    volume = null;
                }
            }
            catch
            {
                // En cas d'erreur, on considère que ce n'est pas un shape
                isShape = false;
                volume = null;
            }
        }
        #endregion
    }

    /// <summary>
    /// Méthodes d'extension pour faciliter la manipulation des éléments TopSolid.
    /// Fournit des méthodes utilitaires pour convertir et trier les éléments.
    /// </summary>
    /// <remarks>
    /// Namespace: OutilsTs  
    /// Assembly: OutilsTs (in OutilsTs.dll)
    /// </remarks>
    /// <example>
    /// <code>
    /// // Convertir une liste d'ElementId en éléments enrichis
    /// List&lt;ElementId&gt; ids = TSH.Shapes.GetShapes(docId);
    /// List&lt;Element&gt; elements = ids.ToElements();
    /// 
    /// // Trier les shapes par volume
    /// List&lt;Element&gt; shapesTries = elements.GetShapesSortedByVolume();
    /// </code>
    /// </example>
    public static class ElementExtensions
    {
        /// <summary>
        /// Convertit une liste d'ElementId en liste d'objets Element enrichis.
        /// Chaque ElementId est encapsulé dans un objet Element qui fournit
        /// un accès simplifié aux propriétés de l'élément.
        /// </summary>
        /// <param name="elementIds">Liste des identifiants d'éléments à convertir.</param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// 
        /// Cette méthode d'extension permet une conversion fluide des ElementId en objets Element.
        /// Si la liste source est null, retourne une liste vide.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Récupérer tous les shapes d'un document
        /// List&lt;ElementId&gt; shapeIds = TSH.Shapes.GetShapes(docId);
        /// 
        /// // Convertir en objets Element
        /// List&lt;Element&gt; elements = shapeIds.ToElements();
        /// 
        /// // Utiliser les propriétés enrichies
        /// foreach (var element in elements)
        /// {
        ///     Console.WriteLine($"{element.FriendlyName}: {element.VolumeMm3} mm³");
        /// }
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{Element}"/>
        /// Liste des objets Element, ou liste vide si elementIds est null.
        /// </returns>
        public static List<Element> ToElements(this List<ElementId> elementIds)
        {
            // Créer une nouvelle liste pour stocker les éléments convertis
            List<Element> elements = new List<Element>();
            
            // Si la liste source est null, retourner une liste vide
            if (elementIds == null) return elements;
            
            // Convertir chaque ElementId en objet Element
            foreach (var id in elementIds)
            {
                elements.Add(new Element(id));
            }
            
            return elements;
        }

        /// <summary>
        /// Filtre uniquement les shapes d'une liste d'éléments et les trie par volume.
        /// Par défaut, le tri est décroissant (du plus gros au plus petit volume).
        /// </summary>
        /// <param name="elements">Liste des éléments à filtrer et trier.</param>
        /// <param name="descending">
        /// <c>true</c> pour un tri décroissant (par défaut), 
        /// <c>false</c> pour un tri croissant.
        /// </param>
        /// <remarks>
        /// Namespace: OutilsTs  
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// 
        /// Cette méthode effectue deux opérations :
        /// 1. Filtre les éléments pour ne garder que les shapes ayant un volume
        /// 2. Trie les shapes par volume (décroissant ou croissant)
        /// 
        /// Les éléments qui ne sont pas des shapes ou qui n'ont pas de volume sont exclus.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Récupérer et convertir les éléments
        /// List&lt;ElementId&gt; ids = TSH.Shapes.GetShapes(docId);
        /// List&lt;Element&gt; elements = ids.ToElements();
        /// 
        /// // Trier par volume décroissant (plus gros en premier)
        /// List&lt;Element&gt; shapesTries = elements.GetShapesSortedByVolume();
        /// 
        /// // Afficher avec numérotation
        /// for (int i = 0; i &lt; shapesTries.Count; i++)
        /// {
        ///     Console.WriteLine($"{i + 1}. {shapesTries[i].FriendlyName} - {shapesTries[i].VolumeMm3:F2} mm³");
        /// }
        /// 
        /// // Trier par volume croissant (plus petit en premier)
        /// List&lt;Element&gt; shapesCroissant = elements.GetShapesSortedByVolume(descending: false);
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="List{Element}"/>
        /// Liste des shapes triés par volume.
        /// </returns>
        public static List<Element> GetShapesSortedByVolume(this List<Element> elements, bool descending = true)
        {
            // Filtrer pour ne garder que les shapes ayant un volume
            // LINQ Where : filtre les éléments selon la condition
            // ToList : convertit le résultat en List<Element>
            var shapes = elements.Where(e => e.IsShape && e.Volume.HasValue).ToList();
            
            // Trier par volume
            if (descending)
            {
                // Tri décroissant : b comparé à a (inversion)
                // Le plus gros volume sera en premier
                shapes.Sort((a, b) => b.Volume.Value.CompareTo(a.Volume.Value));
            }
            else
            {
                // Tri croissant : a comparé à b (ordre normal)
                // Le plus petit volume sera en premier
                shapes.Sort((a, b) => a.Volume.Value.CompareTo(b.Volume.Value));
            }
            
            return shapes;
        }
    }
}