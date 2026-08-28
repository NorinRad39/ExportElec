using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Serialization;
using AutoUpdaterDotNET;

namespace ExportElec
{
    /// <summary>
    /// Logique d'interaction pour App.xaml
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// Constructeur de l'application avec gestion des exceptions non gérées
        /// </summary>
        public App()
        {
            // Capturer les exceptions non gérées de l'UI thread
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;

            // Capturer les exceptions non gérées des threads en arrière-plan
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        /// <summary>Adresse du descripteur de mise a jour, sur le partage reseau.</summary>
        /// <remarks>
        /// Chemin UNC et non une lettre de lecteur, qui designerait pourtant le meme dossier : un
        /// mappage est propre a la session, et sur un poste ou il manque les mises a jour
        /// cesseraient sans que personne ne s'en apercoive.
        ///
        /// Cette adresse ne doit plus bouger. Chaque poste la porte en dur dans l'executable qu'il
        /// a installe : la deplacer coupe des mises a jour tous ceux deja en service, et sans le
        /// moindre signe, puisqu'un partage injoignable laisse l'application demarrer.
        /// </remarks>
        private const string DescripteurMiseAJour =
            @"\\jbtec-be\meca$\topsolid\ExportElec\update.xml";

        /// <summary>
        /// Controle la version avant d'ouvrir la fenetre principale.
        /// </summary>
        /// <remarks>
        /// La fenetre est ouverte ici, et non par <c>StartupUri</c> : WPF traite cet attribut juste
        /// apres <c>OnStartup</c>, donc la fenetre s'afficherait malgre un refus de mise a jour.
        /// L'attribut a ete retire d'App.xaml pour que ce point de passage soit le seul.
        /// </remarks>
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Controle de version avant toute fenetre : une version en retard ne doit meme pas
            // s'ouvrir sur un dossier.
            if (!PeutDemarrer())
            {
                Shutdown();
                return;
            }

            new MainWindow().Show();
        }

        /// <summary>
        /// Indique si l'application est autorisee a demarrer, apres controle de sa version.
        /// </summary>
        /// <remarks>
        /// La mise a jour peut etre refusee, mais l'application ne demarre pas tant qu'elle n'est
        /// pas faite : tous les postes doivent tourner sur la meme version, car l'export ecrit dans
        /// le PDM des donnees qu'une version en retard ne relit pas forcement de la meme facon.
        ///
        /// Le controle est fait ici, et non par <c>AutoUpdater.Start</c> : celui-ci mene le deroule
        /// de bout en bout et ne dit pas si l'utilisateur a refuse. Ses boites de dialogue
        /// s'affichent tres bien sans boucle de messages, elles pompent la leur.
        /// </remarks>
        private static bool PeutDemarrer()
        {
            UpdateInfoEventArgs descripteur;

            try
            {
                descripteur = LireDescripteurMiseAJour();
            }
            catch (Exception ex)
            {
                // Partage injoignable, poste hors reseau : on laisse l'atelier travailler plutot
                // que de l'immobiliser sur un incident de reseau. Ne pas savoir n'est pas la meme
                // chose que savoir qu'une version manque.
                Console.WriteLine($"[Mise a jour] Verification impossible : {ex.Message}");
                return true;
            }

            Version installee = Assembly.GetExecutingAssembly().GetName().Version;
            if (descripteur == null || string.IsNullOrWhiteSpace(descripteur.CurrentVersion)) return true;
            if (new Version(descripteur.CurrentVersion) <= installee) return true;

            MessageBoxResult reponse = MessageBox.Show(
                $"La version {descripteur.CurrentVersion} est disponible ; ce poste utilise la {installee}.\n\n"
                + "ExportElec ne peut pas s'ouvrir tant que la mise a jour n'est pas faite.",
                "Mise a jour requise",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Information);

            if (reponse != MessageBoxResult.OK)
            {
                Console.WriteLine("[Mise a jour] Refusee : l'application ne demarre pas.");
                return false;
            }

            // L'installation se fait par utilisateur, dans %LOCALAPPDATA% : aucune elevation n'est
            // necessaire. Sans ce reglage, AutoUpdater lance le programme d'installation avec le
            // verbe « runas » et declenche une invite UAC pour rien.
            AutoUpdater.RunUpdateAsAdmin = false;

            // Rend la main une fois le programme d'installation lance : celui-ci remplace les
            // fichiers puis relance l'application, il ne faut donc pas la laisser demarrer ici.
            if (AutoUpdater.DownloadUpdate(descripteur)) return false;

            MessageBox.Show(
                "Le telechargement de la mise a jour n'a pas abouti.\n\n"
                + "Verifiez l'acces au reseau, puis relancez ExportElec.",
                "Mise a jour",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            return false;
        }

        /// <summary>
        /// Lit le descripteur publie sur le partage.
        /// </summary>
        /// <remarks>
        /// Meme format et meme mecanisme que ceux d'AutoUpdater — WebClient sait lire une URI
        /// file://, donc un chemin UNC — mais lu ici pour garder la decision de demarrer.
        /// </remarks>
        private static UpdateInfoEventArgs LireDescripteurMiseAJour()
        {
            using (WebClient client = new WebClient())
            {
                // Sans cela, un update.xml fraichement publie peut rester masque par le cache le
                // temps que les postes le voient.
                client.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);

                string xml = client.DownloadString(new Uri(DescripteurMiseAJour));

                XmlSerializer serialiseur = new XmlSerializer(typeof(UpdateInfoEventArgs));
                using (StringReader lecteur = new StringReader(xml))
                {
                    return (UpdateInfoEventArgs)serialiseur.Deserialize(lecteur);
                }
            }
        }

        /// <summary>
        /// Gère les exceptions non gérées sur le thread UI
        /// </summary>
        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            string errorMessage = $"Une erreur non gérée s'est produite:\n\n" +
                                  $"Message: {e.Exception.Message}\n\n" +
                                  $"Type: {e.Exception.GetType().Name}\n\n" +
                                  $"Source: {e.Exception.Source}\n\n" +
                                  $"StackTrace:\n{e.Exception.StackTrace}";

            if (e.Exception.InnerException != null)
            {
                errorMessage += $"\n\nException interne:\n{e.Exception.InnerException.Message}";
            }

            MessageBox.Show(errorMessage, "Erreur non gérée", MessageBoxButton.OK, MessageBoxImage.Error);

            // Marquer l'exception comme gérée pour éviter le crash
            e.Handled = true;
        }

        /// <summary>
        /// Gère les exceptions non gérées sur les threads en arrière-plan
        /// </summary>
        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;

            string errorMessage = "Une erreur fatale s'est produite:\n\n";

            if (ex != null)
            {
                errorMessage += $"Message: {ex.Message}\n\n" +
                               $"Type: {ex.GetType().Name}\n\n" +
                               $"StackTrace:\n{ex.StackTrace}";
            }
            else
            {
                errorMessage += $"Exception: {e.ExceptionObject}";
            }

            MessageBox.Show(errorMessage, "Erreur fatale", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
