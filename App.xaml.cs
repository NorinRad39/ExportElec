using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

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
