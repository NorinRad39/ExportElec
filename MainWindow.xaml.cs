using OutilsTs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using TopSolid.Kernel.Automating;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cad.Drafting.Automating;
using TopSolid.Cam.NC.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;
using System.Runtime.InteropServices;
using System.Text;

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
}

