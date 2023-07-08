using System;
using System.IO;
using System.Net;
using System.Windows.Forms;
using SPClient = Microsoft.SharePoint.Client;

namespace FileRenamer
{
    public partial class MainForm : Form
    {
        private ComboBox comboBox1;
        private ComboBox comboBox2;
        private TextBox textBoxNewName;
        private Panel panelDragDrop;
        

        public MainForm()
        {
            InitializeComponent();
            Load += MainForm_Load;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            GenerateControls();
        }

        private void GenerateControls()
        {
            // Création des ComboBox
            comboBox1 = new ComboBox();
            comboBox2 = new ComboBox();

            // Ajout des choix prédéfinis
            comboBox1.Items.Add("meca");
            comboBox1.Items.Add("as");
            comboBox1.Items.Add("autreTRuc");
            comboBox2.Items.Add("premierr");
            comboBox2.Items.Add("deuxx");
            comboBox2.Items.Add("troiss");
            comboBox2.Items.Add("quatree");
            comboBox2.Items.Add("cinque");
            comboBox2.Items.Add("sixx");

            // Positionnement des ComboBox sur le formulaire
            comboBox1.Location = new System.Drawing.Point(257, 87);
            comboBox2.Location = new System.Drawing.Point(257, 156);

            // Ajout des ComboBox au formulaire
            Controls.Add(comboBox1);
            Controls.Add(comboBox2);

            // Création du TextBox
            textBoxNewName = new TextBox();

            // Positionnement du TextBox sur le formulaire
            textBoxNewName.Location = new System.Drawing.Point(257, 214);

            // Ajout du TextBox au formulaire
            Controls.Add(textBoxNewName);

            // Création du Panel pour le glisser-déposer
            panelDragDrop = new Panel();

            // Positionnement du Panel sur le formulaire
            panelDragDrop.Location = new System.Drawing.Point(165, 273);
            panelDragDrop.Size = new System.Drawing.Size(200, 200);
            panelDragDrop.BorderStyle = BorderStyle.FixedSingle;

            // Activation de la fonctionnalité de glisser-déposer sur le panel
            panelDragDrop.AllowDrop = true;
            panelDragDrop.DragEnter += PanelDragDrop_DragEnter;
            panelDragDrop.DragDrop += PanelDragDrop_DragDrop;

            // Ajout du Panel au formulaire
            Controls.Add(panelDragDrop);
        }
       

        private void buttonRename_Click(object sender, EventArgs e)
        {
            string selectedName = comboBox1.SelectedItem?.ToString();
            string selectedSuffix = comboBox2.SelectedItem?.ToString();
            string newName = textBoxNewName.Text;

            // Vérification si tous les champs sont remplis
            if (string.IsNullOrEmpty(selectedName) || string.IsNullOrEmpty(selectedSuffix) || string.IsNullOrEmpty(newName))
            {
                MessageBox.Show("Veuillez sélectionner un nom, un suffixe et saisir un nouveau nom pour les fichiers.");
                return;
            }

            try
            {
                // SharePoint site URL
                string siteUrl = "https://yoursharepointsiteurl";

                // SharePoint target library relative URL
                string libraryUrl = "/sites/yoursite/Shared Documents/";

                // SharePoint credentials
                string username = "id";
                string password = "mdp";

                // Connect to SharePoint site
                SPClient.ClientContext clientContext = new SPClient.ClientContext(siteUrl);
                clientContext.Credentials = new NetworkCredential(username, password);

                // Get reference to the target library
                SPClient.List targetLibrary = clientContext.Web.Lists.GetByTitle("Documents");
                clientContext.Load(targetLibrary);
                clientContext.ExecuteQuery();


                // Récupération du dossier "Documents" de l'utilisateur
                string documentsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "doc");
                string[] existingFiles = Directory.GetFiles(documentsPath, "*", SearchOption.TopDirectoryOnly);
                

                // Parcours des fichiers dans le dossier "Documents"
                foreach (string filePath in Directory.GetFiles(documentsPath))
                {
                    // Récupération du nom de fichier sans le chemin complet
                    string fileName = Path.GetFileName(filePath);

                    // Vérification si le fichier correspond au nom sélectionné
                    if (fileName.StartsWith(selectedName))
                    {
                        // Construction du nouveau nom de fichier avec le nouveau nom saisi, le suffixe sélectionné et le numéro unique
                        string newFileName = $"{selectedName}_{selectedSuffix}_{newName}{Path.GetExtension(fileName)}";

                        // Chemin complet du nouveau fichier dans le dossier "Documents"
                        string newFilePath = Path.Combine(documentsPath, newFileName);

                        // Lecture du fichier en tant que tableau d'octets
                        byte[] fileContent = File.ReadAllBytes(filePath);

                        // Création du fichier dans SharePoint
                        SPClient.FileCreationInformation fileInfo = new SPClient.FileCreationInformation();
                        fileInfo.Content = fileContent;
                        fileInfo.Url = libraryUrl + newFileName;

                        SPClient.File uploadedFile = targetLibrary.RootFolder.Files.Add(fileInfo);
                        clientContext.Load(uploadedFile);
                        clientContext.ExecuteQuery();

                        // Suppression du fichier local après le téléchargement
                        File.Delete(filePath);
                    }
                }

                MessageBox.Show("Les fichiers ont été renommés et téléchargés avec succès dans SharePoint.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur s'est produite lors du renommage des fichiers et du téléchargement vers SharePoint : " + ex.Message);
            }
        }

        private void PanelDragDrop_DragEnter(object sender, DragEventArgs e)
        {
            // Vérification si l'objet peut être traité en tant que fichier
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void PanelDragDrop_DragDrop(object sender, DragEventArgs e)
        {
            // Récupération du chemin du fichier depuis l'objet déposé
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            // Vérification s'il y a au moins un fichier
            if (files.Length > 0)
            {
                string filePath = files[0];
                string fileName = Path.GetFileName(filePath);

                // Construction du nouveau nom de fichier avec le nouveau nom saisi, le suffixe sélectionné et le numéro unique
                string selectedName = comboBox1.SelectedItem?.ToString();
                string selectedSuffix = comboBox2.SelectedItem?.ToString();
                string newName = textBoxNewName.Text;
                string newFileName = $"{selectedName}_{selectedSuffix}_{newName}{Path.GetExtension(fileName)}";

                // Récupération du dossier "Documents" de l'utilisateur
                string documentsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "doc");


                // Chemin complet du nouveau fichier dans le dossier "Documents"
                string newFilePath = Path.Combine(documentsPath, newFileName);

                try
                {
                    // Renommage et déplacement du fichier vers le nouveau chemin
                    File.Move(filePath, newFilePath);

                    MessageBox.Show("Le fichier a été renommé et enregistré avec succès.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Une erreur s'est produite lors du renommage et de l'enregistrement du fichier : " + ex.Message);
                }
            }
        }
    }
}
