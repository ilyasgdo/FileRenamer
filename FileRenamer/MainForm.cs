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
        private Label lblCombo1;
        private Label lblCombo2;
        private Label lblTxt1;


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
            lblCombo1 = new Label();
            lblCombo2 = new Label();
            lblTxt1 = new Label();




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

            comboBox1.FormattingEnabled = true;
            comboBox2.FormattingEnabled = true;

            // Positionnement des ComboBox sur le formulaire
            comboBox1.Location = new Point(227, 72);
            comboBox2.Location = new Point(227, 140);

            //taille des combo 
            comboBox2.Size = new Size(152, 23);
            comboBox1.Size = new Size(152, 23);
            // 
            // lblCombo1
            // 
            lblCombo1.AutoSize = true;
            lblCombo1.Location = new Point(241, 44);
            lblCombo1.Name = "lblCombo1";
            lblCombo1.Size = new Size(131, 15);
            lblCombo1.TabIndex = 4;
            lblCombo1.Text = "selectionner le prefixe 1";
            // 
            // lblCombo2
            // 
            lblCombo2.AutoSize = true;
            lblCombo2.Location = new Point(241, 111);
            lblCombo2.Name = "lblCombo2";
            lblCombo2.Size = new Size(131, 15);
            lblCombo2.TabIndex = 5;
            lblCombo2.Text = "selectionner le prefixe 2";
            // 
            // lblTxt1
            // 
            lblTxt1.AutoSize = true;
            lblTxt1.Location = new Point(257, 183);
            lblTxt1.Name = "lblTxt1";
            lblTxt1.Size = new Size(93, 15);
            lblTxt1.TabIndex = 6;
            lblTxt1.Text = "saisir le prefixe 3";

            // Ajout des labels au formulaire
            Controls.Add(lblCombo1);
            Controls.Add(lblCombo2);
            Controls.Add(lblTxt1);

            // Ajout des ComboBox au formulaire
            Controls.Add(comboBox1);
            Controls.Add(comboBox2);

            // Création du TextBox
            textBoxNewName = new TextBox();

            // Positionnement du TextBox sur le formulaire
            textBoxNewName.Location = new Point(227, 213);


            // size  du TextBox sur le formulaire
            textBoxNewName.Size = new Size(152, 23);

            // Ajout du TextBox au formulaire
            Controls.Add(textBoxNewName);

            // Création du Panel pour le glisser-déposer
            panelDragDrop = new Panel();

            // Positionnement du Panel sur le formulaire
            panelDragDrop.Location = new Point(145, 287);
            panelDragDrop.Size = new Size(323, 100);
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
                    // Envoyer le fichier vers SharePoint
                    UploadFileToSharePoint(newFilePath);


                    MessageBox.Show("Le fichier a été renommé et enregistré avec succès.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Une erreur s'est produite lors du renommage et de l'enregistrement du fichier : " + ex.Message);
                }
            }
        }
        private void UploadFileToSharePoint(string filePath)
        {
            using (SPClient.ClientContext context = new SPClient.ClientContext("https://your-sharepoint-site-url"))
            {
                // Authentification (remplacez "your-username" et "your-password" par vos informations d'identification SharePoint)
                string userName = "your-username";
                string password = "your-password";

                context.Credentials = new NetworkCredential(userName, password);

                // Obtention du site SharePoint
                SPClient.Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                // Chemin relatif du dossier SharePoint dans lequel vous souhaitez télécharger les fichiers
                string targetFolderUrl = "Shared Documents/Folder/Subfolder";

                // Chargement du dossier cible SharePoint
                SPClient.Folder targetFolder = web.GetFolderByServerRelativeUrl(targetFolderUrl);
                context.Load(targetFolder);
                context.ExecuteQuery();

                // Lecture du contenu du fichier
                byte[] fileContent = File.ReadAllBytes(filePath);

                // Création du fichier dans SharePoint
                SPClient.FileCreationInformation fileInfo = new SPClient.FileCreationInformation();
                fileInfo.Content = fileContent;
                fileInfo.Url = Path.GetFileName(filePath);
                SPClient.File newFile = targetFolder.Files.Add(fileInfo);
                context.Load(newFile);
                context.ExecuteQuery();
            }
        }


    }
}
