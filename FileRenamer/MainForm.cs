using System;
using System.IO;
using System.Windows.Forms;

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
            comboBox1.Items.Add("selectiones");
            comboBox1.Items.Add("un");
            comboBox1.Items.Add("truc");
            comboBox2.Items.Add("hello");
            comboBox2.Items.Add("holla");
            comboBox2.Items.Add("hallo");

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

            // Récupération du dossier "Documents" de l'utilisateur
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            try
            {
                // Parcours des fichiers dans le dossier "Documents"
                foreach (string filePath in Directory.GetFiles(documentsPath))
                {
                    // Récupération du nom de fichier sans le chemin complet
                    string fileName = Path.GetFileName(filePath);

                    // Vérification si le fichier correspond au nom sélectionné
                    if (fileName.StartsWith(selectedName))
                    {
                        // Construction du nouveau nom de fichier avec le nouveau nom saisi et le suffixe sélectionné
                        string newFileName = newName + "_" + selectedSuffix + Path.GetExtension(fileName);

                        // Chemin complet du nouveau fichier dans le dossier "Documents"
                        string newFilePath = Path.Combine(documentsPath, newFileName);

                        // Renommage et déplacement du fichier vers le nouveau chemin
                        File.Move(filePath, newFilePath);
                    }
                }

                MessageBox.Show("Les fichiers ont été renommés avec succès.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur s'est produite lors du renommage des fichiers : " + ex.Message);
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

                // Construction du nouveau nom de fichier avec le nouveau nom saisi et le suffixe sélectionné
                string selectedName = comboBox1.SelectedItem?.ToString();
                string selectedSuffix = comboBox2.SelectedItem?.ToString();
                string newName = textBoxNewName.Text;
                string newFileName = selectedName  + "_" +selectedSuffix  + "_" + newName + Path.GetExtension(fileName);

                // Récupération du dossier "Documents" de l'utilisateur
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

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
