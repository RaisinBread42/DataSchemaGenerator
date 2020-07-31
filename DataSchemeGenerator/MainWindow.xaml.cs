using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DataSchemeGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public Generator Generator { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            if (Generator == null)
                Generator = new Generator();
        }

        private void UploadFileHandler(object sender, RoutedEventArgs e)
        {
            // file dialogue defaults
            OpenFileDialog fileDialogue = new OpenFileDialog();
            fileDialogue.DefaultExt = ".xlsx";
            fileDialogue.Filter = "Microsoft Excel (.xlsx)|*.xlsx";

            if (fileDialogue.ShowDialog() != null)
            {
                try
                {
                    Generator.DocPath = fileDialogue.FileName;
                    Generator.FileName = fileDialogue.FileName.Split("\\").Last();
                    Generator.Contents = new System.IO.StreamReader(fileDialogue.FileName).ReadToEnd();

                    Lbl_UploadedFileName.Content = Generator.FileName;
                }
                catch (Exception)
                {
                    MessageBox.Show("The file can not be accessed because it is either being used by another service or does not exist.");
                }
                
            }
        }

        private void SetGenerationMethod(object sender, RoutedEventArgs e)
        {
            string method = ((CheckBox)sender).Tag.ToString();
            Generator.GenerationMethods.Add(method);
        }

        private void GenerateSchema(object sender, RoutedEventArgs e)
        {
            Generator.NamespaceToUse = TxtBox_Namespace.Text;
            Generator.GenerateSchema();
        }
    }
}
