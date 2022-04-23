using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using ViewAppDocs.ViewModel;
using Path = System.IO.Path;

namespace ViewAppDocs
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Loaded += Window_Loaded;
            string doc = "Rendering Test.docx";
            this.ReadDocx(doc);
        }

        private void MenuItem_Click_OnOpenFile(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog()
            {
                DefaultExt = ".docx",
                Filter = "Word documents (.docx)|*.docx"
            };
            if (openFileDialog.ShowDialog() == true) this.ReadDocx(openFileDialog.FileName);

        }

        private void MenuItem_Click_OnAbout(object sender, RoutedEventArgs e)
        {
            new Window1().ShowDialog();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DataContext = new PersonViewModel();
        }

        private void ReadDocx(string path)
        {
            if (!File.Exists(path)) return;
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read,FileShare.ReadWrite))
            {
                var flowDocConv = new DocxToFlowDocumentConverter(stream);
                flowDocConv.Read();
                this.flowDocumentReader.Document = flowDocConv.Document;
                this.Title = Path.GetFileName(path);
            }
        }

        private void folderTree_SelectionChanged(object sender, EventArgs e)
        {

        }
    }
}
