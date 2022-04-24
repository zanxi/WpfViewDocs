using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        public BackgroundWorker worker;
        public MainWindow()
        {
            InitializeComponent();

            worker = new BackgroundWorker();
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);

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

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(e.Cancelled)
            {
                worker.RunWorkerAsync(folderTree.Selection);
            }
            status.Dispatcher.Invoke(new Action(delegate()
            {
                status.Text = string.Empty;
            }
            ));
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            FolderTreeSelection selection = (FolderTreeSelection)e.Argument;
            foreach(var folder in selection.SelectedFolders)
            {
                status.Dispatcher.Invoke(new Action(delegate ()
                {
                    status.Text = folder.FullName;
                }
                ));
                

                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }
            }

            
        }

        private void folderTree_SelectionChanged(object sender, EventArgs e)
        {
            if(worker.IsBusy)
            {
                worker.CancelAsync();
            }
            else
            {
                worker.RunWorkerAsync();
            }

        }
    }
}
