using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WordExport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {

        #region Implemintation INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            var propertyChanged = PropertyChanged;
            if (propertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion //Implemintation INotifyPropertyChanged

        #region Properties
        private AppDirector _appDirector = new AppDirector();

        private string _templatePath;

        public string TemplatePath
        {
            get 
            { 
                //return _templatePath; 
                return "F:\\temp\\Болванка для фоток.dotx";
            }
            set 
            { 
                _templatePath = value;
                OnPropertyChanged("TemplatePath");
            }
        }

        private string _errorMessage;

        public string ErrorMessage
        {
            get { return _errorMessage; }
            set
            {
                _errorMessage = value;
                OnPropertyChanged("ErrorMessage");
            }
        }

        private string _folderPath;

        public string FolderPath
        {
            get 
            { 
                //return _folderPath; 
                return "F:\\temp\\foto";
            }
            set
            {
                _folderPath = value;
                OnPropertyChanged("FolderPath");
            }
        }

        private int _picWidth;

        public int PicWidth
        {
            get { return _picWidth; }
            set
            {
                _picWidth = value;
                OnPropertyChanged("PicWidth");
            }
        }
        #endregion //Properties

        #region Event handlers
        private void OnBtnChooseTemplate(object sender, RoutedEventArgs e)
        {            
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".png";
            dlg.Filter = "Word 97-2003 Documents (*.doc)|*.doc|Word 2007 Documents (*.docx)|*.docx|Word 2003 Template (*.dot)|*.dot|Word 2007 Template (*.dotx)|*.dotx";
           
            Nullable<bool> result = dlg.ShowDialog();           
            if (result == true)
            {
                TemplatePath = dlg.FileName;
            }
        }

        private void OnBtnChooseFolder(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            dlg.Description = "Выберите папку с изображениями JPG";
            dlg.ShowNewFolderButton = false;
            dlg.RootFolder = Environment.SpecialFolder.MyComputer;

            var result = dlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                FolderPath = dlg.SelectedPath;
                /*foreach (string fileName in Directory.GetFiles(folder, "*.xml", SearchOption.TopDirectoryOnly))
                {
                    
                }*/
            }
        }

        private void OnBtnExitClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void OnExportClick(object sender, RoutedEventArgs e)
        {
            StartExport();              
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            PicWidth = 550;
            ErrorMessage = String.Empty;
        }
        #endregion //Event handlers

        public MainWindow()
        {
            InitializeComponent();
            mainGrid.DataContext = this;

            _appDirector.Main = this;
        }

        private void StartExport()
        {
            try
            {
                ErrorMessage = String.Empty;
                CheckInputs();
                _appDirector.Start(PicWidth, FolderPath, TemplatePath);                
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
            }  
        }

        private void CheckInputs()
        {
            var errorMsg = String.Empty;

            if (String.IsNullOrEmpty(TemplatePath))
            {
                errorMsg += "- Не выбран шаблон документа";
            }
            else
            {
                if (!File.Exists(TemplatePath))
                {
                    errorMsg += "\n- Выбранный файл шаблона документа не сущесвует";
                }
            }

            if (String.IsNullOrEmpty(FolderPath))
            {
                errorMsg += "\n- Не выбрана папка с изображениями";
            }
            else
            {
                if (!Directory.Exists(FolderPath))
                {
                    errorMsg += "\n- Выбранной папки не существует";
                }
                else 
                {
                    
                    var dInfo = new DirectoryInfo(FolderPath);
                    var files = dInfo.GetFiles("*.jpg")
                                     .Concat(dInfo.GetFiles("*.jpeg"))
                                     .Concat(dInfo.GetFiles("*.JPG"))
                                     .Concat(dInfo.GetFiles("*.JPEG"));
                    if (files.Count<FileInfo>() == 0)
                    {
                        errorMsg += "\n- В выбранной папке нет изображений";
                    }
                }
            }

            if (!String.IsNullOrEmpty(errorMsg))
            {
                throw new Exception(errorMsg);
            }
        }
    }
}
