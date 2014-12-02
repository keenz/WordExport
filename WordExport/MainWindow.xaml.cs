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
                return _templatePath; 
            }
            set 
            { 
                _templatePath = value;
                OnPropertyChanged("TemplatePath");
            }
        }

        private string _documentName;

        public string DocumentName
        {
            get
            {
                return _documentName; 
            }
            set
            {
                _documentName = value;
                OnPropertyChanged("DocumentName");
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

        private int _startNum;

        public int StartNum
        {
            get { return _startNum; }
            set
            {
                _startNum = value;
                OnPropertyChanged("StartNum");
            }
        }

        private List<ChooseDirControl> _listControls = new List<ChooseDirControl>();

        #endregion //Properties

        #region Event handlers
        private void OnBtnChooseTemplate(object sender, RoutedEventArgs e)
        {            
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                        
            dlg.Filter = @"Word Templates (.dot;*.dotx)|*.dot;*.dotx";
            Nullable<bool> result = dlg.ShowDialog();           
            if (result == true)
            {
                TemplatePath = dlg.FileName;
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
            StartNum = 1;
            PicWidth = 240;
            ErrorMessage = String.Empty;
            TemplatePath = "F:\\temp\\Болванка для фоток.dotx";
            DocumentName = "Document.docx";

            AddControl();
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
                var list = new List<string>();
                foreach (var item in _listControls)
                {
                    list.Add(item.FolderPath);
                }
                _appDirector.Start(StartNum, PicWidth, list, TemplatePath, DocumentName);                
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

            int filesCount = 0;
            foreach (ChooseDirControl control in _listControls)
            {
                if (String.IsNullOrEmpty(control.FolderPath))
                {
                    continue;
                }
                else
                {
                    if (!Directory.Exists(control.FolderPath))
                    {
                        errorMsg += "\n- Выбранной папки не существует (" + control.TextBlockLabel + ")";
                    }
                    else
                    {

                        var dInfo = new DirectoryInfo(control.FolderPath);
                        var files = dInfo.GetFiles("*.jpg")
                                         .Concat(dInfo.GetFiles("*.jpeg"))
                                         .Concat(dInfo.GetFiles("*.JPG"))
                                         .Concat(dInfo.GetFiles("*.JPEG"));
                        if (files.Count<FileInfo>() == 0)
                        {
                            errorMsg += "\n- В выбранной папке нет изображений (" + control.TextBlockLabel + ")";
                        }
                        else
                        {
                            filesCount += files.Count<FileInfo>();
                        }
                    }
                }
            }

            if (filesCount == 0)
            {
                errorMsg += "\n- В выбранных папках нет изображений";
            }

            if (!String.IsNullOrEmpty(errorMsg))
            {
                throw new Exception(errorMsg);
            }
        }

        private void OnAddBtnClick(object sender, RoutedEventArgs e)
        {
            AddControl();
        }

        private void AddControl()
        {
            var control = new ChooseDirControl();
            control.TextBlockLabel = String.Format("Папка №{0}", _listControls.Count + 1);
            panel.Children.Add(control);
            _listControls.Add(control);
            scroll.ScrollToBottom();
        }
    }
}
