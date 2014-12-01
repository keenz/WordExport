using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    /// Interaction logic for ChooseDirControl.xaml
    /// </summary>
    public partial class ChooseDirControl : UserControl, INotifyPropertyChanged
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

        private string _folderPath;

        public string FolderPath
        {
            get
            {
                return _folderPath;
            }
            set
            {
                _folderPath = value;
                OnPropertyChanged("FolderPath");
            }
        }

        private string _textBlockLabel;

        public string TextBlockLabel
        {
            get
            {
                return _textBlockLabel;
            }
            set
            {
                _textBlockLabel = value;
                OnPropertyChanged("TextBlockLabel");
            }
        }

        public ChooseDirControl()
        {
            InitializeComponent();
            myGrid.DataContext = this; 
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
    }
}
