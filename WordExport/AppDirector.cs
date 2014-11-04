using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;

namespace WordExport
{
    public class AppDirector
    {
        public static string ResizeFolderName = "Resize";

        public string ResizedFolderPath
        {
            get { return Path.Combine(FolderPath, ResizeFolderName); }
        }

        public string FolderPath { get; private set; }

        public string TemplatePath { get; private set; }

        public int Width { get; private set; }

        private ProgressWnd _progressWnd;

        private ImageManager _imageMan;

        private ExportManager _exportMan;

        public MainWindow Main { get; set; }

        /// <summary>
        /// Designer
        /// </summary>
        public AppDirector() 
        {            
        }

        /// <summary>
        /// Start
        /// </summary>
        public void Start(int width, string folderPath, string templatePath)
        {
            Width = width;
            FolderPath = folderPath;
            TemplatePath = templatePath;

            _progressWnd = new ProgressWnd();
            _progressWnd.AppDirector = this;
            _progressWnd.Owner = Main;
            _imageMan = new ImageManager(this);
            _imageMan.Run();
            _progressWnd.ShowDialog();
            
        }

        public void ContinueExport()
        {
            _exportMan = new ExportManager(this);
            _exportMan.Run();
        }

        /// <summary>
        /// Stop
        /// </summary>
        public void SendStopCommand()
        {
            if (_imageMan.IsBusy)
            {
                _imageMan.Stop();
            }

            if (_exportMan.IsBusy)
            {
                _exportMan.Stop();
            }
        }

        public void Stoped()
        {
            _progressWnd.Close();
        }
    }
}
