using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace WordExport
{
    /// <summary>
    /// Class to work with images
    /// </summary>
    public class ImageManager : ManagerBase
    {
        /// <summary>
        /// Designer
        /// </summary>
        public ImageManager(AppDirector appDirector)
            :base (appDirector)
        { 
        }

        protected override void OnWorkerDoWork(object sender, DoWorkEventArgs e)
        {
            Trace.WriteLine("ImageMan - Do work");

            var resizeFolderPath = Path.Combine(AppDirector.FolderPath, AppDirector.ResizeFolderName);
            if (!Directory.Exists(resizeFolderPath))
            {
                Directory.CreateDirectory(resizeFolderPath);
            }

            var dInfo = new DirectoryInfo(AppDirector.FolderPath);            
            var files = dInfo.GetFiles("*.jpg")
                             .Concat(dInfo.GetFiles("*.jpeg"))
                             .Concat(dInfo.GetFiles("*.JPG"))
                             .Concat(dInfo.GetFiles("*.JPEG"));
            foreach (var fileInfo in files)
            {
                if (NeedStop)
                {
                    break;
                }

                var srcImage = Image.FromFile(fileInfo.FullName);
                var srcSize = srcImage.Size;
                double scale = Convert.ToDouble(srcSize.Height) / Convert.ToDouble(srcSize.Width);
                int width = AppDirector.Width;
                int height = Convert.ToInt32(scale * width);

                var image = ResizeImage(srcImage, new Size(width, height));
                var newFilePath = Path.Combine(resizeFolderPath, Path.GetFileName(fileInfo.FullName));
                image.Save(newFilePath);
            }
            
        }

        protected override void OnWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Trace.WriteLine("ImageMan - Completed");
            AppDirector.ContinueExport();
        }

        private static Image ResizeImage(Image imgToResize, Size size)
        {
            return (Image)(new Bitmap(imgToResize, size));
        }

    }
}
