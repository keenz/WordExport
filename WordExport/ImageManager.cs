using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
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
            return;

            foreach (var folderPath in AppDirector.ListFolders)
            {
                var resizeFolderPath = Path.Combine(folderPath, AppDirector.ResizeFolderName);
                if (Directory.Exists(resizeFolderPath))
                {
                    continue;
                }

                Directory.CreateDirectory(resizeFolderPath);

                var dInfo = new DirectoryInfo(folderPath);
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
            
        }

        protected override void OnWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Trace.WriteLine("ImageMan - Completed");
            AppDirector.ContinueExport();
        }

        private static Image ResizeImage(Image imgToResize, Size size)
        {
            Bitmap newImage = new Bitmap(size.Width, size.Height);
            using (Graphics gr = Graphics.FromImage(newImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(imgToResize, new Rectangle(0, 0, size.Width, size.Height));
            }
            return (Image)(newImage);
        }

    }
}
