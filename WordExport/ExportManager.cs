using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace WordExport
{
    public class ExportManager : ManagerBase
    {               

        public ExportManager(AppDirector appDirector)
            : base(appDirector)
        {   
        }
        
        protected override void OnWorkerDoWork(object sender, DoWorkEventArgs e)
        { 	        
            Trace.WriteLine("ExportMan - Do work");

            WordDocument wordDoc = null;
            try
            {
                wordDoc = new WordDocument(AppDirector.TemplatePath);

                for (int j = 0; j < 14; j++)
                {
                    wordDoc.InsertParagraph();
                }
                wordDoc.InsertText("Приложение.", true,
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter);
                wordDoc.InsertBreak();

                var dInfo = new DirectoryInfo(AppDirector.ResizedFolderPath);
                /*var files = dInfo.GetFiles("*.jpg")
                                 .Concat(dInfo.GetFiles("*.jpeg"))
                                 .Concat(dInfo.GetFiles("*.JPG"))
                                 .Concat(dInfo.GetFiles("*.JPEG"));
                 */
                var files = dInfo.GetFiles("*.JPG");

                int i = 1;
                foreach (var fileInfo in files)
                {
                    if (NeedStop)
                    {
                        break;
                    }

                    var text = String.Format("Фото {0}.   ", i);
                    
                    if ((i % 2) == 0) // bottom photo
                    {
                        wordDoc.InsertFotoText(text);
                        wordDoc.InsertImage(fileInfo.FullName);
                        if (files.Count<FileInfo>() > i)
                        {
                            wordDoc.InsertBreak();
                        }
                    }
                    else // top photo
                    {
                        wordDoc.InsertImage(fileInfo.FullName);
                        wordDoc.InsertFotoText(text);
                    }
                    Trace.WriteLine(String.Format("file: {0} count:{1}", fileInfo.FullName, i.ToString()));
                    i++;
                }
            }
            catch (Exception error)
            {
                if (wordDoc != null) 
                { 
                    wordDoc.Close(); 
                }                
                return;
            }

            if (NeedStop)
            {
                return;
            }

            //wordDoc.Visible = true;
            try
            {
                wordDoc.Save();                
                //wordDoc.SaveAs(AppDirector.DocumentPath);
                //wordDoc.Close();

                //Process.Start(AppDirector.DocumentPath);
                //Process.Start(AppDirector.TemplatePath);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
                wordDoc.Close();
            }
            
        }

        protected override void OnWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {            
            Trace.WriteLine("ExportMan - Completed");
            AppDirector.Stoped();
        }
        
    }
}
