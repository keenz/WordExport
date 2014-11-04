using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WordExport
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class ProgressWnd : Window
    {
        public AppDirector AppDirector { get; set; }

        public ProgressWnd()
        {
            InitializeComponent();
        }

        private void OnBtnCancelClick(object sender, RoutedEventArgs e)
        {
            if (AppDirector == null)
            {
                Trace.WriteLine("AppDirector is null");
                return;
            }

            AppDirector.SendStopCommand();
        }
    }
}
