using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace WordExport
{
    public abstract class ManagerBase
    {
        protected volatile bool NeedStop = false;

        protected AppDirector AppDirector { get; private set; }

        private BackgroundWorker _worker = new BackgroundWorker();

        public bool IsBusy
        {
            get { return _worker.IsBusy; }
        }

        public ManagerBase(AppDirector appDirector)
        {
            AppDirector = appDirector;

            _worker.DoWork += OnWorkerDoWork;
            _worker.ProgressChanged += OnWorkerProgressChanged;
            _worker.RunWorkerCompleted += OnWorkerRunWorkerCompleted;
        }

        protected abstract void OnWorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e);
        
        private void OnWorkerProgressChanged(object sender, ProgressChangedEventArgs e)
        {
        }

        protected abstract void OnWorkerDoWork(object sender, DoWorkEventArgs e);

        public void Stop()
        {
            NeedStop = true;
        }

        public void Run()
        {
            if (!_worker.IsBusy)
            {
                _worker.RunWorkerAsync();
            }
        }



    }
}
