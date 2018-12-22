using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Nthrive.TestAutomation.TestExecutor
{
    /// <summary>
    /// Interaction logic for SplashScreen.xaml
    /// </summary>
    public partial class SplashScreen : Window,IDisposable
    {

        public SplashScreen()
        {
            InitializeComponent();
            DataContext = new SplashWork();
        }

        public void Dispose()
        {
        }
    }


    public class SplashWork : BindableBase
    {
        private static int _work;
        public int Progress
        {
            get { return _work; }
            set
            {
                SetProperty(ref _work, value);
                OnPropertyChanged("Progress");
            }
        }
        public SplashWork()
        {
            while(Progress!= 100)
            {
                UpdateBar();
                Thread.Sleep(1000);
            }
            
        }
        public void UpdateBar()
        {
            Progress=Progress+10;
        }
    }
}