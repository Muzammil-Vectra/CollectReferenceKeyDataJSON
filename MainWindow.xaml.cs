using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System;

namespace CollectReferenceKeyDataJSON
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            Main = this;
            InitializeComponent();
        }

        internal static MainWindow Main;
        internal double UpdateProgressBar
        {
            set {Dispatcher.Invoke(new Action(() => { PrgBar.Value = value; }));}
        }
        internal string UpdateProgressBarStatus
        {
            set { Dispatcher.Invoke(new Action(() => { LblStatus.Content = value; })); }
        }
        private async void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            BtnGenerate.IsEnabled = false;
            InventorInteraction inventorInteraction = new InventorInteraction();
            await Task.Run(() => inventorInteraction.CollectReferenceKeyData());         
            BtnGenerate.IsEnabled = true;
            LblStatus.Content = "0.0%";
            PrgBar.Value = 0;
        }
    }
}
