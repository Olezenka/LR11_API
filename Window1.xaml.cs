using LR11.Classes;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace LR11
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        private Classes.APIInteraction apiinteraction;
        public Window1()
        {
            InitializeComponent();
            apiinteraction = new Classes.APIInteraction();
        }

        private void ButtonGetFullName_Click(object sender, RoutedEventArgs e)
        {
            TextBlockFullName.Text = apiinteraction.GetFullName();
        }

        private void ButtonSendResult_Click(object sender, RoutedEventArgs e)
        {
            TextBlockResult.Text = apiinteraction.FillDocument();
        }
    }
}
