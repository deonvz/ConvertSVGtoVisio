using Microsoft.Win32;
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
using System.Windows.Navigation;
using System.Windows.Shapes;



namespace ConvertSVGtoVisio
{
    public partial class SvgtoVsdx : Window
    {
        public SvgtoVsdx()
        {
            InitializeComponent();
        }

        private string[] path;
        private int path_length = 0;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (path_length <= 0)
                MessageBox.Show("Select the files！");
            else
            {
                progressBar1.Maximum = path_length;
                progressBar1.Value = 0;
                for (int i = 0; i < path_length; i++)
                {
                    int b = path[i].LastIndexOf('.');
                    string des = path[i].Substring(0, b);
                    des += ".vsdx";
                    try
                    {
                        ComObj.ConvertSvgtoVsdx(path[i], des);
                    }
                    catch
                    {
                        MessageBox.Show("Is visio installed?" + "Failed:" + path[i], "Failed to Convert this SVG");
                    }
                    progressBar1.Value += 1;      
                    runprocess.Content ="(" + progressBar1.Value.ToString() + "/ " + path_length.ToString() + ")";
                }
                path_length = 0;
                MessageBox.Show("Completed！");
            }
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Files (*.svg)|*.svg"
            };
            openFileDialog.Multiselect = true;
            string text = string.Empty;
            if (openFileDialog.ShowDialog() == true)
            {
                path = openFileDialog.FileNames;
                path_length = path.Length;
                runprocess.Content = "(0/" + path_length.ToString() + ")";
                for (int i = 0; i < path_length; i++)
                {
                    if (i == 0)
                        text = path[0];
                    else
                        text += (", " + path[i]);
                }
                textBox1.Text = text;
            }
        }
    }
}
