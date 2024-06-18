using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace prac9c
{
    /// <summary>
    /// Логика взаимодействия для WORDLOAD.xaml
    /// </summary>
    public partial class WORDLOAD : Window
    {
        public WORDLOAD()
        {
            InitializeComponent();
        }
        void LoadRtfFile(string _fileName)
        {
            if (File.Exists(_fileName))
            {
                TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
                using (FileStream fStream = new FileStream(_fileName, FileMode.Open))
                {
                    range.Load(fStream, DataFormats.Rtf);
                }

            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.Filters.Add(new CommonFileDialogFilter("RTF Files", ".rtf"));
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                string fileName = dialog.FileName;
                LoadRtfFile(fileName);
            }
        }

        void SaveRtbFile(string _fileName)
        {
            TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
            FileStream fStream = new FileStream(_fileName, FileMode.Create);
            range.Save(fStream, DataFormats.Rtf);
            fStream.Close();
        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveRtbFile("C:\\Users\\User\\Desktop\\myRichTextBox.rtf");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SaveFile savefile = new SaveFile();
            savefile.Show();
        }
    }
}
