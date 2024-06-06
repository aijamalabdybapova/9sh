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
    /// Логика взаимодействия для WORD.xaml
    /// </summary>
    public partial class WORD : Window
    {
        public WORD()
        {
            InitializeComponent();
        }

        void SaveRtbFile(string _fileName)
        {
            TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
            FileStream fStream = new FileStream(_fileName, FileMode.Create);
            range.Save(fStream, DataFormats.Rtf);
            fStream.Close();
        }
        void LoadRtfFile(string _fileName)
        {
            if (File.Exists(_fileName))
            {
                TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
                FileStream fStream = new FileStream(_fileName, FileMode.OpenOrCreate);
                range.Load(fStream, DataFormats.Rtf);
                fStream.Close();
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveRtbFile("C:\\Users\\User\\Desktop\\myRichTextBox.rtf");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            LoadRtfFile("C:\\Users\\User\\Desktop\\myRichTextBox.rtf");
        }
    }
}
