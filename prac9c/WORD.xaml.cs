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
using Spire.Doc;

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
            using (FileStream fStream = new FileStream(_fileName, FileMode.Create))
            {
                range.Save(fStream, DataFormats.Rtf);
            }
        }
        
        private void Button_Click(object sender, RoutedEventArgs e)

        {
            CommonSaveFileDialog dialog = new CommonSaveFileDialog();
            dialog.Filters.Add(new CommonFileDialogFilter("RTF Files", ".rtf"));
            dialog.DefaultFileName = "myRichTextBox";
            dialog.DefaultExtension = ".rtf";
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                string fileName = dialog.FileName;
                SaveRtbFile(fileName);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }
}
