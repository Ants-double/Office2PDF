using Microsoft.Office.Interop.Excel;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Office2PDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private String fileFolderPath = System.Environment.CurrentDirectory;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog("请选择一个文件夹");
            dialog.IsFolderPicker = true; //选择文件还是文件夹（true:选择文件夹，false:选择文件）
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string path = dialog.FileName;
                this.filePathText.Text = path;
                fileFolderPath = path;
              //  MessageBox.Show($"当前所选文件夹路径为：{path}");
            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

            ConvertToPdf(fileFolderPath, "");


        }

        public  void ConvertToPdf(string Path, string savePath)
        {
            DirectoryInfo foler = new DirectoryInfo(fileFolderPath);
            string[] files = Directory.GetFiles(fileFolderPath, "*.doc*");
            foreach (string file in files)
            {
                
                string sourcePath = file;
                
              
                string _file = file + @".pdf";
              
                if (Directory.Exists(_file)) continue;
               OperationOffice.word2pdf(file, _file);
            }

            string[] pptfiles = Directory.GetFiles(fileFolderPath, "*.ppt*");
            foreach (string file in pptfiles)
            {

                string sourcePath = file;


                string _file = file + @".pdf";

                if (Directory.Exists(_file)) continue;
                OperationOffice.ppt2pdf(file, _file);
            }
            string[] xlsfiles = Directory.GetFiles(fileFolderPath, "*.xls*");
            foreach (string file in xlsfiles)
            {

                string sourcePath = file;


                string _file = file + @".pdf";

                if (Directory.Exists(_file)) continue;
                OperationOffice.ExportWorkbookToPdf(file, _file);
            }
            MessageBox.Show("转换完成，请在同一目录下查找");
        }
    }
}
