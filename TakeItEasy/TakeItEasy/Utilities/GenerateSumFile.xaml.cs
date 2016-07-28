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
using MahApps.Metro.Controls;
using Microsoft.Win32;
using TakeItEasy.CommonDialog;
using TakeItEasy.Const;

namespace TakeItEasy.Utilities
{
    /// <summary>
    /// Interaction logic for GenerateFile.xaml
    /// </summary>
    public partial class GenerateFile : MetroWindow
    {
        public GenerateFile()
        {
            InitializeComponent();
        }

        private void btn_OK_Click(object sender, RoutedEventArgs e)
        {
            if (tbx_FPath.Text == "" || tbx_DirPath.Text == "")
            {
                //Xceed.Wpf.Toolkit.MessageBox.Show("Please choose input folder and output file path!", 
                //    Consts.WARN, MessageBoxButton.OK, MessageBoxImage.Warning);
                NoticeDialog dialog = new NoticeDialog(Consts.MSG_WARN, Consts.NEED_ENTER_PATH, 
                    "OK", Application.Current.MainWindow, DialogIcons.WARNING);
                dialog.ShowDialog();
                return;
            }

            if (!checkFolderPathValid(tbx_DirPath.Text))
            {
                //Xceed.Wpf.Toolkit.MessageBox.Show("Inputed Folder path does not exist",
                //    Consts.WARN, MessageBoxButton.OK, MessageBoxImage.Warning);
                NoticeDialog dialog = new NoticeDialog(Consts.MSG_WARN, Consts.PATH_NOT_EXIST,
                    "OK", Application.Current.MainWindow, DialogIcons.WARNING);
                dialog.ShowDialog();
                return;
            }
            // transfer GUI to variable
            string dutyName = tbx_Duty.Text == null ? "" : tbx_Duty.Text;
            string projectName = tbx_PJName.Text == null ? "" : tbx_PJName.Text;
            string updateType = tbx_UpdateType.Text == null ? "" : tbx_UpdateType.Text;
            string author = tbx_Author.Text == null ? "" : tbx_Author.Text;
            string date = dpk_Date.DisplayDate == null ?
                DateTime.Now.ToShortDateString() : dpk_Date.DisplayDate.ToShortDateString();
            string description = tbx_Description.Text == null ? "" : tbx_Description.Text;
            string remark = tbx_Remark.Text == null ? "" : tbx_Remark.Text;
            string confirmStt = tbx_ConfirmStt.Text == null ? "" : tbx_ConfirmStt.Text;
            GenerateFileList genFile = new GenerateFileList(dutyName, projectName, updateType,
                author, date, description, remark, confirmStt);
            // generate file
            if (tbx_FPath.Text != null)
                genFile.StartGenerate(tbx_FPath.Text, tbx_DirPath.Text);

            NoticeDialog dlg = new NoticeDialog(Consts.MSG_INFO, Consts.GEN_FILE_SUCCESS,
                    "OK", Application.Current.MainWindow, DialogIcons.INFO);
            dlg.ShowDialog();
            //if (ret == MessageBoxResult.No)
            //{
            //    this.Close();
            //}
        }

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btn_Save_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = Consts.EXT_FILTER;
            if (saveFileDialog.ShowDialog() == true)
                tbx_FPath.Text = saveFileDialog.FileName;
        }
        // Folder Selection dialog
        private void btn_DirChoose_Click(object sender, RoutedEventArgs e)
        {
            //OpenFileOrFolderDialog folderDialog = new OpenFileOrFolderDialog();
            //folderDialog.AcceptFiles = false;
            //folderDialog.Path = @"C:\";
            //DialogResult ret = folderDialog.ShowDialog();
            //if (ret == System.Windows.Forms.DialogResult.OK)
            //    tbx_DirPath.Text = folderDialog.FileNameLabel;
            string[] lst = tbx_DirPath.Text.Split('\\');
            if (lst.Length >= 2)
                tbx_Duty.Text = lst[lst.Length - 2];
        }

        #region CheckPathSyntax
        private bool checkFolderPathValid(string folderPath)
        {
            if (System.IO.Directory.Exists(folderPath))
            {
                return true;
            }
            return false;
        }
        #endregion
    }
}
