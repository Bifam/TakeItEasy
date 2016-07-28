using System;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using TakeItEasy.CommonDialog;
using System.Drawing;
using TakeItEasy.Const;

namespace TakeItEasy.Utilities
{
    class GenerateFileList
    {
        private string dutyName    = "";
        private string projectName = "GUIServer";
        private string updateType  = "新規";
        private string author      = "";
        private string date        = "";
        private string description = "";
        private string remark      = "";
        private string confirmStt  = "●";
        private string dirPath     = "";

        public GenerateFileList(string DutyName, string PjName, string UpdateType, string Author,
            string Date, string Description, string Remark, string ConfirmStt)
        {
            dutyName = DutyName;
            projectName = PjName;
            updateType = UpdateType;
            author = Author;
            date = Date;
            description = Description;
            remark = Remark;
            confirmStt = ConfirmStt;
        }

        public void StartGenerate(string FPath, string DirPath)
        {
            string fExt = Path.GetExtension(FPath);

            if (fExt == ".xls" || fExt == ".xlsx" || fExt == ".xlsm")
            {   
                // Excel file format
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    NoticeDialog dialog = new NoticeDialog(Consts.MSG_ERR, "Excel is not properly installed!",
                    "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                    dialog.ShowDialog();
                    return;
                }

                xlApp.DisplayAlerts = false;
                object misValue = System.Reflection.Missing.Value;
                var xlWorkBook = xlApp.Workbooks.Add(misValue);
                var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
                //format excel
                xlWorkSheet.Cells.Font.Name = "ＭＳ Ｐゴシック";
                xlWorkSheet.Cells.Font.Size = 11;
                //header
                xlWorkSheet.Range["A1", "K1"].Cells.Interior.Color = Color.LightCyan;
                xlWorkSheet.Range["A1", "K1"].Cells.Font.Bold = true;
                xlWorkSheet.Range["A1", "K1"].Cells.HorizontalAlignment
                = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[1, 1] = "No";
                xlWorkSheet.Cells[1, 2] = "業務名称";
                xlWorkSheet.Cells[1, 3] = "プロジェクト名";
                xlWorkSheet.Cells[1, 4] = "パッケージ名";
                xlWorkSheet.Cells[1, 5] = "クラス名";
                xlWorkSheet.Cells[1, 6] = "変更区分";
                xlWorkSheet.Cells[1, 7] = "提出日";
                xlWorkSheet.Cells[1, 8] = "担当者";
                xlWorkSheet.Cells[1, 9] = "説明";
                xlWorkSheet.Cells[1, 10] = "備考";
                xlWorkSheet.Cells[1, 11] = "確認済み";
                //freeze panel
                xlWorkSheet.Activate();
                xlWorkSheet.Application.ActiveWindow.SplitRow = 1;
                xlWorkSheet.Application.ActiveWindow.FreezePanes = true;

                dirPath = DirPath;
                int index = 1;

                ProcessDirectory(DirPath, ref xlWorkSheet, ref index);
                xlWorkSheet.Columns.AutoFit();
                xlWorkSheet.Range["A1", "K1"].AutoFilter(1,
                    XlAutoFilterOperator.xlAnd, XlAutoFilterOperator.xlFilterNoFill, true);

                try
                {
                    xlWorkBook.SaveAs(FPath);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
                catch (Exception ex)
                {
                    NoticeDialog dialog = new NoticeDialog(Consts.MSG_ERR, ex.Message,
                    "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                    dialog.ShowDialog();
                    //Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, Consts.MSG_ERR,
                    //    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                // other file format
                try
                {
                    System.IO.StreamWriter file = new System.IO.StreamWriter(FPath, false, Encoding.UTF8);
                    dirPath = DirPath;
                    int index = 1;
                    if (file != null)
                    {
                        file.WriteLine("No,業務名称,プロジェクト名,パッケージ名,クラス名,変更区分,提出日," +
                            "担当者,説明,備考,確認済み");
                        ProcessDirectory(DirPath, ref file, ref index);
                        file.Close();
                    }
                }
                catch (Exception ex)
                {
                    NoticeDialog dlg = new NoticeDialog(Consts.MSG_ERR, ex.Message,
                    "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                    dlg.ShowDialog();
                    //Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, Consts.MSG_ERR,
                    //    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        #region Generate file
        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        private void ProcessDirectory(string targetDirectory, ref StreamWriter writter, ref int index)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName, ref writter, ref index);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory, ref writter, ref index);
        }
        // Excel format
        private void ProcessDirectory(string targetDirectory, ref Worksheet writter, ref int index)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName, ref writter, ref index);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory, ref writter, ref index);
        }

        // Insert logic for processing found files here.
        private void ProcessFile(string path, ref StreamWriter writter, ref int index)
        {
            string fName = Path.GetFileName(path);
            string fDirPath = Path.GetDirectoryName(path);
            var fName_s = fName;
            var fDirPath1 = fDirPath.Replace(dirPath + "\\", "");
            var fDirPath1_1 = fDirPath1.Replace(dirPath, "");
            var fDirPath_s = fDirPath1_1;

            if (fName.Contains(".java"))
            {
                fName_s = fName.Replace(".java", "");
                var fDirPath2 = fDirPath1.Replace("GUIServer\\WebContent\\", "");
                var fDirPath3 = fDirPath2.Replace("GUIServer\\src\\", "");
                fDirPath_s = fDirPath3.Replace("\\", ".");
            }
            if(fName.Contains(".xml"))
            {
                fDirPath_s = fDirPath1.Replace("GUIServer\\", "");
            }
            if (fName.Contains(".dicon"))
            {
                fDirPath_s = fDirPath1.Replace("GUIServer\\", "");
            }
            if (fName.Contains(".sql"))
            {
                fDirPath_s = fDirPath1.Replace("GUIServer\\", "");
            }
            if ((fDirPath == dirPath) || (fDirPath + "\\" == dirPath))
                fDirPath = dutyName + "\\01_ソース";
            writter.WriteLine(String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10}",
                index, this.dutyName, this.projectName, fDirPath_s, fName_s,
                this.updateType, this.date, this.author,
                this.description, this.remark, this.confirmStt));
            index++;
        }
        //Excel format
        // Insert logic for processing found files here.
        private void ProcessFile(string path, ref Worksheet writter, ref int index)
        {
            string fName = Path.GetFileName(path);
            string fDirPath = Path.GetDirectoryName(path);
            var fName_s = fName;
            var fDirPath1 = fDirPath.Replace(dirPath + "\\", "");
            var fDirPath1_1 = fDirPath1.Replace(dirPath, "");
            var fDirPath_s = fDirPath1_1;

            if (fName.Contains(".java"))
            {
                fName_s = fName.Replace(".java", "");
                var fDirPath2 = fDirPath1.Replace("GUIServer\\WebContent\\", "");
                var fDirPath3 = fDirPath2.Replace("GUIServer\\src\\", "");
                fDirPath_s = fDirPath3.Replace("\\", ".");
            }
            if (fName.Contains(".xml"))
            {
                fDirPath_s = fDirPath1.Replace("GUIServer\\", "");
            }
            if (fName.Contains(".dicon"))
            {
                fDirPath_s = fDirPath1.Replace("GUIServer\\", "");
            }
            if (fName.Contains(".sql"))
            {
                fDirPath_s = fDirPath1.Replace("GUIServer\\", "");
            }
            if ((fDirPath == dirPath) || (fDirPath + "\\" == dirPath))
                fDirPath_s = dutyName + "\\01_ソース";

            writter.Cells[index + 1, 1] = index;
            writter.Cells[index + 1, 2] = this.dutyName;
            writter.Cells[index + 1, 3] = this.projectName;
            writter.Cells[index + 1, 4] = fDirPath_s;
            writter.Cells[index + 1, 5] = fName_s;
            writter.Cells[index + 1, 6] = this.updateType;
            writter.Cells[index + 1, 7] = this.date;
            writter.Cells[index + 1, 8] = this.author;
            writter.Cells[index + 1, 9] = this.description;
            writter.Cells[index + 1, 10] = this.remark;
            writter.Cells[index + 1, 11] = this.confirmStt;
            index++;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, Consts.MSG_ERR,
                //        MessageBoxButton.OK, MessageBoxImage.Error);
                NoticeDialog dlg = new NoticeDialog(Consts.MSG_ERR, ex.Message,
                    "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                dlg.ShowDialog();
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion
    }
}
