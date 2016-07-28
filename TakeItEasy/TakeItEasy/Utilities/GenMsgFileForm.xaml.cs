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
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using TakeItEasy.DatabaseSrc;
using System.Data;
using TakeItEasy.CommonDialog;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using TakeItEasy.Const;

namespace TakeItEasy.Utilities
{
    enum GenType
    {
        MSG_TYPE,
        HEAD_TYPE
    }
    /// <summary>
    /// Interaction logic for GenMsgFile.xaml
    /// </summary>
    public partial class GenMsgFile : MetroWindow
    {
        public GenMsgFile()
        {
            InitializeComponent();

            cb_Type.Items.Add("Duties Message");
            cb_Type.Items.Add("Head Infomation");
            cb_Type.SelectedIndex = 0;
        }

        private void btn_FPath_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = Consts.EXT_FILTER;
            if (saveFileDialog.ShowDialog() == true)
                tb_FPath.Text = saveFileDialog.FileName;
        }

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btn_HeadGen_Click(object sender, RoutedEventArgs e)
        {
            string originalID = tb_Org.Text;
            string newID = tb_New.Text;
            string query = "";

            if (tb_FPath.Text == "" || tb_New.Text == "" || tb_Org.Text == "")
            {
                NoticeDialog dialog = new NoticeDialog("Error", Consts.FIELD_NEED_FILLED,
                "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                dialog.ShowDialog();
                return;
            }

            if (cb_Type.SelectedIndex == 0)
                query = "SELECT * FROM DB_ROCKY_B_001.dbo.TB_DUTIES_MSG_MST WHERE"
                        + " DUTIES_ID='" + tb_Org.Text + "'";
            else
                query = "SELECT * FROM DB_ROCKY_B_001.dbo.TB_RPT_HEAD_INFO WHERE"
                        + " RPT_ID LIKE '%" + tb_Org.Text + "%'";

            System.Data.DataTable lst = GetDBAction.GetTableDetail(query);
            if (lst == null)
                return;
            if (cb_Type.SelectedIndex == 0)
                GenerateDutiesMsg(lst, tb_FPath.Text);
            else
                GenerateHeadInfo(lst, tb_FPath.Text);
        }

        private void GenerateHeadInfo(System.Data.DataTable lst, string path)
        {
            Microsoft.Office.Interop.Excel.Application xlApp 
                = new Microsoft.Office.Interop.Excel.Application();
            // Check if excel is installed
            if (xlApp == null)
            {
                NoticeDialog dialog = new NoticeDialog("Error", Consts.EXCEL_NOT_INSTALL,
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
            xlWorkSheet.Range["A1"].Cells.Font.Size = 14;
            xlWorkSheet.Range["A1"].Cells.Font.Bold = true;
            xlWorkSheet.Range["A1"].Cells.Font.Italic = true;
            xlWorkSheet.Cells[1, 1] = "帳票DB設定";

            //table
            xlWorkSheet.Range["A3"].Cells.Font.Bold = true;
            xlWorkSheet.Cells[3, 1] = "テープル：TB_RPT_HEAD_INFO";
            //table header
            xlWorkSheet.Range["A4", "D4"].Cells.Interior.Color = Color.LightCyan;
            xlWorkSheet.Range["A4", "D4"].Cells.Font.Bold = true;
            xlWorkSheet.Range["A4", "D4"].Cells.Font.Italic = true;
            xlWorkSheet.Range["A4", "D4"].Cells.HorizontalAlignment
                = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[4, 1] = Consts.RPT_ID;
            xlWorkSheet.Cells[4, 2] = Consts.HEAD_KIND;
            xlWorkSheet.Cells[4, 3] = Consts.RPT_HEAD_SEQ_NO;
            xlWorkSheet.Cells[4, 4] = Consts.RPT_HEAD_VAL;
            //table data
            GetHeadData(xlWorkSheet, lst);
            try
            {
                xlWorkBook.SaveAs(path);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                NoticeDialog dlg = new NoticeDialog(Consts.MSG_INFO, Consts.GEN_FILE_SUCCESS,
                "OK", System.Windows.Application.Current.MainWindow, DialogIcons.INFO);
                dlg.ShowDialog();
            }
            catch (Exception ex)
            {
                NoticeDialog dialog = new NoticeDialog("Error", ex.Message,
                "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                dialog.ShowDialog();
                //Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, "Error",
                //    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateDutiesMsg(System.Data.DataTable lst, string path)
        {
            Microsoft.Office.Interop.Excel.Application xlApp
                = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                NoticeDialog dialog = new NoticeDialog("Error", Consts.EXCEL_NOT_INSTALL,
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
            //table
            xlWorkSheet.Cells[2, 1] = Consts.TB_DUTIES_MSG_MST;
            //table header
            xlWorkSheet.Range["A3", "C3"].Cells.Interior.Color = Color.LightCyan;
            xlWorkSheet.Range["A3", "C3"].Cells.Font.Bold = true;
            xlWorkSheet.Range["A3", "C3"].Cells.HorizontalAlignment
                = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[3, 1] = Consts.DUTIES_ID;
            xlWorkSheet.Cells[3, 2] = Consts.DUTIES_MSG_ID;
            xlWorkSheet.Cells[3, 3] = Consts.DUTIES_MSG;
            //table data
            GetMsgData(xlWorkSheet, lst);
            try
            {
                xlWorkBook.SaveAs(path);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                NoticeDialog dlg = new NoticeDialog(Consts.MSG_INFO, Consts.GEN_FILE_SUCCESS,
                "OK", System.Windows.Application.Current.MainWindow, DialogIcons.INFO);
                dlg.ShowDialog();
            }
            catch (Exception ex)
            {
                NoticeDialog dialog = new NoticeDialog("Error", ex.Message,
                "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                dialog.ShowDialog();
                //Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, "Error",
                //    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        // Get data from table TB_RPT_HEAD_INFO
        private void GetHeadData(Worksheet xlWorkSheet, System.Data.DataTable lst)
        {
            int index = 5;
            AllBorders(xlWorkSheet.Range[String.Format("A" + (index - 1)), 
                String.Format("D" + (index + lst.Rows.Count - 1))].Cells.Borders);
            xlWorkSheet.Range[String.Format("A" + (index - 2)), 
                String.Format("D" + (index + lst.Rows.Count - 1))].Columns.AutoFit();
            try {
                foreach (DataRow row in lst.Rows)
                {
                    //rpt_id
                    xlWorkSheet.Cells[index, 1] 
                        = row["RPT_ID"].ToString().Replace(tb_Org.Text, tb_New.Text);
                    //head_kind
                    xlWorkSheet.Cells[index, 2] = row[Consts.HEAD_KIND];
                    //rpt_head_sequence_no
                    xlWorkSheet.Cells[index, 3] = row[Consts.RPT_HEAD_SEQ_NO];
                    //rpt_head_val
                    xlWorkSheet.Cells[index, 4] = row[Consts.RPT_HEAD_VAL];
                    index++;
                }
            }
            catch (Exception ex)
            {
                NoticeDialog dialog = new NoticeDialog("Error", ex.Message,
                "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                dialog.ShowDialog();
                return;
            }
        }
        // get data from table TB_DUTIES_MSG_MST
        private void GetMsgData(Worksheet xlWorkSheet, System.Data.DataTable lst)
        {
            int index = 4;
            AllBorders(xlWorkSheet.Range[String.Format("A" + (index - 1)),
                String.Format("C" + (index + lst.Rows.Count - 1))].Cells.Borders);
            xlWorkSheet.Range[String.Format("A" + (index - 2)),
                String.Format("C" + (index + lst.Rows.Count - 1))].Columns.AutoFit();
            try
            {
                foreach (DataRow row in lst.Rows)
                {
                    //duties_id
                    xlWorkSheet.Cells[index, 1]
                        = row["DUTIES_ID"].ToString().Replace(tb_Org.Text, tb_New.Text);
                    //head_kind
                    xlWorkSheet.Cells[index, 2] = row[Consts.DUTIES_MSG_ID];
                    //rpt_head_sequence_no
                    xlWorkSheet.Cells[index, 3] = row[Consts.DUTIES_MSG];
                    index++;
                }
            }
            catch (Exception ex)
            {
                NoticeDialog dialog = new NoticeDialog("Error", ex.Message,
                "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                dialog.ShowDialog();
                return;
            }
        }
        // Draw border for generated file
        private void AllBorders(Microsoft.Office.Interop.Excel.Borders _borders)
        {
            _borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle
                = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle
                = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle
                = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle
                = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            _borders.Color = Color.Black;
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
                //Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, "Error",
                //        MessageBoxButton.OK, MessageBoxImage.Error);
                NoticeDialog dlg = new NoticeDialog("Error", ex.Message,
                    "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                dlg.ShowDialog();
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
