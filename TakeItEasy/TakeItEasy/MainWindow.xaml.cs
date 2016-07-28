using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using TakeItEasy.CommonDialog;
using TakeItEasy.Const;
using TakeItEasy.DatabaseSrc;
using TakeItEasy.Utilities;

namespace TakeItEasy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class App_Main : MetroWindow
    {
        Dictionary<String, String> TempData;
        Dictionary<String, String> TempNewData;
        String CurrentDB = "";
        String CurrentTbl = "";
        ObservableCollection<ObjectData> listData = new ObservableCollection<ObjectData>();

        public App_Main()
        {
            InitializeComponent();
            //register hotkey
            RoutedCommand SearchSettings = new RoutedCommand();
            SearchSettings.InputGestures.Add(new KeyGesture(Key.F, ModifierKeys.Control));
            CommandBindings.Add(new CommandBinding(SearchSettings, app_ChooseSearch));
        }

        private void frm_MainApp_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Load database list to GUI
                ObservableCollection<ObjectData> dbList = GetDBAction.GetDBList();
                cb_DbList.ItemsSource = dbList;
                if (dbList != null)
                    cb_DbList.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                NoticeDialog dlg = new NoticeDialog("Error", ex.Message, "OK", Application.Current.MainWindow, DialogIcons.ERROR);
                dlg.ShowDialog();
            }
        }

        private void cb_DbList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                //change database
                lv_Table.ItemsSource = null;
                lv_Table.Items.Clear();
                ObjectData it = (ObjectData)cb_DbList.SelectedItem;
                listData = GetDBAction.GetTableList(it);
                lv_Table.ItemsSource = listData;
                CurrentDB = it.Name;
            }
            catch (Exception ex)
            {
                NoticeDialog dlg = new NoticeDialog("Error", ex.Message, "OK", Application.Current.MainWindow, DialogIcons.ERROR);
                dlg.ShowDialog();
            }
        }

        private void lv_Table_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // get db name and table name
            ObjectData data = (ObjectData)lv_Table.SelectedItem;
            if (data == null)
                return;
            CurrentTbl = data.Name;
            string query = "SELECT TOP 1000 * FROM " + CurrentDB + ".dbo.[" + CurrentTbl + "]";
            DataTable dt = GetDBAction.GetTableDetail(query);
            if (dt != null)
            {
                lv_Record.ItemsSource = dt.DefaultView;
            }
        }

        private void tb_Command_Context_Execute(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(Consts.createConnString()))
            {
                connection.Open();
                //get query from textbox
                string query = tb_Command.Text;
                if (query == "")
                    return;
                SqlCommand sqlCmd = new SqlCommand(query, connection);
                try
                {
                    //execute query
                    SqlDataReader reader = sqlCmd.ExecuteReader();
                }
                catch (Exception ex)
                {
                    NoticeDialog dlg = new NoticeDialog("Error", ex.Message, "OK", this, DialogIcons.ERROR);
                    dlg.ShowDialog();
                }
            }
        }

        private void tb_Command_Context_Clipboard(object sender, RoutedEventArgs e)
        {
            //copy to clipboard
            Clipboard.SetText(tb_Command.Text);
        }

        private void tb_Command_Context_Clear(object sender, RoutedEventArgs e)
        {
            //clear command text box
            tb_Command.Text = "";
        }
        
        private void btn_Execute_Click(object sender, RoutedEventArgs e)
        {
            string query = tb_Command.Text;
            
            if (query != "")
            {
                DataTable dt = GetDBAction.GetTableDetail(query);   //execute query
                if (dt != null)
                    lv_Record.ItemsSource = dt.DefaultView;         //display in list view
            }
        }

        private void tb_Menu_AutoGenFile_Click(object sender, RoutedEventArgs e)
        {
            GenerateFile genFile = new GenerateFile();
            genFile.Owner = this;
            genFile.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            genFile.ShowDialog();
        }

        private void tb_Menu_GenMsgFile_Click(object sender, RoutedEventArgs e)
        {
            //generate message file
            GenMsgFile genFile = new GenMsgFile();
            genFile.Owner = this;
            genFile.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            genFile.ShowDialog();
        }

        private void btn_Utilities_Click(object sender, RoutedEventArgs e)
        {
            //create drop down menu for utility button
            (sender as Button).ContextMenu.IsEnabled = true;
            (sender as Button).ContextMenu.PlacementTarget = (sender as Button);
            (sender as Button).ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
            (sender as Button).ContextMenu.IsOpen = true;
        }

        private void tb_Menu_Setting_Click(object sender, RoutedEventArgs e)
        {
            SettingForm settingFrm = new SettingForm();
            settingFrm.Owner = this;
            settingFrm.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            settingFrm.ShowDialog();
        }

        private void lv_Record_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (CurrentDB == "" || CurrentTbl == "")
                return;

            if (e.EditAction == DataGridEditAction.Commit)
            {
                //get cell value
                TextBox tb = e.EditingElement as TextBox;
                string str = tb.Text;
                //get field name
                string fieldName = (string)e.Column.Header;

                if (CheckIsUpdateRecord())
                {
                    TempNewData = null;
                    if (str == TempData[fieldName])
                        return;
                    //create where expression
                    string whereExpTmp = "";
                    foreach (DataGridColumn a in lv_Record.Columns)
                    {
                        string name = (string)a.Header;
                        string value = TempData[name];
                        whereExpTmp += name + "='" + value + "' AND ";
                    }
                    string whereExp = whereExpTmp.Substring(0, whereExpTmp.Length - 1 - 4);
                    //create query
                    string query = String.Format("UPDATE {0}.dbo.{1} SET {2}='{3}' WHERE {4}",
                        CurrentDB, CurrentTbl, fieldName, str, whereExp);
                    //execute query
                    if (!GetDBAction.ExecuteCommand(query))
                        e.Cancel = true;
                }
                else
                {
                    if (TempNewData == null)
                        TempNewData = new Dictionary<string, string>();

                    //get current value and store in tempnewdata
                    string value = "";
                    if (!TempNewData.TryGetValue(fieldName, out value))
                        TempNewData.Add(fieldName, str==null?"":str);
                    //if is not last colum
                    if (e.Column.DisplayIndex != lv_Record.Columns.Count - 1)
                    {
                        e.Cancel = true;
                        return;
                    }
                    //create query
                    string query = String.Format("INSERT INTO {0}.dbo.{1} VALUES(", CurrentDB, CurrentTbl);
                    string strValues = "";
                    foreach (DataGridColumn a in lv_Record.Columns)
                    {
                        string header = (string)a.Header;
                        value = "";
                        if (!TempNewData.TryGetValue(header, out value))
                        {
                            e.Cancel = true;
                            return;
                        }
                        //string value = TempNewData[header];
                        strValues += "'" + value + "', ";
                    }
                    query += strValues.Substring(0, strValues.Length - 2) + ")";
                    //execute query
                    if (!GetDBAction.ExecuteCommand(query))
                        e.Cancel = true;
                    else
                        TempNewData = null;
                }
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void lv_Record_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            TempData = new Dictionary<string, string>();
            //backup all data before edit
            foreach (DataGridColumn a in lv_Record.Columns)
            {
                DataRowView dvr = (DataRowView)e.Row.Item;
                TempData.Add(a.Header.ToString(), dvr[a.DisplayIndex].ToString());
            }
        }

        private void btn_About_Click(object sender, RoutedEventArgs e)
        {
            NoticeDialog dlg = new NoticeDialog("TakeItEasy", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\r\n!!@#$%^&*(!@#$%^&*("
                + "\r\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!", 
                "OK! I see", this, DialogIcons.INFO);
            dlg.ShowDialog();
        }

        private void tbx_Search_MouseOutsideDown(object sender, MouseButtonEventArgs e)
        {
            popup.IsOpen = false;
        }

        private void tbx_SearchTbl_MouseOutsideDown(object sender, MouseButtonEventArgs e)
        {
            popupTbl.IsOpen = false;
        }

        private void tbx_Search_Loaded(object sender, RoutedEventArgs e)
        {
            AddHandler(Mouse.PreviewMouseDownOutsideCapturedElementEvent,
            new MouseButtonEventHandler(tbx_Search_MouseOutsideDown), true);
            Mouse.Capture(tbx_Search);
        }

        private void tbx_SearchTbl_Loaded(object sender, RoutedEventArgs e)
        {
            AddHandler(Mouse.PreviewMouseDownOutsideCapturedElementEvent,
            new MouseButtonEventHandler(tbx_SearchTbl_MouseOutsideDown), true);
            Mouse.Capture(tbx_SearchTbl);
        }

        private void app_DataSearch(object sender, RoutedEventArgs e)
        {
            popup.IsOpen = true;
            tbx_Search.Text = "";
            tbx_Search.Focus();
        }

        private void app_TblSearch(object sender, RoutedEventArgs e)
        {
            popupTbl.IsOpen = true;
            tbx_SearchTbl.Text = "";
            tbx_SearchTbl.Focus();
        }

        private void app_ChooseSearch(object sender, RoutedEventArgs e)
        {
            UIElement focusedControl = Keyboard.FocusedElement as UIElement;
            if (focusedControl.GetType().Name == "lv_Table")
            {
                app_TblSearch(sender, e);
            }
            else if (focusedControl.GetType().Name == "lv_Record")
            {
                app_DataSearch(sender, e);
            }
        }
        

        private void app_TblDrop(object sender, RoutedEventArgs e)
        {
           
        }

        private void app_TblCreate(object sender, RoutedEventArgs e)
        {

        }

        private void lv_Rec_Delete(object sender, RoutedEventArgs e)
        {
            if (CurrentDB == "" || CurrentTbl == "")
                return;

            string tmpQuery = String.Format("DELETE FROM {0}.dbo.{1} WHERE ", CurrentDB, CurrentTbl);
            DataRowView row = (DataRowView)lv_Record.SelectedItems[0];
            foreach (DataGridColumn a in lv_Record.Columns)
            {
                string header = (string)a.Header;
                string value = row[header].ToString();
                tmpQuery += header + "='" + value + "' AND ";
            }
            string query = tmpQuery.Substring(0, tmpQuery.Length - 5);
            GetDBAction.ExecuteCommand(query);

            query = "SELECT TOP 1000 * FROM " + CurrentDB + ".dbo.[" + CurrentTbl+ "]";
            DataTable dt = GetDBAction.GetTableDetail(query);
            if (dt != null)
                lv_Record.ItemsSource = dt.DefaultView;
        }

        private void tbx_Search_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                popup.IsOpen = false;
                return;
            }
            else if (e.Key == Key.Escape)
            {
                tbx_Search.Text = "";
                popup.IsOpen = false;
                return;
            }
        }

        private void tbx_SearchTbl_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                popupTbl.IsOpen = false;
                return;
            }
            else if (e.Key == Key.Escape)
            {
                tbx_SearchTbl.Text = "";
                popupTbl.IsOpen = false;
                return;
            }
        }

        private bool CheckIsUpdateRecord()
        {
            foreach (DataGridColumn a in lv_Record.Columns)
            {
                string name = (string)a.Header;
                string value = TempData[name];
                if (value != "")
                    return true;
            }
            return false;
        }

        private void Command_ShowAll(object sender, ExecutedRoutedEventArgs e)
        {
            if (CurrentDB == "" || CurrentTbl == "")
                return;

            string query = "SELECT * FROM " + CurrentDB + ".dbo.[" + CurrentTbl+ "]";
            DataTable dt = GetDBAction.GetTableDetail(query);
            if (dt != null)
                lv_Record.ItemsSource = dt.DefaultView;
        }

        private void Command_ShowAll_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }


        private void lv_Export_Ins(object sender, RoutedEventArgs e)
        {
            if (CurrentDB == "" || CurrentTbl == "")
                return;

            string query = String.Format("SELECT TOP 0 * FROM {0}.dbo.{1} WHERE 1 = 2;", CurrentDB, CurrentTbl);
            try
            {
                using (SqlConnection connection = new SqlConnection(Consts.createConnString()))
                {
                    connection.Open();
                    SqlCommand sqlCmd = new SqlCommand(query, connection);
                    try
                    {
                        SqlDataReader reader = sqlCmd.ExecuteReader();
                        // This will return false - we don't care, we just want to make sure the schema table is there.
                        reader.Read();

                        var table = reader.GetSchemaTable();
                        for (int i = 0; i < reader.VisibleFieldCount; i++)
                        {
                            System.Type type = reader.GetFieldType(i);
                            reader.GetName(i);
                            switch (Type.GetTypeCode(type))
                            {
                                case TypeCode.DateTime:
                                    break;
                                case TypeCode.String:
                                    break;
                                default: break;
                            }

                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        NoticeDialog dlg = new NoticeDialog("Error", ex.Message, "OK", Application.Current.MainWindow, DialogIcons.ERROR);
                        dlg.ShowDialog();

                    }

                }
            }
            catch (Exception ex)
            {
                NoticeDialog dlg = new NoticeDialog("Error", ex.Message, "OK", Application.Current.MainWindow, DialogIcons.ERROR);
                dlg.ShowDialog();
            }
        }

        private void tbx_SearchTbl_TextChanged(object sender, TextChangedEventArgs e)
        {
            string txtOrig = tbx_SearchTbl.Text;
            string upper = txtOrig.ToUpper();
            string lower = txtOrig.ToLower();

            var tblFiltered = from Temp in listData
                              let ename = Temp.Name
                              where
                               ename.StartsWith(lower)
                               || ename.StartsWith(upper)
                               || ename.Contains(txtOrig)
                              select Temp;

            lv_Table.ItemsSource = tblFiltered;

        }
    }

    /* Data Control Format Block */
    #region ControlData Format

    public class ObjectData
    {
        public string Name { get; set; }
        public int ItemsCount { get; set; }

        public ObjectData(string name, int count)
        {
            Name = name;
            ItemsCount = count;
        }
    }
    public class ToolList
    {
        public string ItemName { get; set; }

        public ToolList(string name) { ItemName = name; }
    }
    #endregion ControlData Format

    #region TableSearch
    public static class DataGridTextSearch
    {
        // Using a DependencyProperty as the backing store for SearchValue.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SearchValueProperty =
            DependencyProperty.RegisterAttached("SearchValue", typeof(string), typeof(DataGridTextSearch),
                new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.Inherits));

        public static string GetSearchValue(DependencyObject obj)
        {
            return (string)obj.GetValue(SearchValueProperty);
        }

        public static void SetSearchValue(DependencyObject obj, string value)
        {
            obj.SetValue(SearchValueProperty, value);
        }

        // Using a DependencyProperty as the backing store for IsTextMatch.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsTextMatchProperty =
            DependencyProperty.RegisterAttached("IsTextMatch", typeof(bool), typeof(DataGridTextSearch), new UIPropertyMetadata(false));

        public static bool GetIsTextMatch(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsTextMatchProperty);
        }

        public static void SetIsTextMatch(DependencyObject obj, bool value)
        {
            obj.SetValue(IsTextMatchProperty, value);
        }
    }

    public class SearchValueConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string cellText = values[0] == null ? string.Empty : values[0].ToString();
            string searchText = values[1] as string;

            if (!string.IsNullOrEmpty(searchText) && !string.IsNullOrEmpty(cellText))
            {
                return cellText.ToLower().StartsWith(searchText.ToLower());
            }
            return false;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

    #endregion

    #region hotkey register
    public static class Commands
    {
        public static readonly RoutedUICommand ShowAll = new RoutedUICommand
                (
                        "ShowAll",
                        "ShowAll",
                        typeof(App_Main)
                );

        //Define more commands here, just like the one above
    }
    #endregion
}
