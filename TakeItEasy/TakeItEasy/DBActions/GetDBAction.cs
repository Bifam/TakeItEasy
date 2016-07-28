using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using TakeItEasy.Const;
using TakeItEasy.CommonDialog;
using System.Collections.ObjectModel;

namespace TakeItEasy.DatabaseSrc
{
    class GetDBAction
    {
        public static bool ExecuteCommand(string query)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Consts.createConnString()))
                {
                    connection.Open();
                    SqlCommand sqlCmd = new SqlCommand(query, connection);

                    sqlCmd.ExecuteNonQuery();
                    return true;

                }
            }
            catch (Exception ex)
            {
                NoticeDialog dlg = new NoticeDialog(Consts.MSG_ERR, ex.Message, "OK", Application.Current.MainWindow, DialogIcons.ERROR);
                dlg.ShowDialog();
                return false;
            }
        }

        public static DataTable GetTableDetail(string query)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Consts.createConnString()))
                {
                    connection.Open();
                    SqlCommand sqlCmd = new SqlCommand(query, connection);
                    try
                    {
                        SqlDataReader reader = sqlCmd.ExecuteReader();
                        DataTable dt = new DataTable();
                        //get all columns
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            dt.Columns.Add(reader.GetName(i));
                        }

                        while (reader.Read())
                        {    //Every new row will create a new dictionary that holds the columns
                            //column = new Dictionary<string, string>();
                            DataRow row = dt.NewRow();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                row[reader.GetName(i)] = reader[reader.GetName(i)];
                            }
                            dt.Rows.Add(row); //Place the dictionary into the list
                        }
                        reader.Close();
                        return dt;
                    }
                    catch (Exception ex)
                    {
                        NoticeDialog dlg = new NoticeDialog(Consts.MSG_ERR, ex.Message, "OK", Application.Current.MainWindow, DialogIcons.ERROR);
                        dlg.ShowDialog();
                        return null;
                    }

                }
            }
            catch (Exception ex)
            {
                NoticeDialog dlg = new NoticeDialog(Consts.MSG_ERR, ex.Message, "OK", Application.Current.MainWindow, DialogIcons.ERROR);
                dlg.ShowDialog();
                //MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(ex.Message, Consts.MSG_ERR,
                //    MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public static ObservableCollection<ObjectData> GetFieldList(string dbName, string tblName)
        {
            using (SqlConnection connection = new SqlConnection(Consts.createConnString()))
            {
                connection.Open();
                string query = "SELECT TOP 0 * FROM " + dbName + ".dbo.[" + tblName + "]" + "WHERE 0=1";
                SqlCommand sqlCmd = new SqlCommand(query, connection);

                SqlDataReader reader = sqlCmd.ExecuteReader();
                DataTable dt = new DataTable();
                if (reader == null)
                    return null;
                ObservableCollection<ObjectData> lst = new ObservableCollection<ObjectData>();
                //get all columns
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    lst.Add(new ObjectData(reader.GetName(i), 0));
                }
                return lst;
            }
        }

        public static ObservableCollection<ObjectData> GetTableList(ObjectData it)
        {
            using (SqlConnection connection = new SqlConnection(Consts.createConnString()))
            {
                connection.Open();
                //get database name
                connection.ChangeDatabase(it.Name);
                //get list of table
                DataTable schema = connection.GetSchema("Tables");
                //sort table list
                DataView dv = schema.DefaultView;
                dv.Sort = "TABLE_NAME ASC";
                DataTable sortedSchema = dv.ToTable();
                ObservableCollection<ObjectData> listData = new ObservableCollection<ObjectData>();
                //add table to list
                foreach (DataRow row in sortedSchema.Rows)
                {
                    listData.Add(new ObjectData(row[2].ToString(), 0));
                }
                return listData;
            }
        }

        public static ObservableCollection<ObjectData> GetDBList()
        {
            using (SqlConnection connection = new SqlConnection(Consts.createConnString()))
            {
                connection.Open();
                DataTable schema = connection.GetSchema("Databases");
                ObservableCollection<ObjectData> dbList = new ObservableCollection<ObjectData>();
                foreach (DataRow row in schema.Rows)
                {
                    //TableNames.Add(row[2].ToString());
                    String strDatabaseName = row["database_name"].ToString();
                    connection.ChangeDatabase(strDatabaseName);
                    DataTable dbTbl = connection.GetSchema("Tables");

                    ObjectData dbData = new ObjectData(strDatabaseName, dbTbl.Rows.Count);
                    dbList.Add(dbData);
                }
                return dbList;
            }
        }
    }
}
