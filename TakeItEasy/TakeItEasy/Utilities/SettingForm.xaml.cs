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
using TakeItEasy.Const;
using TakeItEasy.CommonDialog;

namespace TakeItEasy.Utilities
{
    /// <summary>
    /// Interaction logic for SettingForm.xaml
    /// </summary>
    public partial class SettingForm : MetroWindow
    {
        public SettingForm()
        {
            InitializeComponent();
        }

        private void btn_OK_Click(object sender, RoutedEventArgs e)
        {
            if (tb_Pass.Password != tb_PassConfirm.Password)
            {
                //MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show("Password is not match! Please check again.", Consts.MSG_ERR,
                //    MessageBoxButton.OK, MessageBoxImage.Error);
                NoticeDialog dlg = new NoticeDialog(Consts.MSG_ERR, "Password is not match!Please check again.",
                    "OK", System.Windows.Application.Current.MainWindow, DialogIcons.ERROR);
                dlg.ShowDialog();
                return;
            }
            Consts.SERVER_STRING = tb_Server.Text;
            Consts.USER_STRING = tb_User.Text;
            Consts.PWD_STRING = tb_Pass.Password;

            this.Close();
        }

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void frm_Setting_Loaded(object sender, RoutedEventArgs e)
        {
            tb_Server.Text = Consts.SERVER_STRING;
            tb_User.Text = Consts.USER_STRING;
            tb_Pass.Password = tb_PassConfirm.Password = Consts.PWD_STRING;
        }
    }
}
