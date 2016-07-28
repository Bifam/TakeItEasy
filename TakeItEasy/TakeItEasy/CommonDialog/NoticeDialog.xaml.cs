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
using System.IO;
using System.Drawing;

namespace TakeItEasy.CommonDialog
{
    /// <summary>
    /// Interaction logic for NoticeDialog.xaml
    /// </summary>
    public partial class NoticeDialog : MetroWindow
    {
        private string DialogTitle = "";
        private string DialogContent = "";
        private string DialogBtn = "";
        private Window WinOwner;
        private DialogIcons Type;

        public NoticeDialog()
        {
            InitializeComponent();
        }

        public NoticeDialog(string title, string content, string btnCap, Window owner, DialogIcons type)
        {
            DialogTitle = title;
            DialogContent = content;
            DialogBtn = btnCap;
            WinOwner = owner;
            Type = type;

            this.Owner = WinOwner;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.ShowInTaskbar = false;

            //this.DataContext = this;

            InitializeComponent();

            Loaded += (sender, e) =>
                MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void btn_Action_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void frm_Dialog_Loaded(object sender, RoutedEventArgs e)
        {
            lbl_Content.Text = DialogContent;
            lbl_Title.Content = DialogTitle;
            btn_Action.Content = DialogBtn;

            LoadIcon(Type);
        }

        private void LoadIcon(DialogIcons icon)
        {
            BitmapImage logo = new BitmapImage();
            ImageSourceConverter c = new ImageSourceConverter();
            //get dialog type
            switch (icon)
            {
                case DialogIcons.ERROR:
                    img_Ico.ImageSource = Convert(Properties.Resources.error);
                    break;
                case DialogIcons.INFO:
                    img_Ico.ImageSource = Convert(Properties.Resources.info);
                    break;
                case DialogIcons.WARNING:
                    img_Ico.ImageSource = Convert(Properties.Resources.warning);
                    break;
            }
        }

        #region IValueConverter Members

        public ImageSource Convert(Bitmap value)
        {
            MemoryStream ms = new MemoryStream();
            value.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            BitmapImage image = new BitmapImage();
            //Convert bitmap to bitmapimage
            image.BeginInit();
            ms.Seek(0, SeekOrigin.Begin);
            image.StreamSource = ms;
            image.EndInit();

            return image;
        }
        #endregion
    }

    //Structure for icon
    public enum DialogIcons
    {
        INFO,
        WARNING,
        ERROR
    }
}
