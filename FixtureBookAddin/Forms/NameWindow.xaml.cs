using System;
using System.Windows;

namespace XPFriend.FixtureBook.Forms
{
    /// <summary>
    /// NameWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class NameWindow : Window
    {
        private string connectionName;
        public string ConnectionName { get { return connectionName; } }

        public NameWindow()
        {
            InitializeComponent();
            this.Background = SystemColors.ControlBrush;
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            this.ConnectionNameText.Focus();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            connectionName = this.ConnectionNameText.Text;
            this.DialogResult = !string.IsNullOrEmpty(ConnectionName);
            this.Close();
        }
    }
}
