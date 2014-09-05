/*
 * Copyright 2014 XPFriend Community.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
using System;
using System.Data.Common;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using NetOffice;
using XPFriend.FixtureBook.DB;

namespace XPFriend.FixtureBook.Forms
{
    /// <summary>
    /// ConnectionSettingWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ConnectionSettingWindow : Window
    {
        private DatabaseFactory databaseFactory = DatabaseFactory.GetInstance();

        public ConnectionSettingWindow()
        {
            InitializeComponent();
            this.Background = SystemColors.ControlBrush;
            databaseFactory.ProviderNames.ForEach(s => this.ProviderName.Items.Add(s));
            ConnectionSetting setting = Addin.ConnectionSetting;
            this.ProviderName.SelectedItem = setting.ProviderName;
            this.ConnectionString.Text = setting.ConnectionString;
            this.ConnectionString.IsReadOnly = !databaseFactory.GetDatabase().CanQuery;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateSetting())
            {
                Save();
                this.Close();
            }

        }

        private bool ValidateSetting()
        {
            DebugConsole.WriteLine("ProviderName: " + this.ProviderName.Text);
            DebugConsole.WriteLine("ConnectionString: " + this.ConnectionString.Text);
            if (!databaseFactory.GetDatabase(this.ProviderName.Text).CanQuery)
            {
                return true;
            }

            try
            {
                DbProviderFactory factory = DbProviderFactories.GetFactory(this.ProviderName.Text);
                using (DbConnection connection = factory.CreateConnection())
                {
                    connection.ConnectionString = this.ConnectionString.Text;
                    connection.Open();
                }
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "FixtureBook", MessageBoxButton.OK, MessageBoxImage.Error);
                DebugConsole.WriteException(e);
                return false;
            }
        }

        private void Save()
        {
            ConnectionSetting setting = Addin.ConnectionSetting;
            setting.ProviderName = this.ProviderName.Text;
            setting.ConnectionString = this.ConnectionString.Text;
            XmlSerializer serializer = new XmlSerializer(typeof(ConnectionSetting));
            using (TextWriter writer = new StreamWriter(Addin.DBConfigPath, false, Encoding.UTF8))
            {
                serializer.Serialize(writer, setting);
            }
        }

        private void ProviderName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Database database = databaseFactory.GetDatabase(this.ProviderName.SelectedItem as string);
            this.ConnectionString.Text = database.DefaultConnectionString;
            this.ConnectionString.IsReadOnly = !database.CanQuery;
        }
    }
}
