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
using System.Collections.Generic;
using System.Data.Common;
using System.Windows;
using System.Windows.Controls;
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
        private ConnectionSettingManager manager = Addin.ConnectionSettingManager;
        private List<ConnectionSetting> connectionSettings;
        private bool updatingConnectionName;

        public ConnectionSettingWindow()
        {
            InitializeComponent();
            this.Background = SystemColors.ControlBrush;
            this.databaseFactory.ProviderNames.ForEach(s => this.ProviderName.Items.Add(s));
            this.ProviderName.SelectedIndex = 0;
            this.connectionSettings = manager.ConnectionSettings;
            UpdateConnectionSettings(this.connectionSettings);
        }

        private void UpdateConnectionSettings(List<ConnectionSetting> connectionSettings)
        {
            updatingConnectionName = true;
            try
            {
                this.ConnectionName.Items.Clear();
                connectionSettings.ForEach(s => this.ConnectionName.Items.Add(s.Name));
                this.DeleteButton.IsEnabled = connectionSettings.Count > 1;
                this.ConnectionName.SelectedIndex = 0;
                UpdateConnectionSetting(connectionSettings[0]);
            }
            finally
            {
                updatingConnectionName = false;
            }
        }

        private void UpdateConnectionSetting(ConnectionSetting setting)
        {
            this.ProviderName.SelectedItem = setting.ProviderName;
            this.ConnectionString.Text = setting.ConnectionString;
            this.ConnectionString.IsReadOnly = !databaseFactory.GetDatabase(setting.ProviderName).CanQuery;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateSetting(this.ProviderName.Text, this.ConnectionString.Text))
            {
                UpdateProviderNameAndConnectionString();
                manager.Save(this.connectionSettings);
                this.Close();
            }

        }

        private void UpdateProviderNameAndConnectionString()
        {
            this.connectionSettings[0].ProviderName = this.ProviderName.Text;
            this.connectionSettings[0].ConnectionString = this.ConnectionString.Text;
        }

        private bool ValidateSetting(string providerName, string connectionString)
        {
            DebugConsole.WriteLine("ProviderName: " + providerName);
            DebugConsole.WriteLine("ConnectionString: " + connectionString);
            if (!databaseFactory.GetDatabase(providerName).CanQuery)
            {
                return true;
            }

            try
            {
                DbProviderFactory factory = DbProviderFactories.GetFactory(providerName);
                using (DbConnection connection = factory.CreateConnection())
                {
                    connection.ConnectionString = connectionString;
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

        private void ProviderName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Database database = databaseFactory.GetDatabase(this.ProviderName.SelectedItem as string);
            this.ConnectionString.Text = database.DefaultConnectionString;
            this.ConnectionString.IsReadOnly = !database.CanQuery;
        }

        private void ConnectionName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (updatingConnectionName || this.ConnectionName.SelectedIndex == -1)
            {
                return;
            }
            ConnectionSettingManager.SetAsDefault(this.ConnectionName.SelectedIndex, this.connectionSettings);
            UpdateConnectionSettings(this.connectionSettings);
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            NameWindow window = new NameWindow();
            if (window.ShowDialog().Value)
            {
                UpdateProviderNameAndConnectionString();
                ConnectionSettingManager.AddNewItem(window.ConnectionName, this.connectionSettings);
                UpdateConnectionSettings(this.connectionSettings);
            }
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            ConnectionSettingManager.Remove(this.ConnectionName.SelectedIndex, this.connectionSettings);
            UpdateConnectionSettings(this.connectionSettings);
        }

        private void ConnectionName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.ConnectionName.Text))
            {
                return;
            }
            this.connectionSettings[0].Name = this.ConnectionName.Text;
            UpdateConnectionSettings(this.connectionSettings);
        }
    }
}
