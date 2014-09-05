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
using System.Data;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using XPFriend.FixtureBook.DB;

namespace XPFriend.FixtureBook.Forms
{
    /// <summary>
    /// QueryWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class QueryWindow : Window
    {
        private Database database = DatabaseFactory.GetInstance().GetDatabase();

        public QueryWindow()
        {
            InitializeComponent();
            AddTables();
        }

        public DataTable DataTable { get; set; }

        private void AddTables()
        {
            List<string> tableNames = database.GetTableNames(false);
            tableNames.ForEach(s => this.TableNames.Items.Add(s));
        }

        private void MaxRowCount_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = new Regex("[^0-9]+").IsMatch(e.Text);
        }

        private void TableNames_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string tableName = this.TableNames.SelectedItem as string;
            this.Query.Text = database.GetDefaultQuery(tableName);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ExecuteButton_Click(object sender, RoutedEventArgs e)
        {
            int maxRowCount = GetMaxRowCount();
            this.MaxRowCount.Text = MaxRowCount.ToString();
            if (!string.IsNullOrEmpty(this.Query.Text))
            {
                DataTable dataTable = database.ExecuteQuery(this.Query.Text, maxRowCount);
                dataTable.TableName = this.TableNames.SelectedItem as string;
                this.DataTable = dataTable;
            }
            this.Close();
        }

        private int GetMaxRowCount()
        {
            try
            {
                int count = int.Parse(this.MaxRowCount.Text);
                if (count < 1)
                {
                    return 1;
                }
                return Math.Min(count, 999);
            }
            catch (Exception)
            {
                return 10;
            }
        }
    }
}
