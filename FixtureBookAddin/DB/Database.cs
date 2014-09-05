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
using System.Data.Common;
using NetOffice;

namespace XPFriend.FixtureBook.DB
{
    internal abstract class Database
    {
        private DbProviderFactory factory;
        private List<string> tableNames;
        private bool isAvailable = true;

        protected Database()
        {
            if (CanQuery)
            {
                try
                {
                    factory = DbProviderFactories.GetFactory(ProviderName);
                }
                catch (Exception e)
                {
                    isAvailable = false;
                    DebugConsole.WriteLine(ProviderName + ":" + e.Message);
                }
            }
        }

        public bool IsAvailable { get { return isAvailable; } }

        public abstract bool CanQuery { get; }
        public abstract string ProviderName { get; }
        public abstract string DefaultConnectionString { get; }
        protected abstract DataTable GetSchema(DbConnection connection);

        public List<string> GetTableNames(bool reload)
        {
            if (reload)
            {
                tableNames = null;
            }

            if (tableNames == null)
            {
                using (DbConnection connection = factory.CreateConnection())
                {
                    connection.ConnectionString = Addin.ConnectionSetting.ConnectionString;
                    connection.Open();
                    tableNames = GetTableNames(connection);
                }
            }
            return tableNames;
        }

        protected virtual List<string> GetTableNames(DbConnection connection)
        {
            DataTable schema = GetSchema(connection);
            List<string> tableNames = new List<string>(schema.Rows.Count);
            foreach (DataRow row in schema.Rows)
            {
                tableNames.Add(row["TABLE_NAME"] as string);
            }
            return tableNames;
        }

        public string GetDefaultQuery(string tableName)
        {
            return "SELECT * FROM " + tableName;
        }

        public DataTable ExecuteQuery(string query, int maxRecords)
        {
            using (DbConnection connection = factory.CreateConnection())
            {
                connection.ConnectionString = Addin.ConnectionSetting.ConnectionString;
                connection.Open();
                DbCommand command = factory.CreateCommand();
                command.Connection = connection;
                command.CommandText = query;
                DbDataAdapter adapter = factory.CreateDataAdapter();
                adapter.SelectCommand = command;
                DataTable table = new DataTable();
                adapter.Fill(0, maxRecords, table);
                return table;
            }
        }
    }
}
