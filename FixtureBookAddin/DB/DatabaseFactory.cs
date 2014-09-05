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
using System.Collections.Generic;

namespace XPFriend.FixtureBook.DB
{
    internal class DatabaseFactory
    {
        private static DatabaseFactory instance = new DatabaseFactory();

        public static DatabaseFactory GetInstance()
        {
            return instance;
        }

        private Dictionary<string, Database> databases = new Dictionary<string, Database>();
        private List<string> providerNames = new List<string>();

        private DatabaseFactory()
        {
            AddDatabase(new EmptyDatabase());
            AddDatabase(new SQLServerDatabase());
            AddDatabase(new OracleDatabase());
        }

        public void AddDatabase(Database database)
        {
            if (database.IsAvailable)
            {
                databases.Add(database.ProviderName, database);
                providerNames.Add(database.ProviderName);
            }
        }

        public Database GetDatabase()
        {
            return GetDatabase(Addin.ConnectionSetting.ProviderName);
        }

        public Database GetDatabase(string providerName)
        {
            return databases[providerName];
        }


        public List<string> ProviderNames { get { return providerNames; } }
    }
}
