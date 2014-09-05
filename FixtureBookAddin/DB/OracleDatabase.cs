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
using System.Data;
using System.Data.Common;
using NetOffice;

namespace XPFriend.FixtureBook.DB
{
    internal class OracleDatabase : Database
    {
        public override bool CanQuery
        {
            get { return true; }
        }

        public override string ProviderName
        {
            get { return "Oracle.DataAccess.Client"; }
        }

        public override string DefaultConnectionString
        {
            get { return "User Id=scott;Password=tiger;Data Source=xe"; }
        }

        protected override DataTable GetSchema(DbConnection connection)
        {
            string userId = GetUserId();
            DebugConsole.WriteLine("userId:" + userId);
            if (userId != null)
            {
                return connection.GetSchema("Tables", new string[] { userId, null});
            }
            else
            {
                return connection.GetSchema("Tables");
            }
        }

        private string GetUserId()
        {
            try
            {
                string connectionString = Addin.ConnectionSetting.ConnectionString;
                string[] s = connectionString.Split(';', '=');
                for (int i = 0; i < s.Length; i++)
                {
                    string key = s[i].Trim();
                    if (key.Equals("user id", StringComparison.OrdinalIgnoreCase))
                    {
                        return s[i + 1].Trim().ToUpper();
                    }
                }
            }
            catch (Exception)
            {
            }
            return null;
        }
    }
}
