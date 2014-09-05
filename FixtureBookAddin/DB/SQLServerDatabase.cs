﻿/*
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
using System.Data;
using System.Data.Common;

namespace XPFriend.FixtureBook.DB
{
    internal class SQLServerDatabase : Database
    {
        public override bool CanQuery
        {
            get { return true; }
        }

        public override string ProviderName
        {
            get { return "System.Data.SqlClient"; }
        }

        public override string DefaultConnectionString
        {
            get { return @"Data Source=(LocalDB)\v11.0;integrated security=True;Initial Catalog=test1"; }
        }

        protected override DataTable GetSchema(DbConnection connection)
        {
            return connection.GetSchema("Tables");
        }
    }
}
