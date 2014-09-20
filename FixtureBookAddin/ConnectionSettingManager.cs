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
using System.Configuration;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using XPFriend.FixtureBook.DB;

namespace XPFriend.FixtureBook
{
    internal class ConnectionSettingManager
    {
        private const string DefaultName = "Connection1";
        private static readonly string DefaultProviderName = new EmptyDatabase().ProviderName;

        private List<ConnectionSetting> connectionSettings;

        private ConnectionSettingManager()
        {
        }

        public List<ConnectionSetting> ConnectionSettings 
        { 
            get 
            { 
                return new List<ConnectionSetting>(connectionSettings); 
            } 
        }

        public ConnectionSetting Default
        {
            get
            {
                return connectionSettings[0];
            }
        }

        private static void UpdateConfigurationManager(List<ConnectionSetting> connectionSettings)
        {
            Fixture.FixtureBook.ConnectionStrings.Clear();
            connectionSettings.ForEach(setting =>
                Fixture.FixtureBook.ConnectionStrings.Add(
                    new ConnectionStringSettings(
                        setting.Name, 
                        setting.ConnectionString, 
                        setting.ProviderName)
                )
            );
        }

        public static ConnectionSettingManager Load()
        {
            ConnectionSettingManager manager = new ConnectionSettingManager();
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(List<ConnectionSetting>));
                using (TextReader reader = new StreamReader(Addin.DBConfigPath, Encoding.UTF8))
                {
                    manager.connectionSettings = serializer.Deserialize(reader) as List<ConnectionSetting>;
                }
            }
            catch (Exception)
            {
                manager.connectionSettings = new List<ConnectionSetting>(1);
                AddNewItem(DefaultName, manager.connectionSettings);
            }
            SetDefaultProviderName(manager.connectionSettings);
            UpdateConfigurationManager(manager.connectionSettings);
            return manager;
        }

        private static void SetDefaultProviderName(List<ConnectionSetting> connectionSettings)
        {
            connectionSettings.ForEach(setting =>
            {
                if (string.IsNullOrEmpty(setting.ProviderName))
                {
                    setting.ProviderName = DefaultProviderName;
                }
            });
        }

        public void Save(List<ConnectionSetting> connectionSettings)
        {
            UpdateConfigurationManager(connectionSettings);
            this.connectionSettings = connectionSettings;

            XmlSerializer serializer = new XmlSerializer(typeof(List<ConnectionSetting>));
            using (TextWriter writer = new StreamWriter(Addin.DBConfigPath, false, Encoding.UTF8))
            {
                serializer.Serialize(writer, connectionSettings);
            }
        }

        public static void Remove(int index, List<ConnectionSetting> connectionSettings)
        {
            if (connectionSettings.Count > 1)
            {
                connectionSettings.RemoveAt(index);
            }
        }

        public static void SetAsDefault(int index, List<ConnectionSetting> connectionSettings)
        {
            ConnectionSetting selected = connectionSettings[index];
            connectionSettings.RemoveAt(index);
            connectionSettings.Insert(0, selected);
        }

        internal static void AddNewItem(string name, List<ConnectionSetting> connectionSettings)
        {
            if (!connectionSettings.Exists(setting => setting.Name == name))
            {
                connectionSettings.Insert(0, new ConnectionSetting()
                {
                    Name = name,
                    ProviderName = DefaultProviderName
                });
            }
        }
    }
}
