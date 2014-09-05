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
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Xml.Serialization;
using NetOffice;
using NetOffice.ExcelApi.Tools;
using NetOffice.Tools;
using XPFriend.FixtureBook.Actions;
using XPFriend.FixtureBook.DB;
using XPFriend.FixtureBook.Properties;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;

namespace XPFriend.FixtureBook
{
    [COMAddin("FixtureBookAddin", "Excel addin tool for FixtureBook", 3), CustomUI("XPFriend.FixtureBook.RibbonUI.xml")]
    [GuidAttribute("C75D71C8-E0F3-405A-96E1-8BEF6085F25C"), ProgId("FixtureBookAddin.Addin")]
    public class Addin : COMAddin
    {
        private static string applicationPath;
        private static ConnectionSetting connectionSetting;

        private Excel.Application excel;
        private Office.IRibbonUI ribbon;
        private SheetInsertAction sheetInsertAction;
        private TestCaseInsertAction testCaseInsertAction;
        private TableInsertAction tableInsertAction;
        private OpenDBConfigAction openDBConfigAction;

        static Addin()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            applicationPath = Path.Combine(appData, "XPFriend\\FixtureBook");
            InitializeLogFile();
            connectionSetting = LoadConnectionSetting();
        }

        public Addin()
        {
            this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
            this.OnConnection += new OnConnectionEventHandler(Addin_OnConnection);
            this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);

        }

        private static void InitializeLogFile()
        {
            if (!File.Exists(applicationPath))
            {
                Directory.CreateDirectory(applicationPath);
            }
            string fileName = System.IO.Path.Combine(applicationPath, "log.txt");
            if (IsExpired(fileName))
            {
                File.Delete(fileName);
            }
            DebugConsole.FileName = fileName;
            DebugConsole.AppendTimeInfoEnabled = true;
            DebugConsole.Mode = ConsoleMode.LogFile;
            DebugConsole.WriteLine("");
            DebugConsole.WriteLine("");
            DebugConsole.WriteLine("--------------------------------------------------");
            DebugConsole.WriteLine(DateTime.Today.ToString(CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern));
            DebugConsole.WriteLine("");
        }

        private static bool IsExpired(string fileName)
        {
            return File.Exists(fileName) && File.GetLastWriteTime(fileName).Date != DateTime.Today;
        }

        private static ConnectionSetting LoadConnectionSetting()
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ConnectionSetting));
                using (TextReader reader = new StreamReader(DBConfigPath, Encoding.UTF8))
                {
                    return serializer.Deserialize(reader) as ConnectionSetting;
                }
            }
            catch (Exception)
            {
                return new ConnectionSetting() { ProviderName = new EmptyDatabase().ProviderName };
            }
        }

        internal static void HandleException(string tag, Exception e)
        {
            string message = String.Format(Resources.AnErrorOccured, tag, e.Message, DebugConsole.FileName);
            DebugConsole.WriteLine(message);
            DebugConsole.WriteException(e);

            MessageBox.Show(message, "FixtureBook", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        internal static string ApplicationPath
        {
            get { return applicationPath; }
        }

        internal static string DBConfigPath
        {
            get { return Path.Combine(ApplicationPath, "db.config"); }
        }

        internal static ConnectionSetting ConnectionSetting
        {
            get { return connectionSetting; }
        }


        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            HandleException(methodKind.ToString(), exception);
        }

        [ErrorHandler]
        public void GeneralErrorHandler(ErrorMethodKind methodKind, Exception exception)
        {
            HandleException(methodKind.ToString(), exception);
        }

        public override string GetCustomUI(string ribbonID)
        {
            return Localize(base.GetCustomUI(ribbonID));
        }

        // もうちょっとスマートな方法はないものか...
        private string Localize(string customUI)
        {
            customUI = customUI.Replace("{editGroup}", RibbonUI.editGroup);
            customUI = customUI.Replace("{sheetInsertButton}", RibbonUI.sheetInsertButton);
            customUI = customUI.Replace("{testCaseInsertButton}", RibbonUI.testCaseInsertButton);
            customUI = customUI.Replace("{tableInsertButton}", RibbonUI.tableInsertButton);
            customUI = customUI.Replace("{configGroup}", RibbonUI.configGroup);
            customUI = customUI.Replace("{dbConfigButton}", RibbonUI.dbConfigButton);
            return customUI;
        }

        #region IDTExtensibility2 Members

        void Addin_OnConnection(object application, NetOffice.Tools.ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            excel = (Excel.Application)Factory.CreateObjectFromComProxy(null, application);
            excel.WorkbookOpenEvent += new Excel.Application_WorkbookOpenEventHandler(Application_WorkbookOpen);
            CreateExcelActions(excel);
        }

        private void CreateExcelActions(Excel.Application excel)
        {
            this.sheetInsertAction = new SheetInsertAction(excel);
            this.testCaseInsertAction = new TestCaseInsertAction(excel);
            this.tableInsertAction = new TableInsertAction(excel);
            this.openDBConfigAction = new OpenDBConfigAction(excel);
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            ribbon.InvalidateControl("sheetInsertButton");
        }

        void Addin_OnStartupComplete(ref Array custom)
        {
            CreateUserInterface();
        }

        void Addin_OnDisconnection(NetOffice.Tools.ext_DisconnectMode RemoveMode, ref Array custom)
        {
            RemoveUserInterface();
        }

        #endregion

        #region Classic UI

        private void CreateUserInterface()
        {
            // TODO: create UI items
        }

        private void RemoveUserInterface()
        {

        }

        #endregion

        #region IRibbonExtensibility Members

        public void RibbonLoaded(Office.IRibbonUI ribbon)
        {
            this.ribbon = ribbon;
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "sheetInsertButton":
                        return sheetInsertAction.IsEnabled;
                    case "testCaseInsertButton":
                        return testCaseInsertAction.IsEnabled;
                    case "tableInsertButton":
                        return tableInsertAction.IsEnabled;
                    case "dbConfigButton":
                        return openDBConfigAction.IsEnabled;
                    default:
                        return false;
                }
            }
            catch (Exception e)
            {
                HandleException("GetEnabled(" + control.Id + ")", e);
                return false;
            }
        }

        public void dbConfigButton_OnAction(Office.IRibbonControl control)
        {
            openDBConfigAction.Execute();
        }

        public void SheetInsertButton_OnAction(Office.IRibbonControl control)
        {
            sheetInsertAction.Execute();
        }

        public void TestCaseInsertButton_OnAction(Office.IRibbonControl control)
        {
            testCaseInsertAction.Execute();
        }

        public void TableInsertButton_OnAction(Office.IRibbonControl control)
        {
            tableInsertAction.Execute();
        }

        #endregion
    }
}
