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
using System.Configuration;
using System.Windows;
using XPFriend.FixtureBook.Properties;
using Excel = NetOffice.ExcelApi;

namespace XPFriend.FixtureBook.Actions
{
    internal class SetupAction : ExcelAction
    {
        public SetupAction(Excel.Application excel) : base(excel) { }

        public override bool IsEnabled
        {
            get { return Application.ActiveWorkbook != null; }
        }

        protected override void Run()
        {
            string testCase = FindTestCase();
            if (string.IsNullOrEmpty(testCase))
            {
                return;
            }
            string message = string.Format(Resources.ConfirmSetupMessage, testCase);
            MessageBoxResult result = MessageBox.Show(message, "FixtureBook", MessageBoxButton.OKCancel, MessageBoxImage.Question, MessageBoxResult.OK);
            if (result != MessageBoxResult.OK)
            {
                return;
            }

            Excel.Workbook workbook = Application.ActiveWorkbook;
            if (!workbook.Saved)
            {
                MessageBox.Show(Resources.SaveBookMessage, "FixtureBook", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            Fixture.FixtureBook.ClearCache();
            Fixture.FixtureBook fixtureBook = new Fixture.FixtureBook(typeof(Addin),
                Application.ActiveWorkbook.FullName, (Application.ActiveSheet as Excel.Worksheet).Name, testCase);
            fixtureBook.Setup();
        }

        private string FindTestCase()
        {
            Excel.Worksheet sheet = Application.ActiveSheet as Excel.Worksheet;
            int row = Application.ActiveCell.Row;
            Excel.Range range = sheet.Cells[row, 2];
            while (row > 0)
            {
                if (IsTestCaseSection(range))
                {
                    return GetTestCase(sheet, range);
                }

                row = range.Row;
                range = range.End(Excel.Enums.XlDirection.xlUp);
                if (range.Row == row)
                {
                    break;
                }
            }
            return "";
        }

        private string GetTestCase(Excel.Worksheet sheet, Excel.Range range)
        {
            Excel.Range end = range.End(Excel.Enums.XlDirection.xlDown);
            while (end.Row > range.Row)
            {
                range = sheet.Cells[range.Row, 3].End(Excel.Enums.XlDirection.xlDown);
                if (end.Row > range.Row)
                {
                    return GetValue(range);
                }
            }
            return "";
        }
    }
}
