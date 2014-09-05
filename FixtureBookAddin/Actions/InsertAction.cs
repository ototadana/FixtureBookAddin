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
using Excel = NetOffice.ExcelApi;

namespace XPFriend.FixtureBook.Actions
{
    internal abstract class InsertAction : ExcelAction
    {
        public InsertAction(Excel.Application excel) : base(excel) { }

        public override bool IsEnabled
        {
            get { return Application.ActiveWorkbook != null; }
        }

        protected override void Run()
        {
            Excel.Workbook book = Application.ActiveWorkbook;
            Excel.Worksheet sheet = (Excel.Worksheet)book.ActiveSheet;
            int row = FindInsertPosition(sheet);
            Insert(sheet, row);
        }

        private void Insert(Excel.Worksheet sheet, int row)
        {
            if (!CanInsert(row))
            {
                return;
            }

            Excel.Workbook templateBook = OpenTemplateBook();
            try
            {
                Application.CutCopyMode = Excel.Enums.XlCutCopyMode.xlCopy;
                Excel.Worksheet templateSheet = (Excel.Worksheet)templateBook.Sheets[1];
                Insert(sheet, row, templateSheet);
            }
            finally
            {
                CloseWorkbook(templateBook);
            }
        }

        protected abstract void Insert(Excel.Worksheet sheet, int row, Excel.Worksheet templateSheet);
        protected abstract int FindInsertPosition(Excel.Worksheet sheet);

    }
}
