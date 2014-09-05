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
    internal class SheetInsertAction : ExcelAction
    {
        public SheetInsertAction(Excel.Application excel) : base(excel) { }

        public override bool IsEnabled
        {
            get { return true; }
        }

        protected override void Run()
        {
            Excel.Workbook book = Application.ActiveWorkbook;
            Excel.Worksheet sheet;
            if (book == null)
            {
                book = Application.Workbooks.Add();
                sheet = (Excel.Worksheet)book.ActiveSheet;
            }
            else
            {
                sheet = (Excel.Worksheet)book.Sheets.Add();
            }
            CopyFromTemplate(sheet);
            sheet.Cells[4, 3].Select();
        }

        private void CopyFromTemplate(Excel.Worksheet sheet)
        {
            Excel.Workbook templateBook = OpenTemplateBook();
            try
            {
                sheet.Activate();
                Excel.Worksheet templateSheet = (Excel.Worksheet)templateBook.Sheets[1];
                templateSheet.Cells.Copy(sheet.Cells[1, 1]);
            }
            finally
            {
                CloseWorkbook(templateBook);
            }
        }
    }
}
