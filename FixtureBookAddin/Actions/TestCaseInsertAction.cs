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
using System.Windows;
using NetOffice.ExcelApi.GlobalHelperModules;
using Excel = NetOffice.ExcelApi;

namespace XPFriend.FixtureBook.Actions
{
    internal class TestCaseInsertAction : InsertAction
    {
        public TestCaseInsertAction(Excel.Application excel) : base(excel) { }

        protected override void Insert(Excel.Worksheet sheet, int row, Excel.Worksheet templateSheet)
        {
            Excel.Range copyResource = GetCopySource(templateSheet);
            copyResource.Copy();
            sheet.Cells[row, 1].Insert();
            Move(sheet, row, true);
        }

        private Excel.Range GetCopySource(Excel.Worksheet templateSheet)
        {
            Excel.Range firstCell = templateSheet.Cells[1, 1];
            return templateSheet.Range(firstCell, GetLastCell(firstCell));
        }

        protected override int FindInsertPosition(Excel.Worksheet sheet)
        {
            int maxRowCount = GlobalModule.Rows.Count;
            int row = Application.ActiveCell.Row;
            while (row < maxRowCount)
            {
                Excel.Range range = sheet.Cells[row, 2].End(Excel.Enums.XlDirection.xlDown);
                if (range.Row >= maxRowCount)
                {
                    return GetLastCell().Row + 1;
                }
                row = range.Row;
                if (range.Value.ToString().StartsWith("A."))
                {
                    if (row == 0 || IsValueDefinitionRow(range.Offset(-1)))
                    {
                        return row;
                    }
                    else
                    {
                        return row - 1;
                    }
                }
            }
            return row;
        }
    }
}
