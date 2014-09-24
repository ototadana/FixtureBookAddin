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
using System.Linq;
using System.Text;
using System.Windows;
using NetOffice.ExcelApi.GlobalHelperModules;
using XPFriend.FixtureBook.DB;
using XPFriend.FixtureBook.Forms;
using Excel = NetOffice.ExcelApi;

namespace XPFriend.FixtureBook.Actions
{
    internal class TableInsertAction : InsertAction
    {
        private string section;
        private int row;

        public TableInsertAction(Excel.Application excel) : base(excel) { }

        protected override void Run()
        {
            base.Run();
            Application.ScreenUpdating = true;
            if ((section == "C" || section == "F") && DatabaseFactory.GetInstance().GetDatabase().CanQuery)
            {
                CreateFromDatabase(Application.ActiveSheet as Excel.Worksheet, row, section);
            }
        }

        protected override void Insert(Excel.Worksheet sheet, int row, Excel.Worksheet templateSheet)
        {
            this.row = row;
            this.section = GetCurrentSection(sheet, row);
            if (section == "B" || section == "D" || section == "E" || section == "C" || section == "F")
            {
                CreateFromTemplate(sheet, row, templateSheet, section);
            }
        }

        private void CreateFromDatabase(Excel.Worksheet sheet, int rowIndex, string section)
        {
            QueryWindow window = new QueryWindow();
            window.ShowDialog();
            Application.ScreenUpdating = false;
            using (DataTable dataTable = window.DataTable)
            {
                if (dataTable == null)
                {
                    return;
                }
                sheet.Cells[rowIndex + 1, 3].Value = dataTable.TableName;
                int columnCount = dataTable.Columns.Count;
                int columnRowIndex = rowIndex + 2;
                int columnStartIndex = 4;
                int columnEndIndex = columnStartIndex + columnCount - 1;
                WriteColumnNames(sheet, dataTable, columnCount, columnRowIndex, columnStartIndex, columnEndIndex, section);

                int rowCount = dataTable.Rows.Count;
                int startRowIndex = columnRowIndex + 1;
                int endRowIndex = startRowIndex + Math.Max(rowCount, 2) - 1;
                Excel.Range valueStartCell = sheet.Cells[startRowIndex, columnStartIndex];
                WriteRows(sheet, columnEndIndex, rowCount, startRowIndex, endRowIndex, valueStartCell);

                WriteValues(dataTable, columnCount, rowCount, valueStartCell);
            }
        }

        private void WriteValues(DataTable dataTable, int columnCount, int rowCount, Excel.Range valueStartCell)
        {
            StringBuilder sb = new StringBuilder();
            int tableRowCount = Math.Max(rowCount, 2);
            for (int i = 0; i < tableRowCount; i++)
            {
                if (i > 0)
                {
                    sb.Append("\n");
                }
                DataRow row = (i < rowCount) ? dataTable.Rows[i] : null;
                WriteValues(columnCount, sb, row);
            }
            Clipboard.SetText(sb.ToString());
            valueStartCell.PasteSpecial();
        }

        private void WriteValues(int columnCount, StringBuilder sb, DataRow row)
        {
            for (int i = 0; i < columnCount; i++)
            {
                if (i > 0)
                {
                    sb.Append("\t");
                }
                if (row != null)
                {
                    sb.Append(ToString(row[i]));
                }
            }
        }

        private object ToString(object o)
        {
            if (o == null)
            {
                return "";
            }
            //データ量が多い可能性があるので、あえて byte[] 変換はしない
            //else if (o is byte[])
            //{
            //    return Convert.ToBase64String(o as byte[]);
            //}
            else
            {
                return o.ToString();
            }
        }

        private void WriteRows(Excel.Worksheet sheet, int columnEndIndex, int rowCount, int startRowIndex, int endRowIndex, Excel.Range valueStartCell)
        {
            if (rowCount > 2)
            {
                Application.CutCopyMode = Excel.Enums.XlCutCopyMode.xlCopy;
                sheet.Range((startRowIndex + 2) + ":" + endRowIndex).Insert();
            }
            Application.CutCopyMode = Excel.Enums.XlCutCopyMode.xlCopy;
            valueStartCell.Copy();
            sheet.Range(valueStartCell, sheet.Cells[endRowIndex, columnEndIndex]).PasteSpecial(Excel.Enums.XlPasteType.xlPasteFormats);
        }

        private static void WriteColumnNames(Excel.Worksheet sheet, DataTable dataTable, int columnCount, int columnRowIndex, int columnStartIndex, int columnEndIndex, string section)
        {
            bool addStar = section == "F";
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < columnCount; i++)
            {
                if (i > 0)
                {
                    sb.Append("\t");
                }
                if (addStar && dataTable.PrimaryKey.Contains(dataTable.Columns[i]))
                {
                    sb.Append("*");
                }
                sb.Append(dataTable.Columns[i].ColumnName);
            }
            Clipboard.Clear();
            Clipboard.SetText(sb.ToString());
            Excel.Range columnStartCell = sheet.Cells[columnRowIndex, columnStartIndex];
            columnStartCell.PasteSpecial();
            columnStartCell.Copy();
            sheet.Range(columnStartCell, sheet.Cells[columnRowIndex, columnEndIndex]).PasteSpecial(Excel.Enums.XlPasteType.xlPasteFormats);
        }

        private void CreateFromTemplate(Excel.Worksheet sheet, int row, Excel.Worksheet templateSheet, string section)
        {
            Excel.Range copyResource = GetCopySource(templateSheet, section);
            if (copyResource != null)
            {
                sheet.Range(row + ":" + row).Insert();
                copyResource.Copy();
                sheet.Cells[row + 1, 1].Insert();
                Move(sheet, row + 1, false);
            }
        }

        private Excel.Range GetCopySource(Excel.Worksheet templateSheet, string section)
        {
            int start = FindSection(templateSheet, section);
            if (start == -1)
            {
                return null;
            }
            start++;
            int end = FindTableEnd(templateSheet, start);
            if (end == -1)
            {
                return null;
            }
            return templateSheet.Range(start + ":" + end);
        }

        private int FindSection(Excel.Worksheet sheet, string section)
        {
            string prefix = section + ".";
            Excel.Range currentCell = sheet.Cells[1, 2];
            int maxRowCount = GlobalModule.Rows.Count;
            while (currentCell.Row < maxRowCount)
            {
                currentCell = currentCell.End(Excel.Enums.XlDirection.xlDown);
                string value = Convert.ToString(currentCell.Value);
                if (value.StartsWith(prefix))
                {
                    return currentCell.Row;
                }
            }
            return -1;
        }

        private int FindTableEnd(Excel.Worksheet sheet, int start)
        {
            Excel.Range currentCell = sheet.Cells[start, 2];
            int maxRowCount = GlobalModule.Rows.Count;
            while (currentCell.Row < maxRowCount)
            {
                if (!IsValueDefinitionRow(currentCell))
                {
                    return currentCell.Row - 1;
                }
                currentCell = currentCell.Offset(1, 0);
            }
            return -1;
        }

        protected override int FindInsertPosition(Excel.Worksheet sheet)
        {
            Excel.Range currentCell = Application.ActiveCell;
            if (sheet.Cells[currentCell.Row, 2].Value != null)
            {
                currentCell = currentCell.Offset(1, 0);
            }

            int maxRowCount = GlobalModule.Rows.Count;
            while(currentCell.Row < maxRowCount)
            {
                if (sheet.Cells[currentCell.Row, 2].Value != null)
                {
                    return currentCell.Row;
                }
                if (!IsValueDefinitionRow(sheet.Cells[currentCell.Row, 3]))
                {
                    return currentCell.Row;
                }
                currentCell = currentCell.Offset(1, 0);
            }
            return currentCell.Row;
        }
    }
}
