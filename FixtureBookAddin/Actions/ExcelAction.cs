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
using System.Windows;
using NetOffice.ExcelApi.GlobalHelperModules;
using XPFriend.FixtureBook.Properties;
using Excel = NetOffice.ExcelApi;

namespace XPFriend.FixtureBook.Actions
{
    internal abstract class ExcelAction
    {
        protected const int MaxInsertionRowCount = 100;
        private const string TemplateFileName = "FixtureBookTemplate";

        private Excel.Application excel;
        public Excel.Application Application { get { return excel; } }

        protected ExcelAction(Excel.Application excel)
        {
            this.excel = excel;
        }

        public abstract bool IsEnabled {get;}
        public void Execute()
        {
            if (!IsEnabled)
            {
                return;
            }

            try
            {
                Application.ScreenUpdating = false;
                Run();
            }
            catch (Exception e)
            {
                Addin.HandleException(this.GetType().Name, e);
                throw;
            }
            finally
            {
                Application.CutCopyMode = Excel.Enums.XlCutCopyMode.xlCopy;
                Application.ScreenUpdating = true;
            }
        }

        protected abstract void Run();

        protected string GetTemplateFile()
        {
            string templateFilePath = Path.Combine(Addin.ApplicationPath, GetTemplateFileName());
            if (File.Exists(templateFilePath))
            {
                return templateFilePath;
            }
            File.WriteAllBytes(templateFilePath, Resources.FixtureBookTemplate);
            return templateFilePath;
        }

        protected string GetTemplateFileName()
        {
            return TemplateFileName + "." + CultureInfo.CurrentUICulture.Name + ".xlsx";
        }

        protected Excel.Workbook OpenTemplateBook()
        {
            string templateFileName = GetTemplateFileName();
            try
            {
                return Application.Workbooks[templateFileName];
            }
            catch (Exception)
            {
                Excel.Workbook workbook = Application.Workbooks.Open(GetTemplateFile(), null, true);
                Application.ActiveWindow.Visible = false;
                return workbook;
            }
        }

        protected void CloseWorkbook(Excel.Workbook workbook)
        {
            Application.CutCopyMode = Excel.Enums.XlCutCopyMode.xlCopy;
            workbook.Saved = true;
        }

        protected Excel.Range GetLastCell()
        {
            return GetLastCell(Application.ActiveCell);
        }

        protected Excel.Range GetLastCell(Excel.Range range)
        {
            return range.SpecialCells(Excel.Enums.XlCellType.xlCellTypeLastCell);
        }

        protected bool IsValueDefinitionRow(Excel.Range tmpCell)
        {
            return tmpCell.Value != null ||
                tmpCell.End(Excel.Enums.XlDirection.xlToRight).Column < GlobalModule.Columns.Count ||
                ((Excel.Enums.XlLineStyle)tmpCell.Borders[Excel.Enums.XlBordersIndex.xlEdgeRight].LineStyle)
                != Excel.Enums.XlLineStyle.xlLineStyleNone;
        }

        protected bool CanInsert(int row)
        {
            int maxRowCount = GlobalModule.Rows.Count - MaxInsertionRowCount;
            return row < maxRowCount && GetLastCell().Row < maxRowCount;
        }

        protected string GetCurrentSection(Excel.Worksheet sheet, int row)
        {
            object value = sheet.Cells[row, 2].Value;
            if (value == null)
            {
                value = sheet.Cells[row, 2].End(Excel.Enums.XlDirection.xlUp).Value;
            }
            if (value == null)
            {
                return "";
            }
            string text = value.ToString();
            int index = text.IndexOf('.');
            if (index < 0)
            {
                return "";
            }
            return text.Substring(0, index);
        }

        protected void Move(Excel.Worksheet sheet, int row, bool scroll)
        {
            sheet.Activate();
            Excel.Range target = sheet.Cells[row, 1];
            Application.Goto(target, scroll);
        }
    }
}
