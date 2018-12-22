//using TestAutomation.Interfaces;
using System;
using System.Data;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using TestAutomationFX.Core;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Nthrive.TestAutomation.TestDataProvider
{
    public class OpenXmlDataProvider 
    {
        /// <summary>
        /// Gets the test data.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="rowHeaderIndex">Index of the row header.</param>
        /// <param name="startColumnFrom">The start column from.</param>
        /// <returns></returns>
        public DataSet GetTestData(string filePath, int rowHeaderIndex = 1, int startColumnFrom = 1)
        {
            return LoadExcelFile(filePath, rowHeaderIndex, startColumnFrom);
        }

        private static List<string> _dateFormats = new List<string>() 
        {
            "mm-dd-yy",
            "m/d/yyyy",
            "M/d/yyyy"
        };

        /// <summary>
        /// Loads the excel file.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="headerIndex">Index of the header.</param>
        /// <param name="columnFrom">The column from.</param>
        /// <returns></returns>
        private DataSet LoadExcelFile(string filePath, int headerIndex, int columnFrom)
        {
            DataSet ds = new DataSet();
            try
            {
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (var sheet in xlPackage.Workbook.Worksheets)
                    {
                        int rowHeaderIndex = headerIndex;
                        int startColumnFrom = columnFrom;
                        if (!sheet.Name.StartsWith("_"))
                        {
                            DataTable dt = new DataTable(sheet.Name);

                            var totalRows = sheet.Dimension.End.Row;
                            var totalColumns = sheet.Dimension.End.Column;

                            //handling merged cells
                            List<string> mergedCellsRange = new List<string>();
                            sheet.MergedCells.ToList().ForEach(x => { mergedCellsRange.Add(x.ToString()); });
                            mergedCellsRange.ForEach(y => { sheet.Cells[y].Merge = false; });

                            //adding columns in data table 
                            List<string> rowHeader = sheet.Cells[rowHeaderIndex, startColumnFrom, rowHeaderIndex, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString()).ToList();
                            rowHeader.ForEach(colName =>
                            {
                                if (!string.IsNullOrWhiteSpace(colName))
                                    dt.Columns.Add(colName.Trim());
                            });
                            totalColumns = dt.Columns.Count + (startColumnFrom - 1);

                            //adding rows in data table
                            if (totalColumns > 0)
                            {
                                sheet.Calculate();
                                //Get column with date format
                                List<string> columnWithDateFormat = new List<string>();
                                sheet.Cells[rowHeaderIndex + 1, startColumnFrom, rowHeaderIndex + 1, totalColumns].Cast<OfficeOpenXml.ExcelRangeBase>().ToList().ForEach(cell =>
                                {
                                    if (cell != null && _dateFormats.Contains(cell.Style.Numberformat.Format.ToString()))
                                    {
                                        columnWithDateFormat.Add(Regex.Replace(cell.ToString(), @"[\d-]", string.Empty));
                                    }
                                });

                                for (int rowNum = rowHeaderIndex + 1; rowNum <= totalRows; rowNum++) //select starting row here
                                {
                                    //Update date format column value
                                    columnWithDateFormat.ForEach(column =>
                                    {
                                        var cellValue = sheet.Cells[column + rowNum].Value;
                                        if (cellValue != null)
                                        {
                                            long dateNum;
                                            if (long.TryParse(cellValue.ToString().Trim(), out dateNum))
                                            {
                                                DateTime result = DateTime.FromOADate(dateNum);
                                                sheet.Cells[column + rowNum].Value = result.ToString("MM/dd/yyyy");
                                            }
                                        }
                                    });

                                    var cellValues = sheet.Cells[rowNum, startColumnFrom, rowNum, totalColumns].Value;
                                    var values = cellValues as Object[,];
                                    IEnumerable<string> row = null;
                                    if (values != null)
                                        row = (cellValues as Object[,]).Cast<object>().Select(x => x).Select(y => y == null ? string.Empty : y.ToString().Trim());
                                    else
                                    {
                                        string rowValue = cellValues == null ? string.Empty : cellValues.ToString();
                                        row = new List<string> { rowValue };
                                    }
                                    if (row != null && !string.IsNullOrWhiteSpace(string.Join("", row).Trim()))
                                        dt.Rows.Add(row.ToArray());

                                }
                            }

                            ds.Tables.Add(dt);
                        }
                    }
                }
            }
            catch (Exception )
            {
                throw;// Log.Information("Error: " + ex.Message);
            }
            return ds;
        }
    }
}
