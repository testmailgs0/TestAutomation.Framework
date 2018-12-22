using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nthrive.TestAutomation.Util
{
    public static class DataComparer
    {
        private static List<String> Errors = new List<string>();

        public static List<String> GetMismatchedData(DataTable Source, DataTable Target)
        {
            try
            {
                List<String> SimilarColumns = GetSimilarColumn(Source, Target);
                Errors = RunRecon(Source, Target, SimilarColumns);
            }
            catch (Exception)
            {
                throw;
            }
            return Errors;
        }
        /// <summary>
        /// Runs the recon.
        /// </summary>
        /// <param name="uiData">The uiData.</param>
        /// <param name="excelData">The excelData.</param>
        /// <param name="columns">The columns.</param>
        /// <param name="tolerance">The tolerance.</param>
        /// <param name="toleranceFlag">if set to <c>true</c> [tolerance flag].</param>
        /// <param name="dateTimeFlag">if set to <c>true</c> [date time flag].</param>
        /// <returns></returns>
        public static List<String> RunRecon(DataTable uiData, DataTable excelData, List<String> columns, double tolerance = 0.01, bool toleranceFlag = false, bool dateTimeFlag = false)
        {
            try
            {
                List<String> errors = new List<String>();
                List<String> unmatchedCols = new List<String>();
                double numberColumnValue;
                DateTime dateColumnValue;
                if (columns != null && columns.Count == 0)
                {

                    //To create key if no columns selected as key in order to be used for expression
                    DataRow dataRow = excelData.Rows[0];
                    foreach (DataColumn dc in excelData.Columns)
                    {
                        string colName = dc.ColumnName.Trim();
                        string colValue = dataRow[colName].ToString().Trim();
                        if (!(string.IsNullOrWhiteSpace(colValue) || double.TryParse(colValue, out numberColumnValue) || DateTime.TryParse(colValue, out dateColumnValue)))
                        {
                            if (!columns.Contains(colName))
                                columns.Add(colName);
                        }
                    }
                }

                //Dictionary to contain expression and corresponding matched rows from UI data table
                Dictionary<String, DataRow[]> supersetExpressionWise = new Dictionary<String, DataRow[]>();
                foreach (DataRow dr in uiData.Rows)
                {
                    string expression = string.Empty;
                    //Generates expression for UI data rows
                    foreach (string colName in columns)
                    {
                        string colValue = dr[colName].ToString().Trim();
                        if (columns.Contains(colName))
                        {
                            if (colValue.Equals("#$#"))
                                colValue = string.Empty;
                            expression = string.IsNullOrWhiteSpace(expression) ? expression : expression + " AND ";
                            expression = expression + "[" + colName + "] = '" + colValue + "' ";
                        }
                    }
                    //Finds the matching rows from the table and adds it to the dictionary
                    if (!supersetExpressionWise.ContainsKey(expression))
                    {
                        DataRow[] matchedRows = uiData.Select(expression).ToArray();
                        supersetExpressionWise.Add(expression, matchedRows);
                    }
                }

                //Verifies the excel sheet data to be present on UI
                foreach (DataRow dr in excelData.Rows)
                {
                    //frame the expression for each row
                    string expression = string.Empty;
                    foreach (string colName in columns)
                    {
                        string colValue = dr[colName].ToString().Trim();
                        //Take columns taken in dictionary expression only
                        if (colValue.Equals("$#$"))
                            colValue = string.Empty;
                        expression = string.IsNullOrWhiteSpace(expression) ? expression : expression + " AND ";
                        expression = expression + "[" + colName + "] = '" + colValue + "' ";
                    }
                    if (!supersetExpressionWise.ContainsKey(expression))
                    {
                        errors.Add("Could not find entry : " + expression);
                    }
                    else
                    {
                        DataRow[] matchingDictionary = supersetExpressionWise[expression];
                        HashSet<string> dateColumns = new HashSet<string>();
                        expression = string.Empty;
                        String dateExpression = string.Empty;
                        foreach (DataColumn dc in excelData.Columns)
                        {
                            if (!columns.Contains(dc.ColumnName))
                            {
                                string colValue = dr[dc].ToString().Trim();
                                double T = tolerance;
                                if (colValue == String.Empty)
                                {
                                    continue;
                                }
                                else if (colValue == "$#$")
                                {
                                    colValue = string.Empty;
                                    expression = string.IsNullOrWhiteSpace(expression) ? expression : expression + " AND ";
                                    expression = expression + "[" + dc + "] = '" + colValue + "' ";
                                }
                                else if (double.TryParse(colValue, out numberColumnValue))
                                {
                                    if (toleranceFlag)
                                    {
                                        expression = string.IsNullOrWhiteSpace(expression) ? expression : expression + " AND ";
                                        expression = expression + "[" + dc + "] >= '" + (numberColumnValue - T).ToString() + "' ";
                                        expression = string.IsNullOrWhiteSpace(expression) ? expression : expression + " AND ";
                                        expression = expression + "[" + dc + "] <= '" + (numberColumnValue + T).ToString() + "' ";
                                    }
                                    else
                                    {
                                        expression = string.IsNullOrWhiteSpace(expression) ? expression : expression + " AND ";
                                        expression = expression + "[" + dc + "] = '" + colValue + "' ";
                                    }
                                }
                                else if (DateTime.TryParse(dr[dc].ToString().Trim(), out dateColumnValue))
                                {

                                    if (dateTimeFlag)
                                    {
                                        expression = string.IsNullOrWhiteSpace(expression) ? expression : expression + " AND ";
                                        expression = expression + "[" + dc + "] = '" + dateColumnValue + "' ";

                                    }
                                    else
                                    {
                                        dateExpression = string.IsNullOrWhiteSpace(dateExpression) ? dateExpression : dateExpression + " AND ";
                                        dateColumns.Add(dc.Caption);
                                        dateExpression = dateExpression + "[" + dc + "] = '" + dateColumnValue.Date.ToString("MM/dd/yyyy") + "' ";
                                    }
                                }
                                else
                                {
                                    expression = string.IsNullOrWhiteSpace(expression) ? expression : expression + " AND ";
                                    expression = expression + "[" + dc + "] = '" + colValue + "' ";
                                }
                            }
                        }

                        //Finds the corresponding row exists or not in the dictionary
                        DataTable subsetTable = new DataTable();

                        subsetTable = matchingDictionary.CopyToDataTable();
                        DataRow[] matchedRows = subsetTable.Select(expression);

                        if (subsetTable.Select(expression).ToList().Count <= 0)
                        {
                            errors.Add("Values did not match for " + expression);
                            continue;
                        }

                        //Convert each value of each column in dateColumns to Date in specified Format by parsing each value to datetime Object

                        //Checks if dateTime Flag is true and dateColumns is not empty
                        if (dateColumns.Count > 0)
                        {
                            subsetTable = matchedRows.CopyToDataTable();
                            DateTime columnDateTimeValue = new DateTime();
                            foreach (DataRow row in subsetTable.Rows)
                            {
                                foreach (string columnName in dateColumns)
                                {
                                    DateTime.TryParse(row[columnName].ToString(), out columnDateTimeValue);
                                    row[columnName] = columnDateTimeValue.Date.ToString("MM/dd/yyyy");
                                }
                            }

                            if (subsetTable.Select(dateExpression).ToList().Count <= 0)
                            {
                                errors.Add("Values did not match for " + expression + dateExpression);
                            }
                        }
                    }
                }
                if (errors != null && errors.Count > 0)
                    //CaptureScreenshot(uiData.TableName);
                return errors;
            }
            catch (Exception)
            {
                    throw;
            }
            return null;
        }

        private static List<string> GetSimilarColumn(DataTable source, DataTable target)
        {
            List<String> Cols = new List<String>();
            foreach (DataColumn srccol in source.Columns)
            {
                foreach (DataColumn trgcol in target.Columns)
                {
                    if (srccol.ColumnName.Equals(trgcol.ColumnName))
                    {
                        Cols.Add(srccol.ColumnName);
                    }
                }
            }
            return Cols;
        }
    }
}
