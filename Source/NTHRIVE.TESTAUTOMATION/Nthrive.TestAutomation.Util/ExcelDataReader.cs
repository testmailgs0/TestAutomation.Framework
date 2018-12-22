using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nthrive.TestAutomation.Util
{
   public static class ExcelDataReader
    {
        public static DataSet GetDataFromExcel(string filepath)

        {
            String path = filepath;
            //Instance reference for Excel Application
            Microsoft.Office.Interop.Excel.Application objXL = null;
            //Workbook refrence
            Microsoft.Office.Interop.Excel.Workbook objWB = null;
            DataSet ds = new DataSet();
            try
            {
                //Instancing Excel using COM services
                objXL = new Microsoft.Office.Interop.Excel.Application();
                //Adding WorkBook
                objWB = objXL.Workbooks.Open(path);
                foreach (Microsoft.Office.Interop.Excel.Worksheet objSHT in objWB.Worksheets)
                {
                    objSHT.AutoFilterMode = false;
                    int rows = objSHT.UsedRange.Rows.Count;
                    int cols = objSHT.UsedRange.Columns.Count;
                    DataTable dt = new DataTable();
                    int noofrow = 3;
                    //If 1st Row Contains unique Headers for datatable include this part else remove it
                    //Start
                    for (int c = 1; c <= cols; c++)

                    {

                        string colname = objSHT.Cells[1, c].Text;

                        dt.Columns.Add(colname);

                        noofrow = 2;

                    }
                    //END
                    for (int r = noofrow; r <= rows; r++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int c = 1; c <= cols; c++)
                        {

                            dr[c - 1] = objSHT.Cells[r, c].Text;
                        }
                        dt.Rows.Add(dr);
                    }
                    ds.Tables.Add(dt);
                }
                //Closing workbook
                objWB.Close();
                //Closing excel application
                objXL.Quit();
            }
            catch (Exception )
            {
                objWB.Saved = true;
                //Closing work book
                objWB.Close();
                //Closing excel application
                objXL.Quit();
                //Response.Write("Illegal permission");
                throw;
            }
            return ds;
        }
    }
}
