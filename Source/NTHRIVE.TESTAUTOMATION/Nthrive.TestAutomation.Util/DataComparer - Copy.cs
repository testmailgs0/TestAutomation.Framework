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
        private static DataTable dtMerged = new DataTable("MismatchedData");
        //private static DataTable Source = new DataTable("Source");
        //private static DataTable Target = new DataTable("Target");
        public static bool CompareData(DataTable Source, DataTable Target)
        {
            if (Source.Rows.Count == Target.Rows.Count && Source.Columns.Count == Source.Columns.Count)
            {
                for (int i = 0; i < Source.Rows.Count; i++)
                {
                    for (int c = 0; c < Target.Columns.Count; c++)
                    {
                        String s1 = Source.Rows[i][c].ToString();
                        String s2 = Target.Rows[i][c].ToString();
                        //if (!Equals(Source.Rows[i][c], Target.Rows[i][c]))
                        if (!Equals(s1, s2))
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        public static DataTable GetMismatchedData(DataTable Source, DataTable Target)
        {
            List<String> ValidCols = GetSimilarColumn(Source, Target);
            dtMerged.Clear();
            foreach (DataRow dr in Source.Rows)
            {

                    dtMerged =
           (from a in Source.AsEnumerable()
            join b in Target.AsEnumerable()
                               on
     a["User Name"].ToString() equals b["User Name"].ToString()
                             into g
            where g.Count() >0
            select dtMerged.Rows[dtMerged.Rows.IndexOf(a)]).CopyToDataTable();
                }
            return dtMerged;
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
        private static String compactStringPerfect(String spacesStr)
        {

            String newStr = "";
            String currentStr = "";

            for (int i = 0; i < spacesStr.Length; /*no incrementation*/)
            {
                if (spacesStr.ElementAt(i) != ' ')
                {
                    while (spacesStr.ElementAt(i) != ' ')
                    {
                        currentStr += spacesStr.ElementAt(i);
                        i++;
                    }
                }
                else
                {
                    i++;
                }
                if (!currentStr.Equals(""))
                {
                    newStr += currentStr + " ";
                    currentStr = "";
                }
            }

            return newStr.Trim();

        }
        private static DataTable RemoveTrailingSpaces(DataTable dt)
        {
            foreach (DataColumn dc in dt.Columns)
            {
                dt.Columns[dc.ColumnName].ColumnName = compactStringPerfect(dc.ColumnName);
            }
            return dt;
        }
    }
}
